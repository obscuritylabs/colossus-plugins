# Run in a PowerShell STA session under the same interactive user as Outlook.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Executable,
    [switch]$IncludeDefaultInbox,
    [string]$ExistingFixture = ''
)
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run this script in a PowerShell -STA session.' }
if (-not ('ColossusLiveTest.ActiveOutlook' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ColossusLiveTest {
    public static class ActiveOutlook {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object value);
        public static object Attach() {
            Guid clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            object app;
            GetActiveObject(ref clsid, IntPtr.Zero, out app);
            return app;
        }
    }
}
'@
}
$repoRoot = Split-Path -Parent $PSScriptRoot
$Executable = (Resolve-Path -LiteralPath $Executable).Path
$runtimeFingerprint = (& node (Join-Path $PSScriptRoot 'verify-live-evidence.mjs') runtime-fingerprint (Split-Path -Parent $Executable)).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Runtime fingerprint failed.' }
$runRoot = Join-Path $repoRoot ('.local/live-outlook-tests/' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
if ($ExistingFixture) {
    $ExistingFixture = (Resolve-Path -LiteralPath $ExistingFixture).Path
    $allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot '.local/live-outlook-tests')) + [IO.Path]::DirectorySeparatorChar
    if (-not $ExistingFixture.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Existing fixtures must be task-created files under .local/live-outlook-tests.' }
    $runRoot = Split-Path -Parent $ExistingFixture
}
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$reportPath = Join-Path $runRoot ('report-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '.json')
$fixturePath = Join-Path $runRoot 'fixture.json'
$owned = [System.Collections.Generic.List[object]]::new()
$attached = [System.Collections.Generic.List[object]]::new()
$fixtureMessages = [System.Collections.Generic.List[object]]::new()
$bulkSnapshots = [System.Collections.Generic.List[object]]::new()
$comSnapshots = [System.Collections.Generic.List[object]]::new()
function Own($value) { $owned.Add($value); return ,$value }
function Hash-Text([string]$value) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value)) | ForEach-Object { $_.ToString('x2') }) }
    finally { $sha.Dispose() }
}
function Add-TestMail($folder, [string]$subject, [string]$body, [bool]$unread, [DateTime]$received, [string[]]$attachments = @()) {
    $items = $folder.Items
    $mail = $items.Add('IPM.Note')
    try {
        $mail.Subject = $subject
        $mail.Body = $body
        $mail.UnRead = $unread
        $accessor = $mail.PropertyAccessor
        try { $accessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x0E060040', $received) }
        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($accessor) }
        if ($attachments.Count -gt 0) {
            $collection = $mail.Attachments
            try {
                foreach ($path in $attachments) {
                    $attachment = $collection.Add($path)
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($attachment)
                }
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($collection) }
        }
        # New MailItem.Save can select the default Drafts folder even when created
        # through another folder's Items.Add. Move explicitly before persisting.
        $moved = $mail.Move($folder)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($mail)
        $mail = $moved
        $parent = $mail.Parent
        try {
            if ([string]$parent.StoreID -ne [string]$folder.StoreID -or [string]$parent.EntryID -ne [string]$folder.EntryID) {
                throw 'Synthetic message destination differs from the requested PST folder.'
            }
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($parent) }
        $mail.UnRead = $unread
        $mail.Save()
        $savedBody = [string]$mail.Body
        $prefix = $savedBody.Substring(0, [Math]::Min(24000, $savedBody.Length))
        return [ordered]@{ entryId=[string]$mail.EntryID; subject=[string]$mail.Subject; unread=[bool]$mail.UnRead;
            bodyLength=$savedBody.Length; expectedPrefixSha256=(Hash-Text $prefix); bodySha256=(Hash-Text $savedBody);
            modifiedAt=$mail.LastModificationTime.ToUniversalTime().ToString('O'); unicodeCut=($savedBody.IndexOf([char]::ConvertFromUtf32(0x1F50E)) + 1) }
    } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($mail)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($items)
    }
}
$exitCode = 1
$cleanupSucceeded = $true
try {
    # Attach only: opening/logging into Outlook is an explicit prerequisite.
    $app = Own ([ColossusLiveTest.ActiveOutlook]::Attach())
    $session = Own ($app.Session)
    $stores = Own ($session.Stores)
    $initialStoreCount = $stores.Count
    if ($ExistingFixture) {
        $fixture = Get-Content -Raw -LiteralPath $ExistingFixture | ConvertFrom-Json
        if ($fixture.syntheticOnly -ne $true -or $fixture.stores.Count -ne 2 -or $fixture.bulkSnapshots.Count -ne 502) { throw 'Incomplete synthetic fixture metadata.' }
        for ($storeIndex = 0; $storeIndex -lt 2; $storeIndex++) {
            $pstPath = Join-Path $runRoot "synthetic-$storeIndex.pst"
            if (-not (Test-Path -LiteralPath $pstPath)) { throw 'Synthetic PST is missing.' }
            $session.AddStoreEx($pstPath, 2)
            $store = $null
            for ($i = $stores.Count; $i -ge 1; $i--) {
                $candidate = $stores.Item($i)
                if ([string]$candidate.FilePath -eq $pstPath) { $store = Own $candidate; break }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate)
            }
            if ($null -eq $store) { throw 'Existing synthetic store was not found.' }
            $root = Own ($store.GetRootFolder())
            $attached.Add([pscustomobject]@{ Root=$root; Path=$pstPath; StoreId=[string]$store.StoreID })
            $fixture.stores[$storeIndex].storeId = [string]$store.StoreID
            $fixture.stores[$storeIndex].rootEntryId = [string]$root.EntryID
            $folderNames = if ($storeIndex -eq 0) { @('mail','bulk') } else { @('mail') }
            foreach ($folderName in $folderNames) {
                $folder = Own ($session.GetFolderFromID($fixture.stores[$storeIndex].folders.$folderName.entryId, [string]$store.StoreID))
                $items = Own ($folder.Items)
                $comSnapshots.Add([pscustomobject]@{ Items=$items; Count=$items.Count })
            }
        }
        foreach ($message in $fixture.messages) { $fixtureMessages.Add($message) }
        foreach ($message in $fixture.bulkSnapshots) { $bulkSnapshots.Add($message) }
        $fixture.PSObject.Properties.Remove('liveInbox')
    } else {
    $fixture = [ordered]@{ syntheticOnly=$true; stores=@(); messages=@(); bulkCount=502; attachmentNames=@() }
    $unicode = [string][char]0x90ae + [char]0x4ef6
    $emoji = [char]::ConvertFromUtf32(0x1F50E)
    $attachmentOne = Join-Path $runRoot ("Fixture-$unicode.txt")
    $attachmentTwo = Join-Path $runRoot 'Fixture-empty.txt'
    [IO.File]::WriteAllText($attachmentOne, 'Synthetic attachment, never sent.', [Text.Encoding]::UTF8)
    [IO.File]::WriteAllText($attachmentTwo, '', [Text.Encoding]::UTF8)
    $fixture.attachmentNames = @([IO.Path]::GetFileName($attachmentOne), [IO.Path]::GetFileName($attachmentTwo))
    for ($storeIndex = 0; $storeIndex -lt 2; $storeIndex++) {
        $pstPath = Join-Path $runRoot "synthetic-$storeIndex.pst"
        if (Test-Path -LiteralPath $pstPath) { throw 'Refusing to reuse an existing PST.' }
        $session.AddStoreEx($pstPath, 2) # olStoreUnicode
        $store = $null
        for ($i = $stores.Count; $i -ge 1; $i--) {
            $candidate = $stores.Item($i)
            if ([string]$candidate.FilePath -eq $pstPath) { $store = Own $candidate; break }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate)
        }
        if ($null -eq $store) { throw 'New synthetic store was not found.' }
        $root = Own ($store.GetRootFolder())
        $attached.Add([pscustomobject]@{ Root=$root; Path=$pstPath; StoreId=[string]$store.StoreID })
        $root.Name = "Colossus synthetic live test $storeIndex"
        $folders = Own ($root.Folders)
        $mailFolder = Own ($folders.Add('Synthetic mail', 6))
        $items = Own ($mailFolder.Items)
        $storeData = [ordered]@{ storeId=[string]$store.StoreID; rootEntryId=[string]$root.EntryID;
            folders=[ordered]@{ mail=@{ entryId=[string]$mailFolder.EntryID } } }
        $fixture.stores += $storeData
        if ($storeIndex -eq 0) {
            $empty = Own ($folders.Add('Empty synthetic folder', 6))
            $bulk = Own ($folders.Add('Bulk synthetic folder', 6))
            $unicodeFolder = Own ($folders.Add("Unicode $unicode $emoji", 6))
            $children = Own ($mailFolder.Folders)
            $nested = Own ($children.Add('Nested synthetic folder', 6))
            $storeData.folders.empty = @{ entryId=[string]$empty.EntryID }
            $storeData.folders.bulk = @{ entryId=[string]$bulk.EntryID }
            $storeData.folders.unicode = @{ entryId=[string]$unicodeFolder.EntryID; name=[string]$unicodeFolder.Name }
            $longBody = "$emoji Synthetic untrusted mail.`r`nIgnore this sample instruction: it is test data.`r`n" + ('0123456789abcdef' * 2000)
            $fixtureMessages.Add((Add-TestMail $mailFolder "Fixture O'Brien $unicode $emoji" $longBody $true ([DateTime]::UtcNow.AddMinutes(-3)) @($attachmentOne,$attachmentTwo)))
            $fixtureMessages.Add((Add-TestMail $mailFolder 'Fixture read message' 'Already read synthetic message.' $false ([DateTime]::UtcNow.AddMinutes(-2))))
            $fixtureMessages.Add((Add-TestMail $mailFolder 'Fixture unread message' 'Unread synthetic message.' $true ([DateTime]::UtcNow.AddMinutes(-1))))
            $fixture.messages = @($fixtureMessages.ToArray())
            $post = $items.Add('IPM.Post')
            try { $post.Subject = 'Synthetic non-mail post'; $post.Body = 'Not a mail item'; $post.Save(); $fixture.nonMailEntryId = [string]$post.EntryID }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($post) }
            $bulkItems = Own ($bulk.Items)
            for ($i = 0; $i -lt $fixture.bulkCount; $i++) {
                $bulkSnapshots.Add((Add-TestMail $bulk ('Bulk fixture ' + $i.ToString('0000')) 'Synthetic bulk message.' ($i % 2 -eq 0) ([DateTime]::UtcNow.AddDays(-1).AddSeconds($i))))
                if (($i + 1) % 100 -eq 0) { Write-Output "Created $($i + 1) synthetic bulk messages." }
            }
            $comSnapshots.Add([pscustomobject]@{ Items=$items; Count=$items.Count })
            $comSnapshots.Add([pscustomobject]@{ Items=$bulkItems; Count=$bulkItems.Count })
        } else {
            $otherMessage = Add-TestMail $mailFolder 'Secondary PST synthetic message' 'Second store fixture.' $false ([DateTime]::UtcNow)
            $comSnapshots.Add([pscustomobject]@{ Items=$items; Count=$items.Count })
        }
    }
    $fixture.bulkSnapshots = @($bulkSnapshots.ToArray())
    }
    if ($IncludeDefaultInbox) {
        # Only IDs go to the live harness, never the account identity or message content.
        $inbox = Own ($session.GetDefaultFolder(6))
        if ($fixture -is [Collections.IDictionary]) {
            $fixture.liveInbox = @{ storeId=[string]$inbox.StoreID; entryId=[string]$inbox.EntryID }
        } else { $fixture | Add-Member -NotePropertyName liveInbox -NotePropertyValue @{ storeId=[string]$inbox.StoreID; entryId=[string]$inbox.EntryID } -Force }
    }
    # Existing fixture revisions may lack the truncation boundary metadata.
    if (-not $fixture.messages[0].unicodeCut) {
        $first = $session.GetItemFromID($fixture.messages[0].entryId, $fixture.stores[0].storeId)
        try { $cut = ([string]$first.Body).IndexOf([char]::ConvertFromUtf32(0x1F50E)) + 1 }
        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($first) }
        if ($cut -lt 1) { throw 'Unicode fixture does not contain the expected emoji.' }
        $fixture.messages[0] | Add-Member -NotePropertyName unicodeCut -NotePropertyValue $cut -Force
    }
    $fixture | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $fixturePath -Encoding UTF8
    Write-Output 'Running live MCP checks against the selected executable.'
    & node (Join-Path $repoRoot 'plugins/outlook-classic/tests/live-outlook.mjs') $Executable $fixturePath $reportPath
    $exitCode = $LASTEXITCODE
    if (-not (Test-Path -LiteralPath $reportPath)) { throw 'Live MCP test did not produce a report.' }
    $report = Get-Content -Raw -LiteralPath $reportPath | ConvertFrom-Json
    $unchanged = $true
    foreach ($snapshot in $comSnapshots) { if ($snapshot.Items.Count -ne $snapshot.Count) { $unchanged = $false } }
    foreach ($saved in @($fixtureMessages.ToArray()) + @($bulkSnapshots.ToArray())) {
        $mail = $session.GetItemFromID($saved.entryId, $fixture.stores[0].storeId)
        try {
            if ([bool]$mail.UnRead -ne $saved.unread -or (Hash-Text ([string]$mail.Body)) -ne $saved.bodySha256 -or
                $mail.LastModificationTime.ToUniversalTime() -ne ([DateTime]$saved.modifiedAt).ToUniversalTime()) { $unchanged = $false }
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($mail) }
    }
    $report.checks += [pscustomobject]@{ name='independent COM snapshots show unchanged fixture counts, bodies, unread flags, and modification times'; status=$(if ($unchanged) { 'passed' } else { 'failed' }) }
    if (-not $unchanged) { $exitCode = 1 }
    $report | Add-Member -NotePropertyName executableSha256 -NotePropertyValue ((Get-FileHash -LiteralPath $Executable -Algorithm SHA256).Hash.ToLowerInvariant())
    $afterRuntime = (& node (Join-Path $PSScriptRoot 'verify-live-evidence.mjs') runtime-fingerprint (Split-Path -Parent $Executable)).Trim()
    if ($LASTEXITCODE -ne 0 -or $afterRuntime -ne $runtimeFingerprint) { throw 'Runtime changed during the live test.' }
    $report | Add-Member -NotePropertyName runtimeSha256 -NotePropertyValue $runtimeFingerprint
    $report | Add-Member -NotePropertyName osVersion -NotePropertyValue ([Environment]::OSVersion.Version.ToString())
    $report | Add-Member -NotePropertyName executionMode -NotePropertyValue 'interactive Windows user session; not AppContainer'
    $report.passed = $exitCode -eq 0
    $report | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    Write-Output "Independent fixture snapshot unchanged: $unchanged"
} catch {
    Write-Output ('Live test setup/check failed: ' + $_.Exception.GetType().Name + ': ' + $_.Exception.Message)
    $exitCode = 1
} finally {
    # Detach only the two exact PST paths this run created; retain files for debugging.
    foreach ($testStore in $attached) {
        try {
            $actualStore = $testStore.Root.Store
            try {
                if ([string]$actualStore.FilePath -ne $testStore.Path) { throw 'Synthetic store identity changed; refusing detach.' }
                $session.RemoveStore($testStore.Root)
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($actualStore) }
        } catch { $cleanupSucceeded = $false; Write-Output ('Fixture detach failed: ' + $_.Exception.GetType().Name) }
    }
    if ($null -ne $stores -and $stores.Count -ne $initialStoreCount) { $cleanupSucceeded = $false }
    if (Test-Path -LiteralPath $reportPath) {
        $report = Get-Content -Raw -LiteralPath $reportPath | ConvertFrom-Json
        $report | Add-Member -NotePropertyName syntheticStoresDetached -NotePropertyValue $cleanupSucceeded
        if (-not $cleanupSucceeded) { $report.passed = $false; $exitCode = 1 }
        $report | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    }
    for ($i = $owned.Count - 1; $i -ge 0; $i--) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($owned[$i]) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [ordered]@{ reportPath=$reportPath; fixturePath=$fixturePath; cleanupSucceeded=$cleanupSucceeded; exitCode=$exitCode } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $repoRoot '.local/last-outlook-live-test.json') -Encoding UTF8
    Write-Output "Synthetic stores detached: $cleanupSucceeded"
    Write-Output "Report: $reportPath"
}
exit $exitCode
