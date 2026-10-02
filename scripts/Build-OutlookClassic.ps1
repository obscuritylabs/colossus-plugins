[CmdletBinding()]
param([string]$DotnetPath = '', [string]$Version = '', [switch]$SkipChecks)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceFingerprint = (& node (Join-Path $PSScriptRoot 'verify-live-evidence.mjs') fingerprint).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Source fingerprint failed.' }
if (-not $DotnetPath) {
    $localDotnet = Join-Path $repoRoot '.local\toolchains\dotnet\dotnet.exe'
    $DotnetPath = if (Test-Path -LiteralPath $localDotnet) { $localDotnet } else { 'dotnet' }
}
$env:DOTNET_CLI_HOME = Join-Path $repoRoot '.local\dotnet-home'
$env:NUGET_PACKAGES = Join-Path $repoRoot '.local\nuget-packages'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_GENERATE_ASPNET_CERTIFICATE = 'false'
$pluginRoot = Join-Path $repoRoot 'plugins\outlook-classic'
$manifest = Get-Content -Raw (Join-Path $pluginRoot 'package\plugin.json') | ConvertFrom-Json
if (-not $Version) { $Version = $manifest.version }
if ($Version -notmatch '^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$') { throw 'Invalid release version.' }
$manifest.version = $Version
$project = Join-Path $pluginRoot 'src\OutlookClassicMcp.csproj'
$runId = [Guid]::NewGuid().ToString('N').Substring(0, 12)
$buildRoot = Join-Path $repoRoot "dist\outlook-classic\$($manifest.version)\windows-amd64\$runId"
$packageRoot = Join-Path $buildRoot 'outlook-classic'

Push-Location $repoRoot
try {
    if (-not $SkipChecks) {
        & $DotnetPath run --project (Join-Path $pluginRoot 'tests\OutlookClassicChecks.csproj') -c Release -p:RestoreLockedMode=true
        if ($LASTEXITCODE -ne 0) { throw 'Outlook component checks failed.' }
    }
    New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
    foreach ($name in @('mcp.json', 'skills')) {
        Copy-Item -LiteralPath (Join-Path $pluginRoot "package\$name") -Destination $packageRoot -Recurse
    }
    $manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $packageRoot 'plugin.json') -Encoding utf8
    foreach ($name in @('LICENSE', 'NOTICE')) {
        Copy-Item -LiteralPath (Join-Path $repoRoot $name) -Destination $packageRoot
    }
    Copy-Item -LiteralPath (Join-Path $pluginRoot 'README.md') -Destination $packageRoot
    & $DotnetPath publish $project -c Release -r win-x64 --self-contained true --nologo `
        -p:DebugType=None -p:DebugSymbols=false -p:RestoreLockedMode=true `
        "-p:Version=$Version" `
        -o (Join-Path $packageRoot 'bin')
    if ($LASTEXITCODE -ne 0) { throw 'Outlook publish failed.' }
    Copy-Item -LiteralPath (Join-Path $pluginRoot 'src\packages.lock.json') -Destination (Join-Path $packageRoot 'dependencies.lock.json')
    & (Join-Path $PSScriptRoot 'Write-DependencyNotices.ps1') -PackageRoot $packageRoot -AssetsPath (Join-Path $pluginRoot 'src\obj\project.assets.json')
    $reportedVersion = & (Join-Path $packageRoot 'bin\outlook-classic-mcp.exe') --version
    if ($LASTEXITCODE -ne 0 -or $reportedVersion -ne "outlook-classic-mcp $Version") { throw 'Binary and manifest versions differ.' }
    if (-not $SkipChecks) {
        & node (Join-Path $pluginRoot 'tests\mcp-smoke.mjs') (Join-Path $packageRoot 'bin\outlook-classic-mcp.exe')
        if ($LASTEXITCODE -ne 0) { throw 'MCP protocol checks failed.' }
        & node (Join-Path $pluginRoot 'tests\http-companion-smoke.mjs') (Join-Path $packageRoot 'bin\outlook-classic-mcp.exe')
        if ($LASTEXITCODE -ne 0) { throw 'HTTP companion checks failed.' }
    }
    $portableZip = Join-Path $buildRoot 'outlook-classic-windows-amd64.zip'
    Compress-Archive -LiteralPath $packageRoot -DestinationPath $portableZip -CompressionLevel Optimal
    $runtimeFingerprint = (& node (Join-Path $PSScriptRoot 'verify-live-evidence.mjs') runtime-fingerprint (Join-Path $packageRoot 'bin')).Trim()
    if ($LASTEXITCODE -ne 0) { throw 'Runtime fingerprint failed.' }
    $receipt = [ordered]@{
        sourceFingerprint = $sourceFingerprint
        runtimeSha256 = $runtimeFingerprint
        version = $manifest.version
        packageRoot = $packageRoot
        executable = Join-Path $packageRoot 'bin\outlook-classic-mcp.exe'
        executableSha256 = (Get-FileHash -LiteralPath (Join-Path $packageRoot 'bin\outlook-classic-mcp.exe') -Algorithm SHA256).Hash.ToLowerInvariant()
        portableZip = $portableZip
        portableZipSha256 = (Get-FileHash -LiteralPath $portableZip -Algorithm SHA256).Hash.ToLowerInvariant()
        checksSkipped = [bool]$SkipChecks
    }
    $afterBuild = (& node (Join-Path $PSScriptRoot 'verify-live-evidence.mjs') fingerprint).Trim()
    if ($LASTEXITCODE -ne 0 -or $afterBuild -ne $sourceFingerprint) { throw 'Source changed during the build; rebuild before testing.' }
    $receipt | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $repoRoot '.local\last-outlook-build.json') -Encoding utf8
    $receipt | ConvertTo-Json
} finally { Pop-Location }
