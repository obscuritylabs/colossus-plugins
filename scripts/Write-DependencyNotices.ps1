[CmdletBinding()]
param([Parameter(Mandatory)][string]$PackageRoot, [Parameter(Mandatory)][string]$AssetsPath)
$ErrorActionPreference = 'Stop'
$assets = Get-Content -Raw -LiteralPath $AssetsPath | ConvertFrom-Json
$cache = $env:NUGET_PACKAGES
$licenseRoot = Join-Path $PackageRoot 'licenses'
New-Item -ItemType Directory -Path $licenseRoot -Force | Out-Null
$notices = [System.Collections.Generic.List[string]]::new()
$notices.Add('Third-party components bundled with Outlook Classic MCP. See dependencies.lock.json for NuGet integrity hashes.')
$seen = @{}
foreach ($entry in ($assets.libraries.PSObject.Properties | Sort-Object Name)) {
    if ($entry.Value.type -ne 'package') { continue }
    $path = Join-Path $cache $entry.Value.path
    $nuspec = Get-ChildItem -LiteralPath $path -Filter '*.nuspec' | Select-Object -First 1
    if (-not $nuspec) { throw "Missing dependency license metadata: $($entry.Name)" }
    [xml]$xml = Get-Content -Raw -LiteralPath $nuspec.FullName
    $metadata = $xml.package.metadata
    $license = $metadata.license.InnerText
    if ($license -notin @('MIT', 'Apache-2.0')) { throw "Review new dependency license for $($entry.Name): $license" }
    $notices.Add("`n$($entry.Name)`nLicense: $license`n$($metadata.copyright)`n$($metadata.projectUrl)")
    # Retain unique upstream license/notice texts, including .NET runtime notices.
    foreach ($file in Get-ChildItem -LiteralPath $path -File | Where-Object { $_.Name -match '^(LICENSE|NOTICE|THIRD-PARTY-NOTICES)(\..*)?$' }) {
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        if (-not $seen.ContainsKey($hash)) {
            $name = ($entry.Name -replace '/', '-') + '-' + $file.Name
            Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $licenseRoot $name)
            $seen[$hash] = $name
        }
        $notices.Add("Upstream text: licenses/$($seen[$hash])")
    }
}
$runtimeVersion = ([xml](Get-Content -Raw (Join-Path $PSScriptRoot '../plugins/outlook-classic/src/OutlookClassicMcp.csproj'))).Project.PropertyGroup.RuntimeFrameworkVersion
foreach ($runtime in @('microsoft.netcore.app.runtime.win-x64', 'microsoft.windowsdesktop.app.runtime.win-x64')) {
    $path = Join-Path $cache "$runtime/$runtimeVersion"
    if (-not (Test-Path -LiteralPath $path)) { throw "Runtime license source is missing: $runtime/$runtimeVersion" }
    $notices.Add("`n$runtime/$runtimeVersion`nCopyright Microsoft Corporation and .NET Foundation contributors. See upstream license and notices below.")
    $licenseFiles = @(Get-ChildItem -LiteralPath $path -File | Where-Object { $_.Name -match '^(LICENSE|NOTICE|THIRD-PARTY-NOTICES)(\..*)?$' })
    if ($licenseFiles.Count -eq 0) { throw "Missing upstream runtime license: $runtime" }
    foreach ($file in $licenseFiles) {
        $name = "$runtime-$runtimeVersion-$($file.Name)"
        Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $licenseRoot $name)
        $notices.Add("Upstream text: licenses/$name")
    }
}
$notices.Add("`nApache-2.0 license text: LICENSE. MIT license text: licenses/MIT.txt.")
$notices | Set-Content -LiteralPath (Join-Path $PackageRoot 'THIRD-PARTY-NOTICES.txt') -Encoding utf8
@'
MIT License

Copyright (c) Microsoft Corporation. All rights reserved.
Copyright (c) .NET Foundation and Contributors.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'@ | Set-Content -LiteralPath (Join-Path $licenseRoot 'MIT.txt') -Encoding utf8
