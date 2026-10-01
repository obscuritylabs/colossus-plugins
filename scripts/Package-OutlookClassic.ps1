[CmdletBinding()]
param([Parameter(Mandatory)][string]$ColossusPath)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$build = Get-Content -Raw (Join-Path $repoRoot '.local/last-outlook-build.json') | ConvertFrom-Json
if ($build.checksSkipped) { throw 'Release packaging requires all build checks.' }
$layout = Join-Path (Split-Path -Parent $build.packageRoot) 'outlook-classic.oci'
& $ColossusPath plugins validate $build.packageRoot
if ($LASTEXITCODE -ne 0) { throw 'Colossus validation failed.' }
& $ColossusPath plugins package $build.packageRoot --output $layout
if ($LASTEXITCODE -ne 0) { throw 'Colossus packaging failed.' }
$index = Get-Content -Raw (Join-Path $layout 'index.json') | ConvertFrom-Json
if ($index.manifests.Count -ne 1) { throw 'Expected one plugin manifest.' }
$digest = $index.manifests[0].digest
# Check determinism of archive metadata and OCI packaging on the same staged bytes.
$second = "$layout-repeat"
& $ColossusPath plugins package $build.packageRoot --output $second
if ($LASTEXITCODE -ne 0) { throw 'Repeated packaging failed.' }
$repeat = Get-Content -Raw (Join-Path $second 'index.json') | ConvertFrom-Json
if ($repeat.manifests[0].digest -ne $digest) { throw 'Package digest was not deterministic.' }
$releaseRoot = Join-Path $repoRoot 'dist/release'
if (Test-Path -LiteralPath $releaseRoot) { throw 'dist/release already exists; use a fresh checkout or archive that directory.' }
New-Item -ItemType Directory -Path $releaseRoot | Out-Null
Copy-Item -LiteralPath $layout -Destination (Join-Path $releaseRoot 'outlook-classic.oci') -Recurse
Copy-Item -LiteralPath $build.portableZip -Destination $releaseRoot
Copy-Item -LiteralPath (Join-Path $build.packageRoot 'plugin.json') -Destination $releaseRoot
$sourceCommit = (& git -C $repoRoot rev-parse HEAD).Trim()
[ordered]@{ version=$build.version; digest=$digest; sourceCommit=$sourceCommit; minColossusVersion='0.11.4'; checksPassed=$true } |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $releaseRoot 'release.json') -Encoding utf8
Write-Output "Release staging complete: $releaseRoot ($digest)"
