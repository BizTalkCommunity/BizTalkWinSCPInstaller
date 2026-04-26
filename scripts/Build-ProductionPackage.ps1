<#
.SYNOPSIS
Builds a minimal production bundle for BizTalk WinSCP installer deployment.

.DESCRIPTION
Creates a reduced-surface-area package that contains only the files required to
run InstallWinSCPForBizTalk.ps1 in production.

This build step does NOT require BizTalk to be installed on the machine running
the build. It only copies files and generates package metadata.

.PARAMETER OutputFolder
Target folder where the production package is created.
Default: <repoRoot>\dist\BizTalkWinSCPInstaller-Production

.PARAMETER NuGetPayloadFolder
Optional folder that contains pre-staged nuget.exe and WinSCP package files.
If supplied, this folder is copied to <OutputFolder>\nuget for offline use.

.PARAMETER Clean
Removes OutputFolder before creating the package.

.EXAMPLE
./scripts/Build-ProductionPackage.ps1 -Clean

.EXAMPLE
./scripts/Build-ProductionPackage.ps1 -Clean -NuGetPayloadFolder C:\temp\winscp-cache
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder = (Join-Path (Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..')).Path 'dist') 'BizTalkWinSCPInstaller-Production'),

    [Parameter(Mandatory = $false)]
    [string]$NuGetPayloadFolder,

    [Parameter(Mandatory = $false)]
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputFolder)

$requiredRelativePaths = @(
    'InstallWinSCPForBizTalk.ps1',
    'src/InstallWinSCPForBizTalk.Core.psm1',
    'src/InstallWinSCPForBizTalk.Utils.psm1',
    'src/InstallWinSCPForBizTalk.Workflow.psm1'
)

foreach ($relativePath in $requiredRelativePaths) {
    $sourcePath = Join-Path $repoRoot $relativePath
    if (-not (Test-Path $sourcePath)) {
        throw "Required source file was not found: $sourcePath"
    }
}

if ($Clean -and (Test-Path $resolvedOutput)) {
    Remove-Item -Path $resolvedOutput -Recurse -Force
}

if ((Test-Path $resolvedOutput) -and (Get-ChildItem -Path $resolvedOutput -Force | Measure-Object).Count -gt 0) {
    throw "Output folder already exists and is not empty: $resolvedOutput. Use -Clean to recreate it."
}

New-Item -Path $resolvedOutput -ItemType Directory -Force | Out-Null

foreach ($relativePath in $requiredRelativePaths) {
    $sourcePath = Join-Path $repoRoot $relativePath
    $targetPath = Join-Path $resolvedOutput $relativePath
    $targetDir = Split-Path -Parent $targetPath
    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
    Copy-Item -Path $sourcePath -Destination $targetPath -Force
}

$payloadIncluded = $false
if (-not [string]::IsNullOrWhiteSpace($NuGetPayloadFolder)) {
    $resolvedPayload = (Resolve-Path $NuGetPayloadFolder -ErrorAction Stop).Path
    if (-not (Test-Path $resolvedPayload -PathType Container)) {
        throw "NuGet payload path is not a folder: $resolvedPayload"
    }

    $payloadTarget = Join-Path $resolvedOutput 'nuget'
    New-Item -Path $payloadTarget -ItemType Directory -Force | Out-Null
    Copy-Item -Path (Join-Path $resolvedPayload '*') -Destination $payloadTarget -Recurse -Force
    $payloadIncluded = $true

    if (-not (Test-Path (Join-Path $payloadTarget 'nuget.exe'))) {
        Write-Warning "NuGet payload copied, but nuget.exe was not found under: $payloadTarget"
    }
}

$defaultNuGetFolderLiteral = if ($payloadIncluded) {
    "Join-Path `$PSScriptRoot 'nuget'"
}
else {
    "(Get-Item Env:TEMP).Value + '\\nuget'"
}

$runInstallerScript = @'
<#
.SYNOPSIS
Runs the bundled BizTalk WinSCP installer.
#>
param(
    [Parameter(Mandatory = $false)]
    [string]$NuGetDownloadFolder = __DEFAULT_NUGET__,
    [Parameter(Mandatory = $false)]
    [switch]$ForceInstall,
    [Parameter(Mandatory = $false)]
    [string]$LogFolder,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Info', 'Verbose', 'Debug')]
    [string]$LogLevel = 'Info',
    [Parameter(Mandatory = $false)]
    [switch]$EnableEventLog,
    [Parameter(Mandatory = $false)]
    [string]$EventLogName = 'Application',
    [Parameter(Mandatory = $false)]
    [string]$EventSource = 'BizTalkWinSCPInstaller'
)

$installerPath = Join-Path $PSScriptRoot 'InstallWinSCPForBizTalk.ps1'
if (-not (Test-Path $installerPath)) {
    throw "Bundled installer was not found: $installerPath"
}

& $installerPath -nugetDownloadFolder $NuGetDownloadFolder -ForceInstall:$ForceInstall -LogFolder $LogFolder -LogLevel $LogLevel -EnableEventLog:$EnableEventLog -EventLogName $EventLogName -EventSource $EventSource
'@

$runInstallerScript = $runInstallerScript.Replace('__DEFAULT_NUGET__', $defaultNuGetFolderLiteral)
Set-Content -Path (Join-Path $resolvedOutput 'Run-Installer.ps1') -Value $runInstallerScript -Encoding UTF8

$bundleReadme = @'
# BizTalk WinSCP Installer - Production Bundle

This package is intentionally minimal and contains only runtime essentials.

## Included files

- InstallWinSCPForBizTalk.ps1
- src/InstallWinSCPForBizTalk.Core.psm1
- src/InstallWinSCPForBizTalk.Utils.psm1
- src/InstallWinSCPForBizTalk.Workflow.psm1
- Run-Installer.ps1
- package-manifest.json

## Usage

Run from an elevated PowerShell session on the BizTalk target machine:

```powershell
.\Run-Installer.ps1
```

If this bundle includes an offline payload under .\nuget, the wrapper defaults
to using that local payload.
'@
Set-Content -Path (Join-Path $resolvedOutput 'PRODUCTION-README.md') -Value $bundleReadme -Encoding UTF8

$manifest = [ordered]@{
    generatedUtc = (Get-Date).ToUniversalTime().ToString('o')
    sourceRoot = $repoRoot
    outputFolder = $resolvedOutput
    includesNuGetPayload = $payloadIncluded
    files = @()
}

$packageFiles = Get-ChildItem -Path $resolvedOutput -Recurse -File | Sort-Object FullName
foreach ($file in $packageFiles) {
    $relative = $file.FullName.Substring($resolvedOutput.Length).TrimStart('\\')
    $manifest.files += [ordered]@{
        path = $relative.Replace('\\', '/')
        sha256 = (Get-FileHash -Path $file.FullName -Algorithm SHA256).Hash
        sizeBytes = $file.Length
    }
}

$manifestPath = Join-Path $resolvedOutput 'package-manifest.json'
$manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $manifestPath -Encoding UTF8

Write-Host "Production package created: $resolvedOutput" -ForegroundColor Green
Write-Host "NuGet payload included: $payloadIncluded" -ForegroundColor Green

[pscustomobject]@{
    OutputFolder = $resolvedOutput
    IncludesNuGetPayload = $payloadIncluded
    FileCount = $packageFiles.Count + 1
}
