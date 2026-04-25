[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\docs\help')
)

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$moduleManifests = @(
    (Join-Path $repoRoot 'src\InstallWinSCPForBizTalk.Core.psd1'),
    (Join-Path $repoRoot 'src\InstallWinSCPForBizTalk.Workflow.psd1'),
    (Join-Path $repoRoot 'src\InstallWinSCPForBizTalk.Utils.psd1')
)

if (-not (Get-Module -ListAvailable -Name platyPS)) {
    throw "platyPS is required. Install with: Install-Module platyPS -Scope CurrentUser"
}

Import-Module platyPS -ErrorAction Stop

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

foreach ($manifest in $moduleManifests) {
    if (-not (Test-Path $manifest)) {
        throw "Missing module manifest: $manifest"
    }

    Import-Module $manifest -Force
    New-MarkdownHelp -Module (Split-Path $manifest -LeafBase) -OutputFolder $OutputPath -Force | Out-Null
}

Write-Host "Help markdown generated in: $OutputPath" -ForegroundColor Green