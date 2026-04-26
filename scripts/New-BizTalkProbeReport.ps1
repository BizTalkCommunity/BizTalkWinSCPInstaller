<#
.SYNOPSIS
Creates a BizTalk probe report for offline WinSCP packaging.

.DESCRIPTION
Runs on a BizTalk machine and produces a JSON report with enough metadata for a
separate build machine to package the right WinSCP payload without requiring
BizTalk to be installed there.

.PARAMETER OutputPath
Path to write the probe JSON report.
Default: ./biztalk-probe.json

.EXAMPLE
./scripts/New-BizTalkProbeReport.ps1

.EXAMPLE
./scripts/New-BizTalkProbeReport.ps1 -OutputPath C:\temp\biztalk-probe.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path (Get-Location).Path 'biztalk-probe.json')
)

$ErrorActionPreference = 'Stop'

function Get-RegistryValueSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        return Get-ItemPropertyValue -Path $Path -Name $Name -ErrorAction Stop
    }
    catch {
        return $null
    }
}

function Get-InstalledUpdateDisplayNames {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $displayNames = New-Object System.Collections.Generic.List[string]
    foreach ($path in $paths) {
        try {
            $entries = Get-ItemProperty -Path $path -ErrorAction Stop
            foreach ($entry in $entries) {
                if ($entry.DisplayName) {
                    $displayNames.Add([string]$entry.DisplayName)
                }
            }
        }
        catch {
            # Ignore inaccessible uninstall hives and continue.
        }
    }

    return $displayNames.ToArray()
}

function Resolve-BizTalkVersion {
    param(
        [string]$ProductCode,
        [string]$ProductName
    )

    if ($ProductCode -eq '{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}') { return '2020' }
    if ($ProductCode -eq '{B084F3A7-3E8F-4E7B-B673-EED1715D28ED}') { return '2016' }

    if ($ProductName -match 'BizTalk Server 2020') { return '2020' }
    if ($ProductName -match 'BizTalk Server 2016') { return '2016' }

    return $null
}

function Resolve-WinSCPFrom2020 {
    param(
        [string[]]$DisplayNames
    )

    $kbMap = @(
        @{ Label = 'CU6'; KBs = @('5043408', '5048971'); WinSCP = '6.3.5' }
        @{ Label = 'CU5'; KBs = @('5032870');             WinSCP = '6.1.2' }
        @{ Label = 'CU4'; KBs = @('5009901');             WinSCP = '5.19.2' }
        @{ Label = 'CU3'; KBs = @('5007969');             WinSCP = '5.19.2' }
        @{ Label = 'CU2'; KBs = @('5003151');             WinSCP = '5.17.8' }
        @{ Label = 'CU1'; KBs = @('4538666');             WinSCP = '5.17.6' }
    )

    foreach ($entry in $kbMap) {
        foreach ($kb in $entry.KBs) {
            if ($DisplayNames -match $kb) {
                return [pscustomobject]@{
                    WinSCPVersion = $entry.WinSCP
                    UpdateLabel = $entry.Label
                    MatchedKB = $kb
                    Confidence = 'high'
                }
            }
        }
    }

    return [pscustomobject]@{
        WinSCPVersion = '5.15.4'
        UpdateLabel = 'RTM/no CU detected'
        MatchedKB = 'none'
        Confidence = 'medium'
    }
}

function Resolve-WinSCPFrom2016 {
    param(
        [string[]]$DisplayNames
    )

    $kbMap = @(
        @{ Label = 'CU9 and FP3'; KB = '5005480'; WinSCP = '5.19.2' }
        @{ Label = 'CU9';         KB = '5005479'; WinSCP = '5.19.2' }
        @{ Label = 'CU8 and FP3'; KB = '4590075'; WinSCP = '5.15.9' }
        @{ Label = 'CU8';         KB = '4583530'; WinSCP = '5.15.9' }
        @{ Label = 'CU7 and FP3'; KB = '4536185'; WinSCP = '5.15.9' }
        @{ Label = 'CU7';         KB = '4528776'; WinSCP = '5.15.9' }
        @{ Label = 'CU6 and FP3'; KB = '4294900'; WinSCP = '5.13.1' }
        @{ Label = 'CU6';         KB = '4477494'; WinSCP = '5.13.1' }
        @{ Label = 'CU5 and FP3'; KB = '4103503'; WinSCP = '5.13.1' }
        @{ Label = 'CU5 Hotfix';  KB = '4345385'; WinSCP = '5.13.1' }
        @{ Label = 'CU5';         KB = '4132957'; WinSCP = '5.13.1' }
        @{ Label = 'CU4 and FP2'; KB = '4094130'; WinSCP = '5.7.7' }
        @{ Label = 'CU4';         KB = '4051353'; WinSCP = '5.7.7' }
        @{ Label = 'CU3 and FP2'; KB = '4054819'; WinSCP = '5.7.7' }
        @{ Label = 'CU3 and FU1'; KB = '4014788'; WinSCP = '5.7.7' }
        @{ Label = 'CU3';         KB = '4039664'; WinSCP = '5.7.7' }
        @{ Label = 'CU2 or FU1';  KB = '4021095'; WinSCP = '5.7.7' }
        @{ Label = 'CU1';         KB = '3208238'; WinSCP = '5.7.7' }
    )

    foreach ($entry in $kbMap) {
        if ($DisplayNames -match $entry.KB) {
            return [pscustomobject]@{
                WinSCPVersion = $entry.WinSCP
                UpdateLabel = $entry.Label
                MatchedKB = $entry.KB
                Confidence = 'high'
            }
        }
    }

    return [pscustomobject]@{
        WinSCPVersion = '5.7.7'
        UpdateLabel = 'RTM/no CU detected'
        MatchedKB = 'none'
        Confidence = 'medium'
    }
}

$bizTalkRegPath = 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0'
$productCodeCurrent = Get-RegistryValueSafe -Path $bizTalkRegPath -Name 'ProductCodeCurrent'
$productName = Get-RegistryValueSafe -Path $bizTalkRegPath -Name 'ProductName'
$productVersion = Get-RegistryValueSafe -Path $bizTalkRegPath -Name 'ProductVersion'
$installPath = Get-RegistryValueSafe -Path $bizTalkRegPath -Name 'InstallPath'
$bizTalkVersion = Resolve-BizTalkVersion -ProductCode $productCodeCurrent -ProductName $productName
$displayNames = Get-InstalledUpdateDisplayNames

$selected = $null
if ($bizTalkVersion -eq '2020') {
    $selected = Resolve-WinSCPFrom2020 -DisplayNames $displayNames
}
elseif ($bizTalkVersion -eq '2016') {
    $selected = Resolve-WinSCPFrom2016 -DisplayNames $displayNames
}

$report = [ordered]@{
    schemaVersion = '1.0'
    generatedUtc = (Get-Date).ToUniversalTime().ToString('o')
    computerName = $env:COMPUTERNAME
    bizTalkDetected = [bool]($null -ne $bizTalkVersion)
    bizTalkVersion = $bizTalkVersion
    productCodeCurrent = $productCodeCurrent
    productName = $productName
    productVersion = $productVersion
    installPath = $installPath
    selectedWinSCPVersion = if ($selected) { $selected.WinSCPVersion } else { $null }
    selectedUpdateLabel = if ($selected) { $selected.UpdateLabel } else { $null }
    matchedKB = if ($selected) { $selected.MatchedKB } else { $null }
    confidence = if ($selected) { $selected.Confidence } else { 'low' }
}

$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Parent $resolvedOutput
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

$report | ConvertTo-Json -Depth 5 | Set-Content -Path $resolvedOutput -Encoding UTF8

Write-Host "BizTalk probe report written: $resolvedOutput" -ForegroundColor Green
$report
