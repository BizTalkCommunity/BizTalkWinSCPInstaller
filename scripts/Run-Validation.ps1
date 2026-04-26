<#
.SYNOPSIS
Runs local validation checks for this repository.

.DESCRIPTION
Runs unit tests with code coverage and/or cyclomatic complexity checks.
Designed for repeatable local use before commits and pull requests.

.EXAMPLE
./scripts/Run-Validation.ps1
Runs coverage and complexity checks.

.EXAMPLE
./scripts/Run-Validation.ps1 -CoverageOnly
Runs only unit tests with coverage.

.EXAMPLE
./scripts/Run-Validation.ps1 -ComplexityOnly
Runs only complexity checks.
#>
[CmdletBinding()]
param(
    [switch]$CoverageOnly,
    [switch]$ComplexityOnly,
    [switch]$ShowMissedCommands,
    [switch]$FailOnCoverageTarget,
    [int]$CoverageTarget = 75
)

$ErrorActionPreference = 'Stop'

if ($CoverageOnly -and $ComplexityOnly) {
    throw 'Use either -CoverageOnly or -ComplexityOnly, not both.'
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$runCoverage = -not $ComplexityOnly
$runComplexity = -not $CoverageOnly

function Invoke-InCleanPwsh {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Script,

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    Write-Host "`n=== $Title ===" -ForegroundColor Cyan
    & pwsh -NoProfile -Command $Script
    if ($LASTEXITCODE -ne 0) {
        throw "$Title failed with exit code $LASTEXITCODE."
    }
}

if ($runCoverage) {
    $coverageTemplate = @'
Set-Location '__ROOT__'

$cfg = New-PesterConfiguration
$cfg.Run.Path = 'tests/Unit'
$cfg.Run.PassThru = $true
$cfg.Output.Verbosity = 'None'
$cfg.CodeCoverage.Enabled = $true
$cfg.CodeCoverage.Path = @('InstallWinSCPForBizTalk.ps1','src/InstallWinSCPForBizTalk.Utils.psm1','src/InstallWinSCPForBizTalk.Core.psm1')
$cfg.CodeCoverage.CoveragePercentTarget = __TARGET__

$r = Invoke-Pester -Configuration $cfg
$coverage = [math]::Round($r.CodeCoverage.CoveragePercent, 2)
"Coverage: $coverage% ($($r.CodeCoverage.CommandsExecutedCount)/$($r.CodeCoverage.CommandsAnalyzedCount))"
"Tests: passed=$($r.PassedCount) failed=$($r.FailedCount) skipped=$($r.SkippedCount)"

if (__SHOW_MISSED__ -and $r.CodeCoverage -and ($r.CodeCoverage.PSObject.Properties.Name -contains 'MissedCommands')) {
    $missed = @($r.CodeCoverage.MissedCommands)
    if ($missed.Count -gt 0) {
        'Missed commands:'
        $missed | Select-Object File, Line, Command | Format-Table -AutoSize
    }
}

if (__FAIL_ON_TARGET__ -and $coverage -lt __TARGET__) {
    throw "Coverage $coverage% is below target __TARGET__."
}

if ($r.FailedCount -gt 0) { exit 1 }
'@

    $showMissedLiteral = if ($ShowMissedCommands) { '$true' } else { '$false' }
    $failOnTargetLiteral = if ($FailOnCoverageTarget) { '$true' } else { '$false' }

    $coverageScript = $coverageTemplate
    $coverageScript = $coverageScript.Replace('__ROOT__', $repoRoot.Replace("'", "''"))
    $coverageScript = $coverageScript.Replace('__TARGET__', [string]$CoverageTarget)
    $coverageScript = $coverageScript.Replace('__SHOW_MISSED__', $showMissedLiteral)
    $coverageScript = $coverageScript.Replace('__FAIL_ON_TARGET__', $failOnTargetLiteral)

    Invoke-InCleanPwsh -Script $coverageScript -Title 'Coverage Validation'
}

if ($runComplexity) {
    $complexityTemplate = @'
Set-Location '__ROOT__'

$cfg = New-PesterConfiguration
$cfg.Run.Path = 'tests/Unit/CyclomaticComplexity.Tests.ps1'
$cfg.Run.PassThru = $true
$cfg.Output.Verbosity = 'Normal'

$r = Invoke-Pester -Configuration $cfg
"Complexity tests: passed=$($r.PassedCount) failed=$($r.FailedCount) skipped=$($r.SkippedCount)"
if ($r.FailedCount -gt 0) { exit 1 }
'@

    $complexityScript = $complexityTemplate.Replace('__ROOT__', $repoRoot.Replace("'", "''"))

    Invoke-InCleanPwsh -Script $complexityScript -Title 'Complexity Validation'
}

Write-Host "`nValidation complete." -ForegroundColor Green
