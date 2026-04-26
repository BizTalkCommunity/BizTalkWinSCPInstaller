# Unit Tests for InstallWinSCPForBizTalk

## Setup

Install Pester (if needed):
```powershell
Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -SkipPublisherCheck -Force
```

## Run tests

From the repository root:
```powershell
Invoke-Pester tests/Unit -Verbose
```

Or run a specific test file:
```powershell
Invoke-Pester tests/Unit/Resolve-WinSCPPackageLayout.Tests.ps1 -Verbose
```

## Test Suite

### Current Tests

| Test File | Function | Coverage |
|-----------|----------|----------|
| `Resolve-WinSCPPackageLayout.Tests.ps1` | `Resolve-WinSCPPackageLayout` | Package layout resolution: known paths, missing root, recursive fallback, partial resolution |
| `CuDetection.Tests.ps1` | `Search-BTSCumulativeUpdate`, `Get-BTSCumulativeUpdateByDisplayName` | CU/KB detection for BizTalk 2016/2020, CU6 dual-KB coverage, fallback behavior |
| `AdminCheck.Tests.ps1` | `Test-IsAdministrator` | Elevation detection behavior via dependency-injected identity/principal stubs |
| `InstallFlow.Tests.ps1` | `Get-InstallExecutionPlan` | ForceInstall/WhatIf/admin gating, already-installed skip/reinstall decisions |
| `VersionValidation.Tests.ps1` | `Test-WinSCPVersionString` | Missing/invalid/valid WinSCP version validation and parsing behavior |
| `PackageReadiness.Tests.ps1` | `Get-PackageReadinessState` | NuGet and package artifact readiness (missing/incomplete/ready) |
| `FinalOutcome.Tests.ps1` | `Get-FinalExecutionOutcome` | Final outcome classification (success, dry-run, prerequisite failure, install failure) |
| `Build-ProductionPackage.Tests.ps1` | `scripts/Build-ProductionPackage.ps1` | Minimal production bundle creation and optional offline payload inclusion |
| `New-BizTalkProbeReport.Tests.ps1` | `scripts/New-BizTalkProbeReport.ps1` | Probe report generation for non-BizTalk build workflows |
| `CyclomaticComplexity.Tests.ps1` | Script body and functions | Cyclomatic complexity gates enforced in local runs and CI |

### Planned Tests (Stage 2+)

- Full installer orchestration tests — end-to-end decision flow with mocked registry/download/copy operations
- WhatIf/ShouldProcess interaction tests at command level
- Full installer orchestration tests with mocked command invocations and explicit message assertions

## Current stage

**Stage 4**: Expanded core unit coverage including validation, readiness, and final-outcome classification.

- `src/InstallWinSCPForBizTalk.Core.psm1` - Extracted and testable core functions
- `tests/Unit/Resolve-WinSCPPackageLayout.Tests.ps1` - Package layout resolution tests
- `tests/Unit/CuDetection.Tests.ps1` - CU detection and mapping tests
- `tests/Unit/AdminCheck.Tests.ps1` - Admin/elevation behavior tests
- `tests/Unit/InstallFlow.Tests.ps1` - Install-flow decision tests
- `tests/Unit/VersionValidation.Tests.ps1` - WinSCP version validation tests
- `tests/Unit/PackageReadiness.Tests.ps1` - Package/download readiness classification tests
- `tests/Unit/FinalOutcome.Tests.ps1` - Final outcome classification tests
- `tests/Unit/CyclomaticComplexity.Tests.ps1` - Cyclomatic complexity regression gates for source files

### What's tested

- **Resolve-WinSCPPackageLayout**: Resolves WinSCP files from NuGet package structure
  - Known folder structure (tools/lib paths)
  - Missing package root
  - Fallback recursive search
  - Partial resolution (EXE only)

### Mocking pattern

Tests use Pester `Mock` to override:
- `Test-Path` - file system checks
- `Get-ChildItem` - recursive file search

This allows tests to run without a real NuGet package or file system.

## Next stages

Stage 5: Add integration-style tests for full install flow with mocked registry/download/copy operations and ShouldProcess message-level assertions.
