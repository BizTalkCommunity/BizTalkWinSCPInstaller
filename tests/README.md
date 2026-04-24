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

### Planned Tests (Stage 2+)

- Full installer orchestration tests — end-to-end decision flow with mocked registry/download/copy operations
- Error-path tests — NuGet download failure, package extraction failure, copy failure messaging
- WhatIf/ShouldProcess interaction tests at command level

## Current stage

**Stage 3**: Core unit test coverage for package layout, CU detection, admin checks, and install-flow decisions.

- `src/InstallWinSCPForBizTalk.Core.psm1` - Extracted and testable core functions
- `tests/Unit/Resolve-WinSCPPackageLayout.Tests.ps1` - Package layout resolution tests
- `tests/Unit/CuDetection.Tests.ps1` - CU detection and mapping tests
- `tests/Unit/AdminCheck.Tests.ps1` - Admin/elevation behavior tests
- `tests/Unit/InstallFlow.Tests.ps1` - Install-flow decision tests

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

Stage 4: Add integration-style tests for full install flow with mocked registry/download/copy operations.
