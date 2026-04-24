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

### Planned Tests (Stage 2+)

- `CuDetection.Tests.ps1` — `Get-BTSCumulativeUpdateByDisplayName`, `Search-BTSCumulativeUpdate` — CU/KB mapping, DisplayName parsing, registry detection
- `AdminCheck.Tests.ps1` — `Test-IsAdministrator` — Elevation detection, non-admin context
- `InstallFlow.Tests.ps1` — Installation orchestration — Full flow with mocked downloads, registry checks, file copies

## Current stage

**Stage 1**: Scaffolding core functions and basic unit tests.

- `src/InstallWinSCPForBizTalk.Core.psm1` - Extracted testable functions
- `tests/Unit/Resolve-WinSCPPackageLayout.Tests.ps1` - Basic tests with mocking pattern

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

Stage 2: Add tests for CU detection (Search-BTSCumulativeUpdate, Get-BTSCumulativeUpdateByDisplayName).  
Stage 3: Add integration tests for full install flow with mocked registry/download.
