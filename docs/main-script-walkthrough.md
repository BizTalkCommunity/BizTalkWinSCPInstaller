# Main Script Walkthrough

This document explains what [InstallWinSCPForBizTalk.ps1](../InstallWinSCPForBizTalk.ps1) does, how state flows between phases, and what each variable is responsible for.

## What The Main Script Is Doing

1. Defines user inputs and imports the three modules.
2. Initializes baseline state and default WinSCP metadata.
3. Builds one shared state object called `workflowContext` that all phase functions mutate.
4. Reads BizTalk install/product facts from environment and registry.
5. Runs Phase 1: detect BizTalk and CU mapping.
6. Runs Phase 2: check current install and validate version.
7. Runs Phase 3: prepare folder, acquire package, and copy binaries.
8. Computes final outcome and prints standardized terminal messaging.

## Variable Inventory (Main Script)

### Inputs and module paths

- `nugetDownloadFolder`: Folder used for NuGet EXE and extracted WinSCP package.
- `ForceInstall`: Forces reinstall if required version is already present.
- `coreModulePath`: Path to core functions module.
- `utilsModulePath`: Path to output/snapshot helpers module.
- `workflowModulePath`: Path to phase orchestrator module.

### Global flow flags and constants

- `Continue`: Master gate; `false` stops further execution.
- `PrerequisiteFailure`: Tracks hard-stop prerequisite failures.
- `isAdministrator`: Current elevation status from `Test-IsAdministrator`.
- `winSCPVersion`: Required WinSCP version; starts with a safe default then is replaced by CU logic.
- `winSCPexeFile`: `WinSCP.exe` file name token.
- `winSCPdllFile`: `WinSCPnet.dll` file name token.
- `hashString`: Hash delimiter text for output blocks.
- `bangString`: Bang delimiter text for output blocks.

### Shared mutable state object

- `workflowContext`: Central hashtable carrying state between phase functions.
- `workflowContext.Continue`: Phase-level continuation flag.
- `workflowContext.PrerequisiteFailure`: Phase-level prerequisite failure state.
- `workflowContext.isAdministrator`: Elevation state passed to phases.
- `workflowContext.ForceInstall`: Bool copy of `ForceInstall` switch.
- `workflowContext.WhatIf`: Bool copy of `WhatIfPreference`.
- `workflowContext.winSCPexeFile`: EXE file name for path resolution.
- `workflowContext.winSCPdllFile`: DLL file name for path resolution.
- `workflowContext.hashString`: Banner token used by phase output.
- `workflowContext.bangString`: Banner token used by phase output.
- `workflowContext.nugetDownloadFolder`: Working folder for package acquisition.
- `workflowContext.psCmdlet`: Cmdlet context for `ShouldProcess` compatibility.
- `workflowContext.InvokeWebRequest`: Injectable download callback (testability boundary).

### Raw BizTalk discovery inputs (before Phase 1)

- `bizTalkInstallFolderFromEnv`: `BTSINSTALLPATH` value.
- `bizTalkInstallFolderFromRegistry`: Registry `InstallPath` value.
- `bizTalkProductCodeCurrent`: Registry product code.
- `bizTalkProductName`: Registry product name.
- `bizTalkProductVersion`: Registry product version.

### Phase 1 synced outputs

- `Continue`: Updated continuation flag after detection/prerequisite/CU checks.
- `PrerequisiteFailure`: Updated prerequisite failure state.
- `bizTalkInstallFolder`: Resolved BizTalk install path.
- `bizTalkInstallFolderExists`: Path existence check.
- `bizTalkProductCodeCurrent`: Normalized product code from phase result.
- `bizTalkProductName`: Normalized product name from phase result.
- `bizTalkProductVersion`: Normalized product version from phase result.
- `BizTalkVersion`: Mapped major version label (`2016`/`2020`).
- `winSCPVersion`: CU-mapped required WinSCP version.
- `btsKB`: Matched CU KB id.
- `bizTalkCUVer`: Matched CU label.
- `CUFound`: Whether CU was explicitly detected.

### Phase 2 synced outputs

- `Continue`: Continuation state after existing-install and validation checks.
- `winSCPVersion`: Validated required version.
- `btsWinSCPEXEProductVersionInstalled`: Current target EXE version if present.
- `btsWinSCPDLLProductVersionInstalled`: Current target DLL version if present.
- `btsWinSCPProductInstalledAndCorrect`: Whether current target install already satisfies requirement.
- `installExecutionPlan`: Structured plan object from gating logic.
- `winSCPProductVersionRequired`: Normalized required version for comparison.
- `btsTargetWinSCPExe`: Target EXE path in BizTalk folder.
- `btsTargetWinSCPDll`: Target DLL path in BizTalk folder.

### Phase 3 synced outputs

- `Continue`: Continuation state after package/download/copy operations.
- `winSCPVersion`: Final resolved version used during copy.
- `btsWinSCPProductInstalledAndCorrect`: Final installation correctness flag.
- `WinSCPEXEDownload`: Resolved source EXE path from package.
- `WinSCPDllDownload`: Resolved source DLL path from package.
- `WinSCPTargetEXEExists`: Post-copy target EXE existence.
- `WinSCPDLLTargetExists`: Post-copy target DLL existence.

### Finalization

- `finalOutcome`: Structured terminal classification (`Success`, `DryRun`, `PrerequisiteFailure`, `InstallFailure`/`Unknown`).

## High-Value Understanding Notes

- The script is intentionally a linear orchestrator; real logic lives in modules.
- The most important object is `workflowContext`. If you track that object, script flow becomes straightforward.
- Variable sync blocks after each phase are checkpoints that mirror `workflowContext` into script scope for readability and final reporting.

## Readability Cleanup Status

- Legacy/unused variables `WinSCPexe`, `winSCPdll`, and `upString` were removed from the main script.
- Package path resolution remains centralized in workflow/core logic.
- Behavior was kept identical; this was a readability-only cleanup.
