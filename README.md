# BizTalk WinSCP Installer

This repository provides a PowerShell script that installs the correct WinSCP version for Microsoft BizTalk Server.

The script:
- Detects BizTalk Server version (2016 or 2020).
- Detects installed cumulative update (CU), feature pack (FP), or feature update (FU).
- Maps that result to the required WinSCP version.
- Downloads WinSCP from NuGet when needed.
- Copies `WinSCP.exe` and `WinSCPnet.dll` to the BizTalk installation folder.

## Script

- `InstallWinSCPForBizTalk.ps1`

## Requirements

- Windows with PowerShell.
- Microsoft BizTalk Server 2016 or 2020 installed.
- Write access to the BizTalk installation folder.
- Internet access for online installs, unless using offline mode.
- For actual installation, run from an elevated PowerShell session.

## Parameters

- `-nugetDownloadFolder <path>`
  - Folder used for `nuget.exe` and WinSCP package files.
  - Default: `$env:TEMP\nuget`
  - Folder is not deleted automatically.

- `-ForceInstall`
  - Reinstalls WinSCP even if the required version is already present.

The script also supports PowerShell risk-mitigation parameters:
- `-WhatIf`
- `-Confirm`

## Usage

### Default installation

```powershell
.\InstallWinSCPForBizTalk.ps1
```

### Use a custom download folder

```powershell
.\InstallWinSCPForBizTalk.ps1 -nugetDownloadFolder C:\Temp\WinSCP
```

### Force reinstall

```powershell
.\InstallWinSCPForBizTalk.ps1 -ForceInstall
```

### Dry run

```powershell
.\InstallWinSCPForBizTalk.ps1 -WhatIf
```

### Show verbose diagnostics

```powershell
.\InstallWinSCPForBizTalk.ps1 -Verbose
```

## Offline Workflow (Production-Friendly)

If production servers do not have internet access:

1. Run the script in a connected environment with the same BizTalk/CU level.
2. Use a known folder via `-nugetDownloadFolder`.
3. Copy that folder to the offline target server.
4. Run the script on the target server using the same `-nugetDownloadFolder` path.

If required package files already exist, the script reuses them instead of downloading again.

## Build Minimal Production Bundle

To reduce production attack surface, you can build a minimal deployment bundle that
contains only runtime installer essentials (no tests/docs/dev helpers).

This build step does **not** require BizTalk to be installed on the machine that
creates the production bundle.

### Build minimal bundle

```powershell
.\scripts\Build-ProductionPackage.ps1 -Clean
```

Default output:

- `dist\BizTalkWinSCPInstaller-Production`

### Build minimal bundle with offline payload

```powershell
.\scripts\Build-ProductionPackage.ps1 -Clean -NuGetPayloadFolder C:\Temp\winscp-cache
```

When `-NuGetPayloadFolder` is supplied, the payload is copied into `./nuget` in the
bundle and `Run-Installer.ps1` defaults to using that local folder.

### Run in production

Copy the generated bundle to the target server and run:

```powershell
.\Run-Installer.ps1
```

## Compatibility Matrix

The script includes explicit mapping for these versions.

### BizTalk Server 2020

| BizTalk update | KB | WinSCP |
|---|---|---|
| CU6 | 5043408 or 5048971 | 6.3.5 |
| CU5 | 5032870 | 6.1.2 |
| CU4 | 5009901 | 5.19.2 |
| CU3 | 5007969 | 5.19.2 |
| CU2 | 5003151 | 5.17.8 |
| CU1 | 4538666 | 5.17.6 |
| RTM / no CU detected | n/a | 5.15.4 |

### BizTalk Server 2016

| BizTalk update | KB | WinSCP |
|---|---|---|
| CU9 and FP3 | 5005480 | 5.19.2 |
| CU9 | 5005479 | 5.19.2 |
| CU8 and FP3 | 4590075 | 5.15.9 |
| CU8 | 4583530 | 5.15.9 |
| CU7 and FP3 | 4536185 | 5.15.9 |
| CU7 | 4528776 | 5.15.9 |
| CU6 and FP3 | 4294900 | 5.13.1 |
| CU6 | 4477494 | 5.13.1 |
| CU5 and FP3 | 4103503 | 5.13.1 |
| CU5 Hotfix | 4345385 | 5.13.1 |
| CU5 | 4132957 | 5.13.1 |
| CU4 and FP2 | 4094130 | 5.7.7 |
| CU4 | 4051353 | 5.7.7 |
| CU3 and FP2 | 4054819 | 5.7.7 |
| CU3 and FU1 | 4014788 | 5.7.7 |
| CU3 | 4039664 | 5.7.7 |
| CU2 or FU1 | 4021095 | 5.7.7 |
| CU1 | 3208238 | 5.7.7 |
| RTM / no CU detected | n/a | 5.7.7 |

## Notes

- If the required WinSCP version is already installed, the script makes no changes (unless `-ForceInstall` is used).
- The script validates package layout dynamically for different NuGet package structures.
- If BizTalk cannot be detected, the script exits with error details.

## Documentation Generation

The modules in `src/` now include PowerShell comment-based help and module manifests.

Additional design and flow documentation:

- Main script walkthrough and variable inventory: [docs/main-script-walkthrough.md](docs/main-script-walkthrough.md)
- CU mapping maintenance guide: [docs/cu-mapping-maintenance.md](docs/cu-mapping-maintenance.md)

To generate markdown help from function help comments:

```powershell
Install-Module platyPS -Scope CurrentUser
.\scripts\Generate-ModuleHelp.ps1
```

Generated help files are written to `docs/help` by default.

## Credits

- Thomas Canter
- Sandro Pereira
- Michael Stepensen
- Niclas Oberg
