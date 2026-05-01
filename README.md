# BizTalk WinSCP Installer

This repository provides automation to install the correct WinSCP version for
Microsoft BizTalk Server 2016 or 2020, and to build minimal production-ready
bundles with a reduced attack surface.

## Documentation Map

- Changelog: [CHANGELOG.md](CHANGELOG.md)
- Production readiness release summary: [docs/production-readiness-release.md](docs/production-readiness-release.md)
- Maintainer history and evolution summary: [docs/maintainer-history.md](docs/maintainer-history.md)
- Main script walkthrough: [docs/main-script-walkthrough.md](docs/main-script-walkthrough.md)
- CU mapping maintenance: [docs/cu-mapping-maintenance.md](docs/cu-mapping-maintenance.md)

## Quick Start Decision Table

| Situation | Recommended Path | Command(s) |
|---|---|---|
| You can run scripts on the BizTalk server before packaging | Probe-first (highest success) | `./scripts/New-BizTalkProbeReport.ps1` then `./scripts/Build-ProductionPackage.ps1 -ProbeReportPath ... -FetchNuGetPayload -Clean` |
| You cannot run probe but you know exact WinSCP version | Manual fallback | `./scripts/Build-ProductionPackage.ps1 -TargetWinSCPVersion <version> -FetchNuGetPayload -Clean` |
| You already have a curated offline NuGet payload folder | Pre-staged payload | `./scripts/Build-ProductionPackage.ps1 -NuGetPayloadFolder <folder> -Clean` |
| You just want direct install on a BizTalk server | Direct installer | `./InstallWinSCPForBizTalk.ps1` |

## Direct Installer Process Flow

When you run `InstallWinSCPForBizTalk.ps1`, the installer executes four
phases in order:

1. Detect BizTalk version and cumulative update, then map to the required
  WinSCP version.
2. Check the current BizTalk folder state and decide whether installation is
  required.
3. Prepare the download folder, acquire the NuGet package, and copy
  `WinSCP.exe` and `WinSCPnet.dll` into the BizTalk installation folder.
4. Verify the installed files:
  - Always compare installed file versions with the expected WinSCP version.
  - Optionally compare SHA256 hashes against the source package files when
    `-CheckHash` is supplied and the source payload is available.

The verification phase runs after a successful copy and also runs when the
required WinSCP version is already present, so the install flow can be used as
both an installer and a post-install verifier.

## Probe-First Recommended Workflow

This path has the highest probability of success and does not require BizTalk on
the machine that builds the production package.

1. On the BizTalk server, generate a probe report:

```powershell
.\scripts\New-BizTalkProbeReport.ps1 -OutputPath C:\Temp\biztalk-probe.json
```

2. Copy `biztalk-probe.json` to the build machine.

3. Build minimal package and fetch exact payload:

```powershell
.\scripts\Build-ProductionPackage.ps1 -Clean -ProbeReportPath C:\Temp\biztalk-probe.json -FetchNuGetPayload
```

4. Copy output bundle (default `dist\BizTalkWinSCPInstaller-Production`) to target production server.

5. Run installer wrapper from the bundle:

```powershell
.\Run-Installer.ps1
```

For a fully offline post-install integrity check against the bundled payload,
run:

```powershell
.\Run-Installer.ps1 -CheckHash
```

Notes:
- The probe report is copied into the bundle for traceability.
- The bundle manifest records selected WinSCP version and file hashes.

## Manual Fallback Workflow

Use this path only when probe data is unavailable.

1. Choose WinSCP version manually from the compatibility matrix below.

2. Build package with explicit version:

```powershell
.\scripts\Build-ProductionPackage.ps1 -Clean -TargetWinSCPVersion 6.3.5 -FetchNuGetPayload
```

3. Transfer bundle and run:

```powershell
.\Run-Installer.ps1
```

If the bundle includes a local `nuget` payload and you want to verify the copied
files against that payload on the production server, use:

```powershell
.\Run-Installer.ps1 -CheckHash
```

Alternative: if you already maintain a payload cache, use:

```powershell
.\scripts\Build-ProductionPackage.ps1 -Clean -NuGetPayloadFolder C:\Temp\winscp-cache
```

## Full Parameter Reference (Installer + Builder + Probe)

### InstallWinSCPForBizTalk.ps1

Purpose: install the correct WinSCP binaries into BizTalk install folder.

Parameters:
- `-nugetDownloadFolder <path>`
- `-ForceInstall`
- `-CheckHash`
- `-CheckOnly`
- `-LogFolder <path>`
- `-LogLevel <Info|Verbose|Debug>`
- `-EnableEventLog`
- `-EventLogName <name>`
- `-EventSource <name>`
- PowerShell common risk controls: `-WhatIf`, `-Confirm`

Examples:

```powershell
.\InstallWinSCPForBizTalk.ps1
.\InstallWinSCPForBizTalk.ps1 -nugetDownloadFolder C:\Temp\WinSCP
.\InstallWinSCPForBizTalk.ps1 -CheckHash
.\InstallWinSCPForBizTalk.ps1 -CheckOnly
.\InstallWinSCPForBizTalk.ps1 -ForceInstall -Verbose
.\InstallWinSCPForBizTalk.ps1 -WhatIf
```

### scripts/Build-ProductionPackage.ps1

Purpose: create a minimal production bundle with runtime essentials only.

Parameters:
- `-OutputFolder <path>`
- `-NuGetPayloadFolder <path>`
- `-ProbeReportPath <path>`
- `-TargetWinSCPVersion <version>`
- `-FetchNuGetPayload`
- `-Clean`

Rules:
- Use either `-NuGetPayloadFolder` or `-FetchNuGetPayload`, not both.
- `-FetchNuGetPayload` requires either `-ProbeReportPath` or `-TargetWinSCPVersion`.

### scripts/New-BizTalkProbeReport.ps1

Purpose: run on BizTalk machine to produce a JSON report for external packaging.

Parameters:
- `-OutputPath <path>`

Example:

```powershell
.\scripts\New-BizTalkProbeReport.ps1 -OutputPath C:\Temp\biztalk-probe.json
```

## Compatibility Matrix

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

## Troubleshooting / Verification

The built-in verification phase is the default post-install check. Use the
main installer when you want the standard detection, copy, and verification
flow in one command.

### Quick verification checklist

1. Run a dry run:

```powershell
.\InstallWinSCPForBizTalk.ps1 -WhatIf
```

2. Validate local test and complexity gates:

```powershell
.\scripts\Run-Validation.ps1
```

3. If packaging for production, verify bundle contents include:
- `InstallWinSCPForBizTalk.ps1`
- `src/InstallWinSCPForBizTalk.Core.psm1`
- `src/InstallWinSCPForBizTalk.Init.psm1`
- `src/InstallWinSCPForBizTalk.Utils.psm1`
- `src/InstallWinSCPForBizTalk.Workflow.psm1`
- `Run-Installer.ps1`
- `package-manifest.json`

### Post-install verification in the normal flow

Run the installer as usual to execute the full process including verification:

```powershell
.\InstallWinSCPForBizTalk.ps1
```

Add `-CheckHash` when you want the verification phase to also compare SHA256
hashes of the installed files against the source package payload:

```powershell
.\InstallWinSCPForBizTalk.ps1 -CheckHash -nugetDownloadFolder C:\Temp\WinSCP
```

For offline production bundles, the equivalent entry point is:

```powershell
.\Run-Installer.ps1 -CheckHash
```

### Stand-alone verification

If WinSCP is already installed and you only want to verify the current files,
you can use either the main installer in assessment mode or the lower-level
core verification function directly.

Assessment mode with BizTalk/CU detection but no download or copy:

```powershell
.\InstallWinSCPForBizTalk.ps1 -CheckOnly
```

For offline production bundles, the equivalent entry point is:

```powershell
.\Run-Installer.ps1 -CheckOnly
```

Use the core function directly when you already know the target paths and
expected version.

Version-only verification:

```powershell
Import-Module .\src\InstallWinSCPForBizTalk.Core.psm1

Get-WinSCPInstallVerification `
  -TargetExePath 'C:\Program Files (x86)\Microsoft BizTalk Server\WinSCP.exe' `
  -TargetDllPath 'C:\Program Files (x86)\Microsoft BizTalk Server\WinSCPnet.dll' `
  -ExpectedVersion '5.19.2'
```

Version and SHA256 copy-integrity verification:

```powershell
Import-Module .\src\InstallWinSCPForBizTalk.Core.psm1

Get-WinSCPInstallVerification `
  -TargetExePath 'C:\Program Files (x86)\Microsoft BizTalk Server\WinSCP.exe' `
  -TargetDllPath 'C:\Program Files (x86)\Microsoft BizTalk Server\WinSCPnet.dll' `
  -ExpectedVersion '5.19.2' `
  -SourceExePath 'C:\Temp\WinSCP\WinSCP.5.19.2\tools\WinSCP.exe' `
  -SourceDllPath 'C:\Temp\WinSCP\WinSCP.5.19.2\tools\WinSCPnet.dll' `
  -CheckHash $true
```

The standalone function returns a structured object with version matches,
actual detected versions, optional hash matches, and a final
`VerifiedSuccessfully` result.

Use `-CheckOnly` when you want the script to tell you whether the current
BizTalk host already has the correct WinSCP version or needs an upgrade,
without requiring a local package or internet access.

### Common issues

- BizTalk not detected:
  - Ensure BizTalk registry keys are present on target server.
  - Ensure script is run on a BizTalk 2016/2020 host.

- Non-admin write failures:
  - Run elevated on target server for actual installation.

- Offline payload issues:
  - Ensure `nuget.exe` and expected `WinSCP.<version>` package folder are present under bundle `nuget`.

## Advanced Docs / Tools

Additional documentation:
- Production readiness release summary: [docs/production-readiness-release.md](docs/production-readiness-release.md)
- Main script walkthrough: [docs/main-script-walkthrough.md](docs/main-script-walkthrough.md)
- CU mapping maintenance: [docs/cu-mapping-maintenance.md](docs/cu-mapping-maintenance.md)
- Maintainer history and evolution summary: [docs/maintainer-history.md](docs/maintainer-history.md)

Generate markdown help from comment-based PowerShell help:

```powershell
Install-Module platyPS -Scope CurrentUser
.\scripts\Generate-ModuleHelp.ps1
```

Generated help output defaults to `docs/help`.

## Credits

- Thomas Canter
- Sandro Pereira
- Michael Stepensen
- Niclas Oberg
