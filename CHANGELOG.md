# Changelog

All notable changes to this project will be documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
This project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.0.0] — 2026-04-26

This is a major release. The v1.0 codebase was a monolithic installer script.
v2.0.0 replaces that with a modular, test-driven toolchain that supports both
direct online installation and a probe-driven offline packaging workflow. The
internal architecture is a breaking change: the installer now depends on three
extracted modules that did not exist in v1.0.

### Added

- `src/InstallWinSCPForBizTalk.Core.psm1` — core decision, validation, and
  install helper functions extracted from the monolithic installer script.
- `src/InstallWinSCPForBizTalk.Utils.psm1` — output formatting, logging
  sinks, and semantic message helpers.
- `src/InstallWinSCPForBizTalk.Workflow.psm1` — named workflow phase functions
  that orchestrate the install sequence.
- `scripts/Build-ProductionPackage.ps1` — creates a minimal deployment bundle
  containing only the runtime essentials needed for installation. Supports
  probe-driven version selection, explicit version pinning, and pre-staged
  NuGet payload folders.
- `scripts/New-BizTalkProbeReport.ps1` — inspects a BizTalk server and emits
  a JSON report for use on a separate packaging machine, enabling fully offline
  bundle creation.
- `scripts/Run-Validation.ps1` — local validation runner covering unit tests,
  coverage, and complexity gates.
- `scripts/Generate-ModuleHelp.ps1` — generates platyPS markdown help from
  comment-based PowerShell help in each module.
- Structured file logging via `-LogFolder`, `-LogLevel` (`Info|Verbose|Debug`),
  and an optional Windows Event Log sink via `-EnableEventLog`, `-EventLogName`,
  and `-EventSource` parameters on the main installer.
- 17 Pester unit test files under `tests/Unit` covering CU detection, admin
  checks, install flow, workflow phases, package layout resolution, package
  readiness, version validation, WinSCP target install, output helpers, banner
  output, final outcome, discovery and download plans, logging, production
  package builder, probe report, stochastic property tests, and cyclomatic and
  cognitive complexity gates.
- GitHub Actions workflow (`.github/workflows/unit-tests.yml`) running Pester 5
  on `windows-latest` on every push.
- AST-based cyclomatic and cognitive complexity gates enforced in CI for all
  source files.
- `docs/main-script-walkthrough.md` — step-by-step walkthrough of installer
  execution flow.
- `docs/cu-mapping-maintenance.md` — guidance for updating the CU-to-WinSCP
  version compatibility mapping.
- `docs/maintainer-history.md` — evolution history from v1.0 to v2.0.0,
  including key defect fixes and design rationale.
- `docs/production-readiness-release.md` — release summary, versioning
  rationale, coverage and complexity signals, and supportability impact.
- Probe-first recommended workflow and manual fallback workflow documented in
  `README.md` with a quick-start decision table and full parameter reference.
- BizTalk 2020 CU6 dual-KB detection: both `KB5043408` and `KB5048971` are
  recognised as valid CU6 identifiers in both the direct installer and the
  probe workflow.

### Changed

- `InstallWinSCPForBizTalk.ps1` is now a thin orchestrator. All deep logic
  has moved into the three extracted modules. The script parameters are extended
  with logging options (`-LogFolder`, `-LogLevel`, `-EnableEventLog`,
  `-EventLogName`, `-EventSource`).
- Installer execution is now structured as explicit named phases, making the
  sequence observable and testable.
- All operator-facing output and final outcome messaging is centralised in
  `InstallWinSCPForBizTalk.Utils.psm1` rather than scattered `Write-Host`
  calls.
- Banner delimiters and final outcome messages are now consistent and paired
  correctly.
- `README.md` reorganised around operational workflows, a compatibility matrix,
  a troubleshooting section, and a documentation map.
- Module manifests bumped to `2.0.0`.

### Fixed

- **Boolean assignment defect.** An installer branch used `false` instead of
  `$false`, causing incorrect control flow in install-folder detection. Fixed
  in `66eece7`.
- **CU detection resilience.** CU detection hardened against edge cases in
  installed-update discovery, particularly around BizTalk 2020 CU6. Fixed in
  `fa8b773`.
- **Duplicate helper logic.** `Search-BTSCumulativeUpdate`,
  `Get-BTSCumulativeUpdateByDisplayName`, `Test-IsAdministrator`, and
  `Resolve-WinSCPPackageLayout` existed in both the installer script and the
  core module. The installer copies were removed; the module is the single
  source of truth. Fixed in `b433aa4`.
- **Dead banner footer path.** A legacy output path emitted a banner footer
  that was no longer reachable. Removed and regression-tested in `c0f76f4`.
- **Mismatched bang banner delimiters.** Final outcome and not-installed notice
  messages used unpaired delimiters. Fixed in `66e52b5`.
- **Stray hash banner in existing-install check.** An unintended hash banner
  was emitted after the existing WinSCP status message. Removed in `7f669ef`.

### Removed

- Inline `Search-BTSCumulativeUpdate`, `Get-BTSCumulativeUpdateByDisplayName`,
  `Test-IsAdministrator`, and `Resolve-WinSCPPackageLayout` definitions from
  `InstallWinSCPForBizTalk.ps1`. These now live exclusively in
  `InstallWinSCPForBizTalk.Core.psm1`.

---

## [1.0.0] — 2026-01-23

Informal v1.0 baseline. The repository provided a monolithic
`InstallWinSCPForBizTalk.ps1` script that detected the installed BizTalk
Server version and cumulative update, selected the correct WinSCP version,
downloaded it via NuGet, and copied it to the BizTalk installation folder.

Final commit of this baseline: `0bc63cf00740a85d09a4c0335f081be49bfd8aed`
_(Update BizTalk Server 2020 CU6 and CU5 details)_

No formal git tag was applied to this version at the time.

[2.0.0]: https://github.com/BizTalkCommunity/BizTalkWinSCPInstaller/compare/main...feature/prod-bundle-iteration
[1.0.0]: https://github.com/BizTalkCommunity/BizTalkWinSCPInstaller/commit/0bc63cf00740a85d09a4c0335f081be49bfd8aed
