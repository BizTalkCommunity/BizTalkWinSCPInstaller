# Production Readiness Release — v2.0.0

This document summarizes the repository state represented by the
`prod-bundle-iteration` feature work and the `v2.0.0` release. It is written
for maintainers, reviewers, and operators who need a release-oriented
explanation of what changed and why this branch is a major version increment.

## Versioning Rationale

The team considered the pre-branch state on `main` to be v1.0, even though
it was never formally tagged. This branch is `v2.0.0` for two reasons.

First, the scope of change is major by any reasonable definition. The
monolithic installer script was replaced by a modular architecture. New
operational tooling was introduced. The deployment model expanded from a
direct online install to a probe-driven offline packaging workflow. Structured
logging, production bundle creation, and a test/CI foundation were all absent
in v1.0 and are present in v2.0.0.

Second, the internal interface changed materially. The installer now depends on
three modules that did not exist in v1.0. Anyone running the v1.0 script
directly would need to adopt the new structure. That is a breaking change in
practice.

`v2.0.0` is tagged after merging this feature branch to `main`.

## Release Intent

The goal of this release is to move the repository from a script that can
perform the installation to a production-oriented toolchain that can:

- detect the correct WinSCP target version more reliably,
- validate behavior with automated tests and CI,
- build a minimal deployable bundle,
- support offline and staged production workflows,
- and provide better diagnostics when installation or packaging fails.

## What This Release Adds

### Probe-first offline packaging workflow

The most important new operational capability is the separation of environment
discovery from bundle creation.

Workflow:

1. Run `scripts/New-BizTalkProbeReport.ps1` on the BizTalk server.
2. Move the generated JSON probe report to a build or packaging machine.
3. Run `scripts/Build-ProductionPackage.ps1` to fetch or assemble the exact
   payload and produce a minimal deployment bundle.
4. Transfer the bundle to the target environment and run `Run-Installer.ps1`.

Why it matters:

- Packaging no longer depends on BizTalk being installed on the build machine.
- Deployment can fit environments with tighter network or change-control
  boundaries.
- The selected WinSCP version and bundle contents become easier to trace.

### Minimal production bundle creation

The bundle builder creates a smaller, curated output that contains only the
runtime essentials needed for installation. That supports a reduced attack
surface and makes bundle inspection easier.

### Modular installer architecture

The main installer is now orchestration over dedicated modules for core logic,
output/logging utilities, and workflow phases. This improves testability,
reviewability, and future maintenance.

### Structured diagnostics

The repository now supports structured file logging and an optional Windows
Event Log sink. This improves supportability in real deployment scenarios.

### Validation and CI gates

The repository now has a Pester-based test suite, CI execution, and complexity
gates. The branch is therefore better positioned for production changes without
reintroducing monolithic control flow.

### Coverage and complexity as supportability improvements

This branch should not be described as "just more tests". The increase in test
coverage and the addition of complexity gates directly improved supportability.

Current signals from the repository validation artifacts:

- `tests/Unit` now contains 17 focused unit test files.
- `coverage.xml` shows approximately 98.91% line coverage for
  `InstallWinSCPForBizTalk.ps1`.
- `coverage.xml` shows approximately 99.57% line coverage for the `src`
  package modules.
- Complexity is no longer left to reviewer judgment alone; it is enforced by
  `tests/Unit/CyclomaticComplexity.Tests.ps1`.

Why that matters operationally:

- High coverage on the installer and extracted modules makes hotfixes safer,
  because maintainers can change CU mapping, workflow branching, output, or
  logging with immediate regression feedback.
- Complexity gates reduce the chance that the main installer script regresses
  back into an oversized branch-heavy implementation that is hard to review,
  debug, and support.
- Cognitive complexity limits matter as much as cyclomatic complexity here,
  because support incidents are not only about path count. They are also about
  how difficult it is for a maintainer to reconstruct intent quickly under time
  pressure.
- Together, high coverage and explicit complexity limits move the repository
  toward predictable maintenance rather than heroic debugging.

The current complexity gates also show the intended architectural shape:

- `InstallWinSCPForBizTalk.ps1` is allowed to orchestrate, but within bounded
  script-body and per-function complexity thresholds.
- The extracted modules have near-zero script-body complexity targets, which
  reinforces the rule that meaningful behavior should live in functions and
  dedicated workflow helpers rather than module top-level code.
- Per-function thresholds in `Core`, `Workflow`, and `Utils` keep individual
  units reviewable and testable.

## Important Defects Addressed

This release also closes several defects or reliability gaps that existed in the
earlier implementation.

### CU detection hardening

BizTalk servicing detection is more resilient, especially around BizTalk 2020
CU6 and later probe-driven detection paths. This reduces the risk of selecting
the wrong WinSCP version due to update-detection edge cases.

### Boolean control-flow fix

A real PowerShell boolean assignment defect was fixed by changing `false` to
`$false` in installer logic.

### Duplicate helper drift risk removed

Installer helper functions that had existed in both the main script and module
code were consolidated. That removed a maintenance hazard where behavior could
quietly diverge.

### Output consistency cleanup

Banner formatting, dead footer behavior, and stray console output were cleaned
up and covered by tests, improving operator readability.

## Release Readiness Signals

The strongest indicators that this branch is suitable as a production-readiness
release are:

- current validation artifacts showing approximately 98.91% line coverage for
  the top-level installer and approximately 99.57% line coverage across the
  extracted source modules,
- 17 focused unit test files spanning CU detection, workflow branching, install
  flow, package layout, output, logging, and production packaging,
- broader automated coverage for CU detection, workflow branching, install flow,
  package layout, logging, and production packaging,
- a validation runner for local maintainers,
- explicit cyclomatic and cognitive complexity gates that keep support-critical
  paths within bounded reviewable limits,
- explicit documentation for the production bundle workflow,
- centralized compatibility mapping,
- and post-failure diagnostic improvements.

## Supportability Impact

One of the most important release outcomes is improved supportability.

Before this branch, support depended more heavily on understanding a largely
monolithic script and reproducing failures manually. After this branch:

- maintainers can validate changes locally with a repeatable test and coverage
  path,
- complexity gates help keep future changes understandable,
- logging provides durable evidence after failures,
- and the extracted workflow structure makes incident triage faster because the
  execution phases are clearer.

That combination is what makes the production-readiness claim credible. The
branch does not only add features; it reduces the cost and risk of supporting
them over time.

## Maintainer Notes

If this branch is used as the basis for a production-readiness release, future
changes should preserve these rules:

- Keep compatibility and CU mapping logic centralized.
- Add tests alongside any change to package layout assumptions or workflow
  branching.
- Keep orchestration in `InstallWinSCPForBizTalk.ps1` thin.
- Treat operator-facing output and logging as supported behavior, not incidental
  implementation detail.

## Related Documents

- `docs/maintainer-history.md`
- `docs/main-script-walkthrough.md`
- `docs/cu-mapping-maintenance.md`