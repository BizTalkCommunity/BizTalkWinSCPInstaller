# Maintainer History

This document explains how the repository evolved from the v1.0 baseline
(`0bc63cf00740a85d09a4c0335f081be49bfd8aed`, the last commit on `main` prior
to the `prod-bundle-iteration` feature branch) to the current v2.0.0 state. It
is intended for future maintainers who need to understand why the repository
structure changed, what defects were fixed, and which changes were primarily
capability additions versus refactoring.

The team considered the pre-branch `main` state to be v1.0 informally. This
branch is v2.0.0 because the internal architecture, deployment model, and
operational tooling changed significantly enough to be a breaking major version
increment.

## Baseline

The original repository state was centered on a largely monolithic
`InstallWinSCPForBizTalk.ps1` script. Most of the CU detection, environment
checks, download handling, package layout handling, output, and install flow
were implemented directly in that script.

That baseline worked, but it had three maintainability constraints:

- Critical behavior lived in one file with branching-heavy control flow.
- Some helper behavior later ended up duplicated between the script and module
  implementations.
- There was little automated protection against regressions in CU mapping,
  package discovery, and negative-path behavior.

## What Changed

Between April 23 and April 26, 2026, the repository evolved in four major
stages.

### 1. Detection hardening and early defect fixes

The first wave focused on making CU detection more resilient, especially around
BizTalk 2020 CU6 and related installed-update discovery behavior.

Important outcomes:

- CU detection became more defensive and easier to reason about.
- The codebase later added support for BizTalk 2020 CU6 being represented by
  either `KB5043408` or `KB5048971`, which reduced false negatives when probing
  target environments.
- A real PowerShell control-flow defect was fixed by correcting `false` to
  `$false` in installer logic.

### 2. Test suite and CI introduction

The next wave added a real Pester test suite, plus CI enforcement.

Important outcomes:

- CU detection behavior became covered by unit tests.
- Admin checks, install flow, error paths, workflow branches, package layout,
  output formatting, logging, and production packaging gained targeted tests.
- A GitHub Actions workflow began running the test suite automatically.
- Cyclomatic and cognitive complexity gates were added to keep future changes
  from collapsing back into an unreviewable monolith.

This stage is also where supportability changed materially, not just code
quality on paper. The repository moved from "maintainers must reason about the
script manually" toward "maintainers can rely on repeatable feedback before and
after a change".

### 3. Modularization of the installer

The largest structural change was the extraction of behavior from the main
installer script into dedicated modules:

- `src/InstallWinSCPForBizTalk.Core.psm1`
- `src/InstallWinSCPForBizTalk.Utils.psm1`
- `src/InstallWinSCPForBizTalk.Workflow.psm1`

This shifted the role of `InstallWinSCPForBizTalk.ps1` from "do everything" to
"orchestrate named phases using shared modules".

Important outcomes:

- Core detection, planning, and install helpers became reusable and testable.
- Output handling and final outcome messaging became centralized.
- Workflow phases became explicit, making the execution path easier to follow.
- Duplicate helper definitions were removed, which reduced drift risk between
  script-local and shared implementations.

### 4. Production packaging and offline workflow

The final wave added operational tooling rather than only installer internals.

Important outcomes:

- `scripts/Build-ProductionPackage.ps1` can build a minimal bundle with only the
  runtime essentials needed for deployment.
- `scripts/New-BizTalkProbeReport.ps1` can inspect a BizTalk server and emit a
  JSON report used later on a separate packaging machine.
- The repository now supports a probe-first, offline-friendly workflow rather
  than assuming packaging and installation always happen on the same machine.
- Structured file logging and optional Windows Event Log output improved
  diagnosability in production use.

## Key Defects Fixed

The most important fixes in this evolution were not cosmetic. They changed how
reliably the repository behaves.

### CU detection resilience

CU detection was hardened so the installer and the probe workflow are less
likely to misclassify valid BizTalk environments. This matters because WinSCP
version selection depends on correct BizTalk servicing detection.

### Boolean assignment bug

An installer path used `false` instead of `$false`. In PowerShell that is not a
safe boolean literal substitution. Fixing that corrected control flow in a real
execution branch.

### Duplicate helper logic removal

Functions such as CU detection, admin checks, and package layout resolution were
removed from the installer script and sourced from the core module instead. This
prevented a class of bugs where one copy of the logic would be updated and the
other would silently lag behind.

### Output and banner cleanup

Several commits cleaned up dead footer logic, mismatched banner delimiters, and
stray banner output in the existing-install check. These were smaller defects,
but they improved operator clarity and reduced noisy output during execution.

### Diagnostics improvements

File logging and optional Windows Event Log support do not change the install
result directly, but they fix an operational weakness in the earlier version:
when something failed, troubleshooting depended too heavily on transient console
output.

## Before And After

The practical difference between the baseline and the current repository is:

- Before: a useful but mostly monolithic installer script.
- After: a maintainable installation system with modules, tests, CI, logging,
  production bundle tooling, and an offline probe-driven workflow.

That distinction matters for maintainers. Most of the added files are not noise;
they represent deliberate moves to reduce regression risk and support real-world
operations.

## Coverage and Complexity Significance

The coverage and complexity work deserves explicit credit because it changed the
maintenance profile of the repository.

Evidence in the current branch state:

- The repository now includes 17 focused unit test files under `tests/Unit`.
- A key midpoint commit explicitly raised installer test coverage to 86%.
- Current validation artifacts show approximately 98.91% line coverage for the
  top-level installer script.
- Current validation artifacts show approximately 99.57% line coverage across
  the extracted `src` modules package.
- Cyclomatic and cognitive complexity are enforced by repository tests instead
  of being left to informal review.

Why this matters for maintainers:

- High coverage makes support fixes less risky because the important execution
  paths are no longer protected only by manual reasoning.
- Complexity reduction made the execution path easier to review, explain, and
  troubleshoot when production behavior is under question.
- Complexity gates matter over time because they preserve the benefit of the
  refactor. Without gates, future edits could slowly rebuild the same support
  burden the branch was trying to remove.
- The combination of modular extraction, high coverage, and complexity limits is
  a large part of why this branch can be described as a production-readiness
  milestone rather than only a feature branch.

## Commit Timeline Summary

The high-level timeline across the range was:

1. Harden CU flow and detection behavior.
2. Add Pester test scaffolding, branch coverage, and CI.
3. Fix the boolean control-flow defect.
4. Remove duplicated script helpers and shift behavior into modules.
5. Add complexity gates and continue workflow extraction.
6. Clean up output consistency and banner behavior.
7. Add structured logging and optional Event Log support.
8. Add validation runner, production bundle builder, and probe-first offline
   packaging.

## High-Signal Commit Appendix

This appendix is intentionally selective. It captures the commits that are most
useful when reconstructing the evolution of the production-readiness work.

### `fa8b773` Harden CU6 flow and CU detection resilience

This was the first clear behavior-hardening commit in the range. It improved the
reliability of servicing detection, which is critical because WinSCP version
selection depends on correctly identifying BizTalk update state.

### `66eece7` Increase installer test coverage to 86% and fix boolean assignment bug

This commit combined a large test coverage improvement with a real runtime fix:
an installer branch used `false` instead of `$false`. That corrected control
flow in a path that affects install-folder handling.

### `b433aa4` Remove duplicate functions from installer script and import core module

This commit marks the point where shared logic stopped being maintained in two
places. That reduced the risk of silent divergence between the top-level script
and the reusable implementation.

### `99c484f` Refactor installer output/flow helpers and add cognitive complexity gates

This was a key readability and maintainability milestone. It reinforced the
rule that complex flow and operator output should move into helpers rather than
continue to grow inside the main script.

### `acc2320` Refactor installer workflow extraction and extend complexity gates

This commit pushed the repository further toward explicit workflow phases. It is
one of the clearest transition points between the older monolithic structure and
the current orchestration-based script design.

### `66e52b5` Pair banner delimiters in final and notice output

This is representative of the operator-experience cleanup work. The defect was
small, but it matters because final output is part of how maintainers and
operators judge whether the script behaved correctly.

### `7f669ef` Remove stray hash banner from existing check

This commit removed a noisy, unintended output artifact and locked the behavior
down with tests. It is a good example of using tests to preserve console UX.

### `774ad2d` Add structured installer file logging

This was a major operational step toward production readiness. It made installer
failures diagnosable after the fact instead of relying only on transient console
output.

### `30f9380` Add optional Windows Event Log sink

This extended the logging model to environments where central Windows log
collection is preferred. It is part of the supportability story rather than a
pure functional feature.

### `3687dc6` Add logging coverage tests and validation runner

This commit made the new logging behavior verifiable and added a clearer local
validation path for maintainers before merge or release.

### `370db65` Add minimal production bundle builder and docs

This is where the repository stopped being only an installer and became capable
of producing a curated deployment bundle. That change is central to the
production bundle iteration feature.

### `ae2e445` Add probe-driven offline packaging workflow

This is the capstone commit for the current feature branch. It completes the
probe-first packaging story by separating environment discovery from bundle
assembly, which is a strong fit for real production constraints.

## Guidance For Future Maintainers

When changing this repository, preserve these design goals:

- Keep compatibility logic centralized so BizTalk CU mapping remains consistent
  between direct install and probe workflows.
- Prefer adding or updating tests before modifying CU mapping, package layout,
  or workflow branching.
- Keep `InstallWinSCPForBizTalk.ps1` as orchestration, not as the place where
  new deep behavior accumulates.
- Treat output and logging as operator-facing contracts, because they are part
  of the support experience.

If you need deeper implementation context, use these companion documents:

- `docs/main-script-walkthrough.md`
- `docs/cu-mapping-maintenance.md`
- `docs/production-readiness-release.md`