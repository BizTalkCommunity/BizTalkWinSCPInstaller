# CU Mapping Maintenance Guide

This document describes the core maintenance workflow for keeping BizTalk cumulative update (CU) mappings current in this repository.

## Why This Is A Core Feature

A primary feature of this project is selecting the correct WinSCP version based on installed BizTalk update level.

The mapping logic lives in:
- ../src/InstallWinSCPForBizTalk.Core.psm1
- Function: Get-WinSCPVersionForBizTalk

When Microsoft publishes new BizTalk CUs (for example, CU7 and later), update the mapping so installations continue to select the right WinSCP version.

## Source Of Truth For CU Data

Use Microsoft documentation for BizTalk updates as the reference source:
- https://learn.microsoft.com/en-us/previous-versions/troubleshoot/biztalk/setup-config/biztalk-hotfixes-cumulative-update

For each new CU, collect:
- BizTalk major version (2016 or 2020)
- CU number/label
- KB number(s)
- BizTalk build
- Release date
- Required WinSCP version for this project

## Files To Update

When adding a new CU mapping, update all applicable locations:

1. Mapping logic
- ../src/InstallWinSCPForBizTalk.Core.psm1

2. Unit tests
- ../tests/Unit/CuDetection.Tests.ps1
- ../tests/Unit/InstallerScript.Execution.Tests.ps1 (if scenario coverage needs extension)

3. User-facing matrix
- ../README.md (Compatibility Matrix section)

4. Optional inline comments
- Keep in-function CU comments in ../src/InstallWinSCPForBizTalk.Core.psm1 aligned with code changes.

## How To Add A New BizTalk 2020 CU

In Get-WinSCPVersionForBizTalk, BizTalk 2020 uses two complementary mechanisms:
- DisplayName parsing via Get-BTSCumulativeUpdateByDisplayName
- KB fallback via Search-BTSCumulativeUpdate

Update steps:

1. Add CU entry to the bts2020CUMap hashtable
- Example shape:
  - 7 = @{ KBs = @('NEWKB1', 'NEWKB2'); WinSCP = 'X.Y.Z' }

2. Update CU comment block above the map
- Add the new CU row with build, KBs, release date, WinSCP.

3. Ensure fallback loop checks newest-to-oldest
- Update the ordered CU list in the foreach loop from newest down to oldest.
- Example after adding CU7:
  - foreach ($cuNumber in @(7, 6, 5, 4, 3, 2, 1))

4. Keep defaults unchanged unless required
- Default for BizTalk 2020 should remain RTM/no-CU behavior unless intentionally changed.

## How To Add A New BizTalk 2016 CU/FP/FU Label

BizTalk 2016 currently uses an ordered list of label/KB/WinSCP entries and matches via KB search.

Update steps:

1. Add new entry near the top of bts2016UpdateMap
- Keep newest/highest-priority entries first.

2. Preserve matching order
- The first matched KB wins.
- Ordering matters when multiple labels exist.

3. Update comment matrix above the map
- Keep labels, KBs, and WinSCP values aligned.

## Testing Checklist After Mapping Changes

Run the unit suite:

```powershell
Invoke-Pester -Path .\tests\Unit
```

Minimum checks:
- CuDetection.Tests.ps1 still passes.
- Existing detection tests for older CUs remain green.
- Script execution tests still pass.

Recommended additions when introducing a new CU:
- Add or update tests in CuDetection.Tests.ps1 for:
  - DisplayName format: Cumulative Update N
  - DisplayName shorthand: CUX (for example CU7)
  - Alternate KB IDs if Microsoft lists more than one KB

## Example: Adding BizTalk 2020 CU7

1. Add map entry in bts2020CUMap:
- 7 = @{ KBs = @('KB_A', 'KB_B'); WinSCP = 'NEW_VERSION' }

2. Update fallback CU order:
- @(7, 6, 5, 4, 3, 2, 1)

3. Add/adjust tests in CuDetection.Tests.ps1:
- Mock DisplayName containing CU7 + KB_A
- Mock DisplayName containing CU7 + KB_B
- Assert WinSCP mapping resolves to NEW_VERSION

4. Update README compatibility table with CU7 row.

## Safety Notes

- Do not remove existing CU entries unless they are incorrect.
- Keep mapping backward-compatible; older environments still depend on historical CUs.
- If a CU has multiple KB IDs, include all known KBs in the CU map entry.
- Keep behavior changes isolated to mapping data whenever possible.
