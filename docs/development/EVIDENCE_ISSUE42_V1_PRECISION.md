# Issue #42 V1 Native OOXML Precision — Evidence

## Scope and baseline

- Repository: `C:\MyRD\tmc-processor-public` (`origin` is `https://github.com/bokoboss/tmc-processor.git`).
- Accepted baseline: `63a10ab9acbc33cd81ed851197da440786c2cae9` (`origin/main`).
- Follow-up branch: `codex/issue42-v1-precision`, PR #48 targeting `main`. The branch started this follow-up at `072a86db95fa381da8298ee42225425f8f81d864`; the source-integrity remediation and updated evidence are submitted together on that PR.
- Issue #42’s full discussion was reviewed; the latest approval accepts the five-cell V1 Native OOXML policy with the documented scope boundaries and the requested regression/Excel gates. Accepted-baseline post-merge CI run `38417068041` had all three jobs successful.
- Scope: V1 Native OOXML final totals and the narrow proof for its existing rounded-hour-total fallback. V2, Safe PNG, legacy COM generation, source/Peak calculations, templates and Phase A/B behavior were not changed.
- PR #48 review remediation adds precise per-hour normalized-source reconciliation and duplicate canonical `Movement_Summary` group rejection. The five final formulas, templates, HLOOKUPs, shared formulas, Peak/PHF calculations and all other export paths remain unchanged.
- Existing `docs/development/ISSUE42_INVESTIGATION.md` and the prior ignored `outputs/issue42/` investigation evidence were preserved. Added test/Excel qualification harnesses and output files remain ignored under `outputs/issue42/`.

## Reproduction and RED/GREEN

Before runtime changes, the real-pipeline Native OOXML regressions failed on the existing behavior. The baseline demo produced `Summary!F32 = 23,956` while `W40 = AM22 = 23,959`, against precise reported PCU `23,959.041`. Day 2 produced `F32 = 23,588` while `W40 = AM22 = 23,590`, against `23,590.054`. Fractional cases also showed half-even tie differences, both understatement and overstatement, and native fallback for multi-hour accumulation.

After the correction, the focused suite passed 11/11, including both demo days; minimal `2 × 0.49`; three-contribution `0.49` and `0.51`; `0.5`, `1.5` and `2.5` ties; multi-hour accumulation; alternate confirmed AM/PM windows; report exclusion; Peak exclusion; and the actual V1 Native OOXML route. Additional variable-row/source-contract and corrupt-cache tests passed 2/2. The variable-row test splits an in-memory movement group and confirms the formula expands to the actual V1 support row count.

The PR #48 review added two integrity regressions before runtime changes. The balanced hourly redistribution test changed two precise `Hourly_Summary` values by `+0.1` and `-0.1`; it verified the grand total and displayed integer values stayed unchanged, then failed because the baseline validator accepted inconsistent per-hour sources. The duplicate-group test split one `Movement_Summary` PCU row into two rows with identical canonical keys; grouped and reported totals stayed unchanged, then failed because baseline reconciliation collapsed the duplicate. Both tests passed after remediation. A third regression confirms a legitimate zero-traffic hour remains valid.

## Implementation and formulas

The five approved final Summary cells now use precise sources and one Excel-compatible `ROUND(...,0)`:

| Cells | Formula policy | Demo day 1 | Demo day 2 |
|---|---|---:|---:|
| `F30` | `=ROUND('Peak_PHF'!$F$2,0)` | 2,981 | 2,872 |
| `F31` | `=ROUND('Peak_PHF'!$J$2,0)` | 3,153 | 3,036 |
| `F32`, `W40`, `AM22` | `=ROUND(SUMIFS('Movement_Summary'!$I$2:$I$N,'Movement_Summary'!$G$2:$G$N,TRUE),0)` | 23,959 each | 23,590 each |

`N` is the actual final support row (`len(Movement_Summary) + 1`), not a V2 fixed range. The formula filters the V1 `include_in_report` flag because V1's actual movement schema retains excluded groups. Formula caches use `Decimal` `ROUND_HALF_UP` to match Excel. For the four-row demo, the PCU range is `I2:I5` and the criteria range is `G2:G5`.

The V1 validator checks the actual ordered support schemas, finite numeric PCU values, normalized-to-movement grouping and every group total, report inclusion, normalized/hourly aggregate agreement, and effective AM/PM Peak source values recomputed from `include_in_peak` rows for the selected 60-minute intervals. Invalid schemas, changed template formulas, bad Peak values or source disagreements reject Native OOXML and retain the existing explicit Safe PNG fallback.

Detailed hourly and movement calculations/display values are unchanged, and no rounding residual is allocated. V1 `W40:AJ40` shared-formula dependents are materialized as their existing formulas/caches before replacing only master `W40`. Native verification allowlists the five new formulas/caches and validates the preserved dependents, HLOOKUP Peak bindings, and original package structures.

## Rounding fallback guard

The V1 exception is limited to the aggregate hourly `Total` comparison. It proves each displayed hour equals the existing nearest-even rounding of its precise `Hourly_Summary` PCU and that the support grand total equals the existing nearest-even rounding of the precise normalized total. Only the mathematically explained difference between the sum of those individually rounded hours and the precise once-rounded report total is accepted. Every precise `Hourly_Summary` total is independently reconciled against `Normalized_Data` regrouped by source `time_start` hour, using `rel_tol=1e-12` and `abs_tol=1e-9` only for floating-point accumulation noise. Missing normalized source hours compare as zero, preserving legitimate zero-traffic intervals; duplicate or malformed summary intervals reject. Duplicate canonical `Movement_Summary` keys reject before grouped source reconciliation, while distinct facility and inclusion-flag groups remain valid. The proof also rejects malformed PCU sources, altered displayed hourly values and inconsistent grand totals. Individual movement checks remain exact; the V2 six-PCU tolerance and V2 fixed formula were not changed or applied to V1.

## Validation evidence

- Focused RED: 11 failures on baseline implementation, including the documented two demo mismatches and controlled tie/fallback failures.
- Focused GREEN: 11 passed; later V1 contract/cache negatives: 2 passed.
- Final-state focused suite: `tests/test_v1_pcu_rounding.py` — **13 passed, 12 warnings in 79.54 seconds**.
- PR #48 integrity RED/GREEN: the two required focused tests failed before runtime changes (`DID NOT RAISE` for each corruption), then passed after the fix in 2.16 seconds.
- Updated focused V1 suite: `tests/test_v1_pcu_rounding.py` — **16 passed, 12 warnings in 81.96 seconds**, including hourly redistribution, duplicate-group rejection and a legitimate zero-traffic hour.
- Compatibility selection (V1 rounding, OOXML template export, V2 PCU rounding, PCE, effective Peaks, template integrity and UX-7 Standard/Batch/Safe routing): 103 passed.
- Updated full automated test suite: `.venv/Scripts/python.exe -m pytest -q --basetemp=outputs/issue42/pr48-full-elevated --tb=short` — **636 passed, 43 warnings in 611.91 seconds**. The warnings are the existing openpyxl DrawingML limitation plus one invalid-infinity Peak warning in `test_source_integrity.py`.
- `python -m compileall -q app.py src`: passed.
- `git diff --check`: passed.
- Protected real workbooks, using the pre-existing deterministic software mappings, retained precise reported totals and now exported Native OOXML successfully: case 1, 87,023.550 → 87,024; case 2, 14,032.315 → 14,032; case 3, 6,811.479 → 6,811. Each had matching `F32/W40/AM22`. The previous native fallback is resolved. Source inputs and template hashes were checked unchanged. These are software parity fixtures, not approved engineering-direction mappings.
- Repeated after the PR #48 integrity remediation: all three protected workbooks again passed Native OOXML with those precise totals and matching `F32/W40/AM22`; protected source and template hashes remained unchanged.
- Native OOXML verification and template regression tests check formula/caches, source sheets, all existing HLOOKUP mappings, W40 shared dependents, drawings, chart definitions/caches and package relationships. The original template binaries were not modified. Excel found 33 Summary shapes and two charts in each of the 12 opened outputs.
- Negative coverage: malformed/mismatched Movement_Summary and Peak sources, altered original formula, and deliberately corrupted final cache are rejected.

## Microsoft Excel and remaining gates

Microsoft Excel COM was repeated after the PR #48 integrity change. It opened 12 freshly generated V1 outputs read-only and ran full rebuild/recalculation on Excel 16.0 build 20430: both demo days, fractional over/understatement, `0.5`/`1.5`/`2.5` ties, multi-hour accumulation, alternate confirmed Peaks with Peak exclusion, and the three protected outputs. All five values matched their pre-open OOXML caches after recalculation; the formula-error scan found zero errors; each workbook remained read-only; all hashes were unchanged; and every workbook reported 33 shapes and two charts. Excel COM `.Text` returned `###` for demo day 1 `F32` and day 2 `F30:F32`. The preserved baseline COM evidence reports `###` for those same cells, so this is an existing template display-width condition, not introduced by the precision formulas. COM text results do not constitute visual GUI acceptance.

The Codex Computer Use inventory exposed no desktop apps (only browser surfaces), so GUI inspection of repair prompts, visible cell layout, print area and chart rendering was unavailable. GUI qualification remains **PENDING**, including confirmation of how the inherited `###` cells render in the actual window. Before acceptance, inspect fresh corrected V1 outputs and an unchanged V2 control read-only; check for repair dialogs, verify all five totals are visible and correct without unexpected `###`, force full recalculation, inspect 33 shapes/two charts and their labels/positions, check effective Peak labels and readable Print Preview/PDF, then close without saving and confirm hashes unchanged. Record Excel build, workbook identities, results and non-sensitive screenshots. Excel/GUI acceptance on other builds is also unqualified.

## Changed files and final state

Changed runtime files: `src/tmc_processor/exporter.py`, `src/tmc_processor/ooxml_template_export.py`, `src/tmc_processor/template_write_plan.py`, `src/tmc_processor/template_write_verify.py`. Added regression tests: `tests/test_v1_pcu_rounding.py`. Added this evidence document. No original template, Safe PNG, legacy COM generator, V2 or calculation-core file was changed.

The pre-existing untracked investigation document remains untouched. The accepted base branch was not modified. Final PR head and hosted CI run/status are provided in the PR #48 handoff. Local `git diff --check` and `compileall` passed. GUI qualification remains pending as detailed above.
