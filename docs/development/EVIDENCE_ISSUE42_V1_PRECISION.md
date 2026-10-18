# Issue #42 V1 Native OOXML Precision — Evidence

## Scope and baseline

- Repository: `C:\MyRD\tmc-processor-public` (`origin` is `https://github.com/bokoboss/tmc-processor.git`).
- Accepted baseline and final `HEAD`: `63a10ab9acbc33cd81ed851197da440786c2cae9` (`origin/main`).
- Local branch: `codex/issue42-v1-precision`; implementation remains uncommitted and local.
- Issue #42’s full discussion was reviewed; the latest approval accepts the five-cell V1 Native OOXML policy with the documented scope boundaries and the requested regression/Excel gates. Accepted-baseline post-merge CI run `38417068041` had all three jobs successful.
- Scope: V1 Native OOXML final totals and the narrow proof for its existing rounded-hour-total fallback. V2, Safe PNG, legacy COM generation, source/Peak calculations, templates and Phase A/B behavior were not changed.
- Existing `docs/development/ISSUE42_INVESTIGATION.md` and the prior ignored `outputs/issue42/` investigation evidence were preserved. Added test/Excel qualification harnesses and output files remain ignored under `outputs/issue42/`.

## Reproduction and RED/GREEN

Before runtime changes, the real-pipeline Native OOXML regressions failed on the existing behavior. The baseline demo produced `Summary!F32 = 23,956` while `W40 = AM22 = 23,959`, against precise reported PCU `23,959.041`. Day 2 produced `F32 = 23,588` while `W40 = AM22 = 23,590`, against `23,590.054`. Fractional cases also showed half-even tie differences, both understatement and overstatement, and native fallback for multi-hour accumulation.

After the correction, the focused suite passed 11/11, including both demo days; minimal `2 × 0.49`; three-contribution `0.49` and `0.51`; `0.5`, `1.5` and `2.5` ties; multi-hour accumulation; alternate confirmed AM/PM windows; report exclusion; Peak exclusion; and the actual V1 Native OOXML route. Additional variable-row/source-contract and corrupt-cache tests passed 2/2. The variable-row test splits an in-memory movement group and confirms the formula expands to the actual V1 support row count.

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

The V1 exception is limited to the aggregate hourly `Total` comparison. It proves each displayed hour equals the existing nearest-even rounding of its precise `Hourly_Summary` PCU and that the support grand total equals the existing nearest-even rounding of the precise normalized total. Only the mathematically explained difference between the sum of those individually rounded hours and the precise once-rounded report total is accepted. The proof rejects duplicate/missing/out-of-template hours, malformed PCU sources, altered displayed hourly values and inconsistent grand totals. Individual movement checks remain exact; the V2 six-PCU tolerance and V2 fixed formula were not changed or applied to V1.

## Validation evidence

- Focused RED: 11 failures on baseline implementation, including the documented two demo mismatches and controlled tie/fallback failures.
- Focused GREEN: 11 passed; later V1 contract/cache negatives: 2 passed.
- Final-state focused suite: `tests/test_v1_pcu_rounding.py` — **13 passed, 12 warnings in 79.54 seconds**.
- Compatibility selection (V1 rounding, OOXML template export, V2 PCU rounding, PCE, effective Peaks, template integrity and UX-7 Standard/Batch/Safe routing): 103 passed.
- Final-state full automated test suite: `.venv/Scripts/python.exe -m pytest -q --basetemp=outputs/issue42/pytest-final-full-elevated --tb=short` — **633 passed, 43 warnings in 640.70 seconds**. The warnings are the existing openpyxl DrawingML limitation plus one expected invalid-infinity Peak warning in `test_source_integrity.py`.
- `python -m compileall -q app.py src`: passed.
- `git diff --check`: passed.
- Protected real workbooks, using the pre-existing deterministic software mappings, retained precise reported totals and now exported Native OOXML successfully: case 1, 87,023.550 → 87,024; case 2, 14,032.315 → 14,032; case 3, 6,811.479 → 6,811. Each had matching `F32/W40/AM22`. The previous native fallback is resolved. Source inputs and template hashes were checked unchanged. These are software parity fixtures, not approved engineering-direction mappings.
- Native OOXML verification and template regression tests check formula/caches, source sheets, all existing HLOOKUP mappings, W40 shared dependents, drawings, chart definitions/caches and package relationships. The original template binaries were not modified. Excel found 33 Summary shapes and two charts in each of the 12 opened outputs.
- Negative coverage: malformed/mismatched Movement_Summary and Peak sources, altered original formula, and deliberately corrupted final cache are rejected.

## Microsoft Excel and remaining gates

Microsoft Excel COM opened 12 corrected outputs read-only and ran full rebuild/recalculation on Excel 16.0 build 20430: both demo days, fractional over/understatement, `0.5`/`1.5`/`2.5` ties, multi-hour accumulation, alternate confirmed Peaks with Peak exclusion, and the three protected outputs. All five values matched their pre-open OOXML caches after recalculation; the formula-error scan found zero errors; each workbook remained read-only; all hashes were unchanged; and every workbook reported 33 shapes and two charts. Excel COM `.Text` returned `###` for demo day 1 `F32` and day 2 `F30:F32`. The preserved baseline COM evidence reports `###` for those same cells, so this is an existing template display-width condition, not introduced by the precision formulas. COM text results do not constitute visual GUI acceptance.

The Codex Computer Use inventory exposed no desktop apps, so GUI inspection of repair prompts, visible cell layout, print area and chart rendering was unavailable. These GUI gates remain explicitly incomplete, including confirmation of how the inherited `###` cells render in the actual window. This implementation was not pushed; no post-change hosted CI run is claimed. Excel/GUI acceptance on other builds is also unqualified.

## Changed files and final state

Changed runtime files: `src/tmc_processor/exporter.py`, `src/tmc_processor/ooxml_template_export.py`, `src/tmc_processor/template_write_plan.py`, `src/tmc_processor/template_write_verify.py`. Added regression tests: `tests/test_v1_pcu_rounding.py`. Added this evidence document. No original template, Safe PNG, legacy COM generator, V2 or calculation-core file was changed.

Final `HEAD` remains the accepted baseline above. The pre-existing untracked investigation document remains untouched. Final `git diff --check` passed. The working tree contains only the four listed runtime files, the new regression test and this evidence document, plus the pre-existing untracked investigation document. No commit, push or hosted CI run was performed.
