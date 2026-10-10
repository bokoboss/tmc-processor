# Issue #34 implementation evidence

Date: 2026-10-10 (Asia/Bangkok). Issue: https://github.com/bokoboss/tmc-processor/issues/34.

## Scope and baseline

- Started from verified `origin/main@00b2dfcac89038fd92698a114b23ba207c495f4e` (PR #33 accepted baseline).
- Initial checkout was clean on `codex/issue-32-v2-xlsx-summary@af2469e6f51a95686dac6fb6442be0d63ccb47ae`. Fetched main and created `codex/issue-34-v2-pcu-round-once` at the requested SHA; no user changes were discarded.
- Read PROJECT_PROFILE.md, actual exporter/write-plan/verifier/tests, Issue #34, and installed workflow reference/manifest. The installed governance version is 1.4.1, revision `3547ae260feacf8fc9a102b2abfdb13881e36dab`; no workflow upgrade or legacy `.ai-workflow/` recreation was performed.
- Routing: protected export/calculation presentation change, bounded implementation in the existing chat. Considered the installed policy's Luna-first recommendation; retained the current executor for protected OOXML/shared-formula investigation and adjudication. No additional agents or model changes. Escalation trigger: unexplained source disagreement, template drift, or failed fidelity/calculation gates.
- Gates defined before implementation: precise grand-total/confirmed-Peak formulas and caches; raw audit precision; five-cell override allow-list; unrelated formula/layout/style preservation; DrawingML/Chart fidelity; V1/Single/Batch/package compatibility; focused tests, compileall, whitespace; real Excel recalculation; branch-only PR for independent review.

## Implementation

Only these five V2 Summary formulas change their calculation:

| Cells | Exported formula | Source |
|---|---|---|
| F32, W40, AM22 | `ROUND(SUM('Movement_Summary'!$G$2:$G$17),0)` | Precise 16 canonical movement PCU totals |
| F30 | `ROUND('Peak_PHF'!$F$2,0)` | Precise effective/confirmed AM PCU |
| F31 | `ROUND('Peak_PHF'!$J$2,0)` | Precise effective/confirmed PM PCU |

The write plan validates the original authored rollup formulas and support-column/16-code contracts. Grand totals must reconcile with precise Normalized_Data and Hourly_Totals. Cached presentation values use Decimal ROUND_HALF_UP, matching Excel's half-away-from-zero policy. Recalculation flags remain automatic/full/forced and stale calcChain removal remains unchanged.

W40 is the master of the authored shared formula group W40:AJ40 (si=8). Before replacing W40, X40:AJ40 are materialized with their original expanded SUM formulas and existing caches; their calculation, style and value are unchanged. The verifier validates the exact original shared stubs and expanded formulas, plus the five final formulas/caches. All other protected formulas retain their original XML; all movement HLOOKUP bindings retain the established effective-Peak contract.

Export_Metadata documents that independently rounded movement/hourly subtotals may not add to the independently rounded grand total. No residual allocation, PCU/PCE calculation, normalization, mapping, QC, Peak selection, WorkflowState, UI, routing, ZIP/PNG or template redesign is introduced. Production uses the existing direct OOXML path without Excel COM or openpyxl saving.

Changed paths:

- `src/tmc_processor/template_write_plan.py`: validated V2 final-total plan and Excel-compatible integer caches.
- `src/tmc_processor/ooxml_template_export.py`: bounded formula/cache overrides and shared-formula materialization.
- `src/tmc_processor/template_write_verify.py`: allow-list, original/expanded formula and cache verification.
- `src/tmc_processor/exporter.py`: rounding traceability metadata.
- `tests/test_v2_pcu_rounding.py`: demo, eight fraction/tie fixtures, five invalid-source cases and three corrupted-output cases.
- This evidence document.

## Reproduction and focused validation

The initial six regression cases failed before implementation: demo F32 cached 23,956 instead of 23,959; fractional fixtures also reproduced cumulative rounding errors. An intermediate verifier failure exposed W40's shared-formula dependencies and was fixed by preserving their expanded formulas. The demo fixture was corrected to the specified confirmed periods, AM 08:00-09:00 and PM 17:00-18:00.

Final commands (repository .venv Python 3.12):

```powershell
.venv/Scripts/python.exe -m pytest tests/test_v2_pcu_rounding.py -q --tb=short --basetemp=outputs/issue34/rounding-pytest
# 17 passed, 9 warnings, 85.05 seconds

.venv/Scripts/python.exe -m pytest tests/test_phase_f_v2_dry_run.py tests/test_ooxml_template_export.py tests/test_template_integrity.py tests/test_effective_peak_export.py tests/test_export_package.py tests/test_batch.py -q --tb=short --basetemp=outputs/issue34/compat-pytest
# 119 passed, 17 warnings, 255.26 seconds

.venv/Scripts/python.exe -m compileall -q app.py src
# Exit 0
git diff --check
# Exit 0
```

No local full-suite run. Initial focused runs using pytest's default sandbox temp directory had fixture setup errors (`could not create numbered dir ... after 10 tries`); the final runs above resolved that environment issue by placing temporary files under the workspace. The warnings are openpyxl's known incomplete DrawingML read support; generated output is written directly through OOXML, and protected drawings are byte checked.

Fraction fixtures cover 0.49, 0.51, 0.5, 2.5 and -0.5 across movements/hours, plus isolated +0.5/+2.5/-0.5 grand-total ties. Assertions cover all five final caches/formulas, precise support totals, confirmed Peaks, recalc flags and protected drawing/media bytes. Negative cases reject altered original formulas, movement order, normalized/hourly disagreements, nonfinite Peak values, stale final caches/formulas and changed vehicle formulas. Existing focused tests cover V1, V2 Single/Batch template export, per-file ZIP PNG, support-sheet contracts and COM-independent routing.

## Demo and package fidelity

| Measure | Baseline | Fixed |
|---|---:|---:|
| Precise total PCU | 23,959.041 | 23,959.041 |
| F32 | 23,956 | 23,959 |
| W40 / AM22 | 23,959 / 23,959 | 23,959 / 23,959 |
| F30 AM (precise 2,980.827) | 2,981 | 2,981 |
| F31 PM (precise 3,153.213) | 3,153 | 3,153 |

An isolated source snapshot under ignored outputs/issue34/baseline_src used `git show <baseline>:<changed source path>` and the current immutable template to generate comparison exports, with identical setup and generation timestamp. Baseline and fixed exports were compared for the demo and all three local protected workbooks. In each case:

- Identical ZIP part names/order; only Summary (`sheet1.xml`) and Export_Metadata (`sheet2.xml`) differ.
- Every detailed/support/audit sheet is byte-identical, including Normalized_Data, Movement_Summary, Hourly_Totals, Peak_PHF and Movement_Source_Stream_Audit.
- All DrawingML, Chart, chart style/color/relationship and media parts are byte-identical to the accepted baseline export. The existing data-dependent native chart-cache refresh is unchanged; chart1/chart2 were already refreshed by baseline export, so their bytes are not expected to equal the blank template's caches.
- Summary XML outside sheetData is identical; 30 merges, 33 native shapes and two chart frames remain.
- DrawingML and its relationships also match the source template byte-for-byte.

Demo output hashes:

```text
xl/drawings/drawing1.xml ce0adabc8de01d6789152b16956bab3a5548e181d9470c8ca37bcef112d49077
xl/charts/chart1.xml     0f3a32407a9bd90496ccf8c93d84b9876cf924c556abe73639bce5ac7c95eaed
xl/charts/chart2.xml     c069e7f35da732f42d55071ac79dca6df3ba7b124d3df1dedfd7d34303938301
```

The authoritative `templates/four_leg_tmc_report_template_approach_v2.xlsx` exactly matches its baseline Git blob. SHA-256: `409965c21f9ea4872af9d3e3ca416592bf31b2f9821b14305a76334e847df5d9`. Neither template nor template map was edited or saved.

## Protected local workbook regression

Kabin Buri, Bo Phloi and Nong Prue were read only from ignored samples/raw/. Temporary deterministic qualification mappings cycled the 16 canonical V2 codes across detected source-sheet names and confirmed fixed-hour AM/PM periods. These checks qualify software/export preservation, not physical-direction engineering acceptance. Raw workbooks, generated reports and source snapshots stay ignored and are not committed.

| Workbook | Normalized rows | Precise PCU | Fixed F32 / W40 / AM22 |
|---|---:|---:|---|
| Kabin Buri | 13,824 | 87,023.550 | 87,024 / 87,024 / 87,024 |
| Bo Phloi | 8,640 | 14,032.315 | 14,032 / 14,032 / 14,032 |
| Nong Prue | 6,912 | 6,811.479 | 6,811 / 6,811 / 6,811 |

Local supplementary commands: `.venv/Scripts/python.exe outputs/issue34/qualify.py --baseline` then `.venv/Scripts/python.exe outputs/issue34/qualify.py`. The script uses identical baseline/fixed inputs and verifies package bytes/layout as described above. Detailed local results: outputs/issue34/qualification.json and qualification.log; focused logs: outputs/issue34-rounding-tests-fixed.log and outputs/issue34-compat-tests.log.

## Remaining gates and limitations

- Real Excel recalculation remains blocked. Following explicit approval for ONE bounded, non-elevated check, the checker was inspected and hardened: fixed list of 12 generated qualification XLSX files, separate hidden DispatchEx instance, ReadOnly=True, UpdateLinks=0, AskToUpdateLinks=False, EnableEvents=False, AutomationSecurity=3, CorruptLoad=0, repair-mode/formula/cache/shape/chart checks, Close(SaveChanges=False), and cleanup restricted to its own instance. A 180-second watchdog targets only that instance's PID. Input hashes guard against modification; evidence JSON uses exclusive creation, not overwrite. No deletion, network access or global Excel process termination is present. Production code was not changed.
- The single approved command, `.venv/Scripts/python.exe outputs/issue34/excel_recalc.py`, ran without elevation and exited 1 at line 14 (`DispatchEx('Excel.Application')`), before any workbook opened: `pywintypes.com_error: (-2147024891, 'Access is denied.', None, None)`. No Excel instance handle/PID was obtained, no calculation or workbook save occurred, and no retry/elevation was attempted. Earlier Excel 16.0 launch evidence required elevation and does not qualify this approved check. No real Excel recalculation/viewer/no-repair success is claimed.

| Requested real Excel verification | Expected | Observed in approved run |
|---|---|---|
| Demo F32 / W40 / AM22 | 23,959 / 23,959 / 23,959 PCU | Not observed; Excel instance creation denied |
| Demo F30 / F31 | 2,981 / 3,153 PCU | Not observed; no workbook opened |
| Vehicle-class formulas and caches | Original formulas and unchanged values | Not observed |
| Native shapes / charts; repair status | 35 shape objects / 2 charts; no repair warning | Not observed |
| Eight fractional edge cases and three qualified local workbooks | Recalculated totals equal recorded expectations | Not observed |

The existing 136 focused passes and package/cache comparisons remain the available evidence; they do not substitute for the blocked real Excel gate.

## User acceptance and independent workbook review

- User Acceptance: the user opened the generated `demo.xlsx` in Microsoft Excel and confirmed it works. This records user-attested usability; no screenshot or direct observation of Excel's recalculated in-memory state was supplied.
- Independent Workbook Review: the actual uploaded `demo.xlsx` was inspected as raw ZIP/OOXML without opening or saving it through an Excel library. ZIP CRC passed; the workbook contains 24 worksheets, with Summary first. The five formulas and saved cached values were correct: F30=2,981; F31=3,153; F32/W40/AM22=23,959. Precise source totals reconcile to 23,959.041 PCU; AM/PM Peak sources are 2,980.827 / 3,153.213 PCU. Vehicle-class formulas/caches are intact. The drawing and chart structures remain present (33 shape objects, two connectors, two graph frames, and two charts); DrawingML and chart hashes match the existing evidence above.
- The independent review verifies saved formulas, caches, workbook structure, and the user's reported Excel usability. It does not directly prove that Excel performed a new calculation in memory after opening. Automated Excel COM recalculation remains unverified because `DispatchEx('Excel.Application')` failed with Access Denied before opening any workbook, as recorded above. No COM retry was performed.

- The prior GitHub Actions run #38018859554 passed all three Windows/Linux jobs at the accepted implementation HEAD. The documentation-only follow-up does not rerun the local full suite; GitHub Actions for the pushed commit is the regression gate before merge. PR #35 is Ready for Review. No merge, tag, release or issue closure is performed.
- Cloud compatibility is covered by the unchanged direct OOXML path and focused COM-independent tests; no live Cloud redeployment/browser acceptance was performed.
- Legacy diagnostic V2 COM export is outside this direct OOXML fix. V1 compatibility is covered by the existing focused tests.
