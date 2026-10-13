# Data Integrity QC Hardening — Phase A

## Independent-review remediation: blank QC severity

Review finding: ordinary `source_blank_count` records were `warning`, causing Batch's unchanged clean-file eligibility rule to reject files with otherwise valid suggested Peaks. The Phase A fix changes only this severity to `info`; source count status, message/detail, sheet/time/class/cell provenance, exports, calculations, and eligibility rules remain unchanged. Blank info stays visible in Single QC, Batch review, `QC_Check`, and `Batch_QC`. Invalid/non-finite counts and interval integrity defects remain errors; existing warning/error eligibility gates are unchanged.

RED/GREEN: before the change, the added blank-only Batch regression failed because the item had nonzero QC warnings and was ineligible. After the change, the focused suite passed (63 passed, 1 pre-existing PHF RuntimeWarning). It confirms blank-only eligibility, QC info severity/provenance, and exported blank visibility. Existing Batch eligibility tests continue to prove warnings and errors disqualify bulk acceptance. Source integrity regressions cover both automatically detected and explicitly confirmed Peaks.

Three-workbook validation after the remediation passed exact baseline frame/export-content parity and 3/3 Batch analyses. Blank finding counts are unchanged; the former 960 / 720 / 576 warnings are now 960 / 720 / 576 info findings (2,256 total). No errors were introduced for those workbooks. Compile and whitespace checks are rerun for this remediation.

The current full-suite invocation was attempted but did not complete in this environment: the pre-existing repeated Streamlit `AppTest` case (`tests/test_app_runtime.py::test_apptest_fresh_loads_and_reruns_preserve_module_classes`) remained active without returning within the extended wait; it was interrupted. A second suite run excluding that one case also remained in long-running UI integration tests and was interrupted before summary. This is not recorded as a passing full-suite result. The prior Phase A 593-test pass predates this severity-only remediation; hosted CI for the new remediation commit is the remaining full-suite gate.

Date: 2026-10-13 (Asia/Bangkok). Scope: source-data diagnostics only.

## Provenance and scope

- Repository: `C:\MyRD\tmc-processor-public`, remote `https://github.com/bokoboss/tmc-processor.git`.
- Initial branch: `codex/issue-43-post-release-docs`; initial HEAD: `55942e2774b45481464806783ee3f425eb060fb6`; working tree clean.
- `git fetch origin` advanced the cached main from the accepted v1.1.0 commit `1514b142e566f1edf64ac1ba5aaf6b5a1523541b` to `29bcfc00d05de1a1316278a939d608d0ee4ea457`.
- Current task branch: `codex/data-integrity-qc-phase-a`, created from fetched `origin/main` at `29bcfc00d05de1a1316278a939d608d0ee4ea457`. Runtime files were identical between the initial checkout and fetched main; their only file difference was `.devcontainer/devcontainer.json`.
- GitHub read-only inspection found no open PRs and only Issue #42 open (the protected V1 Summary discrepancy). No Phase A issue was supplied or found; this work does not change Issue #42.
- At the pre-commit snapshot, HEAD equaled the fetched `origin/main` base and the implementation was an uncommitted working-tree diff. Final commit and PR identifiers are reported in the PR metadata.
- `PROJECT_PROFILE.md`, `AGENTS.md`, the installed workflow reference, execution/evidence templates, model-routing policy, and scrutiny guidance were inspected. Installed workflow version: 1.4.1. No workflow installation or legacy `.ai-workflow/` directory was created.
- Execution used the existing Codex GPT-6 session and its current reasoning settings, with one worker and no model switch. Protected engineering/QC scope calls for strict validation. Luna was considered for bounded implementation; existing context continuity and the source/runtime/export integration favored continuing this session. No stronger-model escalation was needed. Writes stayed inside this repository. Git fetch/branch creation and full-suite execution needed sandbox escalation, which was approved automatically.

Scope permits new source-integrity QC and additive parser provenance. PCU/PCE, Peak selection, PHF, movement aggregation, WorkflowState, mapping semantics, normalized output schema, V1 templates, V2 calculations, and export routing remain unchanged. No interpolation, duplicate removal, or mandatory export blocking was added.

## Pre-implementation gates and scrutiny

Decision: GO WITH CONDITIONS for source diagnostics using the original observations. Acceptance requires RED/GREEN reproduction of all six requested defects, preservation of valid-input results, Single/Batch and QC_Check coverage, full regression, compile and whitespace checks, and representative local-workbook parity. Independent review remains a separate acceptance step; no merge is authorized.

The key failure modes checked before implementation were loss of count state during coercion, loss of source coverage during aggregation, false duplicate detection from classes/mapping fan-out, and accidental changes to export eligibility or protected formulas. Keeping provenance on parsed source frames and QC in the existing QC schema avoids a second readiness system or export contract.

## Root causes and reproductions

1. `importer._numeric_count` returns zero for both blanks and non-numeric input, losing the distinction from measured zero.
2. `pcu.apply_pce_factors` also coerces counts and fills missing values with zero for direct wide/long inputs. Its accepted calculation behavior is preserved; QC now inspects original values.
3. Peak functions group all included sources by start minute. A complete sheet hides a missing slot in another sheet; repeated rows inflate sums; a malformed/off-grid source interval is accepted into those totals.
4. Existing QC inspected aggregate survey bounds and volumes, without per-source interval coverage, duration, duplicate, alignment, or original count-state checks.
5. Excel loading treated literal text `NA`, `N/A`, `NULL`, and `nan` as empty cells. The loader now preserves text markers while retaining actual empty cells as missing values so existing header filling works.

Synthetic fixtures contain two sheets, two vehicle classes, four 15-minute intervals, and deliberate stream aggregation to one movement. Each defect is introduced in only one sheet. All six are asserted for automatic and user-confirmed peaks.

RED command (before production edits):

```text
.venv/Scripts/python.exe -m pytest tests/test_source_integrity.py -q
```

Result: **14 failed, 1 passed**. Twelve failures reproduced the six absent QC checks across both peak paths; two proved missing count-state/provenance and missing presentation/export diagnostics. The valid-input control passed (96 vehicles, 85.328 PCU, both PHF values 1.0). Local log: `outputs/phase_a_red.txt`.

Additional focused RED checks found **4 failures** for literal Excel NA markers (`outputs/phase_a_na_red.txt`) and **1 failure** for repeated copies of the same source cell retaining its original row identity (`outputs/phase_a_repeated_cell_red.txt`). Both were fixed within the count/duplicate-integrity scope.

Baseline full-suite command (production code still loaded from baseline):

```text
.venv/Scripts/python.exe -X utf8 -m pytest --basetemp=outputs/phase_a_baseline_tmp -vv
```

Result: **552 passed, 14 expected RED failures**, 30 warnings, 460.38 seconds. This includes all **551 pre-existing tests passing** plus the new valid-input control. No pre-existing test failed or skipped. Log: `outputs/phase_a_baseline_suite.txt`.

Environment notes: the unqualified `python` resolves to an inaccessible Windows Store alias, so the existing `.venv/Scripts/python.exe` was used. A restricted baseline invocation stalled in the existing Streamlit runtime test and was interrupted; the approved execution-permission rerun completed. The first local snapshot script hit the Windows console encoding on Thai filenames; rerunning with `-X utf8` resolved that logging-only failure.

## Implementation and changed files

| File | Change |
| --- | --- |
| `src/tmc_processor/importer.py` | Preserve source row, count column, original count value, and blank/zero/numeric/invalid state alongside the unchanged parsed count; prevent text NA markers from becoming blanks during workbook loading. |
| `src/tmc_processor/normalizer.py` | Expose source count observations for parsed long and supported wide inputs, reusing existing time/class parsing conventions. Normalization/calculation behavior is unchanged. |
| `src/tmc_processor/qc.py` | Inspect sources before mapping fan-out can affect integrity checks; append findings in the existing QC schema. |
| `tests/test_source_integrity.py` | Defect, provenance, valid-input, no-false-duplicate, Single/Batch display/export and native-template regressions. |
| `docs/development/EVIDENCE_DATA_INTEGRITY_PHASE_A.md` | This review/evidence record. |

New checks:

| Check | Severity | Evidence/action |
| --- | --- | --- |
| `source_missing_interval` | error | Sheet/time lacks a valid quarter-hour present in another included report source or required by its selected peak; verify the source, without imputation. |
| `source_duplicate_interval` | error | Repeated physical source rows at the same interval; inspect/remove unintended duplicates upstream. |
| `source_invalid_interval_duration` | error | Duration differs from 15 minutes; correct source times. |
| `source_misaligned_interval` | error | Start is off the quarter-hour grid; correct source times. |
| `source_blank_count` | info | Original blank, with class and cell when available; verify whether explicit zero was intended. Info remains visible in Single/Batch QC and exports, while preserving Batch bulk-accept eligibility. |
| `source_invalid_count` | error | Non-numeric/non-finite source value, class and cell; correct the source value. |

Duplicate detection uses physical row identity for imported/wide inputs, treating all classes/columns in one row as one interval observation. Legacy long frames without row identity use repeated class records within an interval. Mapping rows are never used as source observations. Report-excluded sheets do not impose survey coverage on included sheets. Peak-required coverage applies only to sources included in peak calculation.

Repeated copies of a physical source cell also produce a duplicate error even when all copies retain the same original row number. Distinct columns for the same class on one physical row remain legitimate observations.

## Validation

| Gate | Command/method | Result |
| --- | --- | --- |
| Remediation focused GREEN | `.venv/Scripts/python.exe -m pytest tests/test_source_integrity.py tests/test_qc_checks.py tests/test_ux4_ux5.py -q` | **63 passed**, 1 existing non-finite PHF RuntimeWarning. Includes blank-only eligibility and info visibility/export regressions. |
| First GREEN | `.venv/Scripts/python.exe -X utf8 -m pytest tests/test_source_integrity.py -q` | 15 passed; `outputs/phase_a_green.txt`. |
| Expanded focused | `.venv/Scripts/python.exe -X utf8 -m pytest tests/test_source_integrity.py tests/test_qc_checks.py -q` | 44 passed, 1 warning, 13.37 seconds; `outputs/phase_a_focused.txt`. |
| Native template QC | `.venv/Scripts/python.exe -X utf8 -m pytest tests/test_source_integrity.py::test_native_template_export_retains_actionable_source_qc -q` | 1 passed, 6.95 seconds; `outputs/phase_a_native_qc.txt`. |
| Final focused before blank-severity remediation | `.venv/Scripts/python.exe -X utf8 -m pytest tests/test_source_integrity.py tests/test_qc_checks.py -q` | 50 passed, 1 warning, 19.20 seconds; `outputs/phase_a_focused_final.txt`. |
| Initial full GREEN | `.venv/Scripts/python.exe -X utf8 -m pytest --basetemp=outputs/phase_a_green_tmp -vv` | 587 passed, 0 failed/skipped, 31 warnings, 503.86 seconds; `outputs/phase_a_full_suite.txt`. This invocation predates the last loader/duplicate changes; the final run qualifies those. |
| Final full suite before blank-severity remediation | `.venv/Scripts/python.exe -X utf8 -m pytest --basetemp=outputs/phase_a_final_tmp -vv` | **593 passed, 0 failed/skipped**, 31 warnings, 519.26 seconds; `outputs/phase_a_final_suite.txt`. Includes all 551 pre-existing tests and 42 new regression cases. |
| Compile | `.venv/Scripts/python.exe -X utf8 -m compileall -q app.py src` | Passed. |
| Whitespace | `git diff --check` | Passed. |
| Real baseline/fixed | `.venv/Scripts/python.exe -X utf8 outputs/phase_a_real_validation.py baseline`, then `green` | Exact frame/export-content parity; 3/3 Batch analysis passed. Repeated after the loader fix with the same result (`outputs/phase_a_real_final.txt`). |

The non-finite input test intentionally exposes the existing Peak `inf/inf` RuntimeWarning while proving an actionable invalid-count error. Its numeric/PHF behavior was not changed. No tests were disabled.

Final local scrutiny: GO for PR creation. User authorization now covers commit, push to the task branch, and PR creation. Hosted CI and independent review remain pending; merge, release and deployment are out of scope. No scope expansion or formula/template edits are needed for this Phase A change.

Single validation exercises both peak paths and confirms error-level QC still exports. Existing display helpers preserve severity, sheet, time, cell and original value in detail. Six Batch cases exercise analysis suggestions, explicit confirmation, per-file exported QC_Check, Batch_QC in the summary workbook, summary severity counts, and successful ZIP generation despite QC. Native OOXML export includes the invalid-value evidence and native chart part.

UI presentation was verified through the actual Single/Batch display helpers and runtime regression tests, not a new manual browser acceptance session. No UI code was changed. No Excel GUI/recalculation acceptance is claimed.

## Protected real-workbook parity

Original local workbooks were read without modification from ignored `samples/raw/`. Qualification-only mapping cycles the existing V1 canonical movements across the union of detected sheet names. This proves software preservation, not physical-direction engineering acceptance. The mapping is identical before/after. Raw workbooks, snapshots, detailed values and generated artifacts remain ignored and are not included in the public diff.

`assert_frame_equal` passed for normalized data, hourly, movement and vehicle summaries, and Peak/PHF frames. A SHA-256 digest of all exported worksheet cell contents excluding the intentionally changed QC_Check passed exactly, with fixed metadata timestamps. Batch analysis vehicle/PCU totals equal the corresponding Single results for all three files.

SHA-256 comparison of the original source workbooks before/after also passed for all three files. The final comparison was repeated after both final runtime changes and still passed (`outputs/phase_a_real_final.txt`).

| Workbook | Sheets | Normalized rows | Vehicles | PCU | New blank info | New interval/invalid errors |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Kabin Buri | 24 | 13,824 | 82,621 | 87,023.550 | 960 | 0 |
| Bo Phloi | 15 | 8,640 | 16,905 | 14,032.315 | 720 | 0 |
| Nong Prue | 12 | 6,912 | 7,908 | 6,811.479 | 576 | 0 |

Both protected templates exactly match their HEAD Git blobs:

- V1 template SHA-256: `1b9b8ddffe1a44799b5ba9012b627a8c0f88750bcf8cc2937674d51d9fc2c66b`.
- V2 template SHA-256: `409965c21f9ea4872af9d3e3ca416592bf31b2f9821b14305a76334e847df5d9`.

`git diff v1.1.0 -- src/tmc_processor/pcu.py src/tmc_processor/peaks.py src/tmc_processor/summaries.py src/tmc_processor/workflow_state.py src/tmc_processor/exporter.py` is empty. Text content equality also passed for these modules (accounting for the checkout's CRLF line endings).

Tracked runtime diff: three files, 146 insertions and 3 deletions. The new regression and evidence files are additional. At review preparation, `git status --short` contains only the three modified runtime files and these two new files; transient test exporter folders are removed by their existing cleanup as tests finish. Ignored `outputs/` and real raw workbooks remain outside the diff. Whitespace checks also passed explicitly for the two untracked files using `git diff --no-index --check -- NUL <file>`.

## Known limitations and remaining approvals

- QC is advisory. The accepted zero fallback and aggregation still operate; invalid counts/duplicates/incomplete coverage can still yield biased exported totals. Operators must fix source workbooks and re-analyze.
- Many legacy workbooks encode empty counts as blanks. Informational findings can be numerous (consolidated per sheet/interval/state, with affected cells in detail); they remain visible in QC but do not disqualify Batch bulk acceptance. Existing warnings and errors still disqualify bulk acceptance; individual confirmation/export remains available.
- Coverage uses the observed union of included report sources plus selected peak intervals. A slot absent from every sheet outside a selected peak cannot be inferred without an authoritative expected survey schedule. Deliberately staggered source coverage needs separate policy if it should be accepted.
- Physical row/cell provenance is exact for workbook imports; direct wide frames report one-based data-row/column positions, not guaranteed original Excel row offsets. Long frames created without provenance cannot recover original cells or distinguish repeated same-class columns from repeated observations.
- Existing time parsing/truncation, skipped unparseable source rows, multi-day handling and blank/invalid coercion were not redesigned. No automatic interpolation or deletion was introduced.
- Excel error cells and formula cells without cached values can be surfaced by pandas as missing values and receive a blank-count info finding. Literal text NA markers are covered; reconstructing Excel error codes/formula-cache state is outside this parser change.
- Hosted CI and independent review are pending. No release, deployment or merge is authorized.
- Commit, push to `codex/data-integrity-qc-phase-a`, and PR creation were explicitly authorized after the local gates passed. No other branch or main history is in scope.
