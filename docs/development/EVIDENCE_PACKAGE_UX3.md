# UX-3 Evidence Package — Explicit Peak Review Confirmation

## Classification

READY FOR REVIEW

This package records the UX-3 implementation and qualification evidence. It does not authorize merge or release. The branch remains unmerged.

## Repository and Git state

- Repository: 'C:\MyRD\tmc-processor-public'
- Authoritative remote: 'https://github.com/bokoboss/tmc-processor'
- Branch: 'codex/ux-3-explicit-peak-review'
- Baseline 'main' / 'origin/main': 'aabc6a65c478a8bc6e86e247f3ab30272d2d0849'
- Implementation commit: 'e9aa8dcf1b15da24bba3b6ffc5f6e6bb89f73327'
- Remediation commit: '171988230b00776247fd499ef6e8e884d7aa887b'
- Final source validation HEAD for this record: '171988230b00776247fd499ef6e8e884d7aa887b'
- Previous PR HEAD before this focused remediation: 'ce31d3d1fbd086bea9d36462c137e4d62176eedf'
- Working tree: clean after the committed evidence updates
- Qualified runtime: 'C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe' — Python 3.12.10
- Python 3.14 was not used.

## Commands executed

Baseline and implementation validation used the qualified interpreter only:

~~~text
.venv\Scripts\python.exe -m pytest
.venv\Scripts\python.exe -m compileall -q app.py src
git diff --check
~~~

Focused regression command:

~~~text
.venv\Scripts\python.exe -m pytest tests/test_explicit_peak_review.py tests/test_batch.py tests/test_effective_peak_export.py tests/test_workflow_state.py tests/test_workflow_contract.py tests/test_workflow_contract_adapter.py tests/test_setup_persistence.py
~~~

Additional qualification commands included direct protected-workbook Batch analysis/export through the Python API, 'probe_excel_com()', and real-browser Streamlit UAT on ports 8508 and 8509. Temporary browser snapshots, servers, logs, and deterministic mapping fixtures were removed after UAT.

## Changed files and diff summary

Commit 'e9aa8dc' changed seven files: 'app.py' (+345/-107), 'src/tmc_processor/batch.py' (+20), the existing Batch/effective-export/setup/workflow tests, and new 'tests/test_explicit_peak_review.py' (216 lines). Remediation commit '1719882' changed only 'src/tmc_processor/peaks.py', 'src/tmc_processor/exporter.py', 'src/tmc_processor/batch.py', and 'tests/test_batch.py'. The implementation keeps the existing peak mathematics, mapping model, QC model, export surface, Project Session schema, and WorkflowState authority.

## Implementation summary

Single and Batch now separate:

1. Suggested — calculated by Analyze and displayed as read-only evidence.
2. Draft — the current Review widget selection; it may default to Suggested but has no confirmation meaning.
3. Confirmed — written only by the explicit 'Confirm Peak Review' action and marked with user-confirmed provenance.

Single confirmation stores AM/PM values and 'peak_selection_source=user_confirmed'. Draft edits do not change confirmed values, effective export values, review fingerprints, or export state. A changed explicit confirmation keeps Analysis current, keeps Review ready, and invalidates the existing export. Reconfirming the same values produces no semantic revision churn.

Batch 'analyze_batch_files()' now leaves successful items at blank confirmed AM/PM values. The Review screen stores per-file drafts and requires an explicit per-file confirmation; no bulk confirm, exception-first queue, or QC-driven automatic acceptance was added. The legacy one-shot 'process_batch_files()' path materializes suggestions only to preserve its pre-review immediate-export API contract; the interactive analysis path remains unconfirmed.

## Reviewer remediation — Batch provenance

The reviewer identified that the compatibility path 'process_batch_files()' materialized Suggested Peaks and then routed them through the reviewed export path with 'peak_selection_source=user_confirmed_batch'. That could falsely claim a human confirmation.

The focused remediation preserves the immediate-export API while passing the existing automatic provenance convention through the one-shot path:

- Explicit interactive Batch confirmation and the default reviewed export use 'user_confirmed_batch' (named 'PEAK_SELECTION_USER_CONFIRMED_BATCH').
- Legacy automatic one-shot export uses 'auto_suggested' (the existing 'PEAK_SELECTION_AUTO' value).
- Workbook 'Export_Metadata', Project Session 'peaks', and per-file export summary text now retain the correct source for both paths.
- The shared exporter recognizes the established Batch explicit value so explicit reviewed Peak periods and provenance are preserved.

No Peak mathematics, WorkflowState architecture, Suggested -> Draft -> Confirmed behavior, UX-4, or UX-5 behavior was changed.

Source, Mapping, PCE, and peak-window semantic changes continue to use the existing WorkflowState invalidation seams. PCE invalidation also clears single confirmed Peak state. Batch review state is cleared on source/mapping/analysis invalidation and remapped with output-stem metadata changes.

## Regression evidence

### Single state tests

'tests/test_explicit_peak_review.py' covers:

- fresh analysis produces suggestions without confirmation or Review readiness;
- draft defaults can show suggestions without writing confirmed state;
- explicit confirmation stores AM/PM and provenance while preserving Analysis;
- unconfirmed draft edits preserve confirmed/effective Peaks and an existing export;
- changed confirmation stales export without re-analysis;
- same-decision reconfirmation has no semantic churn;
- compatible Project Session confirmed Peaks remain confirmed;
- mapping change invalidates the explicit confirmation.

### Batch state tests

The dedicated tests and existing 'tests/test_batch.py' cover blank confirmed values after analysis, 0/N confirmation readiness, draft-only edits, per-file confirmation, the confirmed-value export gate, one-shot compatibility, and Batch ZIP summary/export behavior.

### Project Session compatibility

The existing session implementation already persists 'am_peak_start', 'am_peak_end', 'pm_peak_start', 'pm_peak_end', and 'peak_selection_source'. The new state helper overlays explicit session values without destructive migration; the dedicated reload test verifies a compatible confirmed session remains confirmed.

### WorkflowState invalidation

Existing workflow contract and adapter tests remain green. Explicit UX-3 tests verify that draft-only changes do not change the review decision, first confirmation preserves Analysis and enables Review, changed confirmation invalidates Export only, and Mapping changes invalidate Analysis/Review/Export. PCE and peak-window invalidation behavior remains covered by the existing workflow/state regression suite.

### Effective-Peak/export propagation

'tests/test_effective_peak_export.py' remains green. The effective confirmed Peak path continues to feed Export Metadata, Peak/PHF, Diagram Data, summary text, and session/export provenance. The implementation did not modify 'resolve_effective_peak_periods()' or peak calculation methodology.

## Test and static results

- Baseline: 334 passed.
- Focused UX-3/remediation command: **131 passed, 2 existing openpyxl drawing warnings**.
- Full suite: **344 passed, 11 warnings** in 2:53.76 on Python 3.12.10.
- 'compileall -q app.py src': passed.
- 'git diff --check': passed. Git emitted only the repository’s existing LF/CRLF normalization warnings for tracked Python files.

The warnings are existing 'openpyxl.reader.drawings' warnings about incomplete DrawingML support; no test failed.

## Protected real-workbook qualification

The three protected workbooks under 'samples/raw' were analyzed with a temporary deterministic 24-row 'from_to' mapping generated from detected sheet names. This fixture is qualification-only and explicitly states that it is not physical engineering approval. The physical-direction Mapping Presets requested for engineering validation were not available, so these results prove runtime/software behavior only.

| Workbook | Status | Matched sheets | Vehicles | PCU | Suggested AM / PM | QC error / warning / info |
|---|---:|---:|---:|---:|---|---:|
| '03.TMC1 แยกกบินทร์บุรี_พุธ 17 กค67f.xlsx' | success | 24, missing 0 | 82,621 | 87,023.550 | 07:00-08:00 / 17:00-18:00 | 0 / 0 / 4 |
| '03.TMC1 แยกบ่อพลอย_พุธ 28 พค68.xlsx' | success | 15, missing 0 | 16,905 | 14,032.315 | 07:30-08:30 / 15:45-16:45 | 0 / 0 / 4 |
| '03.TMC4 แยกหนองปรือ_พุธ 28 พค68.xlsx' | success | 12, missing 0 | 7,908 | 6,811.479 | 10:45-11:45 / 15:45-16:45 | 0 / 0 / 4 |

## Windows Streamlit UAT — Single

The first protected workbook was uploaded through Data, matched against the temporary deterministic mapping, and analyzed in the real Streamlit UI.

- Analyze showed AM/PM suggestions, a 0-confirmed Review state, hourly PCU chart, QC summary/details, movement evidence, PHF/movement review sections, and the 'Confirm Peak Review' action.
- Before confirmation, the confirmed AM/PM metrics were blank and Export was not ready.
- After explicit confirmation, the UI showed user-confirmed AM '07:00-08:00' and PM '17:00-18:00'; Analysis remained current and Safe PNG export was enabled.
- Safe PNG report generation completed successfully; the UI exposed the report and package downloads.
- After changing only the AM draft to '08:00-09:00', the confirmed/effective AM remained '07:00-08:00' and the UI stated the draft was not applied until confirmation.
- After confirming the changed draft, Analysis remained current, Export showed the new effective AM, and the previous output was no longer present/stale.

## Windows Streamlit UAT — Batch

All three protected workbooks were uploaded and processed with one shared temporary deterministic mapping.

- Fresh Batch Review showed '0/3' confirmed with suggestions and draft selectors present.
- The three files were confirmed individually; the UI progressed through '1/3', '2/3', and '3/3' and only then reported Batch ready for export.
- Safe PNG 'Generate Batch ZIP' completed successfully.
- After changing the third file’s AM draft from '11:00-12:00' to '12:00-13:00' and explicitly reconfirming, the prior Batch output was marked stale/no longer available without re-analysis.

## Safe PNG and Batch ZIP results

Direct API qualification assigned each analyzed file’s AM/PM suggestion as an explicit confirmed decision and generated a Safe PNG Batch ZIP in memory:

- ZIP size: 4,523,437 bytes; 22 entries.
- Three successful per-file reports, three export summaries, three Project Session files, three mapping JSON files, nine PNG chart/diagram assets, and one 'batch_summary.xlsx'.
- 'batch_summary.xlsx' contained 'Batch_Summary' and 'Batch_QC' with three summary records and the confirmed/suggested Peak columns.
- Each report contained 'Export_Metadata', 'Peak_PHF', 'Diagram_Data', 'Hourly_Summary', 'Movement_Summary', 'Charts', and 'Diagram' among the expected report sheets.
- No raw input workbook was included in the ZIP.

## Excel COM/native-template qualification

'probe_excel_com()' and the live UI reported:

~~~text
available: false
reason: PYWIN32_MISSING
detail: No module named 'win32com'; pywintypes import warning: No module named 'pywintypes'
~~~

Native Excel COM/template behavior was therefore not executable on this machine and is not claimed as qualified. Safe PNG fallback behavior was qualified successfully. No COM/template code was redesigned.

## PR and CI

- Pull request: [#19](https://github.com/bokoboss/tmc-processor/pull/19), open and non-draft; no merge performed.
- Remediation PR head for this CI record: '171988230b00776247fd499ef6e8e884d7aa887b'.
- Required PR body marker: 'Closes #6' — present.
- GitHub workflow: [CI run 35615750782](https://github.com/bokoboss/tmc-processor/actions/runs/35615750782) — completed successfully.
- [pytest (Python 3.10)](https://github.com/bokoboss/tmc-processor/actions/runs/35615750782/job/106385887165) — success.
- [pytest (Python 3.12)](https://github.com/bokoboss/tmc-processor/actions/runs/35615750782/job/106385887408) — success.

## Known limitations and scope

- Physical-direction engineering approval is not claimed because approved Mapping Presets were unavailable; the deterministic mapping proves software/runtime behavior only.
- Excel COM/native-template export was not qualified because 'win32com'/'pywintypes' are unavailable.
- UX-4 Export redesign, UX-5 batch exception-first/bulk acceptance, UX-6 architecture split, peak/PCE/QC/movement methodology changes, mapping redesign, template redesign, and Python 3.14 work were not implemented.

