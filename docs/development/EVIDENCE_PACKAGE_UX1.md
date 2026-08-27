# Evidence Package — UX-1 Remediation

## Provenance

- Repository: `C:\MyRD\tmc-processor-public` (execution worktree: `C:\MyRD\tmc-processor-public-ux-1`)
- Issue: [#4 — UX-1: Replace competing navigation with canonical five-stage workflow](https://github.com/bokoboss/tmc-processor/issues/4)
- Existing PR: [#15](https://github.com/bokoboss/tmc-processor/pull/15), against `main`; no new PR was created and no merge was performed.
- Branch: `codex/ux-1-five-stage-navigation`
- Accepted validation baseline: `b75efcff3f359267906137f70ce5b1d90eeb9cd3`
- Remediation commits: `233a8ca` and `5b50cae`
- Evidence updated: 2026-08-26 (Asia/Bangkok)

## Scope and protected behavior

UX-1 preserves the canonical five-stage workflow for Single and Batch modes:
`Data -> Mapping -> Analyze -> Review -> Export`.

The remediation covers the two known PR #15 lifecycle regressions:

- Analyze settings are rehydrated from canonical setup state after stage navigation, so time inputs do not reset to the current clock.
- Explicit upload clearing resets the corresponding Streamlit uploader widget as well as canonical state, so reselecting the same Batch files does not append duplicates.

Calculation, mapping, QC, Peak, export, and the `WorkflowState` revision/invalidation contract remain protected. The three local raw workbooks under `samples/raw/` were read from the original checkout and were not added to public Git history.

## Automated validation

| Gate | Command / method | Result |
|---|---|---|
| Accepted baseline | `python -m pytest` at `b75efcff3f359267906137f70ce5b1d90eeb9cd3` | PASS — 297 passed, 11 expected openpyxl DrawingML warnings |
| Focused UX/state suite | `python -m pytest tests/test_upload_persistence.py tests/test_setup_persistence.py tests/test_workflow_contract.py tests/test_workflow_contract_adapter.py tests/test_workflow_state.py -q --disable-warnings` | PASS — 88 passed, 1 warning in 10.53s |
| Full suite | `python -m pytest` | PASS — 316 passed, 11 expected openpyxl DrawingML warnings in 171.91s |
| Compile | `python -m compileall -q app.py src` | PASS |
| Diff hygiene | `git diff --check` | PASS |
| Streamlit smoke | `python -m streamlit run app.py --server.headless true --server.port 8508 --browser.gatherUsageStats false`; `GET /_stcore/health` | PASS — `ok` |

The focused tests cover uploader persistence and clear/replacement behavior, canonical Analyze setup hydration, stage ownership, navigation-only stability, readiness, artifact preservation, and intentional invalidation when setup changes.

## Single browser/runtime qualification

Final-head UAT was run against the Windows Streamlit server in the in-app Browser.

- Uploaded the protected Kabin Buri workbook, cleared it, and confirmed the canonical source and prompt were cleared.
- Reuploaded the same bytes under `ux1-kabin-renamed.xlsx`; the loaded identity changed to the new filename and the old name disappeared. The original filename was then restored for processing.
- Loaded the deterministic qualification Mapping Preset: 24 valid rows, 24 matched, 0 missing, 0 extra.
- Set `rolling_60min` with AM `06:00-10:00` and PM `15:00-19:00`; navigated through Mapping -> Data -> Analyze and confirmed all values persisted.
- Before analysis, unchanged navigation left the analysis non-stale. Changing AM start to `06:15` produced the expected stale/re-analysis-required state. Restoring `06:00` and reanalyzing returned the workflow to ready.
- Analysis completed with 13,824 normalized rows, 82,621 vehicles, 87,023.6 PCU, 0 QC errors, 0 QC warnings, and 8 informational QC findings. Effective peaks were AM `07:00-08:00` and PM `17:00-18:00`.
- Excel Template Mode export completed successfully. The downloaded artifact was `C:\Users\kitti\Downloads\tmc_processor_output (2).xlsx`.

## Batch browser/runtime qualification

Final-head UAT was run against the same Windows Streamlit server in the in-app Browser.

- Uploaded the protected Kabin Buri, Bo Phloi, and Nong Prue workbooks; Data retained `Loaded sources: 3 workbook(s)` through Mapping navigation.
- `Clear Batch sources` cleared the canonical inventory and readiness. Reselecting the same three files produced exactly three loaded sources, not six.
- Loaded the deterministic qualification Mapping Preset: 24 rows, 3 source files, ready.
- `Clear Mapping Preset` removed the preset and readiness; reupload restored the 24-row, 3-file ready state.
- Set `rolling_60min` with AM `06:00-10:00` and PM `15:00-19:00`; navigation away and back preserved the values. The intentional setting change produced the expected stale state before analysis.
- Batch analysis completed for all 3 files: `3/3` Peak assignments, `0` failed files. The selected Kabin result matched the Single runtime totals: 82,621 vehicles, 87,023.6 PCU, 0 QC errors, 0 warnings, 8 info findings.
- Safe PNG Export Mode generated a ZIP successfully: `C:\Users\kitti\Downloads\UX-1_deterministic_qualification_only_batch_package.zip` (3,672,543 bytes, 22 entries, including 4 `.xlsx` files and 9 `.png` chart/diagram images).
- Visual inspection of the extracted `hourly_pcu.png` and `tmc_movement_diagram.png` showed readable labels, highlighted peaks, movement arrows, and summary tables.

## Protected real-workbook runtime metrics

Read-only current-head processing used the three protected raw workbooks with the deterministic qualification mapping, `rolling_60min`, search windows `06:00-10:00` and `15:00-19:00`, and generated workbook output. These numbers qualify ingestion, analysis, Peak, QC, and export mechanics; they do not approve physical direction mapping.

| Workbook | Source sheets | Normalized rows | Hourly rows | Movement rows | Vehicle rows | Effective AM peak | Effective PM peak | QC errors / warnings | Workbook bytes |
|---|---:|---:|---:|---:|---:|---|---|---|---:|
| Kabin Buri | 24 | 13,824 | 12 | 16 | 12 | 07:00-08:00 | 17:00-18:00 | 0 / 0 | 1,727,916 |
| Bo Phloi | 15 | 8,640 | 12 | 9 | 12 | 07:30-08:30 | 15:45-16:45 | 0 / 0 | 1,261,544 |
| Nong Prue | 12 | 6,912 | 12 | 8 | 12 | 07:45-08:45 | 15:45-16:45 | 0 / 0 | 1,110,862 |

## Excel visual qualification

Microsoft Excel COM was available at version 16.0. The final Single export was opened through the Windows Excel UI and visually inspected on `Diagram_Data`, `Report_Text`, and `Peak_PHF`.

- `Diagram_Data` rendered populated movement totals and AM/PM peak PCU columns.
- `Report_Text` rendered the Thai report text and totals/peak values.
- `Peak_PHF` rendered the rolling-60-minute mode, configured search windows, effective peak windows, and PHF values.
- The artifact contains 19 sheets and opens successfully.

Known limitation: openpyxl exposed zero chart objects in the COM/template artifact, and the workbook did not contain `TMC_Report`, `Charts`, or `Diagram` sheets. Native COM chart/template fidelity is therefore not claimed. Safe PNG output is the verified fallback and was visually inspected above.

## Mapping qualification limitation

The approved physical Mapping Preset for the protected real workbooks was not available in the execution environment. The qualification runs used a temporary deterministic `from_to` mapping fixture generated from the protected workbook structure. It was not committed and is not an engineering approval of physical direction mapping. A release decision still requires the approved Mapping Preset and its normal mapping/QC review.

## CI and review status

- GitHub Actions CI was queued for code head `5b50cae` as run `#71` at the time of this evidence capture; the final handoff records the latest remote result after the evidence-only commit is pushed.
- No new lifecycle defect outside the two known upload/Analyze remediation scopes was found. The pre-patch same-file Batch duplication was part of the upload-clear regression and was fixed by uploader widget revision reset.
- Final review classification: **READY FOR CODE REVIEW** after CI settles; **NOT READY FOR FINAL RELEASE / MERGE** until the approved physical Mapping Preset is available and the native COM chart/template decision is resolved.

## Changed files

- `app.py`
- `PROJECT_PROFILE.md`
- `tests/test_setup_persistence.py`
- `tests/test_upload_persistence.py`
- `tests/test_workflow_contract_adapter.py`
- `tests/test_workflow_state.py`
- `docs/development/EVIDENCE_PACKAGE_UX1.md`

UX-2 sheet-centric Mapping redesign, UX-3 explicit Peak confirmation redesign, UX-4 export simplification, and UX-5 Batch exception-first review remain out of scope.
