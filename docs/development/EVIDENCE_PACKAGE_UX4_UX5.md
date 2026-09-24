# UX-4/UX-5 Evidence Package

## Classification

READY FOR REVIEW

This package records the bounded UX-4/UX-5 implementation cycle. It does not authorize merge or release. The branch remains unmerged.

## Repository and Git state

- Repository: `C:\MyRD\tmc-processor-public`
- Authoritative remote: `https://github.com/bokoboss/tmc-processor`
- Branch: `codex/ux-4-5-export-batch-closure`
- Baseline main/origin SHA: `c342c4b9460c2718da50debf8b344895ee6f07f7`
- Previous reviewed HEAD: `8d84161930aa749362b3eb9a78d51fc8cc6ab965`
- Remediation commit: `3efb7bf7b60c5bd1fe36a9bbec07abf9dbd008f1`
- Final remediation HEAD: `3efb7bf7b60c5bd1fe36a9bbec07abf9dbd008f1`
- Qualified runtime: `C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe` — Python 3.12.10
- Python 3.14 was not used.

## Changed files

- `app.py` — Standard report and Advanced export controls; readiness/fallback display; exception-first Batch queue, filter, bulk acceptance, exclusion/restore, and disposition display.
- `src/tmc_processor/batch.py` — Batch review states, clean-file eligibility, explicit batch bulk confirmation, exclusion disposition, readiness, and succeeded/failed/excluded summary rows.
- `src/tmc_processor/exporter.py` — Standard report backend decision and requested/used/fallback export provenance metadata, including runtime fallback updates.
- `src/tmc_processor/pipeline.py` — Propagates export provenance into generated workbooks.
- `tests/test_ux4_ux5.py` — Focused UX-4/UX-5 contracts.

## UX-4 outcome

Normal Export presents `Standard report — Recommended` as the operator-facing choice. Backend terminology remains under Advanced export options. Standard report prefers native Excel Template behavior only when Excel COM is available and the template path is compatible; otherwise it selects the validated Safe PNG/generated fallback.

Unsupported template paths are not selected by the Standard report decision. Fallback is communicated before generation and the workbook `Export_Metadata` records `export_mode_requested`, `export_mode_used`, and `export_fallback_notice`. Runtime COM/template fallback also updates the actual used backend metadata. Existing explicit Excel Template and Safe PNG paths remain available.

Readiness shows current Analysis, Peak Review confirmation, report readiness, intended report outcome, and fallback/technical details. COM diagnostics remain in the advanced area.

## Reviewer state-sync defect and remediation

The reviewer defect was that Standard report resolved its backend only in a local Export-tab variable. `report_export_mode` and `tmc_batch_export_mode` could therefore retain the previous Advanced value, while Batch export signatures were computed before the Standard decision. A Standard native Excel Batch ZIP could be marked stale on the post-generation rerun, and Advanced controls could display a backend different from the generated artifact.

The remediation routes both Standard decisions through the existing authoritative mode-change adapters. Single persists the resolved labeled backend in `report_export_mode` and its Advanced control key; Batch persists the resolved technical backend in `tmc_batch_export_mode` and its Advanced control key. Effective-mode changes use the existing export invalidation path; resolving Standard to the already-active effective backend does not invalidate. The UI reruns after a Standard transition, so the next Batch signature is computed from the resolved backend. Standard preference, native Excel preference, Safe PNG fallback, export provenance, Peak Review, bulk acceptance, and exclusion/restore behavior remain unchanged.

## UX-5 outcome

Batch Review now exposes an exception-first queue with file, survey date, analysis status, review state, suggested AM/PM, confirmed AM/PM, QC error/warning/info counts, and disposition reason. `Needs review` is the default view; failed files remain visible. Detailed per-file chart and Peak review remains available.

`Accept suggested Peaks for clean files` is explicit and uses the existing per-file confirmation semantics with `user_confirmed_batch` provenance. Eligibility requires successful analysis, zero QC errors, zero QC warnings, both suggestions, and no prior confirmation/exclusion. QC info-only files remain eligible. Warning/error files remain in `Needs review`.

Successful files may be intentionally excluded with a retained reason and can be restored to review. Excluded files satisfy Batch review readiness, are not exported as successful report folders, and appear as `excluded` rather than `failed` in Batch summary output. Batch readiness requires each successful file to be confirmed or excluded; failed analysis retains failed disposition behavior.

Draft Peak edits do not alter Analysis or confirmed values. Confirmation and exclusion changes preserve Analysis and invalidate stale Batch Export when downstream package content changes.

## Focused tests

Executed during implementation:

```text
.venv\Scripts\python.exe -m pytest tests/test_ux4_ux5.py -q
12 passed

.venv\Scripts\python.exe -m pytest tests/test_ux4_ux5.py tests/test_workflow_state.py tests/test_workflow_contract_adapter.py tests/test_explicit_peak_review.py -q
72 passed

.venv\Scripts\python.exe -m pytest tests/test_ux4_ux5.py tests/test_explicit_peak_review.py -q
15 passed

.venv\Scripts\python.exe -m pytest tests/test_phase_l0_dual_scheme.py::test_single_file_export_default_is_template_when_excel_com_available tests/test_phase_l0_dual_scheme.py::test_single_file_export_default_is_safe_png_when_excel_com_unavailable tests/test_phase_l0_dual_scheme.py::test_approach_movement_batch_export_default_is_safe_png -q
3 passed
```

The existing Batch, explicit Peak Review, effective export, UI helper, dual-scheme, workflow contract, and setup persistence coverage was preserved.

## Final validation

- Full pytest: `356 passed, 11 warnings` in 3:08. The warnings are existing openpyxl DrawingML warnings.
- Compile: `.venv\Scripts\python.exe -m compileall -q app.py src` passed.
- `git diff --check` passed.
- Final remediation working tree was clean at `3efb7bf7b60c5bd1fe36a9bbec07abf9dbd008f1`; this evidence update is documentation-only.

## Focused real-workbook UAT

The three protected workbooks under `samples/raw/` were exercised through the Python API using a temporary deterministic `from_to` mapping generated from detected sheet names. This proves software/runtime behavior only and is not physical-direction engineering approval.

### Single

- Representative workbook: `03.TMC1 แยกกบินทร์บุรี_พุธ 17 กค67f.xlsx`.
- Confirmed AM `07:00-08:00` and PM `17:00-18:00`.
- Standard report with COM unavailable resolved to Safe PNG and communicated the fallback.
- Generated workbook contained 21 sheets; metadata recorded requested `Standard report — Recommended`, used `Safe PNG Export Mode`, fallback notice, and `user_confirmed` Peak provenance.

### Batch

- All three protected workbooks analyzed successfully with suggestions:
  - Kabin Buri: `07:00-08:00` / `17:00-18:00`, QC `0/0/8`.
  - Bo Phloi: `07:30-08:30` / `15:45-16:45`, QC `0/1/2`.
  - Nong Prue: `10:45-11:45` / `15:45-16:45`, QC `0/0/0`.
- Clean-file bulk acceptance accepted Kabin Buri and Nong Prue only.
- Bo Phloi warning case was not bulk accepted; it remained unresolved until intentionally excluded with a reason.
- Exclusion made readiness pass; restoring it returned Batch to `Needs review`.
- Generated Safe PNG Batch ZIP contained two succeeded report folders plus one excluded summary row, `batch_summary.xlsx`, and no raw input workbooks.

## Safe PNG and COM status

Safe PNG runtime was qualified in Single and Batch UAT. Native Excel COM/template execution was not qualified on this machine because `win32com`/`pywintypes` are unavailable. No pywin32 or system software was installed. Exact native Excel template fidelity remains the preferred Standard report path when COM and a compatible template are available; Safe PNG is the validated fallback.

## Known limitations

- Deterministic UAT mapping is software qualification only; no physical-direction engineering approval is claimed.
- Native Excel COM/template execution remains unexecuted and is mandatory for the UX-7/release gate.
- UX-6 architecture split, Peak/PCE/QC/movement methodology changes, mapping redesign, template redesign, and Python 3.14 compatibility were not implemented.

## PR and CI

- Pull request: [#20](https://github.com/bokoboss/tmc-processor/pull/20), opened against `main`.
- Required PR body markers: `Closes #7` and `Closes #8`.
- GitHub CI for final remediation HEAD: [run 36006632881](https://github.com/bokoboss/tmc-processor/actions/runs/36006632881) passed; both Python 3.10 and Python 3.12 jobs completed successfully.
