# UX-8 P1 Operator UI cleanup — evidence

## Baseline and scope

- Accepted `main`: `67c645baf1323b3d8716079574504c2a167dce61`.
- Branch: `codex/ux-8-operator-ui-cleanup`.
- Linked worktree: `C:\MyRD\TMC Processor\ux8-operator-ui-cleanup`.
- Changed source: `src/tmc_processor/ui/workflows/single.py`, `src/tmc_processor/ui/workflows/batch.py`, and `src/tmc_processor/ui/app_shell.py`.
- Changed tests: `tests/test_setup_persistence.py` and `tests/test_ux8_p1_ui_hierarchy.py`.
- This note is the only documentation addition. There are no changes to parsers, domain calculations, Peak algorithms, PCE calculations, QC, Project Session or Mapping Preset schemas, exporters, Excel template, or COM backend. No P2 work is included.

## Operator hierarchy

| Stage | Before | After |
| --- | --- | --- |
| Single Data | Project and report fields shared the primary view; Survey period appeared in Analyze. | Source identity and detected sheets, project fields, and Survey period lead. Weather, responsible party, direction labels, road names, diagram caption, and U-turn display remain in one closed report-details disclosure. |
| Single Mapping | Six summary metrics, advanced scheme selection, and reuse tools competed with the editor. | Four summary metrics and the Basic mapping editor lead. Scheme selection and Mapping Excel/Preset controls remain accessible in separate closed disclosures. |
| Single Analyze | Survey period and Peak mode appeared alongside the AM/PM windows. | AM/PM windows lead. Peak mode and the existing PCE editor remain accessible in closed disclosures. |
| Single Review | Technical metrics and multiple technical table expanders competed with Peak/QC review. | Peak confirmation and QC lead. Technical metrics, parser details, and data tables are grouped in one closed disclosure. |
| Batch Data | Later-stage PCE, Peak, and ZIP readiness text appeared beside source setup. | Shared project/source information leads; report detail is collapsed. Duplicate source and output names surface as warnings. |
| Batch Mapping | Readiness also listed PCE and per-file metadata. | Source and shared Mapping Preset readiness lead; preset and sheet-matching controls remain accessible. |
| Batch Analyze | Survey period repeated in Analyze and Peak mode preceded time windows. | Survey period is Data-owned; AM/PM windows lead, with Peak mode and PCE settings still accessible. |
| Batch Review | Export status and a static ZIP checklist appeared in review. | Per-file analysis, Peak, and QC state lead. Export state remains in Export. |

The `survey_period_input` widget and `survey_period` setup key are unchanged. Removing `survey_period` from `ANALYZE_SETUP_FIELDS` keeps its Data-owned value through stage reruns and processing. Workflow invalidation rules and stored schemas were not changed. UX-7 Standard/Safe PNG choices, labels, and routing were not changed.

## Batch readiness remediation

The first populated smoke found a disabled Analyze Batch action despite two persisted uploads and a loaded preset. `_remember_batch_uploads` stores the files under `tmc_batch_source_uploads`; `_remember_upload` stores the preset under `tmc_batch_mapping_preset_upload`. `_sync_batch_workflow_from_state` already writes `WorkflowReadiness(source=bool(batch_uploads), mapping=mapping_preset is not None, ...)` to the Batch workflow state. The Batch `WorkflowContext` constructor omitted its existing `uploaded_ready` and `mapping_ready` fields, so the renderer received `None` for both. Its existing `batch_inputs_ready` function therefore kept Analyze disabled.

The app shell now passes `batch_workflow_state.readiness.source` and `.mapping` into those two existing context fields. No readiness formula, session-state source, validation, movement scheme rule, PCE gate, metadata handling, or Batch processing function changed. The regression test drives two demo uploads and a preset through the real app shell and renderer, checks persisted state, and verifies both the enabled positive case and the disabled missing-preset case.

## Batch analysis propagation and context audit

The next populated smoke completed `application_analyze_batch` and stored its one result in `tmc_batch_analysis_result`, but Review's Peak queue displayed the empty state. Its first section read `context.batch_analysis`, which the Batch context constructor had omitted; its QC section re-read the saved result directly. The app shell now passes that saved object into the existing `batch_analysis` field, and Review uses that same context value in both sections. No second state key, analysis rerun, or Peak rule was introduced.

A bounded audit compared every `context` field read by the five Batch renderers with the Batch constructor. It found three more already-defined fields needed for the existing flow: `batch_export_mode` from `tmc_batch_export_mode`, `batch_stale` from `tmc_batch_stale`, and `batch_export_stale` from `tmc_batch_export_stale`. They are now passed through without changing their derivation or invalidation rules. The remaining unpopulated reads are local placeholders overwritten before use, unused display placeholders, or `batch_process_block_reason`, for which there is no current block (the existing reason function returns an empty string for valid schemes).

| Batch renderer | Context fields read | Populated by constructor | Remaining unpopulated classification |
| --- | ---: | ---: | --- |
| Data | 18 | 8 | 10 local or unused placeholders |
| Mapping | 25 | 7 | 17 local or unused placeholders; 1 empty block reason |
| Analyze | 24 | 14 | 9 local or unused placeholders; 1 empty block reason |
| Review | 56 | 7 | 49 local placeholders, including the saved export result reloaded before use |
| Export | 35 | 15 | 20 local placeholders, including the saved analysis and export result reloaded before use |

No new `WorkflowContext` field was added. The Review regression exercises the real app shell and renderer, asserts the context's analysis is the identical saved object, verifies the populated Peak queue and QC, and checks the pre-confirmation Export action remains disabled. A no-analysis case still shows the empty Review state.

## Verification

- Focused readiness regression: `C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe -m pytest -q tests/test_ux8_p1_ui_hierarchy.py -k 'batch_analyze_receives or batch_analyze_stays'` — **2 passed** after the wiring fix. Before the fix, the positive case failed while the negative case passed.
- Focused Review regression: `C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe -m pytest -q tests/test_ux8_p1_ui_hierarchy.py -k 'batch_review_receives or batch_review_without'` — **2 passed** after the wiring fix. Before the fix, the populated case failed because `context.batch_analysis` was `None`.
- Expanded focused suite: `C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe -m pytest -q tests/test_ux8_p1_ui_hierarchy.py tests/test_setup_persistence.py tests/test_application_architecture.py tests/test_workflow_contract_adapter.py tests/test_workflow_contract.py tests/test_workflow_state.py tests/test_batch.py tests/test_upload_persistence.py tests/test_explicit_peak_review.py tests/test_effective_peak_export.py tests/test_ux7_p0_ui_cleanup.py tests/test_ux7_standard_ooxml_routing.py` — **179 passed, 3 known openpyxl DrawingML warnings**.
- Full regression: `C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe -m pytest` — **431 passed, 0 failures, 14 warnings in 254.35 s**. All warnings are the existing openpyxl DrawingML limitation; the warning count matches the pre-Review-fix run.
- Compile: `C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe -m compileall -q app.py src` — passed.
- `git diff --check` — passed.
- `tests/conftest.py` inserts this worktree's `src` at the front of `sys.path`; the test run used this linked worktree as its working directory.

## Live smoke

- Single: previously completed on the edited worktree with `samples/demo/DEMO_TMC1_FourLeg.xlsx` and `samples/demo/DEMO_TMC1_FourLeg.mapping.json`. Data showed four detected sheets; Mapping Preset loaded; Analyze, Review, Peak confirmation, and Export were usable. The accepted UX-7 Standard action and wording remained present. No later Single UI edits require repeating this flow.
- Batch empty state: Data, Mapping, Analyze, Review, and Export were visited.
- Populated Batch: used `samples/demo/DEMO_TMC1_FourLeg.xlsx`, `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`, and `samples/demo/DEMO_TMC1_FourLeg.mapping.json`. Data retained both workbooks. Mapping loaded the shared preset and displayed compatible `from_to` readiness. Analyze was enabled with default PCE factors and completed with two successful files, zero failed files, and six QC info entries.
- Review: the Peak queue and QC summary showed the same two-file analysis. The first and second file's suggested AM/PM windows were each explicitly confirmed using the existing per-file action; the confirmation count advanced from `0/2` to `1/2` to `2/2`. The failed and excluded summaries both remained zero for these demo inputs.
- Export: after both confirmations, the existing checklist showed source, preset, analysis, Peak, and output stem ready. Standard Batch ZIP generation completed. The per-file export status section and ZIP download action appeared. Advanced options still listed Safe PNG Export Mode as the fallback backend. No failed or excluded file outcome was generated by this all-success sample; those presentation paths were not live exercised.
- The live server was launched with the linked worktree's `src` first on `PYTHONPATH`; `tmc_processor.__file__` resolved to `C:\MyRD\TMC Processor\ux8-operator-ui-cleanup\src\tmc_processor\__init__.py`. The shared virtualenv otherwise resolves the package from the authoritative checkout. No shared environment or Office/Windows configuration was changed. The browser session and server were closed, and temporary browser snapshots were removed.

## Change control and limitation

The current worktree is intentionally uncommitted. No push, PR, or merge was performed. The demo smoke verifies the operator flow and does not qualify Batch domain behavior beyond the bundled all-success sample.
