# UX-6 Evidence Package

## Scope

- Repository: `https://github.com/bokoboss/tmc-processor`
- Authoritative local path: `C:\\MyRD\\tmc-processor-public`
- Branch: `codex/ux-6-application-shell`
- Previous reviewed HEAD: `d3e30f9b775d32836e39f07f1d07496a79cd7c9e`
- Remediation commit: final implementation commit on this branch
- Objective: final architecture-only UX-6 remediation with no intentional operator-visible redesign.

Canonical workflow remains `Data -> Mapping -> Analyze -> Review -> Export`.

## Final architecture

`app.py` remains the minimal executable/compatibility facade. It delegates execution and attribute access to `tmc_processor.ui.app_shell` and contains no workflow implementation.

`ui/app_shell.py` owns Streamlit bootstrap, mode/sidebar setup, navigation, shared state preparation, compatibility adapters, named workflow-context construction, and canonical-stage dispatch. It no longer passes `globals()`/`locals()` or an equivalent whole-shell namespace to workflows.

`ui/workflow_context.py` defines the explicit boundary:

- `WorkflowContext`: active stage, mode, session state, and named workflow data/state fields.
- `WorkflowOperations`: named callbacks, constants, imported helpers, and application service seams exposed by the shell.

The shell constructs this contract explicitly at the Batch and Single dispatch points. Stage modules do not import `app_shell`, and the context module does not import the shell.

### Stage ownership

- `workflows/data.py` owns the canonical Data entrypoint and dispatches only to `render_single_data` or `render_batch_data`.
- `workflows/mapping.py` owns the canonical Mapping entrypoint and dispatches only to `render_single_mapping` or `render_batch_mapping`.
- `workflows/analyze.py` owns the canonical Analyze entrypoint and dispatches only to `render_single_analyze` or `render_batch_analyze`.
- `workflows/review.py` owns the canonical Review entrypoint and dispatches only to `render_single_review` or `render_batch_review`.
- `workflows/export.py` owns the canonical Export entrypoint and dispatches only to `render_single_export` or `render_batch_export`.
- `workflows/single.py` contains the five Single renderers and no stage selector.
- `workflows/batch.py` contains the five Batch renderers and no stage selector.

There is no `render_single_stage`, `render_batch_stage`, or `active_tab == ...` selector in the mode-specific workflow modules. Existing rendering bodies were mechanically separated; domain, mapping, Peak, QC, export, provenance, and session behavior were not redesigned.

Pure state/workflow/service seams remain under `src/tmc_processor/application/`; reusable display boundaries remain under `src/tmc_processor/ui/components/`.

## Services and dependency direction

`application/services.py` remains the orchestration seam for Single analysis, Single dry-run analysis, generated Single export, Single template/COM export, Batch analysis, and reviewed Batch ZIP export. The shell uses those service aliases and does not bypass them with direct pipeline, batch, or reviewed-export calls.

`application/state.py` and `application/workflow.py` remain Streamlit-free. Accepted provenance keys remain `tmc_confirmed_peak_selection_source` and `tmc_batch_peak_selection_source`, with `user_confirmed`, `user_confirmed_batch`, and `auto_suggested` values preserved.

The enforced dependency direction is:

`domain -> application -> UI workflows/components -> app.py`

Application modules do not import Streamlit, `app`, or `tmc_processor.ui`. UI components do not import the engineering calculation modules tested by the architecture suite.

## Changed files

```text
docs/development/EVIDENCE_PACKAGE_UX6.md
src/tmc_processor/ui/app_shell.py
src/tmc_processor/ui/workflow_context.py
src/tmc_processor/ui/workflows/__init__.py
src/tmc_processor/ui/workflows/data.py
src/tmc_processor/ui/workflows/mapping.py
src/tmc_processor/ui/workflows/analyze.py
src/tmc_processor/ui/workflows/review.py
src/tmc_processor/ui/workflows/export.py
src/tmc_processor/ui/workflows/single.py
src/tmc_processor/ui/workflows/batch.py
tests/test_application_architecture.py
```

## Validation

- Architecture tests: `11 passed`.
- Streamlit setup/navigation regression group: `12 passed, 1 warning`.
- Final full suite at the remediation implementation head: `369 passed, 11 warnings`.
- `.venv\\Scripts\\python.exe -m compileall -q app.py src`: passed.
- `git diff --check`: passed.
- Streamlit AppTest covered Single and Batch stage navigation and setup persistence; the required Data -> Mapping -> Analyze -> Review -> Export path remained executable without app exceptions.
- Streamlit startup/health smoke remains passed at `http://localhost:8501/_stcore/health`.
- Protected workbook UAT and COM qualification were not repeated because this change is structural and the existing regression suite remained green. Native Excel COM remains unavailable on this host.

## PR and CI

- Pull request: [#21](https://github.com/bokoboss/tmc-processor/pull/21)
- Baseline CI #96: [successful](https://github.com/bokoboss/tmc-processor/actions/runs/36238711238), Python 3.10 and 3.12, `369 passed, 11 warnings`, on the previous reviewed head.
- Final remediation CI: recorded after the final implementation push.

## Known limitations

- Native Excel COM/template fidelity remains unexecuted on this host because the COM runtime is unavailable.
- Protected real-workbook smoke uses deterministic qualification mapping and is not physical-direction engineering approval.
- No merge or new pull request was performed.

## Classification

READY FOR RE-REVIEW, contingent on final-head CI passing on Windows Python 3.10 and 3.12.
