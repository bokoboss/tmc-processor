# UX-6 Evidence Package

## Scope

- Repository: `https://github.com/bokoboss/tmc-processor`
- Authoritative local path: `C:\\MyRD\\tmc-processor-public`
- Branch: `codex/ux-6-application-shell`
- Baseline: `096132f9a6269e8b91e96f809f3499aaf69d7d8f`
- Previous reviewed HEAD: `6002577cae824419a3eb24968607413a7d437c6e`
- Remediation commit: `f65176ff8ca70bc3d54d5634ba852b0432b37414`
- Objective: architecture-only UX-6 split with no intentional operator-visible redesign.

Canonical workflow preserved: `Data -> Mapping -> Analyze -> Review -> Export`.

## Architecture before and after

Before, `app.py` was a 7,380-line mixed Streamlit module containing bootstrap, navigation, session-state adapters, workflow revisions/transitions, mapping and Peak review semantics, service orchestration, reusable rendering, and all single/batch stage rendering.

After, `app.py` remains an executable compatibility facade. `ui/app_shell.py` now owns only Streamlit bootstrap, mode/sidebar setup, navigation, shared context preparation, compatibility adapters, and canonical-stage dispatch. The complete Single and Batch stage implementations are under `ui/workflows/single.py` and `ui/workflows/batch.py`; `data.py`, `mapping.py`, `analyze.py`, `review.py`, and `export.py` are the canonical registry owners that select the appropriate mode implementation. Pure state/workflow/service seams live under `src/tmc_processor/application/`; reusable display boundaries live under `src/tmc_processor/ui/components/`.

Dependency direction is enforced by tests:

`domain -> application -> UI workflows/components -> app.py`

Application modules do not import Streamlit, `app`, or `tmc_processor.ui`. UI components do not import calculation, Peak, QC, PCU, or pipeline engines. `workflow_state.py` remains the sole semantic revision/invalidation authority.

## Changed files and module tree

```text
app.py
src/tmc_processor/application/
  __init__.py
  state.py
  workflow.py
  services.py
src/tmc_processor/ui/
  __init__.py
  app_shell.py
  components/{__init__,mapping,peak,qc,status,export}.py
  workflows/{__init__,data,mapping,analyze,review,export,single,batch}.py
tests/test_application_architecture.py
tests/test_application_services.py
docs/development/EVIDENCE_PACKAGE_UX6.md
```

`application/state.py` owns plain-mapping adapters for Single confirmed/draft Peak state, Batch confirmed/draft/disposition state, export backend provenance, and legacy session-key synchronization. `application/workflow.py` owns pure revision composition and transition orchestration around `WorkflowState`, `WorkflowReadiness`, `WorkflowRevisions`, `WorkflowTransition`, and existing fingerprints. `application/services.py` only composes existing domain pipeline, batch, and exporter functions.

The UI package owns reusable display components and canonical stage entrypoints. The app shell dispatches with `render_workflow_stage(active_tab, context=...)` and no longer contains the full Data/Mapping/Analyze/Review/Export implementations. `application/services.py` is used for Single analyze, Single dry-run analyze, generated export, template/COM export, Batch analyze, and reviewed Batch ZIP export. Accepted provenance keys remain `tmc_confirmed_peak_selection_source` and `tmc_batch_peak_selection_source`, with `user_confirmed`, `user_confirmed_batch`, and `auto_suggested` values preserved.

Compatibility is retained through `app.py` attribute delegation and a module assignment bridge so established tests that access or monkeypatch `app.*` continue to target the one shell implementation. No duplicate business/state implementation was retained in `app.py`.

## Protected behavior

The refactor did not change Peak windows/search methodology, PCE/PCU, QC, movement aggregation or code derivation, physical Mapping semantics, Mapping Preset or Project Session contracts, effective-Peak propagation, workbook template mappings, Excel COM behavior, or Safe PNG/generated export behavior.

UX-3 characterization remains green for Suggested -> Draft -> Explicit Confirmed, draft-only edits, changed confirmation, identical reconfirmation, and Single/Batch provenance. UX-4 Standard/Advanced backend state and UX-5 exception-first Batch confirmation, exclusion/restore, and export readiness remain green.

## Validation

- Focused architecture/service tests: `13 passed`.
- Required focused regression groups: `131 passed, 1 warning`.
- Final full suite at final implementation head: `369 passed, 11 warnings`.
- `.venv\Scripts\python.exe -m compileall -q app.py src`: passed.
- `git diff --check`: passed.
- Streamlit startup: `python -m streamlit run app.py` started successfully at `http://localhost:8501` and stopped cleanly.
- Streamlit AppTest navigation: Single -> Batch -> Single mode switching passed with no app exceptions; the headless health endpoint returned `ok`.
- Protected Single smoke: representative Kabin Buri workbook analyzed with 13,824 normalized rows, 8 QC rows, and a 1,690,248-byte Safe PNG/generated workbook. COM probe reported unavailable.
- Protected Batch smoke: all three protected workbooks analyzed successfully; all three were clean-file bulk eligible under the deterministic qualification mapping; exclusion and restore worked; a reviewed Safe PNG Batch ZIP was generated with 3,752,788 bytes and three confirmed rows.
- The protected-workbook mapping fixture was deterministic `from_to` mapping generated from detected sheet names. It proves wiring/runtime only and is not physical-direction engineering approval.
- Native Excel COM was not executed because `win32com`/`pywintypes` are absent. No pywin32 or system software was installed. Native COM/template execution remains a UX-7/release gate.

## Scope intentionally not changed

No frontend framework migration, domain-engine redesign, calculation-method change, template redesign, Python 3.14 work, or deep engineering workbook requalification was performed. Existing real-workbook and COM release gates remain applicable.

## PR and CI

- Pull request: [#21](https://github.com/bokoboss/tmc-processor/pull/21)
- PR body includes `Closes #9`.
- Superseded pre-remediation workflow: [CI run #93](https://github.com/bokoboss/tmc-processor/actions/runs/36235120662)
- Final-head workflow: to be recorded after pushing `f65176ff8ca70bc3d54d5634ba852b0432b37414`; Python 3.10 and Python 3.12 jobs remain the release gate.

## Known limitations

- Native Excel COM/template fidelity remains unexecuted on this host.
- Protected real-workbook smoke uses deterministic qualification mapping only.
- Final review readiness is gated on the external CI run for remediation commit `f65176ff8ca70bc3d54d5634ba852b0432b37414`.
