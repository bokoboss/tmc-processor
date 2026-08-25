# Evidence Package

## Provenance
- Repository: `C:\MyRD\tmc-processor-public` (execution worktree: `C:\MyRD\tmc-processor-public-ux-1`)
- Issue: [#4 — UX-1: Replace competing navigation with canonical five-stage workflow](https://github.com/bokoboss/tmc-processor/issues/4)
- PR: To be opened against `main` after the focused commit is pushed.
- Branch: `codex/ux-1-five-stage-navigation`
- Head SHA: Recorded in the final PR handoff (`git rev-parse HEAD`). Validation baseline: `b75efcff3f359267906137f70ce5b1d90eeb9cd3`.
- Date: 2026-08-25 (Asia/Bangkok)

## Execution route
- Model / effort: GPT-5.6 Luna / Max
- Orchestrator / workers if used: Codex desktop; no worker delegation.
- Retries / escalations: One targeted test repair cycle after the initial UX-1 edits; no external escalation.
- Reason for escalation if any: N/A.

## Validation evidence
| Gate | Command / Method | Result | Evidence / Artifact |
|---|---|---|---|
| Clean baseline | `python -m pytest` at `b75efcff3f359267906137f70ce5b1d90eeb9cd3` | PASS — 297 passed, 11 warnings | Baseline captured before implementation on clean `origin/main`. |
| UX-1/state regression suite | `python -m pytest -q --disable-warnings` | PASS — 301 passed, 11 warnings in 172 seconds | Includes canonical stage, stage ownership, navigation-only invalidation, shell/readiness, mapping/PCE/Peak/export contract tests, and existing regression coverage. |
| Compile | `python -m compileall -q app.py src` | PASS | No compile errors. |
| Diff hygiene | `git diff --check` | PASS | No whitespace errors. |
| Streamlit process smoke | `python -m streamlit run app.py --server.headless true --server.port 8507`; `GET /_stcore/health` | PASS — `ok` | Windows process started and stopped cleanly. |
| Single UX runtime | Streamlit `AppTest`: upload demo workbook in Data; Mapping; Analyze TMC; Review; return to Data; Export; generate workbook | PASS | `tests/test_setup_persistence.py::test_streamlit_apptest_setup_values_survive_processing_and_export`. |
| Single stage ownership | Streamlit `AppTest` stage navigation | PASS | Analyze TMC is absent in Mapping and present in Analyze. |
| Batch stage ownership | Streamlit `AppTest` mode switch and stage navigation | PASS | Batch Mapping exposes the Mapping Preset uploader; Analyze exposes `วิเคราะห์ Batch`; Mapping does not expose the analysis action. |
| Navigation invalidation | Adapter tests across all five stages for Single and Batch | PASS | Active-stage changes leave revisions and artifacts unchanged. |
| Excel COM probe | `_probe_excel_com_for_ui(force=True)` | PASS — available, Excel version 16.0 | Local Windows host has Microsoft Excel COM available. |
| Excel Template / COM export | Real demo processing with `use_template_report_layout=True`, `use_excel_com_native_charts=True` | PASS — 269,174-byte workbook returned | Expected protected-formula warnings were emitted by the existing COM guard. Manual visual inspection remains outstanding. |
| Safe PNG / generated export | Real-workbook processing with Safe PNG-compatible generated export | PASS | Three local raw workbooks produced valid workbook bytes in memory. |

## Engineering / reference evidence
- Canonical stage model is shared by both modes: `Data -> Mapping -> Analyze -> Review -> Export`.
- Primary Single/Batch workbook uploaders are stage-owned by Data; the Batch Mapping Preset uploader is stage-owned by Mapping.
- Single Analyze TMC and Batch Analyze actions are stage-owned by Analyze; Mapping has no analysis action.
- Review retains existing QC, result, and effective Peak behavior; Export generation logic remains in place.
- `WorkflowState`, `WorkflowReadiness`, semantic fingerprints, and invalidation rules in `src/tmc_processor/workflow_state.py` were not changed.
- Legacy navigation values are normalized only as compatibility adapters; navigation keys are not included in semantic revisions.
- `PROJECT_PROFILE.md` now accepts `b75efcff3f359267906137f70ce5b1d90eeb9cd3` as the post-governance baseline.

## Browser / UX / real-data evidence
- AppTest covered the canonical navigation and the end-to-end Single workflow with the public demo workbook.
- The Streamlit server health endpoint passed on Windows.
- Protected local raw-workbook ingestion and Safe PNG processing were run in memory for:
  - `03.TMC1 แยกกบินทร์บุรี_พุธ 17 กค67f.xlsx`: 24 source sheets, 13,824 normalized rows, AM/PM peaks, 1,714,721-byte generated workbook.
  - `03.TMC1 แยกบ่อพลอย_พุธ 28 พค68.xlsx`: 15 source sheets, 8,640 normalized rows, AM/PM peaks, 1,284,160-byte generated workbook.
  - `03.TMC4 แยกหนองปรือ_พุธ 28 พค68.xlsx`: 12 source sheets, 6,912 normalized rows, AM/PM peaks, 1,130,790-byte generated workbook.
- The Batch engine then analyzed all three files successfully and generated a 7,116,394-byte Safe PNG ZIP in memory.
- The local raw workbooks are intentionally not in the public worktree. Because no approved Mapping Preset accompanies them, the real-workbook processing check used a temporary deterministic `from_to` qualification mapping only to exercise ingestion, analysis, Peak, and export mechanics. It is not an engineering approval of the physical direction mapping.

## CI evidence
- Local CI-equivalent test and compile gates passed.
- GitHub Actions / PR CI: pending push and PR creation; no claim is made until the remote run completes.

## Unverified claims
- Manual browser UAT of the complete real-workbook Single and Batch flows has not been completed.
- Physical-direction correctness for the three protected raw workbooks has not been re-approved in this UX-1 change because their approved Mapping Preset is not present locally.
- Manual Excel visual inspection of COM-native charts/template preservation is still required even though the COM export returned a valid workbook.

## Known limitations / follow-up
- UX-2 sheet-centric physical Mapping redesign, UX-3 explicit Peak confirmation redesign, UX-4 export simplification, and UX-5 Batch exception-first review remain out of scope.
- The application still contains compatibility definitions for the former status/shell helpers, but they are no longer rendered as competing primary workflow UI.
- Final review status: **READY FOR CODE REVIEW** based on automated and local runtime evidence; **NOT READY FOR FINAL RELEASE QUALIFICATION** until approved real-workbook Mapping Presets, manual browser UAT, visual Excel inspection, and PR CI are completed.

