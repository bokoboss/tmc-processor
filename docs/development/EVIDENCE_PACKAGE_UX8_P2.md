# UX-8 P2 evidence package — workflow guidance and workspace cleanup

## Starting checkpoint and scope

- Accepted P1 checkpoint: `2d447825a2b0b6669bcac6aa40887f111c1aa8c1` on `codex/ux-8-operator-ui-cleanup`, with a clean working tree at the start of P2.
- P1 parent / accepted UX-7 baseline: `67c645baf1323b3d8716079574504c2a167dce61`.
- P2 is uncommitted. No push, pull request, or merge was performed.
- P2 files: `src/tmc_processor/ui/app_shell.py`, `src/tmc_processor/ui/workflows/single.py`, `src/tmc_processor/ui/workflows/batch.py`, `tests/test_setup_persistence.py`, `tests/test_ux8_p1_ui_hierarchy.py`, `tests/test_ux8_p2_workflow_guidance.py`, and this evidence file.

## Workspace and workflow presentation

- The sidebar **โครงการ** section presents **เปิด Project Session** and **บันทึก Project Session** on every Single stage. It reuses the existing session build, parse, apply, and download paths. The duplicate Project Session download in Single Export was removed; the existing report package still includes the session file.
- The Project Session schema stores one source and one mapping. At the operator's direction, Batch shows the same controls disabled with a clear scope note. Batch files and its Mapping Preset therefore cannot be silently saved as a partial Project Session.
- A compact textual status line covers Data → Mapping → Analyze → Review → Export. It reads `WorkflowState.readiness` and existing stale flags; it does not persist its labels or introduce a workflow state machine. Data follows source readiness; Mapping follows source and mapping readiness; Analyze follows current analysis and its existing stale flag; Review follows current analysis and review readiness; Export follows export readiness and existing Batch export stale/signature state. Before the first Batch ZIP, a stale flag alone displays **พร้อมสร้าง**, not **ต้องสร้างใหม่**.
- Each of Data, Mapping, Analyze, and Review shows at most one forward navigation CTA when its existing readiness allows progression. A blocked stage gets a short reason instead. The CTA only calls the existing active-stage setter and reruns the page. Analyze, Peak confirmation, and Export remain explicit domain actions.
- Stage headings and common action guidance use Thai-first wording, retaining TMC, Mapping, Peak, PCE, QC, Project Session, Standard, and Safe PNG where useful. Redundant success/readiness notices were removed from Single Analyze/Review and Batch Mapping/Review/Export. Actionable blockers, QC warnings, provenance, and technical details remain.

## State, session, and domain boundaries

- No edits were made to `workflow_state.py`, session serialization/schema, Mapping Preset schema, source parsing, calculation, Peak/PCE/QC, Batch processing, exporter, template, or Excel COM modules.
- The new status tests apply existing `transition_workflow` and `readiness_after_transition` results for source, Mapping, analysis configuration (PCE and Peak search), confirmed Peak, report metadata, and export configuration changes. Analysis and Review remain current after changes that only invalidate Export, as dictated by the existing contract.
- A Single Project Session save/open round-trip test verifies current Data and Mapping presentation is reconstructed after load, without serializing UI labels or the active stage. It checks the existing `CURRENT_SCHEMA_VERSION` and that no new UI-only fields appear in the session.

## Verification

- Focused P2/P1/UX-7/workflow regression selection: **199 passed**, 3 known openpyxl DrawingML warnings.
- Full repository gate: `C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe -m pytest` → **451 passed**, 14 known openpyxl DrawingML warnings, zero failures.
- `C:\MyRD\tmc-processor-public\.venv\Scripts\python.exe -m compileall -q app.py src` → passed.
- `git diff --check` → passed (Git emitted only an LF-to-CRLF working-copy notice for an existing test file).
- With `PYTHONPATH` set to the worktree `src`, `tmc_processor.__file__` resolved to `C:\MyRD\TMC Processor\ux8-operator-ui-cleanup\src\tmc_processor\__init__.py`.

## Live operator smoke

- Streamlit was launched with the worktree `src` explicitly on `PYTHONPATH`; a real browser was used against the running app.
- **Single:** uploaded bundled `DEMO_TMC1_FourLeg.xlsx`, applied its compatible Mapping Preset, then traversed Data → Mapping → Analyze → Review → Export. Project Session controls stayed visible. Stage status and CTA advanced as readiness changed. Entering Analyze through the CTA did not run analysis. The explicit Analyze action produced a current result; explicit Peak confirmation enabled the Review CTA. Export showed the accepted Standard report controls, and the explicit Excel report action completed Export. The existing Safe PNG choice and routing were covered by the UX-7 regression tests.
- **Batch:** uploaded bundled `DEMO_TMC1_FourLeg.xlsx` and `DEMO_TMC1_FourLeg_Day2.xlsx`, applied the compatible Mapping Preset, and traversed all five stages. The Project Session controls were visibly disabled with their scope note. The Mapping CTA did not run Batch analysis. Explicit Batch analysis produced two successful files; both required individual Peak confirmation before Export became ready. Explicit ZIP generation completed Export. The saved Batch analysis context, queue, and QC presentation remained usable.
- **Safe invalidation:** after the Single export, changing only the project/report name kept Analyze and Review complete while Export returned to ready-to-create, matching existing export-only invalidation. No protected workbook was modified.
- During smoke, the first Batch ZIP was initially described as a regeneration because the pre-existing stale flag was set before any ZIP existed. The presentation was corrected to consult the existing export signature; the focused first-export test passed after that correction. The corrected wording was not repeated in a fresh live browser run.

## Limits

- Batch Project Session Save/Open remains disabled until a format that round-trips Batch uploads and its Mapping Preset exists; this scope was selected by the operator.
- Live smoke used successful bundled demo files and did not exercise failed or excluded Batch export paths. Existing Batch regression tests remained green.
