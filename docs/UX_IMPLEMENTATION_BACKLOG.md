# TMC Processor UX Modernization Backlog

Baseline: `793c12c688faa0ab7bd89e31e5929fb8c4171370`  
Design authority: `docs/UX_UAT_ARCHITECTURE_REVIEW.md`

This backlog intentionally separates UX/state modernization from engineering-method changes. Existing validated calculation behavior is a constraint, not a target for redesign.

## Priority model

- **P0** — required before visual redesign can be considered trustworthy.
- **P1** — major operator workflow improvement.
- **P2** — scalability/maintainability improvement after the primary workflow is stable.

## Phase UX-0 — Workflow/state contract harness

Priority: **P0**

### Objective

Create testable contracts for workflow readiness, state transitions, and invalidation before moving UI structure.

### Required work

- Define one application-level workflow state for Single and Batch.
- Define fingerprints/revisions for source input, mapping, analysis configuration, review decision, and export configuration.
- Centralize invalidation rules.
- Add golden workflow tests that do not depend on visual layout.

### Acceptance criteria

- Raw workbook change invalidates Mapping-derived analysis, Peak review, and Export.
- Mapping change invalidates Analysis, Peak review, and Export.
- PCE or Peak-window change invalidates Analysis, Peak review, and Export.
- Confirmed Peak change does not recalculate Analysis but invalidates Export.
- Report-only metadata change does not invalidate Analysis.
- View-only state changes do not invalidate engineering results.
- Existing effective-Peak export regression tests continue to pass.
- Existing 257-test baseline must not regress; any changed count must be explained by added tests, not deleted coverage.

### Explicit non-goals

- No major visual redesign yet.
- No calculation-engine rewrite.
- No MCP work.

---

## Phase UX-1 — Canonical five-stage shell

Priority: **P0**

### Objective

Replace competing navigation/progress models with one canonical flow:

`Data → Mapping → Analyze → Review → Export`

### Required work

- Use one stage navigation as both navigation and workflow position.
- Remove or demote duplicated stepper/status representations that repeat the same workflow.
- Simplify top header/status area.
- Simplify sidebar to workspace-level controls.
- Move Upload into Data.
- Move Batch Mapping Preset action into Mapping.
- Add a dedicated Analyze stage in both modes.

### Acceptance criteria

- Single and Batch display the same five conceptual stages.
- No `Process` step exists without a corresponding place where the action is performed.
- `Analyze TMC` exists in Analyze, not Mapping.
- `Analyze Batch` exists in Analyze, not Review.
- Required current-stage actions are not hidden in sidebar.
- Operator can complete a normal Single journey without opening Advanced diagnostics.
- Navigation state changes do not invalidate engineering state.

---

## Phase UX-2 — Mapping workspace redesign

Priority: **P1**

### Objective

Make Mapping sheet-centric and physical-semantics-first while preserving all existing mapping capabilities.

### Required work

- Treat detected raw sheet/source as the anchor.
- Basic mode asks for physical direction/movement semantics.
- Derive canonical movement code and derived legs/turn type where deterministic.
- Show row-level status/validation near the affected row.
- Keep Mapping Preset reuse available.
- Move scheme/source-stream/aggregation/legacy fields to Advanced unless essential.
- Remove redundant full-table derived-summary presentation when equivalent status is available inline.

### Acceptance criteria

- Valid Basic mapping can be completed without opening Advanced.
- Existing valid movement codes still derive `from_leg`, `to_leg`, and turn/movement type correctly.
- Invalid/ambiguous rows identify the specific row and corrective action.
- Loading legacy Mapping Excel/Preset remains compatible.
- Changing/clearing Mapping triggers the workflow invalidation contract from UX-0.
- Real Kabin Buri, Bo Phloi, and Nong Prue samples remain processable after equivalent mappings are entered.

---

## Phase UX-3 — Explicit engineering review and Peak confirmation

Priority: **P0**

### Objective

Separate system suggestion from human engineering approval.

### Required work

- Represent suggested AM/PM Peaks separately from selected Peaks.
- Add explicit `Confirm Peak Review` action.
- Persist confirmation provenance.
- Put QC, hourly PCU, movement summary, PHF, and diagram in Review.
- Move parser/template/COM internals under Advanced Audit/Diagnostics.

### Acceptance criteria

- A default selectbox value alone does not mark Peak review as human-confirmed.
- Export readiness follows the chosen review policy and is explicit in tests.
- Modified Peak followed by confirmation is propagated consistently to all export artifacts.
- Changing a confirmed Peak invalidates stale export output and returns Export to ready-to-generate state.
- QC errors/warnings are visible without opening developer diagnostics.

---

## Phase UX-4 — Report-oriented Export

Priority: **P1**

### Objective

Make the operator choose the report outcome, not the implementation backend.

### Required work

- Default to `Standard report — Recommended`.
- Automatically use native Excel template mode when available/compatible.
- Automatically use the validated fallback when necessary.
- Keep explicit backend choice under Advanced export options.
- Present final provenance/readiness compactly before generation.

### Acceptance criteria

- A normal operator can export without understanding Excel COM.
- The actual backend used is still recorded in export metadata.
- No unsupported template path is silently attempted.
- Fallback is communicated clearly when used.
- Existing Excel Template and Safe PNG output contracts remain valid.

---

## Phase UX-5 — Batch exception-first review

Priority: **P1**

### Objective

Make Batch practical for larger file sets.

### Required work

- Add review queue/table.
- Show file-level QC severity, suggested Peaks, and review state.
- Add `Needs review` filtering.
- Support bulk acceptance of suggested Peaks for clean files.
- Preserve per-file explicit review provenance.
- Keep failed files visible and explain their disposition.

### Acceptance criteria

- Clean files can be confirmed in bulk without opening each individually.
- Files with warnings/errors remain excluded from bulk acceptance unless the rule is deliberately overridden and audited.
- Individual file detail remains available for chart/Peak review.
- Batch export reports which files succeeded, failed, or were excluded.
- Any edit to a confirmed file Peak invalidates only the appropriate downstream Batch export state, not unrelated completed analysis.

---

## Phase UX-6 — Split UI orchestration from `app.py`

Priority: **P2**

### Objective

Reduce regression risk and prepare a stable application service boundary.

### Required work

- Extract workflow modules for Data, Mapping, Analyze, Review, Export.
- Extract reusable status/QC/Peak/Mapping components.
- Move readiness/invalidation/orchestration to application-level modules.
- Keep `src/tmc_processor` domain/engine behavior unchanged unless a separate defect is proven.

### Acceptance criteria

- `app.py` becomes an application shell/composition root rather than the owner of all workflow logic.
- State-transition logic can be tested without rendering Streamlit pages.
- UI components do not directly encode calculation methodology.
- Existing test suite and real-workbook qualification remain green.

---

## Phase UX-7 — Workflow qualification and release gate

Priority: **P0 for release**

### Objective

Prove the UX modernization did not change validated engineering behavior.

### Automated gates

- Full pytest suite.
- Golden Single workflow.
- Golden Batch workflow.
- State invalidation matrix.
- Effective Peak export consistency.
- Session/Preset compatibility tests.

### Local/manual Windows gates

- All three real sample workbooks end-to-end.
- Excel COM native-template export when Excel is available.
- Safe PNG fallback.
- Verify Summary, Export_Metadata, Peak_PHF, Diagram_Data/equivalent outputs use the effective confirmed Peak.
- Verify changing Mapping/PCE/Peak creates no stale-output presentation.

### Release decision

Do not close UX modernization solely on automated test success. Real workbook + Excel qualification remains mandatory.

---

# Codex implementation guardrails

When implementation begins, Codex should be instructed to:

1. Start from synchronized `main` and create an isolated branch/worktree.
2. Read `docs/UX_UAT_ARCHITECTURE_REVIEW.md` and this backlog before editing.
3. Implement one phase at a time; do not collapse UX-0 through UX-7 into a single unreviewable rewrite.
4. Preserve validated engine outputs and public data formats unless a separately documented migration is approved.
5. Add tests before or with each state/workflow behavior change.
6. Do not weaken validation to make UX tests pass.
7. Run the full suite after every phase and report exact counts.
8. Run real-workbook/Excel COM gates before declaring release readiness.
9. Keep MCP out of the implementation until application service contracts are stable.
10. Report residual limitations explicitly.
