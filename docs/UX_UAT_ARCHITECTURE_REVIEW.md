# TMC Processor UX/UAT Architecture Review

Status: Proposed design baseline  
Reviewed baseline: `main` at `793c12c688faa0ab7bd89e31e5929fb8c4171370`  
Scope: Streamlit workflow, information architecture, operator UX, state semantics, batch review, export semantics, and workflow-harness implications.  
Non-goal: change validated calculation logic, movement mathematics, Excel formulas, or engineering outputs unless later implementation uncovers a defect.

## 1. Executive conclusion

The processing engine is already modular and should be preserved. The primary product risk is now the application shell around it: competing workflow models, mapping cognitive load, implicit Peak confirmation, batch review scalability, export implementation details exposed to operators, and distributed Streamlit session-state invalidation.

The recommended target is one canonical five-stage workflow shared by Single and Batch modes:

1. Data
2. Mapping
3. Analyze
4. Review
5. Export

The UI should expose operator decisions and engineering evidence. Implementation details such as Excel COM internals, parser diagnostics, legacy movement-code representation, and template formula diagnostics should remain available but move behind progressive disclosure.

## 2. Strengths to preserve

Do not rewrite these areas merely as part of UX modernization:

- Modular processing package under `src/tmc_processor/`.
- Mapping validation and normalization logic.
- PCE override traceability and stale-result protection.
- Effective Peak propagation into exported outputs.
- Excel COM and fallback export paths.
- QC generation and audit outputs.
- Project Session and Mapping Preset formats unless a migration is explicitly designed.
- Batch processing engine and export package generation.
- Existing unit/regression tests, especially effective-peak export, mapping, PCE, QC, Batch, Excel COM, and dual-scheme tests.
- Real-workbook validation as an independent qualification gate.

## 3. Primary UX findings

### F1 — Competing workflow models

The current UI exposes overlapping representations of progress through sidebar actions, a top status bar, workflow stepper, workflow summary chips, five top-level tabs, and per-page readiness/status cards.

They are not semantically identical. `ประมวลผล` appears as a workflow step but has no dedicated page; Single processing is triggered from Mapping, while Batch analysis is triggered from Peak Review. `ตรวจสอบข้อมูล` is a top-level tab but is absent from the workflow stepper.

**Impact:** operators must learn the implementation layout rather than follow one consistent task model.

**Required direction:** one canonical navigation model must represent the actual workflow. Status information may supplement that navigation but must not create a second workflow model.

### F2 — Analyze action is misplaced

Single-file Analyze is currently an action inside Mapping. Batch Analyze is currently an action inside Peak Review.

**Impact:** the same conceptual operation appears in different stages depending on mode.

**Required direction:** create a dedicated Analyze stage for both modes.

### F3 — Mapping workspace has excessive cognitive load

The Mapping page currently combines scheme selection, Mapping Excel import/export, Mapping Preset import/export, readiness metrics, Basic editor, Advanced editor, derived summaries, aggregation audit, validation messages, and the Process action.

**Required direction:** make the raw workbook sheet/source the primary anchor and ask the operator to define its physical meaning. Derive internal canonical fields where deterministic.

### F4 — Internal movement representation is too prominent

`movement_code_scheme`, source streams, aggregation fields, legacy from-to codes, and related fields are useful for traceability but are not all primary operator decisions.

**Required direction:** use progressive disclosure:

- Basic = physical interpretation and report inclusion.
- Advanced = internal code representation, aggregation/source-stream behavior, compatibility controls.

Legacy schemes remain supported but should not dominate the normal workflow.

### F5 — Peak selection and Peak review are not semantically separated

A selectbox having a default value does not prove that an engineer reviewed and accepted that value.

**Required direction:** model at least three concepts:

- suggested Peak,
- selected Peak,
- user-confirmed review.

The operator should explicitly accept suggested Peaks or confirm modified selections. Export readiness should depend on a clear review policy rather than merely the existence of selected values.

### F6 — Engineering review and developer diagnostics are mixed

Current `ตรวจสอบข้อมูล` combines engineering evidence with parser internals, template diagnostics, and Excel COM internals.

**Required direction:** engineering review belongs in the Review stage. Developer/runtime diagnostics belong under Advanced diagnostics / troubleshooting.

### F7 — Export exposes implementation choices as primary decisions

Operators are asked to understand Excel Template Mode, Safe PNG Export Mode, Excel COM availability, native charts, and fallback behavior.

**Required direction:** default to a report-oriented choice such as `Standard report — Recommended`, automatically choosing the best compatible implementation. Keep explicit export-engine choice under Advanced options.

### F8 — Batch review does not scale operationally

Current Batch flow focuses on selecting one file and reviewing it. This is acceptable for a small Batch but inefficient as file count grows.

**Required direction:** provide an exception-first review queue/table with bulk acceptance for clean files and focused review for warnings/errors.

### F9 — UI orchestration is concentrated in `app.py`

`app.py` currently owns global CSS, workflow derivation, mapping UI, Batch UI, session-state orchestration, export behavior, diagnostics, and many state invalidation rules.

**Impact:** future UX changes have elevated regression risk; stale-state defects become easier to introduce.

**Required direction:** split presentation/workflow modules and centralize application-state dependency rules without changing validated engine behavior.

## 4. Target information architecture

### Canonical five-stage navigation

Use the same stage names in Single and Batch modes.

| Stage | Single-file | Batch |
|---|---|---|
| Data | Upload workbook, inspect source, project metadata | Upload workbooks, inventory, shared/per-file metadata |
| Mapping | Map detected source sheets to physical movements | Apply/validate shared mapping against each workbook |
| Analyze | PCE, Peak search windows, Analyze TMC | Shared settings, Analyze all files |
| Review | Results, Peak decision, QC, engineering audit | Review queue, per-file Peak/QC, bulk accept clean results |
| Export | Generate report/package | Generate Batch package |

### Navigation rule

The stage navigation itself is the workflow indicator. Do not maintain a separate visual stepper that can disagree with navigation.

Each stage may show one compact state marker:

- Not started
- Needs attention
- Ready
- Complete

### Sidebar target

Keep the sidebar focused on workspace-level actions only:

- Work mode: Single / Batch
- New/Open Project Session
- Current project/file identity
- Advanced tools / diagnostics

Move stage-specific Upload and Mapping Preset actions into the corresponding stage.

## 5. Target Single-file workflow

### Stage 1 — Data

Primary tasks:

- Upload TMC workbook in the main content area.
- Show detected workbook/sheet summary immediately.
- Enter project/report metadata.
- Show parser failure only when blocking.

Primary CTA: `Continue to Mapping`.

### Stage 2 — Mapping

Primary tasks:

- Display detected raw sheets/sources.
- Assign physical direction/movement semantics.
- Validate completeness and ambiguity inline.
- Optionally import/reuse Mapping Preset.

Basic-view principle:

> Ask what the source means physically; derive canonical/internal representation where deterministic.

Suggested Basic columns:

- Source sheet
- Raw label/direction
- Physical approach/direction
- Movement
- Include in report
- Include in Peak
- Status

Advanced-only fields may include canonical movement code, legacy scheme, source stream, aggregation method, facility type, derived legs/turn type, and compatibility metadata.

Do not show a second full derived table underneath the editor when equivalent information can be rendered inline.

Primary CTA: `Continue to Analyze`.

### Stage 3 — Analyze

Primary tasks:

- PCE factors.
- Peak search windows.
- Analysis-readiness summary.
- Run calculation.

Primary CTA: `Analyze TMC`.

On success, advance to Review and preserve result provenance/fingerprint.

### Stage 4 — Review

Primary tasks:

- hourly PCU profile,
- suggested Peaks,
- Peak decision,
- QC errors/warnings,
- movement summary,
- PHF and relevant engineering evidence,
- movement diagram where applicable.

Peak decision sequence:

1. Show suggested AM/PM Peaks.
2. Allow alternative selection.
3. Require explicit `Confirm Peak Review` action.
4. Store review provenance.

Do not treat a selectbox default as proof of human review.

Advanced Audit may contain normalized data, aggregation audit, parser details, movement-code reference, template diagnostics, and Excel COM diagnostic detail.

### Stage 5 — Export

Primary task: generate final report/package.

Default operator choice: `Standard report — Recommended`.

Implementation chooses native Excel Template mode when available and compatible; otherwise uses validated fallback behavior.

Advanced export options may expose explicit export-engine selection.

Before generation, show compact provenance:

- source workbook,
- Mapping revision/fingerprint,
- PCE status,
- confirmed AM/PM Peaks,
- QC status,
- output mode that will be used.

## 6. Target Batch workflow

Batch uses the same five stages.

### Data

- Upload multiple workbooks in the main stage.
- Show file inventory.
- Configure shared metadata plus per-file date/output stem.

### Mapping

- Open/apply Mapping Preset in the main Mapping stage.
- Show matching status for every file.
- Sort files needing attention first.

### Analyze

- Shared PCE and Peak-search settings.
- `Analyze Batch` primary CTA.
- Display success/failure counts.

### Review

Use a review queue, not only a file selector.

Recommended columns:

- File
- Survey date
- Analysis status
- QC severity
- Suggested AM
- Suggested PM
- Review state

Support:

- `Accept suggested Peaks for all clean files`
- filter `Needs review`
- open an individual file for chart/detail review
- explicit per-file confirmation provenance

Failed files remain visible but do not silently disappear.

### Export

- Compact readiness summary.
- Generate Batch ZIP.
- Clearly state which files succeeded/failed and what is included.

## 7. State model and invalidation contract

State invalidation should be centralized and testable.

| Change | Analysis | Peak review | Export |
|---|---|---|---|
| Raw workbook | invalidate | invalidate | invalidate |
| Mapping | invalidate | invalidate | invalidate |
| PCE factors | invalidate | invalidate | invalidate |
| Peak search windows | invalidate | invalidate | invalidate |
| Confirmed Peak | keep | update | invalidate |
| Project/report metadata used only in export | keep | keep | invalidate |
| Export mode | keep | keep | invalidate |
| View mode / accordion / selected stage | keep | keep | keep |

Move toward explicit fingerprints/revisions for source input, mapping, analysis configuration, analysis result, review decision, and export configuration.

A UI rerun or display-only change must never invalidate engineering state.

## 8. Proposed UI architecture refactor

Do not change framework in this phase.

Suggested responsibility split:

```text
ui/
  app_shell.py
  workflows/
    data.py
    mapping.py
    analyze.py
    review.py
    export.py
  components/
    status.py
    mapping_editor.py
    peak_review.py
    qc.py
    export.py
application/
  state.py
  workflow.py
  services.py
```

Exact paths may differ; responsibility boundaries matter more than names.

### UI layer

Responsible for rendering, collecting operator input, navigation, and presenting validation/evidence.

### Application/workflow layer

Responsible for readiness, state transitions, invalidation, orchestration, provenance, and fingerprints.

### Domain/engine layer

Existing `tmc_processor` calculation/export modules remain authoritative unless explicitly refactored with regression proof.

## 9. Harness requirements

Existing tests remain. Add workflow-level contracts testing sequences of engineering decisions.

### Golden Single journey

`load → map → analyze → confirm review → export`

Verify:

- deterministic result for unchanged workbook/configuration,
- effective Peak appears consistently in all exported artifacts,
- Peak change invalidates export but not analysis,
- Mapping/PCE/Peak-window changes invalidate analysis,
- display-only UI actions do not invalidate engineering state,
- Project Session reload restores compatible configuration without silently reusing incompatible result state.

### Golden Batch journey

`load N files → apply mapping → analyze → review → export`

Verify:

- per-file matching status,
- partial failure behavior,
- clean-file bulk confirmation,
- per-file Peak provenance,
- Batch summary/QC consistency,
- stale package behavior after metadata/Peak/export changes.

### Real workbook qualification

The existing real workbook suite remains a release gate. UX refactoring must not weaken this qualification merely because automated tests pass.

## 10. UX invariants

1. Exactly one canonical workflow navigation is visible.
2. The primary CTA of each stage performs the operation implied by that stage.
3. A user never needs to know which sidebar contains a required action for the current stage.
4. A selected Peak is not represented as human-confirmed until an explicit confirmation action occurs.
5. A stale result is never displayed as current/ready.
6. Developer diagnostics never obscure or replace engineering review evidence.
7. Internal movement representation is not required in Basic mapping when it can be derived.
8. Export chooses a safe compatible implementation without requiring normal users to understand Excel COM.
9. Batch review prioritizes exceptions and supports efficient confirmation of clean files.
10. Existing validated engineering results remain numerically unchanged for unchanged inputs/configuration.

## 11. First-phase non-goals

Do not combine the initial UX refactor with:

- React/Next.js migration,
- MCP server implementation,
- redesign of calculation formulas,
- new traffic engineering methodology,
- new Excel report template unless required by a confirmed UX requirement,
- removal of legacy movement support,
- Project Session schema break without migration support.

## 12. MCP and future interfaces

MCP should be an adapter, not the architecture layer.

First stabilize application-level contracts such as:

- inspect workbook,
- validate mapping,
- analyze TMC,
- confirm Peak review,
- export report,
- analyze/export Batch.

Streamlit, CLI, automation, and a future MCP adapter can then call the same application service layer.

## 13. Recommended implementation sequence

1. Establish UX/state contract and workflow-state tests.
2. Introduce canonical five-stage navigation and simplify header/sidebar.
3. Move Analyze into its own stage for both Single and Batch.
4. Redesign Basic Mapping around physical semantics with inline validation.
5. Add explicit Peak Review confirmation/provenance.
6. Merge engineering QC/results into Review and move developer diagnostics to Advanced.
7. Simplify Export to report-oriented defaults with automatic backend selection.
8. Redesign Batch Review into an exception-first queue/bulk workflow.
9. Refactor `app.py` orchestration into workflow/components/application-state modules while preserving engine behavior.
10. Run the full automated suite plus real-workbook and Excel COM qualification before release.

## 14. Definition of done

UX modernization is complete only when:

- operators can describe the application flow as `Data → Mapping → Analyze → Review → Export`,
- the same conceptual stages apply to Single and Batch,
- every blocking condition appears in the stage where it can be resolved,
- Peak review requires an explicit engineering decision,
- Batch review is practical beyond a few files,
- stale-state rules are centralized and tested,
- diagnostics remain available without dominating the normal workflow,
- existing validated real-workbook outputs remain correct,
- automated tests and Windows/Excel qualification pass.
