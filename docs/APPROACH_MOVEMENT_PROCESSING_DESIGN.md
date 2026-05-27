# Approach-Movement Processing Design

## Scope

This document records the design path and release-current behavior for dual-scheme processing. Earlier phase sections are historical notes; the release-current source of truth is:

- `from_to` and `approach_movement` are both supported movement-code schemes.
- `movement_code_scheme` is the source of truth for validation, UI labels, processing path selection, and export routing.
- Existing mapping files without `movement_code_scheme` metadata default to `from_to`.
- `approach_movement` single-file processing supports Excel Template Mode through the Excel COM/native path when COM is available, and Safe PNG/generated export.
- `approach_movement` Batch supports Safe PNG/generated ZIP export.
- `approach_movement` Batch Excel Template Mode remains intentionally unsupported/blocked in this version.
- Single-file workflow is recommended for report generation and template-output review because it gives the most complete mapping, peak, layout, and export review loop.

## Phase L0 Dual-Scheme Consistency Rule

`movement_code_scheme` is the source of truth for mapping validation, UI labels, processing path selection, and export routing.

- `from_to` remains the compatibility scheme for existing mapping presets, Mapping Excel files, project sessions, v1 processing, and v1 report exports. Old mapping files that do not declare `movement_code_scheme` load as `from_to`.
- `approach_movement` uses travel direction plus movement type. `N`, `S`, `E`, and `W` mean travel direction; `L`, `T`, `R`, and `U` mean movement type. For example, `NT` means Northbound Through / มุ่งเหนือ-ตรง, not a north-leg-to-target-leg relationship.
- `approach_movement` mappings may only use `NL`, `NT`, `NR`, `NU`, `SL`, `ST`, `SR`, `SU`, `EL`, `ET`, `ER`, `EU`, `WL`, `WT`, `WR`, and `WU`.
- Codes such as `NE`, `NS`, `WE`, and `EN` can remain valid under `from_to`, but are invalid under `approach_movement`.
- The app must never silently convert one scheme into the other.

Routing by scheme:

- `from_to` uses the v1 processing/export/template path with `templates/four_leg_tmc_report_template.xlsx` and `templates/four_leg_tmc_report_template_map.json`.
- `approach_movement` uses the v2 processing/export/template paths with `templates/four_leg_tmc_report_template_approach_v2.xlsx` and `templates/four_leg_tmc_report_template_approach_v2_map.json`.
- `approach_movement` Safe PNG/generated mode uses the v2 generated workbook/package path.
- `approach_movement` single-file Excel Template Mode is native Excel COM-only.
- `approach_movement` Batch Safe PNG/generated ZIP export is supported.
- `approach_movement` Batch Excel Template Mode remains blocked until explicit support is added.

The Mapping workspace exposes a `ระบบรหัส Movement` selector before mapping rows are created. Once rows exist, changing scheme is blocked by workflow behavior rather than reinterpreting the existing rows. Loading a Mapping Preset or Mapping Excel with declared scheme metadata updates the active scheme from that file. Files without scheme metadata default to `from_to`.

## Phase F Implementation Note

Phase F adds a dedicated internal dry-run entry point, `process_tmc_dry_run_v2()`, for `approach_movement` normalization and summaries only. The normal `process_tmc()` application/report path remains guarded for v2 and still raises before processing.

Supported in the Phase F dry-run:

- Load and validate v2 Mapping Preset / Mapping Excel inputs with `movement_code_scheme = "approach_movement"`.
- Normalize records with canonical `movement_code` and compatibility alias `output_movement_code` carrying v2 values such as `NT`, `WT`, and `ET`.
- Preserve v2 normalized metadata columns: `movement_code_scheme`, `approach_direction`, and `movement_type`.
- Produce hourly totals, movement summary, vehicle composition, QC rows, and AM/PM peak suggestions using the existing peak method.
- Produce hourly movement PCU summaries ordered exactly by `APPROACH_MOVEMENT_CODES`.

Still unsupported after Phase F:

- Final Excel workbook/report export for v2.
- Excel Template Mode for v2.
- v2 movement diagram export.
- Batch v2 processing/export.
- Any silent conversion between v2 approach-movement codes and v1 from-to codes.

Remaining blockers before v2 export/report:

- Generated workbook v2 support with explicit movement-code scheme metadata.
- A verified v2 report layout that does not use v1 diagram assumptions.
- Either a verified v2 diagram or an enforced no-diagram report path.
- UI enablement that clearly labels v2 as experimental and disables unsupported export modes.
- Batch support, if desired, with explicit v2 dry-run or export tests.

Phase F test coverage includes v2 preset/XLSX loading, dry-run normalization, v2 movement-code validation, hourly movement ordering, AM/PM peak suggestions, export blocking, mixed-code rejection, and batch v2 blocking. The existing v1 smoke/export tests remain unchanged.

## Phase G Implementation Note

Phase G adds generated workbook export for v2 dry-run results through `export_v2_generated_workbook()`. This is a generated-only / Safe PNG-style export path: it writes workbook data sheets directly and does not load the v1 template map, the v1 template workbook, Excel COM/native charts, or the v1 movement diagram path.

Implemented in Phase G:

- `Export_Metadata` records `movement_code_scheme = "approach_movement"`, `template_version = "generated_approach_movement_v2"`, `export_template = "generated_approach_movement_v2"`, `export_mode_used`, and limitation notes.
- The generated workbook includes `PCE_Factors`, `Normalized_Data`, `Hourly_Totals`, `Hourly_Movement_PCU`, `Movement_Summary`, `Vehicle_Composition`, `Peak_Summary`, `QC_Check`, `Movement_Code_Reference`, and `Mapping_Scheme_Info`; it also includes `Mapping` when a mapping table is supplied.
- `Hourly_Movement_PCU` follows `APPROACH_MOVEMENT_CODES` order exactly: `NL`, `NT`, `NR`, `NU`, `SL`, `ST`, `SR`, `SU`, `EL`, `ET`, `ER`, `EU`, `WL`, `WT`, `WR`, `WU`.
- `Movement_Code_Reference` enumerates all 16 approach-movement codes with direction, movement type, readable labels, and display labels.
- A v2 generated package helper, `create_v2_generated_export_package_zip()`, can package the generated workbook and `export_summary.txt` without including raw input Excel files.

Still unsupported after Phase G:

- Normal `process_tmc()` v2 processing/export remains blocked.
- `export_workbook()` still rejects `approach_movement` v2 input and therefore keeps Excel Template Mode and Excel COM/native template export blocked for v2.
- v2 diagram export remains unsupported/omitted.
- UI enablement remains deferred to Phase J so the application does not expose a broad additional v2 export workflow before the unsupported modes are clearly handled.
- Batch v2 processing/export remains blocked.

Remaining blockers before UI enablement/full release:

- Decide whether v2 single-file processing should call the dry-run plus generated export path directly or gain a separate explicit processing entry point.
- Add or formally omit v2 diagram support with user-facing messaging.
- Finalize and verify a real v2 Excel template before enabling Excel Template Mode.
- Add UI controls that only allow the generated workbook/Safe PNG-style v2 path and keep all template/native paths disabled.
- Decide whether batch v2 should support generated exports or remain dry-run only.

The current production scheme is `from_to` v1. Examples are `NS`, `WE`, and `EN`, where the code represents from leg to destination leg. The target v2 scheme is `approach_movement`. Examples are `NL`, `NT`, `NR`, and `NU`, where the first token is travel direction and the second token is movement type:

- `N`, `S`, `E`, `W`: northbound / มุ่งเหนือ, southbound / มุ่งใต้, eastbound / มุ่งตะวันออก, westbound / มุ่งตะวันตก.
- `L`, `T`, `R`, `U`: left, through, right, U-turn.

## Design Decision

### 1. Canonical movement key

Recommendation: use both `movement_code` and `output_movement_code`, with a clear rule.

- `movement_code` is the canonical internal movement key for the active movement scheme.
- `output_movement_code` remains a backward-compatible export and serialized-file alias.
- For v1 rows, both fields should normally contain v1 codes such as `NS`, `WE`, and `EN`.
- For v2 rows, both fields should normally contain v2 codes such as `NL`, `NT`, and `NR`.
- Consumers that need to group, pivot, detect peaks, or produce summaries should use `movement_code`.
- Loaders may continue accepting existing files that only have `output_movement_code`, but should copy it into `movement_code` during cleaning.

This fits the current code better than creating a third canonical field. `normalizer.py` already emits both fields, and `summaries.py` already prefers `movement_code` after compatibility handling. Keeping `output_movement_code` as an alias avoids breaking existing exports, project sessions, mapping presets, and tests.

### 2. Pipeline model

Options evaluated:

- A. Branch inside the existing pipeline based on `movement_code_scheme`.
- B. Normalize v1 and v2 into a common `MovementRecord` model.
- C. Convert v2 approach-movement codes to v1 from-to codes internally.
- D. Convert v1 from-to codes to v2 approach-movement codes internally.

Recommendation: use option B, with a very small scheme branch at the pipeline boundary.

The safe model is:

1. `process_tmc` reads and validates `movement_code_scheme`.
2. A scheme-aware mapping/normalization layer produces a common normalized record shape.
3. Downstream summaries and exports read canonical `movement_code`, `movement_code_scheme`, and optional scheme-specific metadata.
4. Export/report/diagram support is enabled only for scheme and mode combinations that have explicit tests.

Do not silently convert v2 to v1. The relationship between approach direction, road leg, left/right, U-turn, and local geometry is not proven enough to reinterpret `NT` as a destination-leg code. Silent conversion is the highest-risk option because it can produce plausible but wrong reports.

Do not silently convert v1 to v2 either. `NS` can be a from-to movement but is not a valid v2 approach-movement code. Any future conversion must be explicit, reviewable, and geometry-aware.

### 3. Movement ordering

v1 keeps the existing order:

`NE`, `NS`, `NW`, `NU`, `SW`, `SN`, `SE`, `SU`, `WN`, `WE`, `WS`, `WU`, `ES`, `EW`, `EN`, `EU`

v2 uses `APPROACH_MOVEMENT_CODES`:

`NL`, `NT`, `NR`, `NU`, `SL`, `ST`, `SR`, `SU`, `EL`, `ET`, `ER`, `EU`, `WL`, `WT`, `WR`, `WU`

Future summary/export code should choose the movement order from `movement_code_scheme`. Unknown compatibility extras may still be appended after the standard list for v1 compatibility, but v2 should reject invalid approach-movement codes before processing.

## Summary Table Behavior Under v2

All v2 tables should keep their current calculation methods. Only movement identity and ordering change.

- Normalized data: include `movement_code_scheme`, canonical `movement_code`, alias `output_movement_code`, and v2 metadata `approach_direction` and `movement_type`. Existing v1 columns such as `from_leg`, `to_leg`, and `turn_type` may remain blank or contextual for v2, but must not be inferred from the v2 code unless explicitly mapped.
- Hourly movement PCU: pivot by canonical `movement_code`; columns follow `APPROACH_MOVEMENT_CODES`; `Total` remains the row sum.
- Peak movement summary: peak selection still uses included records and PCU totals. Any movement-level breakdown must use v2 codes directly.
- PHF data: no formula change; respect `include_in_peak` and aggregate v2 rows by time interval.
- Vehicle composition: no movement-code dependency; should behave identically for v1 and v2.
- QC rows: include the scheme where useful. Invalid v2 codes, mismatched `approach_direction` / `movement_type`, and mixed schemes should be errors.
- Batch summary: record `movement_code_scheme` in metadata and ideally in `Batch_Summary`.

## Diagram Behavior

The current diagram cannot be assumed correct for v2 directly.

Current assumptions:

- `diagram.py` hard-codes v1 `MOVEMENT_CODES`.
- `APPROACH_TABLES` use from-to codes, for example north table `NU`, `NW`, `NS`, `NE`.
- The generated arrows and table labels are visually arranged around a four-leg intersection, but the code semantics are still v1 from-to.
- `exporter.py` and `report_template.py` build `Diagram_Data` and formulas around the same v1 movement codes.

What must change:

- Add a scheme-aware movement list and table layout.
- Define whether each v2 table is grouped by travel direction or physical approach leg.
- Rework labels and arrows so `N` means northbound / มุ่งเหนือ, not "north leg" / ขาทิศเหนือ.
- Verify left/right orientation for every approach.
- Verify U-turn placement and totals.

Phase H implementation: v2 diagram support is table-based. `build_v2_movement_diagram_data()` emits one row for each `APPROACH_MOVEMENT_CODES` value in deterministic order (`NL`, `NT`, `NR`, `NU`, `SL`, ... `WU`). The sheet is named `Movement_Diagram_Data` and includes `approach_direction`, `approach_direction_label`, `movement_type`, `movement_type_label`, `display_label`, total count/PCU, AM/PM peak PCU when available, and diagram ordering fields.

Semantic assumptions:

- `N`, `S`, `E`, and `W` mean travel direction: northbound, southbound, eastbound, and westbound.
- `L`, `T`, `R`, and `U` mean left, through, right, and U-turn.
- `NL` is labeled "Northbound Left turn" and is not interpreted as "from north leg".
- No v2 diagram code converts or aliases v2 codes to v1 `from_to` codes.

Phase H2 implementation: generated v2 export packages also include `diagram/movement_diagram.png`. The PNG is rendered from `Movement_Diagram_Data` with a deterministic 2x2 table-style visual: Northbound, Southbound, Eastbound, and Westbound groups, each with L/T/R/U tiles showing movement code, movement label, total PCU, and count. It uses `matplotlib` with the headless `Agg` backend and adds no additional dependency.

Current limitation: Phase H/H2 does not add a v2 PNG arrow diagram or enable the v1 `Diagram_Data` / `Diagram` sheet path for v2. The existing v1 PNG and generated/template diagram behavior remains unchanged. A visual arrow diagram should only be added after a tested coordinate model verifies left/right/U-turn placement for every travel direction.

## Export and Report Behavior

Relevant paths:

- Safe PNG/generated workbook: `exporter.py` writes data sheets, generated `TMC_Report`, PNG charts, and static PNG diagram.
- Excel Template Mode: `exporter.py` loads template resources and `report_template.py` writes into a template workbook/map.
- Draft v2 map: `templates/four_leg_tmc_report_template_approach_v2_map.json`.

Recommendation: v2 export should first support Safe PNG/generated workbook only, with diagram disabled or explicitly replaced by a simple movement table. Excel Template Mode should remain disabled for v2 until a v2 `.xlsx` template is finalized and verified.

Rationale:

- Generated workbook output is easier to parameterize and test.
- The current native template workbook is v1 and contains drawings/charts that can be fragile under `openpyxl`.
- The draft v2 map alone is not enough; the matching workbook must have verified headers, formulas, drawings, and chart sources.
- `report_template.py` formulas and `Diagram_Data` ranges currently assume 16 v1 rows and v1 movement headers.

## UI Behavior

Later UI enablement should:

- Auto-detect `movement_code_scheme` from Mapping Preset or Mapping Excel metadata.
- Show clear scheme status, for example `from_to / process-compatible` or `approach_movement / experimental`.
- Keep v2 processing blocked until the relevant pipeline/export phase is complete.
- Block unsupported export modes by scheme. For example, Excel Template Mode should be disabled for v2 until Phase I.
- Prevent mixed v1/v2 mapping rows in one active mapping.
- Validate that v2 rows preserve `approach_direction` and `movement_type`.
- Warn that v2 processing is experimental until release hardening is complete.
- Avoid UI wording that confuses road leg with travel direction.

## Batch Behavior

Batch v2 rules:

- All files in a batch run must use the same `movement_code_scheme`.
- The same Mapping Preset scheme applies to every file in the batch.
- v2 batch processing should remain blocked until the v2 pipeline and v2 export support are complete.
- Mixed v1/v2 schemes must fail before analysis, not during export.
- `batch_summary.xlsx` should record `movement_code_scheme` in metadata and ideally in each summary row.
- Per-file project sessions and mapping presets inside the ZIP should preserve the scheme.

## Project Session and Mapping Preset Compatibility

Project sessions:

- Save `movement_code_scheme` under the mapping/session data.
- Old sessions without the field load as `from_to`.
- v2 sessions may load safely for preview/editing, but processing remains blocked until v2 support exists.
- A v2 session should not silently change export mode to a v1-only template mode.

Mapping presets:

- Preserve top-level `movement_code_scheme`.
- Old presets without the field load as `from_to`.
- v2 presets must preserve `approach_direction` and `movement_type`.
- For v2, never infer `from_leg` or `to_leg` from the second character of `movement_code`; in v2 the second character is `L`, `T`, `R`, or `U`.

## Code Impact Map for Phase F+

### `src/tmc_processor/pipeline.py`

- Current: reads `setup["movement_code_scheme"]`, calls `mapping_processing_block_reason()`, and raises for v2.
- Future: keep the guard until a feature flag or explicit v2 dry-run path is added. Then branch only enough to call scheme-aware validation/normalization and restrict unsupported exports.
- Phase F: add a v2 dry-run path with `generate_workbook=False`, summaries only, and no diagram/template export.

### `src/tmc_processor/normalizer.py`

- Current: `clean_mapping()` is scheme-neutral, then `_append_normalized_row()` copies mapping `movement_code` into both `movement_code` and `output_movement_code`.
- Future: accept scheme-aware cleaned mapping and emit `movement_code_scheme`, `approach_direction`, and `movement_type` for v2.
- Do not change v1 row shape or v1 values.

### `src/tmc_processor/summaries.py`

- Current: `STANDARD_MOVEMENT_ORDER` is v1; `_movement_columns()` uses v1 order plus extras; pivots group by `movement_code`.
- Future: select movement order by scheme, preferably from `movement_scheme.py`.
- Phase F: add v2 order for `hourly_movement_pcu()` and movement summaries without changing formulas.

### `src/tmc_processor/diagram.py`

- Current: hard-coded v1 codes, v1 approach tables, and v1 movement arrows.
- Future: add an explicit scheme parameter. For v2, either return a supported simple table artifact or raise/skip diagram generation until orientation is verified.
- Phase H: implement verified v2 diagram support.

### `src/tmc_processor/exporter.py`

- Current: generated workbook, generated report sheet, `Diagram_Data`, charts, template mode, and diagram all assume current v1 movement layout.
- Future: carry `movement_code_scheme` into `Export_Metadata`; choose movement order by scheme; suppress v2 diagram/template paths until supported.
- Phase G: generated workbook / Safe PNG mode for v2, with v2 movement headers and no full diagram unless Phase H is done.

### `src/tmc_processor/report_template.py`

- Current: default template and map are v1; movement formulas are driven by map movement codes and `Diagram_Data`.
- Future: load v2 template resources only when both v2 workbook and v2 map are verified.
- Phase I: enable v2 Excel Template Mode after workbook integrity checks.

### `src/tmc_processor/mapping.py`

- Current: detects v2 in mapping Excel metadata, validates v2 codes, and blocks v2 processing.
- Future: make required fields scheme-aware. V1 requires `from_leg`, `to_leg`, `turn_type`; v2 requires valid `movement_code`, `approach_direction`, and `movement_type`.
- Keep `mapping_processing_block_reason()` until later phases explicitly remove or narrow it.

### `src/tmc_processor/mapping_preset.py`

- Current: preserves `movement_code_scheme`, v2 optional fields, and defaults existing presets without scheme metadata to `from_to`.
- Future: preserve v2 fields through apply/build paths and avoid v1 field inference for v2 rows.
- Add tests for mixed/invalid scheme rows before enabling processing.

### `src/tmc_processor/session.py`

- Current: stores `movement_code_scheme` in session mapping and defaults existing sessions without scheme metadata to `from_to`.
- Future: ensure v2 sessions restore UI state safely, including blocked export modes and v2 mapping metadata.

### `src/tmc_processor/batch.py`

- Current: detects preset scheme, inserts it into setup, and v2 fails through the processing guard. Batch inputs also block v2 readiness.
- Future: add explicit batch scheme validation and record scheme in `batch_summary.xlsx`.
- Phase F/G: keep v2 batch blocked until single-file v2 summaries and generated exports are verified.

### `app.py`

- Current: shows scheme status, blocks processing buttons for v2, and supports loading/previewing v2 presets/mapping Excel.
- Future: expose v2 as experimental only when the dry-run path exists; disable unsupported export modes; make scheme status visible in single-file and batch flows.

### `tests/`

- Current: v1 smoke/export coverage plus Phase B-D v2 validation and blocking tests.
- Future: add focused v2 tests before removing any guard.

### `samples/demo/`

- Current: contains v1 demo workbook/preset and v2 mapping/demo artifacts.
- Future: add a v2 raw-processing demo only when Phase F can dry-run summaries safely.

### `templates/`

- Current: v1 `.xlsx` template plus v1 map, a draft v2 map, and Phase I1 template design/integrity documentation.
- Future: create/finalize a v2 `.xlsx` template manually in Excel, then verify it read-only before enabling v2 Excel Template Mode.

## Risk Assessment

- Northbound ambiguity: `N` must mean northbound / มุ่งเหนือ, not "north leg" / ขาทิศเหนือ.
- Left/right interpretation depends on approach direction and must not be inferred from v1 from-to codes without geometry rules.
- U-turn handling can look deceptively similar between v1 and v2 because both use `U`, but the first token meaning changes.
- Existing v1 formulas and HLOOKUP headers are tied to v1 movement codes.
- Excel template drawings/charts can be damaged by `openpyxl` load/save cycles.
- v2 diagram correctness is unproven and could mislead users even if totals are numerically correct.
- Old mapping presets could be silently misused if the scheme is hidden or ignored.
- Mixed schemes in batch could produce partial failures or wrong exports unless blocked up front.
- Users may confuse road leg with travel direction, especially where labels say north/south/east/west.

## Test Plan Before Enabling v2 Processing

- v2 `process_tmc` dry-run succeeds with `generate_workbook=False`.
- v2 normalized data uses `NL`/`NT`/`NR` style codes in `movement_code` and `output_movement_code`.
- v2 normalized data includes correct `movement_code_scheme`, `approach_direction`, and `movement_type`.
- v2 summary columns follow `APPROACH_MOVEMENT_CODES` order.
- v2 peak calculation uses v2 codes correctly and does not reinterpret them as from-to codes.
- v2 export workbook contains v2 movement headers in `Hourly_Movement_PCU`.
- v2 generated workbook records `movement_code_scheme`.
- v2 batch blocks mixed schemes before analysis.
- v2 batch summary records `movement_code_scheme`.
- v1 `scripts/smoke_demo.py` output remains unchanged.
- v1 export workbook headers, diagram data, template mode, and batch behavior remain unchanged.

## Recommended Implementation Phases

### Phase F: v2 dry-run summaries

Add an internal v2 processing dry-run path and summaries only. No Excel Template Mode. No diagram. Keep v1 unchanged and keep normal user-facing v2 processing blocked unless a test-only or explicit dry-run route is used.

### Phase G: v2 generated workbook support

Implemented generated workbook / Safe PNG-style export support for v2 dry-run results. Use v2 movement headers and metadata. Excel native/template export remains disabled for v2 because the workbook template is not verified. The full diagram remains disabled unless Phase H is done.

### Phase H: v2 diagram support

Implemented table-based v2 movement visualization support. Generated v2 workbooks include `Movement_Diagram_Data`, and generated v2 ZIP packages include `diagram/movement_diagram_data.csv`. This is intentionally not a PNG arrow diagram; it preserves correct approach-movement semantics while full v2 template/native/batch workflows remain blocked.

### Phase H2: v2 visual diagram PNG

Implemented a simple package PNG at `diagram/movement_diagram.png`. The visual is a quadrant/table-style movement tile layout rather than an intersection arrow drawing, so it does not imply v1 `from_to` geometry. Excel Template Mode, native template export, and v2 batch processing remain blocked.

### Phase I: v2 Excel template finalization

Create or finalize the v2 Excel template manually or programmatically after verifying workbook integrity, formulas, charts, drawing preservation, and template map alignment.

### Phase I1: v2 template design and integrity tests

Implemented the v2 template design document and map integrity tests. The v2 map now declares `four_leg_approach_movement_v2`, uses `APPROACH_MOVEMENT_CODES` order across hourly, diagram, and approach-table headers, and excludes v1-only movement codes from v2 movement areas. A read-only template integrity helper verifies workbook/map alignment without comparing fragile chart or drawing binaries.

No v2 `.xlsx` draft was created. The current v1 template contains native charts and many shape/drawing elements, and `openpyxl` warns those drawings may be lost on save. v2 template XLSX finalization is deferred to manual Excel editing. v2 Excel Template Mode and Excel COM/native template export remain blocked; v2 generated workbook export remains the supported export artifact for v2 dry-run results.

Next step: Phase I2 can implement v2 template export only after a manually finalized `four_leg_tmc_report_template_approach_v2.xlsx` passes the integrity checks.

### Phase I2C: v2 openpyxl structural template helper

Implemented a limited v2 openpyxl structural helper through `export_v2_template_workbook()`. This helper is not visual/native-template-preserving and must not be treated as UI Excel Template Mode because openpyxl can drop or damage Excel-authored drawings, charts, shapes, lines, and arrows when saving. The helper always selects the validated v2 resources:

- Workbook: `templates/four_leg_tmc_report_template_approach_v2.xlsx`
- Map: `templates/four_leg_tmc_report_template_approach_v2_map.json`

The helper writes to an in-memory output workbook copy and does not overwrite or rewrite the source template file. It is useful for structural validation of mapped cells, v2 headers, and formulas only. The `Summary` sheet keeps the v2 movement headers in `W9:AL9` using `APPROACH_MOVEMENT_CODES`, writes mapped setup metadata, populates the mapped hourly movement and vehicle-class tables, preserves template formulas such as `HLOOKUP(...,$W$9:$AL$22,...)`, and adds a `Movement_Diagram_Data` sheet for the table-based v2 diagram data.

Native Excel COM template export remains pending for v2. UI Excel Template Mode and v2 batch export are blocked until a native Excel/COM preservation path is implemented. The normal v1 `from_to` Excel Template Mode continues to use `templates/four_leg_tmc_report_template.xlsx` and `templates/four_leg_tmc_report_template_map.json`.

Remaining blockers before full UI/batch parity:

- Decide which user-facing v2 export modes should be exposed and how to label the experimental support.
- Add v2 batch processing/export design and mixed-scheme safeguards for batch inputs.
- Verify native Excel chart/drawing behavior manually before any v2 COM/native path is considered.
- Keep validating that `N`, `S`, `E`, and `W` mean travel direction, not source leg.

### Phase I3: v2 native Excel COM Template Mode

Implemented a native Excel COM-only v2 Template Mode path through `export_v2_template_workbook_com()`. This path is intentionally separate from the openpyxl structural helper. It selects only:

- Workbook: `templates/four_leg_tmc_report_template_approach_v2.xlsx`
- Map: `templates/four_leg_tmc_report_template_approach_v2_map.json`

The COM path copies the manually authored v2 workbook to a temporary output file, opens that copy in Microsoft Excel, writes mapped metadata, labels, hourly movement values, support sheets, and `Movement_Diagram_Data`, triggers Excel recalculation, saves the copy, and returns the saved workbook bytes. The source v2 template workbook is not rewritten.

UI routing for v2 is now COM-or-block:

- v2 Excel Template Mode uses native Excel COM when COM is available.
- If COM is unavailable, the UI keeps the clear Thai blocking message and does not fall back to openpyxl template save.
- v2 Safe PNG Export Mode continues to use the generated workbook/package path.
- The openpyxl v2 template helper remains an internal structural validation helper, not visual Template Mode.

Known limitation: automated tests verify cells, metadata, movement headers, template/map selection, and support-sheet data without requiring Excel COM in CI. Native chart cache refresh and visual preservation of chart/drawing objects require a local Windows Excel smoke inspection. v2 batch support remains blocked.

### Phase J: limited Streamlit UI enablement

Implemented limited UI enablement for `approach_movement` v2 in the single-file workflow only.

Supported UI scope:

- Single-file v2 processing is enabled when the loaded mapping declares `movement_code_scheme = "approach_movement"` and validates successfully.
- The UI routes v2 processing through `process_tmc_dry_run_v2()` and keeps the v2 result separate from unsafe v1 `ProcessingResult` assumptions.
- Peak Review uses the v2 hourly movement summary and shows approach-movement codes directly.
- Data Review exposes v2 QC, `Normalized_Data`, `Hourly_Movement_PCU`, `Movement_Summary`, `Movement_Diagram_Data`, and `Movement_Code_Reference`.
- Export supports v2 Generated Workbook / Safe PNG mode. v2 Excel Template Mode is blocked in the UI until a native Excel/COM preservation path is implemented. The openpyxl helper remains available only for internal structural validation with `templates/four_leg_tmc_report_template_approach_v2.xlsx` and `templates/four_leg_tmc_report_template_approach_v2_map.json`.
- The v2 export package includes the workbook, `export_summary.txt`, mapping artifacts, `diagram/movement_diagram_data.csv`, and `diagram/movement_diagram.png` when the package helper can render it. Raw input Excel files are not included.

Still blocked:

- Batch analysis/export for `approach_movement` remains blocked. Loading/detecting a v2 Mapping Preset is allowed, but Analyze Batch must not run.
- Excel Template Mode remains blocked for `approach_movement` with the UI message: `Excel Template Mode สำหรับ approach_movement ต้องใช้ Excel COM เพื่อรักษากราฟและ diagram จาก template ขณะนี้ไม่สามารถใช้โหมดนี้ได้ กรุณาใช้ Safe PNG Export Mode หรือเปิดใช้งาน Excel COM`.
- v2 is not silently converted into v1, and v2 exports must not use the v1 template or v1 template map.

## Current Recommendation Summary

Use `movement_code` as the canonical key and keep `output_movement_code` as an alias. Build toward a common normalized movement record model instead of converting between schemes. v2 single-file and v2 Batch now support generated/Safe PNG exports; v2 Batch Template Mode remains the main parity gap.

### Phase K: v2 Batch Safe PNG/generated support

Implemented `approach_movement` Batch analysis and Safe PNG/generated ZIP export.

Supported v2 Batch scope:

- Batch runs use one shared `movement_code_scheme` from the Mapping Preset/setup.
- `from_to` and `approach_movement` mapping codes are validated as separate schemes; mixed or mismatched mappings fail per file during analysis and are recorded as failed rows.
- v2 Batch analysis routes through `process_tmc_dry_run_v2()` and does not call v1 `process_tmc()`.
- Per-file suggested AM/PM peaks are produced and defaulted as confirmed peaks; user overrides are carried into final export.
- v2 hourly movement PCU uses `APPROACH_MOVEMENT_CODES` order.
- Batch QC rows and Batch Summary rows include `movement_code_scheme`.
- v2 Safe PNG/generated Batch export writes generated v2 workbooks and table-style movement diagram artifacts without raw input workbooks.

v2 Batch ZIP contents:

- `batch_summary.xlsx`
- `file_01_<safe_output_stem>/<output_stem>_report.xlsx`
- `file_01_<safe_output_stem>/<output_stem>_export_summary.txt`
- `file_01_<safe_output_stem>/<output_stem>_session.tmcproj.json`
- `file_01_<safe_output_stem>/<output_stem>.mapping.json`
- `file_01_<safe_output_stem>/diagram/movement_diagram_data.csv`
- `file_01_<safe_output_stem>/diagram/movement_diagram.png`
- `file_01_<safe_output_stem>/charts/*.png` when chart rendering succeeds

`batch_summary.xlsx` includes `metadata`, `Batch_Summary`, and `Batch_QC`. Metadata records the app version, template version, generated timestamp, mapping preset name, and `movement_code_scheme`. `Batch_Summary` records suggested/confirmed peaks, export mode requested/used, export status/error, generated report file name, and QC counts.

Excel Template Mode decision:

- v2 Batch Excel Template Mode is intentionally blocked for Phase K.
- The UI/blocker message is: `Excel Template Mode สำหรับ Batch approach_movement ยังไม่รองรับในเวอร์ชันนี้ กรุณาใช้ Safe PNG Export Mode`.
- There is no silent fallback from Template Mode to generated workbook for v2 Batch.

Remaining blockers before full parity:

- Implement and test per-file v2 COM Template Mode for Batch, if needed.
- Keep COM-dependent tests optional/skipped so CI does not require Microsoft Excel.
- Manually smoke-test native chart/drawing preservation before enabling Template Mode for v2 Batch.
