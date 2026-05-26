# Approach-Movement Processing Design

## Scope

This is the Phase E design for eventual `approach_movement` v2 processing. It does not enable v2 processing. The active processing guard must remain in place until the later phases described here are complete and tested.

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
- Loaders may continue accepting old files that only have `output_movement_code`, but should copy it into `movement_code` during cleaning.

This fits the current code better than creating a third canonical field. `normalizer.py` already emits both fields, and `summaries.py` already prefers `movement_code` after compatibility handling. Keeping `output_movement_code` as an alias avoids breaking old exports, project sessions, mapping presets, and tests.

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

Future summary/export code should choose the movement order from `movement_code_scheme`. Unknown legacy extras may still be appended after the standard list for v1 compatibility, but v2 should reject invalid approach-movement codes before processing.

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

Recommendation: disable the full movement diagram for v2 initially. Phase F should use v2 summaries only; Phase G can export generated workbook tables; Phase H should add either a verified v2 diagram or an explicit limitation. Before diagram support, a simpler v2 movement summary table is safer than a visually confident but semantically unproven diagram.

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

- Current: preserves `movement_code_scheme`, v2 optional fields, and defaults old presets to `from_to`.
- Future: preserve v2 fields through apply/build paths and avoid v1 field inference for v2 rows.
- Add tests for mixed/invalid scheme rows before enabling processing.

### `src/tmc_processor/session.py`

- Current: stores `movement_code_scheme` in session mapping and defaults old sessions to `from_to`.
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

- Current: v1 `.xlsx` template plus v1 map, and draft v2 map only.
- Future: create/finalize a v2 `.xlsx` template after generated workbook v2 support is stable.

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

Add v2 generated workbook / Safe PNG export support. Use v2 movement headers and metadata. Keep Excel native template disabled for v2 if the workbook template is not verified. Keep the full diagram disabled unless Phase H is done.

### Phase H: v2 diagram support or limitation

Either add verified v2 diagram support with visual tests/manual review, or document and enforce an explicit v2 diagram limitation.

### Phase I: v2 Excel template finalization

Create or finalize the v2 Excel template manually or programmatically after verifying workbook integrity, formulas, charts, drawing preservation, and template map alignment.

### Phase J: UI enablement and release notes

Enable v2 processing in the UI only for supported paths. Add release notes, experimental warnings if applicable, and compatibility guidance for v1 sessions/presets.

## Current Recommendation Summary

Use `movement_code` as the canonical key and keep `output_movement_code` as an alias. Build toward a common normalized movement record model instead of converting between schemes. Keep v2 processing blocked until Phase F+ tests prove the v2 path. Support v2 summaries first, generated workbook second, diagram/template last.
