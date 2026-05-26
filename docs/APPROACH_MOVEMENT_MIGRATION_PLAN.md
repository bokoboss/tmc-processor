# Approach-Movement Migration Plan

## Executive summary

The current TMC processor is built around a leg-to-leg movement-code scheme. Codes such as `NS`, `WE`, and `EN` are treated as output movement codes, where the first character is the source/from leg and the second character is the destination/to leg. The application also stores explicit `from_leg`, `to_leg`, and `turn_type` fields, but `movement_code` / `output_movement_code` is the effective grouping and reporting key across normalization, summaries, QC, diagrams, exports, templates, mapping presets, project sessions, demo files, and tests.

The target scheme should use approach-movement notation: `NL`, `NT`, `NR`, `NU`, `SL`, `ST`, `SR`, `SU`, `EL`, `ET`, `ER`, `EU`, `WL`, `WT`, `WR`, `WU`. In this convention, the first letter is travel direction / approach movement direction (`N` = northbound / มุ่งเหนือ, `S` = southbound / มุ่งใต้, `E` = eastbound / มุ่งตะวันออก, `W` = westbound / มุ่งตะวันตก), and the second letter is the turn movement (`L`, `T`, `R`, `U`).

Recommendation: implement Option C. Create a v2 approach-movement template and v2 mapping preset/session schema while preserving v1 compatibility. Do not replace the v1 scheme in place. The repo has enough hard-coded template, diagram, and test assumptions that an in-place replacement would risk silently wrong reports for old presets and sessions.

## Current leg-to-leg assumptions

- `src/tmc_processor/constants.py` defines `MOVEMENT_CODE_OPTIONS` as `NE`, `NS`, `NW`, `NU`, `SW`, `SN`, `SE`, `SU`, `WN`, `WE`, `WS`, `WU`, `ES`, `EW`, `EN`, `EU`.
- `MAPPING_COLUMNS` includes `movement_code`, `from_leg`, `to_leg`, and `turn_type`. The required mapping fields are `movement_code`, `from_leg`, `to_leg`, `turn_type`, and `facility_type`.
- `NORMALIZED_COLUMNS` persists both `movement_code` and `output_movement_code`, plus `from_leg`, `to_leg`, and `turn_type`.
- `clean_mapping()` treats `output_movement_code` as a backward-compatible alias for `movement_code`, then canonicalizes known codes against the current leg-to-leg list.
- `normalizer.py` copies mapping `movement_code` into both normalized `movement_code` and `output_movement_code`, then carries `from_leg`, `to_leg`, and `turn_type`.
- `summaries.py` sorts movement columns using `STANDARD_MOVEMENT_ORDER`, which matches the leg-to-leg constants.
- `diagram.py` uses `MOVEMENT_CODES` and `APPROACH_TABLES` with current codes, for example north approach `["NU", "NW", "NS", "NE"]`.
- `exporter.py`, `report_template.py`, and `excel_com_export.py` build `Diagram_Data`, diagram formulas, and report blocks using the same codes.
- `mapping_preset.py` defaults `movement_code_scheme` to `"from_to"` and serializes `movement_code` as `output_movement_code`.
- `session.py` stores cleaned mapping rows under schema version 1 without a movement scheme field.
- `samples/demo/DEMO_TMC1_FourLeg.mapping.json` declares `"movement_code_scheme": "from_to"` and uses `NS`, `WE`, `EW`.
- `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx` has mapping rows with `movement_code` values `NS`, `NS`, `WE`, `EW` and explicit `from_leg` / `to_leg`.
- Tests use current codes throughout, especially `NS`, `EW`, and duplicate `NS` aggregation scenarios.

## Target approach-movement convention

Target standard movement order:

`NL`, `NT`, `NR`, `NU`, `SL`, `ST`, `SR`, `SU`, `EL`, `ET`, `ER`, `EU`, `WL`, `WT`, `WR`, `WU`.

Recommended internal fields:

- `movement_scheme`: `from_to` for v1, `approach_movement` for v2.
- `approach_direction`: one of `N`, `S`, `E`, `W`; means travel direction / approach movement direction.
- `movement_type`: one of `L`, `T`, `R`, `U`; means left, through, right, U-turn.
- `movement_code`: canonical normalized code for the selected scheme.
- `output_movement_code`: retained as an export/backward-compatibility alias during migration.
- `source_direction`: raw/vendor source direction label, kept separate from approach direction.
- `source_stream`: unchanged; supports many-to-one aggregation such as mainline plus frontage.
- `raw_movement_label`: unchanged; human-readable source movement label.

Recommendation on canonical field: introduce `movement_code` as the canonical internal field and keep `output_movement_code` as a compatibility/export alias. The code already mostly uses `movement_code` internally and only falls back to `output_movement_code` for old files. Keeping that pattern avoids a broad rename while still making the canonical meaning explicit.

## Code impact map

| Area | Current assumption | Migration impact |
| --- | --- | --- |
| `app.py` | Mapping editor exposes `movement_code` under the label `output_movement_code`, plus `from_leg`, `to_leg`, and `turn_type`. It uses `MOVEMENT_CODE_OPTIONS`, `LEG_OPTIONS`, and `TURN_TYPE_OPTIONS`. Batch preset preview looks for `output_movement_code` or `movement_code`. | Add movement scheme selection/status. For v2 mappings, editor should expose `approach_direction` and `movement_type`, or derive them from `movement_code`. Keep v1 editor behavior for old sessions/presets. |
| `src/tmc_processor/constants.py` | Hard-coded leg-to-leg movement options and normalized/mapping columns. | Add scheme constants, v1 and v2 movement orders, validators, and optional v2 columns. Avoid removing v1 constants immediately. |
| `src/tmc_processor/mapping.py` | Cleans `output_movement_code` into `movement_code`; validates against current movement options; requires `from_leg`, `to_leg`, `turn_type`. | Make cleaning scheme-aware. Add v2 validation for `approach_direction`, `movement_type`, `movement_code`. For v1, retain existing required fields. |
| `src/tmc_processor/normalizer.py` | Writes mapping `movement_code` to normalized `movement_code` and `output_movement_code`; copies `from_leg`, `to_leg`, `turn_type`. | Add v2 fields to normalized output. Continue writing `output_movement_code` for consumers until all exports are updated. |
| `src/tmc_processor/summaries.py` | Uses leg-to-leg `STANDARD_MOVEMENT_ORDER`; pivots by `movement_code`. | Select movement order from `movement_scheme`. Ensure extras are still appended for legacy/unknown codes. |
| `src/tmc_processor/diagram.py` | Diagram tables and value extraction depend on current codes. Table titles say `North approach`, but codes are currently from-to. | Add v2 diagram layout using approach-movement code headers and verify orientation. Decide whether existing diagram arrows represent approach direction or destination leg before drawing v2 arrows. |
| `src/tmc_processor/exporter.py` | `Diagram_Data`, generated `TMC_Report`, `Diagram` sheet, and fallback report formulas use current codes and `Diagram_Data!$B$2:$B$17`. | Parameterize movement code list/order and report movement blocks by scheme. Consider separate v2 generated report layout. |
| `src/tmc_processor/report_template.py` | Template map approach tables provide current codes; formulas are written from those codes into `Diagram_Data` lookups. | Make template selection scheme-aware. Use v2 map with v2 codes. Do not mutate v1 map at runtime. |
| `src/tmc_processor/excel_com_export.py` | Mirrors template and diagram-data logic with `MOVEMENT_CODES` and current map. | Update in parallel with `report_template.py` / `exporter.py`; avoid a v2 path that only works in safe PNG mode. |
| `src/tmc_processor/mapping_preset.py` | Schema version 1; default `movement_code_scheme` is `from_to`; rows include `output_movement_code`, `from_leg`, `to_leg`, `turn_type`. `_rows_to_mapping_frame()` infers `from_leg` and `to_leg` from two-letter codes. | Create preset schema v2 with explicit scheme. For v2 rows, do not infer `to_leg` from the second code character because it will mean turn type. |
| `src/tmc_processor/session.py` | Project session schema version 1 stores cleaned mapping rows but no movement scheme. | Add session schema v2 with `movement_scheme` in mapping and export settings. Load v1 as `from_to`. |
| `src/tmc_processor/qc.py` | Groups and reports by `movement_code`; duplicate output movement checks assume same code means aggregation. | Mostly reusable, but messages should reference scheme and v2 validators should catch invalid combinations. |
| `src/tmc_processor/batch.py` | Builds/applies mapping presets per file and carries sessions through batch ZIPs. | Batch must reject or warn on scheme mismatches between preset, template, and loaded session. |
| `tests/` | Test fixtures use v1 codes and fields. | Keep v1 tests, then add v2 tests for cleaning, normalization, summaries, template export, batch, and preset/session compatibility. |
| `samples/demo/` | Demo preset and mapping workbook are v1/from-to. | Add approach-movement demo mapping files rather than overwriting the current demo. |

## Template impact map

Files inspected:

- `templates/four_leg_tmc_report_template.xlsx`
- `templates/four_leg_tmc_report_template_map.json`

Workbook findings:

- The workbook contains one sheet: `Summary`.
- The used area is `A1:BG218`.
- The workbook has 121 formulas.
- Movement table headers are in `W9:AL9`: `NE`, `NS`, `NW`, `NU`, `SW`, `SN`, `SE`, `SU`, `WN`, `WE`, `WS`, `WU`, `ES`, `EW`, `EN`, `EU`.
- Diagram movement header cells include:
  - North: `J13:M13` = `NU`, `NW`, `NS`, `NE`
  - South: `J30:M30` = `SW`, `SN`, `SE`, `SU`
  - West: `D19:D22` = `WN`, `WE`, `WS`, `WU`
  - East: `R21:R24` = `EU`, `EN`, `EW`, `ES`
- Diagram values depend on those headers through formulas such as `HLOOKUP(J13,$W$9:$AL$22,$U$22,FALSE)`.
- Hourly movement totals are formula-driven in row 22, for example `W22=SUM(W10:W21)` through `AM22=SUM(AM10:AM21)`.
- Summary cells depend on diagram ranges, for example `F30=SUM(G19:G22,J16:M16,O21:O24,J27:M27)`, `F31=SUM(J15:M15,F19:F22,J28:M28,P21:P24)`, and `F32=SUM(J14:M14,E19:E22,J29:M29,Q21:Q24)`.
- The map file binds the current movement codes into `movement_diagram_cells.approach_tables` and `hourly_movement_table.columns`.
- The map has schema version 2 but is still a v1 movement-code layout.

Does the template use movement codes directly in formulas?

Mostly indirectly. The formulas typically use movement-code header cells with `HLOOKUP` rather than hard-coding `NE` / `NS` inside each formula. However, the header cells and the map are movement-code-specific. Changing to approach-movement requires replacing headers, approach table code lists, and ensuring formulas still point to the correct rows/columns.

Can it be migrated programmatically with openpyxl?

Partially. Simple cell values, formulas, and the JSON map can be migrated with `openpyxl`. However, this workbook includes drawings/charts and the repo already documents caution around programmatic template edits. `openpyxl` warns that DrawingML support is incomplete and shapes/drawings may be lost when loading/saving. Programmatic migration is acceptable for generating a new candidate workbook, but it should be manually reviewed in Excel and regression-tested before becoming the shipped template.

Is a new template version safer than overwriting the existing template?

Yes. Create a v2 template and v2 map, for example:

- `templates/four_leg_tmc_report_template_v2_approach_movement.xlsx`
- `templates/four_leg_tmc_report_template_v2_approach_movement_map.json`

Keep the current files as v1/from-to compatibility assets.

## Data/schema impact

Project sessions:

- Current schema version is 1 and has no movement scheme field.
- v1 sessions should load as `movement_scheme = "from_to"`.
- v2 sessions should store `movement_scheme = "approach_movement"` and either store v2 mapping fields or a scheme-tagged mapping payload.

Mapping presets:

- Current preset schema version is 1 and defaults to `"movement_code_scheme": "from_to"`.
- v2 should make `movement_code_scheme` required and validated.
- v2 rows should include `approach_direction` and `movement_type`.
- v2 rows may retain `from_leg`, `to_leg`, and `turn_type` only as legacy/context fields. Do not infer them from v2 code characters.

Normalized data:

- Add nullable columns for `movement_scheme`, `approach_direction`, and `movement_type`.
- Continue emitting `movement_code` and `output_movement_code`.
- For v1 rows, `approach_direction` / `movement_type` can be blank or derived only if explicitly reviewed.

Export workbooks:

- Existing exported reports are static artifacts and do not need migration.
- New exports should include the movement scheme in `Export_Metadata`, `Setup`, `Mapping`, and possibly `Movement_Aggregation_Audit`.

## Migration strategy

Option A: replace the existing leg-to-leg scheme in place.

- Pros: simplest final mental model and fewer UI branches after migration.
- Cons: high risk. Old presets and sessions would silently reinterpret codes such as `NS`: currently north-to-south, but under approach-movement the second character `S` is invalid. Existing tests, demos, templates, and diagrams would all break or need simultaneous replacement.

Option B: add approach-movement as a new scheme while preserving old files through `schema_version`.

- Pros: safer load path; supports old project files; lets validators prevent accidental cross-scheme use.
- Cons: more conditional logic in mapping, summary, export, and template selection.

Option C: create v2 template and v2 mapping preset schema while keeping v1 compatibility.

- Pros: safest for this repo. It acknowledges that report layout and movement semantics are coupled. It prevents old presets from producing plausible but wrong v2 reports. It also gives demos/tests a clean v2 fixture set.
- Cons: requires maintaining v1/v2 template selection and compatibility tests.

Recommended: Option C, with Option B-style schema loading underneath. Keep v1 compatibility, introduce v2 assets, and make scheme mismatches explicit errors or strong warnings.

## Backward compatibility plan

- Existing Project Session files: load as schema v1/from-to. Do not auto-convert unless the user explicitly selects a migration flow and reviews every movement.
- Existing Mapping Preset files: load as v1/from-to. Show scheme in UI. Prevent applying a v1 preset to a v2 approach-movement template without review.
- Existing demo files: keep current `DEMO_TMC1_FourLeg*` files as v1 demos. Add separate v2 demo mapping/preset files.
- Old exported reports: leave untouched. They are output artifacts, not source state.
- Tests: preserve v1 tests and add v2 test coverage. Avoid rewriting v1 tests in a way that hides compatibility regressions.
- Batch workflow: require one movement scheme per batch run. Validate that each preset/session/template combination agrees before analysis/export.

## Key risks

- Northbound ambiguity: `N` must mean travel direction / มุ่งเหนือ, not merely "the north leg" / ขาทิศเหนือ. This needs UI copy, docs, and tests.
- U-turn handling: v1 U-turns already use second letter `U`, but v2 also uses `U`; the first letter meaning changes from from-leg to travel direction. Existing U-turn rows still need review.
- Left/right conversion: converting from leg-to-leg to `L`/`R` depends on approach orientation and local driving-side assumptions. Do not auto-convert without deterministic geometry rules and review.
- Diagram orientation: current diagram labels approach tables by north/south/east/west but uses from-to codes. V2 arrows and table ordering must be verified visually.
- Source survey sheets with non-cardinal labels: raw sheets such as Thai `ทิศ 1` do not encode movement direction. They need explicit mapping.
- Many-to-one aggregation: duplicate `output_movement_code` is allowed for aggregation. V2 must preserve this behavior for mainline/frontage/source streams.
- Silent wrong reports: old mapping presets could look structurally valid but semantically wrong if interpreted under v2. Scheme mismatch should block export.
- Template chart/drawing preservation: openpyxl edits can disturb native drawings/charts. Use new versioned templates and visual verification.

## Implementation phases

### Phase B: movement scheme constants and validators

- Add `MovementScheme` constants: `from_to`, `approach_movement`.
- Add v1 and v2 movement code lists/orders.
- Add validators for scheme, approach direction, movement type, and code consistency.
- Add metadata/export fields for movement scheme.

### Phase C: approach-movement demo mapping/template

- Create v2 demo mapping JSON/XLSX using `NL`/`NT`/etc.
- Create v2 template workbook/map files.
- Verify template formulas, diagram cells, chart ranges, and generated exports.

### Phase D: mapping editor and preset schema

- Add scheme-aware mapping editor columns.
- Create mapping preset schema v2.
- Load v1 presets as from-to and v2 presets as approach-movement.
- Block or require explicit review for scheme mismatches.

### Phase E: normalizer, summaries, export, report

- Add v2 normalized columns.
- Make movement ordering scheme-aware.
- Update `Movement_Aggregation_Audit` to include scheme.
- Update exporter/report template/Excel COM paths to select v1 or v2 movement layout.

### Phase F: diagram and tests

- Update static diagram generation for v2 movement codes and orientation.
- Add unit tests for v2 mappings, summaries, template exports, batch, QC, and sessions.
- Add visual/smoke verification for v2 diagram and report workbook.

### Phase G: migration, compatibility, release notes

- Add user-facing migration notes.
- Add clear UI labels explaining northbound / มุ่งเหนือ.
- Add compatibility warnings for v1 sessions/presets.
- Document known non-auto-convert cases.

## Test plan

- Unit test v2 constants and validators.
- Unit test mapping cleaning for both `movement_code` and `output_movement_code`.
- Unit test v1 preset/session load remains unchanged.
- Unit test v2 preset/session round trips with `movement_code_scheme = "approach_movement"`.
- Unit test that v1 presets cannot silently feed v2 reports.
- Unit test normalization emits `movement_scheme`, `approach_direction`, `movement_type`, `movement_code`, and `output_movement_code`.
- Unit test summaries use v2 movement order.
- Unit test duplicate v2 `output_movement_code` aggregation remains allowed.
- Export smoke test for v1 demo remains passing.
- Add v2 demo smoke test that verifies workbook sheets, movement headers, diagram data, and expected peak behavior.
- Template test: load v2 map and assert all 16 v2 movement codes are present in diagram cells and hourly table columns.
- Visual/manual check: open the v2 template-generated report in Excel and verify chart/drawing preservation.

## Open questions / decisions needed

- Should v2 mapping retain `from_leg` / `to_leg` as optional legacy fields, or remove them from the v2 editor while keeping them in serialized v1 rows?
- Should `turn_type` remain a word field (`left`, `through`, `right`, `u_turn`) while `movement_type` is `L`/`T`/`R`/`U`?
- Should the app offer an assisted v1-to-v2 conversion, or require users to remap manually?
- What exact diagram table order is preferred for each approach in v2: `L,T,R,U` or a layout-specific order matching physical arrows?
- Should old template filenames keep `four_leg` naming while adding scheme, or should v2 introduce a clearer naming convention?
- How should non-cardinal/diagonal legs currently allowed by `LEG_OPTIONS` (`NE`, `NW`, `SE`, `SW`, `CUSTOM`) map into the v2 four-cardinal approach scheme?

## Files inspected

- `app.py`
- `src/tmc_processor/constants.py`
- `src/tmc_processor/mapping.py`
- `src/tmc_processor/normalizer.py`
- `src/tmc_processor/pipeline.py`
- `src/tmc_processor/summaries.py`
- `src/tmc_processor/diagram.py`
- `src/tmc_processor/exporter.py`
- `src/tmc_processor/report_template.py`
- `src/tmc_processor/excel_com_export.py`
- `src/tmc_processor/mapping_preset.py`
- `src/tmc_processor/session.py`
- `src/tmc_processor/qc.py`
- `src/tmc_processor/batch.py`
- `templates/four_leg_tmc_report_template.xlsx`
- `templates/four_leg_tmc_report_template_map.json`
- `samples/demo/DEMO_TMC1_FourLeg.mapping.json`
- `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx`
- `samples/demo/DEMO_TMC1_FourLeg.xlsx`
- `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`
- `scripts/smoke_demo.py`
- `tests/`
