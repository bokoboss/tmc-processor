# Approach Movement v2 Excel Template Design

## Goal

Phase I1 prepares the `approach_movement` v2 Excel template design and integrity checks without enabling Excel Template Mode. The template must report v2 movement codes directly:

- `N` = northbound / มุ่งเหนือ
- `S` = southbound / มุ่งใต้
- `E` = eastbound / มุ่งตะวันออก
- `W` = westbound / มุ่งตะวันตก

The first letter is the travel direction. It is not the leg where the vehicle comes from. The template must not convert v2 codes to v1 `from_to` codes.

The canonical v2 movement header order is:

```text
NL, NT, NR, NU,
SL, ST, SR, SU,
EL, ET, ER, EU,
WL, WT, WR, WU
```

## v1 Template Dependencies

The current production template remains:

- Workbook: `templates/four_leg_tmc_report_template.xlsx`
- Map: `templates/four_leg_tmc_report_template_map.json`
- Export code: `src/tmc_processor/report_template.py` and `src/tmc_processor/exporter.py`

Read-only inspection of the v1 workbook found:

- Sheets present: `Summary` only.
- Used range: `A1:BG218`.
- Print area: `Summary!$B$2:$AK$64`.
- Sheet protection: disabled.
- Excel tables: none.
- Merged ranges: 30 ranges, including `B2:AN2`, `E4:S4`, `C8:H8`, `C9:H9`, `E35:R36`, `V8:V9`, `W8:AM8`, `W26:AJ26`, and `B61:U62`.
- Native charts/drawings: the package contains `xl/drawings/drawing1.xml`, `xl/charts/chart1.xml`, and `xl/charts/chart2.xml`.
- Drawing audit: `drawing1.xml` has 33 two-cell anchors, 2 chart graphic frames, 33 shape elements, and 2 connector shape elements.
- `openpyxl` warning on load: DrawingML support is incomplete, and shapes/drawings may be lost.

Movement header cells in v1:

- Diagram north block: `J13:M13` = `NU`, `NW`, `NS`, `NE`.
- Diagram south block: `J30:M30` = `SW`, `SN`, `SE`, `SU`.
- Diagram west block: `D19:D22` = `WN`, `WE`, `WS`, `WU`.
- Diagram east block: `R21:R24` = `EU`, `EN`, `EW`, `ES`.
- Hourly movement table: `W9:AL9` = `NE`, `NS`, `NW`, `NU`, `SW`, `SN`, `SE`, `SU`, `WN`, `WE`, `WS`, `WU`, `ES`, `EW`, `EN`, `EU`.

Diagram and diagram-related cells/ranges:

- Title/date: `C8:H8`, `C9:H9`.
- Direction labels: north `G12:H12`, south `O32`, east `Q17:R17`, west `D26`.
- Road labels: north `K11`, south `K32`, east `R19`, west `D24`.
- Movement values: north `J14:M16`, west `E19:G22`, east `O21:Q24`, south `J27:M29`.
- Caption: `E35:R36`.
- Summary box: `D30:F32`, with label range `D33:E33`.

Formula ranges depending on movement headers:

- Diagram lookups: `J14:M16`, `E19:G22`, `O21:Q24`, `J27:M29`.
- Summary totals: `F30:F32`.
- Hourly movement totals: `W22:AM22`.
- Vehicle-class linked formulas: `W28:AJ40`.
- Vehicle composition chart source helpers: `AP38:BA38`, `AP40:BA40`.
- Template title/date formulas: `B2`, `C8:C9`.

Metadata mapped cells:

- Report title: `B2:AN2`.
- Project: `E4:S4`.
- Survey point: `E5`.
- Survey date: `K5:M5`.
- Weather: `Q5:R5`.
- Responsible party: `E6`.
- Survey period: `K6:M6`.

The v1 template map also defines chart anchors for the native charts and PNG fallback placement:

- Hourly PCU chart: native source `V10:V21` and `AM10:AM21`, PNG anchor `B37`.
- Vehicle composition chart: native source `AP39:BA39` and `AP40:BA40`, PNG anchor `U41`.

## Proposed v2 Sheet Structure

The first v2 template should remain one report worksheet named `Summary`, matching the existing mapped export surface. Supporting generated data sheets may be added by export code later, but Phase I1 does not enable template export.

Recommended structure:

- `Summary`: Excel-authored report layout using v2 headers.
- Generated export data sheets, when Phase I2 enables the path: existing export sheets plus `Movement_Diagram_Data`.
- No hidden conversion sheet from v2 to v1.

## Proposed Movement Summary Area

Use the v2 map draft at `templates/four_leg_tmc_report_template_approach_v2_map.json`.

The hourly movement table should use:

- Time column: `V`.
- Header row: `9`.
- Data rows: `10:21`.
- Total row: `22`.
- Movement columns: `W:AL`.
- Code order: `NL`, `NT`, `NR`, `NU`, `SL`, `ST`, `SR`, `SU`, `EL`, `ET`, `ER`, `EU`, `WL`, `WT`, `WR`, `WU`.
- Total column: `AM`.

## Proposed Movement Diagram Area

Keep the same broad cell layout as v1 for the draft map, but replace headers with v2 codes:

- Northbound group: `J13:M13` = `NL`, `NT`, `NR`, `NU`.
- Southbound group: `J30:M30` = `SL`, `ST`, `SR`, `SU`.
- Eastbound group: `R21:R24` = `EL`, `ET`, `ER`, `EU`.
- Westbound group: `D19:D22` = `WL`, `WT`, `WR`, `WU`.

This design is a movement summary layout, not a proven arrow-geometry diagram. Any arrow drawing must be manually checked against v2 travel-direction semantics before it is shown as a production diagram.

## Proposed Metadata Cells

Reuse the v1 metadata cell layout:

- Report title: `B2`.
- Project: `E4`.
- Survey point: `E5`.
- Survey date: `K5`.
- Weather: `Q5`.
- Responsible party: `E6`.
- Survey period: `K6`.
- Direction labels: `G12`, `O32`, `Q17`, `D26`.
- Road labels: `K11`, `K32`, `R19`, `D24`.
- Caption: `E35`.

## Proposed Formula Strategy

For Phase I2, formulas should remain map-driven and should reference v2 headers directly. The template should not translate `NL` to any v1 `from_to` code.

Recommended formula strategy:

- Keep movement totals as `HLOOKUP` or `INDEX/MATCH` against the hourly v2 movement table.
- Keep summary total formulas summing the visible v2 diagram value cells.
- Keep hourly total formulas on the v2 hourly table.
- Keep vehicle-class and vehicle-composition formulas independent from movement-code semantics.
- Preserve `summary_formula_write_mode = preserve_template` unless the final workbook intentionally moves formula ownership to the exporter.

## Relationship to Movement_Diagram_Data and PNG

`Movement_Diagram_Data` and `diagram/movement_diagram.png` are generated v2 artifacts from Phases H/H2. They should relate to the template as follows:

- `Movement_Diagram_Data` is the semantic source of truth for v2 diagram values.
- The current package PNG is a table-style visual, not an intersection arrow diagram.
- A future v2 template may embed the PNG as a static visual if native shapes are too risky.
- A future native Excel diagram may be built manually, but only after its labels and arrow positions are verified against v2 travel-direction semantics.

## Native Chart and Drawing Decision

Native Excel charts and drawings should be deferred for v2 template automation.

Reasoning:

- The v1 workbook contains native chart parts and many drawing/shape elements.
- `openpyxl` warns that unsupported shapes/drawings can be lost on load/save.
- Phase I1 should not create a production-looking workbook by copying v1 and saving through `openpyxl`.
- If native charts/drawings are needed, the v2 `.xlsx` should be finalized manually in Microsoft Excel and then inspected read-only.

The safer automated path is to keep generated workbook export and package PNG support separate until the v2 template workbook is manually finalized.

## Risks Under openpyxl

- Shape and connector loss when saving a shape-heavy workbook.
- Native chart part changes or chart source drift.
- Merged-cell write mistakes if map ranges are edited manually.
- Formula references that still point to v1 headers after a visual edit.
- Silent semantic errors if a header such as `NS` appears in a v2 movement area.

Tests must not compare chart/drawing binaries. They should check workbook readability, mapped cells/ranges, expected sheet names, movement header values, and absence of v1-only movement codes.

## Manual Excel Finalization Checklist

Before enabling v2 Excel Template Mode:

- Create `templates/four_leg_tmc_report_template_approach_v2.xlsx` manually in Excel, not by copying and saving through `openpyxl`.
- Confirm the workbook opens in Excel with no repair prompt.
- Confirm the `Summary` sheet exists and the print area is intentional.
- Confirm all v2 movement headers match `APPROACH_MOVEMENT_CODES` in the required order.
- Confirm no v1-only movement headers such as `NS`, `WE`, `EN`, or `EW` remain in v2 movement areas.
- Confirm formulas reference v2 headers or `Movement_Diagram_Data`, not v1 `Diagram_Data` semantics.
- Confirm native charts still display and use intended source ranges, or replace them with generated PNG chart anchors.
- Confirm any arrows/shapes represent travel direction, not from-leg semantics.
- Run `python -m pytest tests/test_template_integrity.py`.

Phase I1 decision: v2 template XLSX finalization is deferred to manual Excel editing. No v2 `.xlsx` draft is created in this phase because the current source workbook has shapes/drawings that are unsafe to round-trip with `openpyxl`.

## Phase I2B/I2C Status

The v2 workbook has now been manually finalized in Excel and validated against the v2 map:

- Workbook: `templates/four_leg_tmc_report_template_approach_v2.xlsx`
- Map: `templates/four_leg_tmc_report_template_approach_v2_map.json`
- Template version: `four_leg_approach_movement_v2`
- Movement code scheme: `approach_movement`

Phase I2C keeps the Excel-authored workbook as the v2 visual template source, but the current openpyxl helper is limited to structural/internal validation. It writes only to an output copy in memory and preserves template formula ownership where the map declares `preserve_template`, but it is not native-template-preserving: openpyxl can drop or damage Excel-authored drawings, charts, shapes, lines, and arrows when saving.

User-facing v2 Excel Template Mode must therefore require a native Excel/COM preservation path. Before Phase I3, the Streamlit UI blocked v2 Excel Template Mode with a clear message and directed users to Safe PNG Export Mode. Safe PNG/generated workbook export remains the supported fallback for v2 when COM is unavailable.

## Phase I3 Native Excel COM Export

Phase I3 enables v2 Excel Template Mode only through Microsoft Excel COM/native automation on Windows. The export API is `export_v2_template_workbook_com()`.

Behavior:

- Uses only `templates/four_leg_tmc_report_template_approach_v2.xlsx` and `templates/four_leg_tmc_report_template_approach_v2_map.json`.
- Copies the source workbook to a temporary file before opening it in Excel.
- Writes mapped metadata, road/destination labels, caption text, hourly v2 movement values, vehicle-class data, support sheets, and `Movement_Diagram_Data`.
- Preserves Excel-authored charts, drawings, lines, arrows, formulas, styles, and page layout by letting Excel save the temporary workbook copy.
- Recalculates before save using the same COM calculation pattern as the v1 native path.
- Returns workbook bytes from the saved temporary output.
- Never rewrites the source v2 template workbook.

If Excel COM is unavailable, v2 Excel Template Mode remains blocked with the Thai UI message and does not silently fall back to openpyxl template save or generated chart workbook. Safe PNG Export Mode is still available and continues to use the generated v2 workbook/package path.

The openpyxl helper `export_v2_template_workbook()` remains limited to structural/internal validation. It is not visual Template Mode because saving the template through openpyxl may drop or damage Excel-authored charts, drawings, shapes, lines, and arrows.

Formula and chart notes:

- `Summary!W9:AL9` must remain `NL, NT, NR, NU, SL, ST, SR, SU, EL, ET, ER, EU, WL, WT, WR, WU`.
- Existing template formulas own the visual summary calculations where the map declares `preserve_template`.
- CI tests do not require Excel COM and should test workbook cells/formulas and resource selection rather than native chart cache pixels.
- Final visual confirmation of native chart/drawing preservation requires a local Windows Excel smoke test.

Phase K update: v2 Batch now supports Safe PNG/generated exports only. Excel Template Mode for v2 Batch is still blocked with the Thai message `Excel Template Mode สำหรับ Batch approach_movement ยังไม่รองรับในเวอร์ชันนี้ กรุณาใช้ Safe PNG Export Mode` until a per-file COM batch path is implemented and tested.
