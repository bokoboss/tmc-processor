# TMC Processor

Streamlit MVP for processing Turning Movement Count (TMC) Excel files into report-ready outputs.

## Features

- Detects raw direction sheets named like `ทิศ 1`, `ทิศ 2`, or `ทิศ 2+3`.
- Previews detected raw sheets before processing.
- Provides an editable mapping table for raw sheet to movement metadata, with dropdowns for standard movement codes, legs, turn types, and facility types.
- Exports and reloads mapping templates so the same intersection mapping can be reused across team members and future surveys.
- Normalizes raw count sheets into the long format defined in `AGENTS.md`.
- Calculates PCU using the default PCE factors.
- Produces hourly, movement, vehicle-composition, AM/PM peak, and PHF summaries.
- Supports fixed hourly or rolling 60-minute peak detection with configurable AM/PM peak search windows.
- Runs QC checks for unmapped sheets, vehicle/PCE issues, zero-volume mapped movements, and peak-time issues.
- Exports an Excel workbook with `Setup`, `Mapping`, `Normalized_Data`, `QC_Check`, `Hourly_Summary`, `Movement_Summary`, `Vehicle_Composition`, `Peak_PHF`, and `Report_Text`.
- Recommends Microsoft Excel COM automation on Windows for final reports with native template charts, formulas, and workbook layout.
- Keeps openpyxl with generated PNG charts available as the compatibility fallback export path.

Automatic intersection diagram generation is intentionally out of scope for this MVP.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

## Run The App

```powershell
streamlit run app.py
```

If Windows reports that `streamlit` is not on PATH, use:

```powershell
python -m streamlit run app.py
```

Upload a raw TMC Excel workbook. Each raw direction sheet should be named with the Thai direction prefix, for example `ทิศ 1`, `ทิศ 2`, or `ทิศ 2+3`.

## Mapping Workflow

After upload, the app detects every raw direction sheet and creates one mapping row per sheet. Review every row before processing.

Required fields:

- `movement_code`: choose one of the standard movement codes such as `NS`, `NE`, `SW`, `OVP_NS`, `UND_EW`, or `CUSTOM`.
- `from_leg` and `to_leg`: choose `N`, `S`, `E`, `W`, `NE`, `NW`, `SE`, `SW`, or `CUSTOM`.
- `turn_type`: choose `Through`, `Left`, `Right`, `U-turn`, `Combined`, or `Other`.
- `facility_type`: choose `At-grade`, `U-turn`, `Overpass`, `Underpass`, or `Other`.

Use `Download mapping template` to save the current mapping table to Excel. This is useful after one person has reviewed the vendor sheet directions and wants to share the mapping with the rest of the team.

Use `Load saved mapping Excel` to apply a previous template to the currently detected raw sheets. Rows are matched by `raw_sheet`; sheets present in the current workbook but missing from the saved mapping stay in the table with blank required fields so they can be completed manually.

The app validates the mapping before processing. Every detected raw sheet must have `movement_code`, `from_leg`, `to_leg`, `turn_type`, and `facility_type`; incomplete mappings are shown in a validation table and processing stops until they are fixed.

## Expected Raw Sheet Shape

The MVP supports straightforward interval tables where each detected sheet contains:

- A time column such as `time`, `interval`, `ช่วงเวลา`, or `เวลา`, with values like `07:00-07:15`.
- Or separate start/end columns such as `time_start` and `time_end`.
- Vehicle-class columns matching `Bicy`, `MC`, `PC<7`, `PC>7`, `LB`, `MB`, `HB`, `LT`, `MT`, `HT`, `TR`, `STR`.

Raw direction numbers are vendor-defined. The app does not infer movement direction automatically; users must review and edit the mapping table.

## Tests

```powershell
pytest
```

## Excel Export Modes

On Windows machines with Microsoft Excel available, the Streamlit export section defaults to Excel Template Mode. This is the recommended final report path because it uses Excel COM automation to preserve native template charts, template formulas, and workbook layout.

Safe PNG Export Mode remains available as the compatibility fallback. It uses openpyxl and generated static PNG charts, avoids Excel COM automation, and does not require Microsoft Excel.

Excel Template Mode requires Microsoft Excel and pywin32. It opens a copy of `templates/four_leg_tmc_report_template.xlsx`, writes report data into mapped cells/ranges, lets the native template charts update from their linked source ranges, and saves a new output workbook. The source template is never overwritten.

If Excel COM is unavailable, the app warns and automatically falls back to Safe PNG Export Mode. If a COM-generated workbook opens with an Excel repair warning, do not use that output; report the issue and use the safe export instead.

Manual COM smoke test:

```powershell
python scripts/export_with_excel_com_smoke_test.py
```

## Expected Validation Results

Create or refresh the validation baseline workbook with:

```powershell
python scripts/create_expected_results.py
```

The script writes `samples/expected/expected_results.xlsx` with one row per raw workbook in `samples/raw`. See `samples/expected/README.md` for the manual review workflow and verification status rules.

Validate eligible baseline rows with:

```powershell
python scripts/validate_expected_results.py
```

The validation runner includes rows marked `verified`, `needs_manual_check`, or `partial_report_baseline`. It loads each row's raw workbook and mapping workbook when available, runs the processing pipeline, and writes `samples/expected/validation_report.xlsx`.

By default, validation searches AM peaks in `07:00-12:00` and PM peaks in `15:00-19:00`. To match report-specific peak definitions, fill `am_peak_window_start`, `am_peak_window_end`, `pm_peak_window_start`, or `pm_peak_window_end` in `expected_results.xlsx`; blank window fields use the defaults. The Streamlit app exposes the same peak search window settings, and exported workbooks record them in `Setup`.

Report-derived expected PCU values may be rounded presentation values rather than exact raw calculations. Validation therefore allows `total_pcu` to pass within the larger of 10 PCU or 0.1%, AM/PM peak PCU within 2 PCU, and PHF within 0.005. Peak time fields still must match exactly after normalization.

Blank expected fields are reported as `skipped_blank_expected` with `pass_fail = SKIPPED`, not filled or treated as failures. Rows without a mapping file are reported as `skipped_missing_mapping`. The validation report includes `tolerance_used` and `relative_difference_pct` columns for numeric comparisons. Use `--fail-on-mismatch` if you want a non-zero exit code when any compared metric fails.
