# Synthetic Demo Dataset

All files in this folder are synthetic. They do not contain real survey counts,
real project names, real survey points, client data, or private files.

## Files

- `DEMO_TMC1_FourLeg.xlsx` - synthetic Day 1 raw TMC workbook.
- `DEMO_TMC1_FourLeg_Day2.xlsx` - synthetic Day 2 raw TMC workbook for batch testing.
- `DEMO_TMC1_FourLeg_mapping.xlsx` - mapping workbook for the compatible "load saved mapping Excel" workflow.
- `DEMO_TMC1_FourLeg.mapping.json` - Mapping Preset for the same synthetic intersection. Use this for the recommended preset workflow and for batch mode.
- `DEMO_TMC1_FourLeg_session.tmcproj.json` - optional Project Session with synthetic setup metadata, mapping, PCE factors, peak settings, and export settings. It does not embed raw Excel content.
- `DEMO_TMC1_FourLeg_approach_v2_mapping.xlsx` - v2 approach-movement mapping workbook draft for migration/demo review.
- `DEMO_TMC1_FourLeg_approach_v2.mapping.json` - v2 approach-movement Mapping Preset draft for migration/demo review.

## Single-File Demo

1. Start the app with `start_tmc_processor.bat`.
2. Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx`.
3. Load either:
   - `samples/demo/DEMO_TMC1_FourLeg.mapping.json`, or
   - `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx`.
4. Process the workbook.
5. Review the dashboard and confirm Peak periods.
6. Generate the Excel report or Export Package ZIP.

The synthetic data is designed to show an AM peak around 08:00-09:00 and a PM
peak around 17:00-18:00.

## Batch Demo

1. Start the app with `start_tmc_processor.bat`.
2. Choose `ประมวลผลหลายไฟล์`.
3. Upload both demo workbooks:
   - `samples/demo/DEMO_TMC1_FourLeg.xlsx`
   - `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`
4. Load `samples/demo/DEMO_TMC1_FourLeg.mapping.json` as the shared Mapping Preset.
5. Set survey date and output stem per file if needed.
6. Analyze the batch.
7. Review or adjust the per-file AM/PM confirmed peaks.
8. Generate and download the Batch ZIP.

## approach_movement Batch Demo

`approach_movement` Batch analysis and Safe PNG/generated ZIP export are supported.

1. Start the app with `start_tmc_processor.bat`.
2. Choose Batch mode.
3. Upload both demo workbooks:
   - `samples/demo/DEMO_TMC1_FourLeg.xlsx`
   - `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`
4. Load `samples/demo/DEMO_TMC1_FourLeg_approach_v2.mapping.json` as the shared Mapping Preset.
5. Analyze the batch.
6. Review or adjust per-file AM/PM confirmed peaks.
7. Use Safe PNG Export Mode to generate the Batch ZIP.

The v2 Batch ZIP includes `batch_summary.xlsx`, one per-file generated report workbook, export summary text, session JSON, mapping JSON, `diagram/movement_diagram_data.csv`, `diagram/movement_diagram.png`, and generated charts when available. Raw input Excel files are excluded.

Excel Template Mode for `approach_movement` Batch is intentionally blocked in this version with the message: `Excel Template Mode สำหรับ Batch approach_movement ยังไม่รองรับในเวอร์ชันนี้ กรุณาใช้ Safe PNG Export Mode`.

## Mapping Excel Vs Mapping Preset

The mapping Excel file is kept for compatibility with the saved mapping Excel
workflow.

The Mapping Preset (`.mapping.json`) is the recommended reusable mapping format.
It stores mapping rows only, so it can be shared across survey dates without
carrying full project setup or raw workbook content.

## Movement-Code Schemes

The standard demo artifacts use the `from_to` movement-code scheme, with
codes such as `NS`, `WE`, and `EW`. Existing mapping files remain supported as
`from_to` when they do not include `movement_code_scheme` metadata.

The `*_approach_v2*` demo artifacts use the `approach_movement` scheme, with
codes such as `NT`, `WT`, and `ET` from the standard v2 set `NL`, `NT`, `NR`,
`NU`, `SL`, `ST`, `SR`, `SU`, `EL`, `ET`, `ER`, `EU`, `WL`, `WT`, `WR`, `WU`.
In this scheme `NT` means Northbound Through / มุ่งเหนือ-ตรง: `N` is travel
direction and `T` is movement type.
They are supported for single-file Excel Template Mode through the Excel
COM/native path when available, single-file generated/Safe PNG export, and Batch
generated/Safe PNG export. `approach_movement` Batch Excel Template Mode remains
blocked until a safe per-file COM template export path is implemented.

All demo data in this folder is synthetic.
