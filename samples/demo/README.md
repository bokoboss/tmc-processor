# Synthetic Demo Dataset

All files in this folder are synthetic. They do not contain real survey counts,
real project names, real survey points, client data, or private files.

## Files

- `DEMO_TMC1_FourLeg.xlsx` - synthetic Day 1 raw TMC workbook.
- `DEMO_TMC1_FourLeg_Day2.xlsx` - synthetic Day 2 raw TMC workbook for batch testing.
- `DEMO_TMC1_FourLeg_mapping.xlsx` - mapping workbook for the older "load saved mapping Excel" workflow.
- `DEMO_TMC1_FourLeg.mapping.json` - Mapping Preset for the same synthetic intersection. Use this for the recommended preset workflow and for batch mode.
- `DEMO_TMC1_FourLeg_session.tmcproj.json` - optional Project Session with synthetic setup metadata, mapping, PCE factors, peak settings, and export settings. It does not embed raw Excel content.

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

## Mapping Excel Vs Mapping Preset

The mapping Excel file is kept for compatibility with the saved mapping Excel
workflow.

The Mapping Preset (`.mapping.json`) is the recommended reusable mapping format.
It stores mapping rows only, so it can be shared across survey dates without
carrying full project setup or raw workbook content.
