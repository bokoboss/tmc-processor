# Synthetic Demo Dataset

All files in this folder are fully synthetic. They do not contain real project
names, real survey points, real traffic counts, client data, or private files.

## Files

- `DEMO_TMC1_FourLeg.xlsx` - synthetic raw TMC workbook with four direction
  sheets: `ทิศ 1`, `ทิศ 2`, `ทิศ 3`, and `ทิศ 4`.
- `DEMO_TMC1_FourLeg_Day2.xlsx` - second synthetic raw TMC workbook with the
  same sheet structure and different deterministic volumes for Basic Batch v1.
- `DEMO_TMC1_FourLeg_mapping.xlsx` - saved mapping workbook for the current
  "Load saved mapping Excel" workflow. It includes a simple aggregation example
  where a mainline through stream and a frontage through stream both map to
  `NS`.
- `DEMO_TMC1_FourLeg.mapping.json` - Mapping Preset for the same synthetic
  intersection. It stores mapping rows only, so it can be reused separately from
  the full Project Session.
- `DEMO_TMC1_FourLeg_session.tmcproj.json` - optional synthetic Project Session
  with demo metadata, labels, default PCE factors, peak windows, export settings,
  and mapping rows. It does not embed raw Excel content.

## Quick Demo Workflow

1. Start the app with `start_tmc_processor.bat`.
2. Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx`.
3. In the mapping section, load either `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx`
   or `samples/demo/DEMO_TMC1_FourLeg.mapping.json`.
4. Click Process.
5. Review the Dashboard and Peak Review. The synthetic data is designed to show
   an AM peak around 08:00-09:00 and a PM peak around 17:00-18:00.
6. Generate the Excel Report.

The Project Session file can be loaded as a convenience after uploading the demo
workbook, but the workbook plus either mapping file is enough for the basic demo.

## Batch Demo Workflow

1. Start the app with `start_tmc_processor.bat`.
2. Open the `ประมวลผลหลายไฟล์` tab.
3. Upload both `samples/demo/DEMO_TMC1_FourLeg.xlsx` and
   `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`.
4. Load `samples/demo/DEMO_TMC1_FourLeg.mapping.json` as the shared Mapping
   Preset.
5. Click Process Batch.
6. Review the status table and download the Batch ZIP.
