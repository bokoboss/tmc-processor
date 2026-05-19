# Synthetic Demo Dataset

All files in this folder are fully synthetic. They do not contain real project
names, real survey points, real traffic counts, client data, or private files.

## Files

- `DEMO_TMC1_FourLeg.xlsx` - synthetic raw TMC workbook with four direction
  sheets: `ทิศ 1`, `ทิศ 2`, `ทิศ 3`, and `ทิศ 4`.
- `DEMO_TMC1_FourLeg_mapping.xlsx` - saved mapping workbook for the current
  "Load saved mapping Excel" workflow. It includes a simple aggregation example
  where a mainline through stream and a frontage through stream both map to
  `NS`.
- `DEMO_TMC1_FourLeg_session.tmcproj.json` - optional synthetic Project Session
  with demo metadata, labels, default PCE factors, peak windows, export settings,
  and mapping rows. It does not embed raw Excel content.

## Quick Demo Workflow

1. Start the app with `start_tmc_processor.bat`.
2. Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx`.
3. In the mapping section, load `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx`.
4. Click Process.
5. Review the Dashboard and Peak Review. The synthetic data is designed to show
   an AM peak around 08:00-09:00 and a PM peak around 17:00-18:00.
6. Generate the Excel Report.

The Project Session file can be loaded as a convenience after uploading the demo
workbook, but the workbook and mapping file are enough for the basic demo.
