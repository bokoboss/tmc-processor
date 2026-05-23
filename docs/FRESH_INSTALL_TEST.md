# Fresh Install Test Checklist

Use this checklist before publishing a public beta release.

## Setup

- Start from a clean clone or downloaded ZIP.
- Confirm Python 3.10 or newer is installed and available on `PATH`.
- Double-click `start_tmc_processor.bat`.
- Confirm `.venv` is created.
- Confirm dependency installation completes.
- Confirm the app opens in a browser.

## Single-File Demo

- Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx`.
- Load `samples/demo/DEMO_TMC1_FourLeg.mapping.json`.
- Process the workbook.
- Review and confirm Peak periods.
- Generate an Excel report or Export Package ZIP.

## Batch Demo

- Choose `ประมวลผลหลายไฟล์`.
- Upload:
  - `samples/demo/DEMO_TMC1_FourLeg.xlsx`
  - `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`
- Load `samples/demo/DEMO_TMC1_FourLeg.mapping.json`.
- Set survey date and output stem per file.
- Analyze the batch.
- Review Peak periods per file.
- Generate the Batch ZIP.
- Open `batch_summary.xlsx` and confirm `Batch_QC` exists.

## Export Package Safety

- Open the single-file Export Package ZIP.
- Open the Batch ZIP.
- Confirm raw input Excel workbooks are not included by default.
- Confirm local raw file paths are not exposed in generated package summaries.

## Basic Troubleshooting

- If `python` is not found, reinstall Python and enable `Add python.exe to PATH`.
- If first run is slow, wait for `.venv` creation and dependency installation to finish.
- If Excel Template Mode fails, confirm Microsoft Excel desktop app is installed, then retry.
- If Excel COM remains unavailable, use Safe PNG Export Mode.
- If a browser does not open automatically, copy the local Streamlit URL from the launcher window into a browser.
