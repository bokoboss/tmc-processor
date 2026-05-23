# TMC Processor

TMC Processor is a Streamlit tool for processing Turning Movement Count (TMC)
Excel workbooks. It helps traffic teams turn raw survey workbooks into mapped
movement tables, PCU/PCE summaries, peak-hour review outputs, charts, Excel
reports, and export packages.

Version: `0.2.0` public beta.

## What It Does

- Reads TMC Excel workbooks with directional survey sheets.
- Maps raw survey directions and source streams to standard movement codes.
- Supports many-to-one movement aggregation, with audit output for traceability.
- Calculates PCU values from editable PCE factors.
- Lets users review and confirm AM/PM peak periods before export.
- Exports Excel reports using either Excel Template Mode or Safe PNG Export Mode.
- Saves and loads Mapping Presets (`.mapping.json`) and Project Sessions (`.tmcproj.json`).
- Processes one file at a time or multiple days for the same survey point in batch mode.
- Creates Export Package ZIP files and Batch ZIP files with `batch_summary.xlsx` and `Batch_QC`.

## Quick Start For Windows Users

1. Download the repository ZIP from GitHub, or clone it with Git.
2. Extract the ZIP to a local folder, for example `C:\MyRD\tmc-processor`.
3. Double-click `start_tmc_processor.bat`.
4. Wait for setup to finish. The first run may take several minutes because the launcher creates `.venv`, installs dependencies, and checks optional Excel COM support.
5. Use the app in the browser window that opens. Later runs are faster because the virtual environment already exists.

Requirements:

- Windows
- Python 3.10 or newer, with `python.exe` available on `PATH`
- Microsoft Excel desktop app only if you want Excel Template Mode with COM/native chart preservation

If Excel COM is unavailable, use Safe PNG Export Mode.

## Single-File Workflow

1. Start the app with `start_tmc_processor.bat`.
2. Choose the single-file work mode.
3. Upload one raw TMC workbook.
4. Fill in setup metadata such as project name, survey point, survey date, road labels, and weather.
5. Create or load mapping:
   - load a Mapping Preset (`.mapping.json`), or
   - load a saved mapping Excel file, or
   - edit the mapping table in the app.
6. Review or adjust editable PCE factors if needed.
7. Process the workbook.
8. Review the dashboard and confirm AM/PM Peak periods.
9. Generate an Excel report or an Export Package ZIP.

Export packages include processed outputs, charts, summary text, mapping/session metadata, and report files. They do not include raw input Excel files by default.

## Batch Workflow

Batch v1 is intended for the same survey point or same intersection surveyed across multiple days, using one shared Mapping Preset.

1. Start the app with `start_tmc_processor.bat`.
2. Choose `ประมวลผลหลายไฟล์`.
3. Upload multiple raw or demo workbooks for the same survey point.
4. Load one shared Mapping Preset, for example `samples/demo/DEMO_TMC1_FourLeg.mapping.json`.
5. Set survey date and output stem for each file.
6. Analyze the batch.
7. Review and confirm Peak periods per file.
8. Generate the Batch ZIP.

The Batch ZIP contains `batch_summary.xlsx`, including `Batch_QC`, plus one sanitized output folder per successful workbook. Raw input Excel files and local raw paths are not included.

## Demo Workflow

Demo files live in `samples/demo/`:

- `DEMO_TMC1_FourLeg.xlsx`
- `DEMO_TMC1_FourLeg_Day2.xlsx`
- `DEMO_TMC1_FourLeg.mapping.json`
- `DEMO_TMC1_FourLeg_mapping.xlsx`
- `DEMO_TMC1_FourLeg_session.tmcproj.json`

All demo files are synthetic. They do not contain real survey counts, project names, client data, or private locations.

Single-file demo:

1. Start the app.
2. Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx`.
3. Load `samples/demo/DEMO_TMC1_FourLeg.mapping.json` or `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx`.
4. Process the workbook.
5. Review Peak periods.
6. Generate an Excel report or Export Package ZIP.

Batch demo:

1. Choose `ประมวลผลหลายไฟล์`.
2. Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx` and `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`.
3. Load `samples/demo/DEMO_TMC1_FourLeg.mapping.json`.
4. Set survey dates and output stems if needed.
5. Analyze the batch.
6. Review Peak periods per file.
7. Generate the Batch ZIP.

## Mapping Preset Vs Project Session

Mapping Preset (`.mapping.json`) stores reusable mapping rows only: raw sheet,
source stream, movement label, output movement code, include flags, and
aggregation fields. Use it when the same intersection mapping should be reused
across dates or workbooks.

Project Session (`.tmcproj.json`) stores broader job setup: metadata, mapping,
PCE factors, peak settings, and export settings. It does not embed the raw Excel
input file.

## Export Modes

Excel Template Mode uses Microsoft Excel COM on Windows to fill the report
template while preserving native charts, formulas, layout, and formatting.

Safe PNG Export Mode uses openpyxl and static chart images. It is the fallback
when Excel COM is unavailable.

## Privacy And Data Safety

This is a public repository. Do not commit real survey files, client data,
private project names, generated reports from real projects, private mapping
files, or private Project Sessions.

The repository is configured so these work areas stay ignored:

- `samples/raw/`
- `outputs/`
- generated `.tmcproj.json` files outside `samples/demo/`
- generated ZIP files

Export packages do not include raw input Excel files by default. Keep real survey
inputs and real outputs on local or private storage.

## Developer Verification

Useful checks before release:

```powershell
python -m py_compile app.py
python scripts/smoke_demo.py
python -m pytest
```

The smoke demo uses synthetic files from `samples/demo/`.

## License

This project is released under the MIT License. See [LICENSE](LICENSE).
