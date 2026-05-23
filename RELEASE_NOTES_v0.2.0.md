# TMC Processor v0.2.0 - Public Beta

## Highlights

- Single-file workflow for TMC Excel processing, mapping, PCU/PCE calculation, Peak review, and Excel report export.
- Batch workflow for processing multiple days for the same survey point with one shared Mapping Preset.
- Per-file Peak review in batch mode.
- Mapping Preset and Project Session save/load support.
- Editable PCE factors.
- Excel Template Mode and Safe PNG Export Mode.
- Export Package ZIP for single-file outputs.
- Batch ZIP with `batch_summary.xlsx` and `Batch_QC`.
- Synthetic demo workbooks and demo mapping files.
- Windows one-click launcher.

## How To Try It

1. Download the repository ZIP from GitHub or clone the repository.
2. On Windows, double-click `start_tmc_processor.bat`.
3. Wait for first-run setup to create `.venv` and install dependencies.
4. Try the single-file demo with `samples/demo/DEMO_TMC1_FourLeg.xlsx` and `samples/demo/DEMO_TMC1_FourLeg.mapping.json`.
5. Try the batch demo with both demo workbooks and the same Mapping Preset.

All demo data is synthetic.

## Known Limitations

- Batch v1 is intended for the same survey point or same intersection across multiple days using one shared Mapping Preset.
- No per-file mapping assignment yet.
- Excel Template Mode requires local Microsoft Excel / COM on Windows.
- Safe PNG mode is available as fallback when Excel COM is not available.
- Users still need to verify mapping and Peak periods before relying on final reports.

## Notes For Windows Users

- The first run of `start_tmc_processor.bat` may take several minutes.
- Later runs are faster because `.venv` already exists.
- Install Python 3.10 or newer and make sure `python.exe` is on `PATH`.
- Microsoft Excel desktop app is only required for Excel Template Mode.

## Privacy Note

Do not upload or commit real survey workbooks, client files, generated reports
from private projects, private mapping files, or private Project Sessions to the
public repository. Export packages do not include raw input Excel files by
default.

## Verification Performed

- `python -m py_compile app.py`
- `python scripts/smoke_demo.py`
- `python -m pytest`
- `git status`
- tracked-file scan for Excel, ZIP, project session, raw, legacy, output, and starter paths
- app launcher smoke via `start_tmc_processor.bat`
