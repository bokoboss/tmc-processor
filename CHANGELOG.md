# Changelog

## v0.2.0 - Public beta

### Added

- Single-file TMC workflow for uploading, mapping, processing, Peak review, and Excel report export.
- Batch workflow for processing the same survey point or intersection across multiple days with one shared Mapping Preset.
- Per-file Peak review in batch mode before generating final batch outputs.
- Mapping Preset support for reusable `.mapping.json` mapping files.
- Project Session support for saving and loading broader job setup in `.tmcproj.json` files.
- Editable PCE factors for PCU calculations.
- Excel Template Mode and Safe PNG Export Mode.
- Export Package ZIP for single-file outputs.
- `batch_summary.xlsx` with `Batch_QC` in Batch ZIP outputs.
- Synthetic demo workbooks, demo mapping Excel, demo Mapping Preset, and demo Project Session.
- Windows one-click launcher through `start_tmc_processor.bat`.

### Improved

- Unified single/batch work-mode layout so users can move between workflows more clearly.
- Release documentation for non-technical Windows users, demo workflows, privacy expectations, and batch limitations.
- Demo guidance for Mapping Preset versus mapping Excel usage.

### Fixed / Hardened

- Export package guidance clarifies that raw input Excel files are not included by default.
- Batch workflow documentation now calls out same-intersection/shared-preset assumptions.
- Version metadata now targets `0.2.0` for the public beta.

### Notes

- Batch v1 is designed for multiple days at one survey point or intersection using one shared Mapping Preset.
- Excel Template Mode requires local Microsoft Excel/COM on Windows.
- Safe PNG Export Mode remains available as a fallback.
