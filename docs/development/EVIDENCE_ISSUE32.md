# Issue #32 V2 XLSX Summary Evidence

## Scope

Add a generated V2 `Summary` worksheet for Single and Batch exports. It contains an embedded four-leg conceptual movement schematic, confirmed AM/PM peak results, and source-stream-level movement values. Existing worksheets, V2 calculations and mappings, and the V1 export path remain in place.

## Baseline

- Task branch: `codex/issue-32-v2-xlsx-summary`
- Base: verified `main` / `origin/main` at `48db746a93c8b20f25079fa1839bffae1e8fab34`
- Working data for the generated example: synthetic demo workbook and V2 mapping; no sensitive raw workbooks used.

## Validation

- `python -m pytest tests/test_phase_f_v2_dry_run.py tests/test_batch.py::test_v2_batch_confirmed_peak_override_is_used_in_export_summary -q`: 31 passed, 4 pre-existing openpyxl DrawingML warnings.
- `python -m pytest tests/test_ooxml_template_export.py tests/test_phase_j_ui_helpers.py -q --basetemp=outputs/pytest-issue32-v1-20261009`: 42 passed, 2 pre-existing openpyxl DrawingML warnings.
- `python -m compileall -q app.py src`: passed.
- `git diff --check`: passed.
- Synthetic XLSX inspection: `ZipFile.testzip()` returned no corrupt member; openpyxl loaded all 14 worksheets; `Summary` is first and contains one embedded image; drawing, worksheet relationship, and PNG media entries are present. All 16 canonical movement codes reconcile to stream details, including zero-valued movements; source-stream details reconcile to normalized movement totals and AM/PM peak totals. Both peaks are marked `user_confirmed`.
- Excel desktop: opened the generated workbook and visually inspected the Summary schematic. The screenshot is `outputs/issue32_synthetic_v2_summary_excel.jpg`.

## Changed files

- `src/tmc_processor/diagram.py`
- `src/tmc_processor/exporter.py`
- `src/tmc_processor/ui/components/export.py`
- `src/tmc_processor/ui/workflows/batch.py`
- `src/tmc_processor/ui/workflows/single.py`
- `tests/test_batch.py`
- `tests/test_phase_f_v2_dry_run.py`
- `docs/development/EVIDENCE_ISSUE32.md`

## Known limitations and follow-up

- The intersection is a conceptual, not-to-scale rendering derived from V2 direction and L/T/R/U movement semantics. It does not claim measured road geometry.
- Source stream is preserved in the movement table so duplicate movement codes from mainline and frontage remain distinguishable.
- GitHub CI and the cloud/deployed smoke check require the reviewed branch to be pushed and deployed; they are pending and were not run locally.
- Full 513-test local suite intentionally not run per task instructions.
- The generated test workbook is `outputs/issue32_synthetic_v2_summary_final.xlsx`.
- Final commit SHA is reported with the delivery; this evidence file is committed with that implementation.
