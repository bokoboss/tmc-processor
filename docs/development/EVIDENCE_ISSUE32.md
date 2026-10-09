# Issue #32 V2 Authored Summary Evidence

## Scope

Populate the original `Summary` worksheet from `templates/four_leg_tmc_report_template_approach_v2.xlsx` for Single and Batch V2 exports. The implementation writes through direct OOXML package patching, keeps the authored intersection shapes, arrows, charts, layout, styles, and formulas, and adds a source stream audit for duplicate movement codes. V1 export behavior remains on its existing path.

## Baseline

- Branch: `codex/issue-32-v2-xlsx-summary`
- Correction started from `ef09106f73c318581a26d94bc8bb68d3d31b9231`, the rejected Issue #32 implementation on the existing task branch.
- The branch was clean before this correction. Its implementation parent is `48db746a93c8b20f25079fa1839bffae1e8fab34` (`main` verified at task start).
- The evidence workbook uses reduced synthetic counts derived from the repository demo input. No sensitive workbook data was used.

## Evidence artifacts

- Original template opened in Excel: `outputs/issue32_original_v2_template_summary_excel.jpg`
- Populated output opened in Excel: `outputs/issue32_populated_v2_template_summary_excel.jpg`
- Actual synthetic XLSX: `outputs/issue32_synthetic_v2_template_summary_output.xlsx`
- OOXML structural and peak reconciliation report: `outputs/issue32_v2_template_structure_and_reconciliation.txt`
- Per-movement and per-stream reconciliation: `outputs/issue32_v2_movement_peak_reconciliation.csv`

The output opened in desktop Excel without a repair prompt. The output Summary retains 33 authored shapes, two connectors/arrows, and two chart frames. Drawing XML, drawing relationships, and Summary worksheet relationships are byte-identical to the source template. Chart definitions match after removing their updated data caches. All 5,959 original Summary cell style assignments, all 30 merged ranges, and original style records are preserved. The map-authorized report title in `Summary!B2` replaces its template formula; the other 120 formulas remain unchanged. The output ZIP passes integrity checks. `xl/calcChain.xml` is removed and 19 support worksheets are added.

The synthetic source stream audit reconciles 727 counts and 554.914 PCU to normalized movement rows. Duplicate `NT` values remain separately attributed to `frontage` and `mainline`. The confirmed AM Peak is 08:00–09:00 at 83.654 PCU; PM is 17:00–18:00 at 87.987 PCU. The authored Summary displays whole-number values (84 and 89 in the roll-up), while `Peak_PHF` and `Movement_Source_Stream_Audit` retain precise values.

## Validation

- `python -m pytest -q tests/test_phase_f_v2_dry_run.py`: 30 passed, 8 openpyxl DrawingML read warnings.
- `python -m pytest -q tests/test_batch.py::test_v2_batch_template_zip_contains_authored_summary_without_raw_inputs tests/test_batch.py::test_v2_batch_confirmed_peak_override_is_used_in_export_summary tests/test_batch.py::test_v2_batch_excel_template_mode_is_supported`: 3 passed, 3 openpyxl DrawingML read warnings.
- `python -m pytest -q tests/test_phase_j_ui_helpers.py::test_v2_single_ui_export_uses_authored_template_without_excel_com tests/test_phase_j_ui_helpers.py::test_v2_single_template_mode_does_not_call_excel_com tests/test_phase_j_ui_helpers.py::test_v2_single_export_never_selects_excel_com_backend`: 3 passed after the workflow metadata assertion.
- `python -m pytest -q tests/test_workflow_state.py`: 22 passed.
- `python -m pytest -q tests/test_ooxml_template_export.py tests/test_phase_j_ui_helpers.py --basetemp=outputs/pytest-issue32-v1-rerun`: 42 passed, 2 openpyxl DrawingML read warnings.
- `python -m compileall -q app.py src`: passed.
- `git diff --check`: passed (Git emitted only its LF-to-CRLF working copy warning for the evidence Markdown file).
- The full 513-test suite was not run, per request.

## Changed files

- `src/tmc_processor/batch.py`
- `src/tmc_processor/diagram.py`
- `src/tmc_processor/export_package.py`
- `src/tmc_processor/exporter.py`
- `src/tmc_processor/ooxml_template_export.py`
- `src/tmc_processor/template_write_plan.py`
- `src/tmc_processor/template_write_verify.py`
- `src/tmc_processor/ui/app_shell.py`
- `src/tmc_processor/ui/components/export.py`
- `src/tmc_processor/ui/workflows/batch.py`
- `src/tmc_processor/ui/workflows/single.py`
- `tests/test_batch.py`
- `tests/test_phase_f_v2_dry_run.py`
- `tests/test_phase_j_ui_helpers.py`
- `docs/development/EVIDENCE_ISSUE32.md`

## Limitations

- The approved template displays movement and roll-up cells as whole PCU values. Precise totals and Peaks remain available in `Movement_Source_Stream_Audit` and `Peak_PHF`; the Summary display can differ from exact source totals within the twelve hourly rows' rounding bound.
- `openpyxl` warns that its reader does not fully support DrawingML. It was used read-only for data validation; the output is written and verified as OOXML, and was opened in Excel for visual inspection.
- GitHub CI and public deployment verification remain pending because this correction has not been pushed, merged, or released.
- Full local regression was not run as requested.
