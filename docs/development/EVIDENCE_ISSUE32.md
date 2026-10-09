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

## Final review remediation (2026-10-09)

Authoritative review: https://github.com/bokoboss/tmc-processor/issues/32#issuecomment-6084311740.
This correction addresses only the missing supplemental package PNG and the incorrect diff baseline.

Baseline: clean `codex/issue-32-v2-xlsx-summary` at `d21f47afd55dfa5516d4fbbee577b67f31526f6f`.
The immutable Issue #32 base is `48db746a93c8b20f25079fa1839bffae1e8fab34`.
`git merge-base --is-ancestor 48db746a93c8b20f25079fa1839bffae1e8fab34 HEAD` succeeds,
and `git merge-base` returns that exact base. The stale local `main` pointer is not used or changed.
The existing executor context is retained for this small correction; no additional coding agent is invoked.

Success gates: valid supplemental PNGs at the original Single and per-file Batch ZIP paths;
truthful metadata and package previews; byte-preserved native Summary drawings and workbook
bytes during packaging; focused V1/V2 package/export regressions; compile and whitespace checks.

Restored the existing `render_v2_movement_diagram_png` renderer and
`diagram/movement_diagram.png` member in Single and Batch packages. This remains the existing
supplemental table-style PNG. The native Excel Summary remains the approved authored intersection.
Restored `diagram_png_package_path` metadata and Single/Batch package previews. No template,
Summary population, native shapes/arrows/charts, calculation, Peak confirmation, or stream-audit
implementation was changed. The exporter change is limited to the two package-path metadata values.

### Focused validation

First updated the Single and Batch package assertions: both failed on the missing PNG ZIP member.
After restoring the implementation, selected 54 package/export tests with this command:

```text
.venv/Scripts/python.exe -m pytest -q tests/test_export_package.py tests/test_metadata_export.py tests/test_ooxml_template_export.py tests/test_phase_f_v2_dry_run.py::test_v2_generated_package_excludes_raw_inputs_and_includes_summary tests/test_phase_f_v2_dry_run.py::test_v2_template_helper_uses_direct_ooxml_and_preserves_authored_drawing tests/test_phase_f_v2_dry_run.py::test_v2_template_export_preserves_hlookup_formula_structure tests/test_phase_j_ui_helpers.py::test_v2_single_ui_export_uses_authored_template_without_excel_com tests/test_phase_j_ui_helpers.py::test_v2_single_template_mode_does_not_call_excel_com tests/test_phase_j_ui_helpers.py::test_v2_single_export_never_selects_excel_com_backend tests/test_batch.py::test_v2_batch_template_zip_contains_authored_summary_without_raw_inputs tests/test_batch.py::test_v2_batch_confirmed_peak_override_is_used_in_export_summary tests/test_batch.py::test_v2_batch_excel_template_mode_is_supported tests/test_batch.py::test_batch_zip_contains_summary_and_one_folder_per_success_without_raw_inputs tests/test_batch.py::test_batch_zip_contents_preview_lists_expected_artifacts tests/test_batch.py::test_custom_confirmed_peak_overrides_suggested_in_final_zip --tb=short
```

That run had 24 passes, 28 setup errors because pytest could not create its sandbox temporary
directory, and two new assertion failures referencing a nonexistent `Mapping_Scheme_Info` sheet.
Corrected the assertions to check the existing `Export_Metadata` worksheet. Reran only the
affected files/tests with a repository-local temporary directory:

```text
.venv/Scripts/python.exe -m pytest -q tests/test_ooxml_template_export.py tests/test_phase_f_v2_dry_run.py::test_v2_generated_package_excludes_raw_inputs_and_includes_summary tests/test_batch.py::test_v2_batch_template_zip_contains_authored_summary_without_raw_inputs --tb=short --basetemp=outputs/pytest-issue32-package-remediation -p no:cacheprovider
```

Result: **33 passed**, three known openpyxl DrawingML read warnings. Combined with the first
run's 24 passes (three overlap), **all 54 unique selected tests pass**. Tests validate actual
Single/Batch ZIPs, decode and verify their PNGs, preserve packaged XLSX bytes, verify metadata
paths and Batch previews, and compare the native Summary drawing to the template. Existing
focused tests cover confirmed Peaks, chart/formula fidelity and unchanged V1 package exports.

- `.venv/Scripts/python.exe -m compileall -q app.py src`: PASS.
- `git diff --check`: PASS.
- `git diff --quiet d21f47afd55dfa5516d4fbbee577b67f31526f6f -- templates src/tmc_processor/diagram.py src/tmc_processor/ooxml_template_export.py src/tmc_processor/template_write_plan.py src/tmc_processor/template_write_verify.py`: PASS.

Correction files: `src/tmc_processor/batch.py`, `src/tmc_processor/export_package.py`,
`src/tmc_processor/exporter.py`, `src/tmc_processor/ui/workflows/single.py`,
`tests/test_batch.py`, `tests/test_phase_f_v2_dry_run.py`, `tests/test_phase_j_ui_helpers.py`,
and this evidence document.

The regenerated binary patch is `outputs/issue32_48db746a_to_head.patch`, created with
`git diff --binary 48db746a93c8b20f25079fa1839bffae1e8fab34 HEAD` after the correction commit.
`outputs/issue32_diff_summary.txt` records exact HEAD, branch, ancestry, status, changed paths,
and diff statistics. The final review ZIP replaces the superseded local-main patch with these
immutable-base artifacts and updated evidence. The whole Issue #32 diff has 14 changed paths;
PR #31 Support UI, QR asset, requirements and CI changes are absent from that diff.

### Evidence limits

The previously reviewed synthetic XLSX, Excel screenshots, structural TXT and reconciliation CSV
are preserved from the `d21f47a` checkpoint; their workbook metadata predates this PNG restoration.
Current metadata and package behavior are validated by the focused tests above. Excel screenshots
were not recaptured because the native Summary implementation is unchanged. The synthetic
`NT/frontage` row's raw label still contains `mainline`: the fixture overrides `source_stream`
to demonstrate stream separation without renaming that raw label; it is not a production mapping rule.
No sensitive input was used. No full local suite, push, merge or release was performed.

## Legacy worksheet compatibility correction (2026-10-10)

Review: https://github.com/bokoboss/tmc-processor/issues/32#issuecomment-6084971985.
Started from clean `f614abbd4535caf0724881eab5a35095a9647b6e` on the existing Issue #32 branch.
The user authorized this bounded correction and, after focused PASS, pushing only
`codex/issue-32-v2-xlsx-summary` and opening a PR to `main`. Merge, tags and release remain outside scope.
Remote `main` was verified at immutable `48db746a93c8b20f25079fa1839bffae1e8fab34` before editing.

Compared the generated V2 exporter at that immutable base. Its `Hourly_Totals` uses the
original hourly dataframe (`hour_start`, `hour_end`, `count`, `pcu`); `Peak_Summary` uses
resolved effective Peak rows; `Movement_Code_Reference` uses the canonical V2 reference helper;
and `Mapping_Scheme_Info` uses the mapping metadata helper (`field`, `value`). The reference
and Peak helpers remain unchanged from that base. Mapping metadata retains the original fields
and mapping counts while truthfully describing the approved native OOXML export and restored PNG.

Implementation: four additional entries in `_v2_template_export_sheets`; no other runtime changes.
The regression test first failed on the four missing sheets, then passed after their restoration.
It checks all 13 original V2 sheet names plus `Summary`, the four restored column contracts,
hourly count/PCU reconciliation, confirmed AM/PM Peak values and provenance, canonical movement
reference values, mapping counts and the existing PNG package path.

Success gates: original sheet-name/column/data compatibility, retention of all existing sheets
and native Summary/package artifacts, focused export regressions, compile and whitespace checks.

```text
.venv/Scripts/python.exe -m pytest -q tests/test_phase_f_v2_dry_run.py tests/test_export_package.py tests/test_batch.py::test_v2_batch_template_zip_contains_authored_summary_without_raw_inputs tests/test_batch.py::test_v2_batch_confirmed_peak_override_is_used_in_export_summary tests/test_batch.py::test_v2_batch_excel_template_mode_is_supported tests/test_phase_j_ui_helpers.py::test_v2_single_ui_export_uses_authored_template_without_excel_com tests/test_phase_j_ui_helpers.py::test_v2_single_template_mode_does_not_call_excel_com tests/test_phase_j_ui_helpers.py::test_v2_single_export_never_selects_excel_com_backend --tb=short --basetemp=outputs/pytest-issue32-legacy-focused -p no:cacheprovider
```

Result: **43 passed**, 11 known openpyxl DrawingML read warnings, 136.36 seconds.
`.venv/Scripts/python.exe -m compileall -q app.py src` and `git diff --check`: **PASS**.
The full local suite was not run; GitHub PR CI will provide full regression evidence.

An identical-input comparison with the exporter loaded from `f614abb` produced 20 prior and
24 current worksheets. Only the four requested sheets were added. All 20 existing worksheet
values/formulas are unchanged. Summary XML, its relationships, drawing XML/relationships,
both chart XMLs and styles are byte-identical; package PNG and movement CSV are byte-identical;
both XLSX ZIPs pass integrity checks. Report: `outputs/issue32_backward_compatibility_comparison.txt`.
Existing approved screenshots and synthetic workbook are historical fidelity evidence and were
not regenerated for this four-sheet addition.

Correction files: `src/tmc_processor/exporter.py`, `tests/test_phase_f_v2_dry_run.py`, and this
evidence document. The immutable-base patch and `outputs/issue32_diff_summary.txt` are regenerated
after committing. The full Issue #32 diff still has 14 changed paths; no unrelated PR #31 files.
No changes to templates, native Summary writer/verifier, calculations, stream audit, ZIP packaging,
PNG renderer or UI are included in this correction. CI and Cloud smoke remain acceptance gates;
this evidence does not authorize merge or release.

## PR #33 Linux startup correction (2026-10-10)

Review: https://github.com/bokoboss/tmc-processor/pull/33#issuecomment-6091117631.
Started from clean `faa32ea0b47c549a5b6d4b78812e2e1bca3c894b` on the existing PR branch.
The real `_workflow_operations()` factory reproduced CI's unexpected-keyword TypeError.
Added exactly `V2_TEMPLATE_MAP_PATH`, `V2_TEMPLATE_PATH`, and `MOVEMENT_SCHEME_V1` to
`WorkflowOperations`, using the existing `Any = None` field convention. A focused regression
constructs the actual factory and verifies all three values. No export/template/worksheet,
PNG packaging, calculation or other UI implementation changed.

Success gates: real factory construction, actual `app.py` entrypoint without exceptions,
focused context/architecture regressions, compile and whitespace checks.

- `.venv/Scripts/python.exe -m pytest -q tests/test_application_architecture.py tests/test_workflow_contract_adapter.py --tb=short --basetemp=outputs/pytest-issue32-startup-focused -p no:cacheprovider`: **41 passed**.
- The exact CI AppTest lines ran against `app.py`:

```python
from streamlit.testing.v1 import AppTest
app = AppTest.from_file("app.py", default_timeout=60).run()
assert not app.exception, [error.message for error in app.exception]
print("app.py executed without exceptions")
```

Result: **PASS**, no app exceptions. The Windows sandbox initially blocked the harness's asyncio
socket-pair creation before executing the app (confirmed by a stack trace). The assertion passed
outside the sandbox. Optional `win32com`, `pythoncom`, and `pywintypes` imports were unavailable
only within this test process, matching CI's no-COM condition; no installed packages or source
were changed. Actual Linux confirmation remains the new GitHub CI run on the pushed SHA.

- `.venv/Scripts/python.exe -m compileall -q app.py src`: **PASS**.
- `git diff --check`: **PASS**.

Correction files: `src/tmc_processor/ui/workflow_context.py`,
`tests/test_application_architecture.py`, and this evidence document. No full local suite rerun.
Push is limited to the existing PR branch; no merge, tag or release is authorized.
