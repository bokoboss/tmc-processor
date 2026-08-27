# Evidence Package - UX-2 Sheet-Centric Physical-Semantics Mapping

## Provenance

- Repository: `C:\MyRD\tmc-processor-public`
- Issue: [#5 - UX-2: sheet-centric physical-semantics Mapping workflow](https://github.com/bokoboss/tmc-processor/issues/5)
- PR: [#17 - Implement UX-2 sheet-centric physical-semantics Mapping workflow](https://github.com/bokoboss/tmc-processor/pull/17)
- Branch: `codex/ux-2-sheet-centric-mapping`
- Baseline SHA: `e70066e651c0c0d182307614a6ea03b7f241a416`
- Head SHA at implementation and CI capture: `c2d33286de525224f761a4e368d5bb7313baa77a`
- Date: 2026-08-27 (Asia/Bangkok)
- Working tree after evidence commit: clean; no protected raw workbook is tracked or staged.

## Execution route

- Model / effort: Codex GPT-5 execution on the local Windows workstation; no delegated workers.
- Orchestrator / workers if used: none.
- Retries / escalations: one browser-runtime reconnect after long-running analysis/export calls exceeded the 30-second tool wait; the Streamlit session and server state continued successfully.
- Reason for escalation if any: not applicable.

## Scope and protected behavior

UX-2 changes the Basic Mapping presentation to the workflow:

`actual source sheet -> operator physical interpretation -> derived canonical mapping`

Basic now presents actual source-sheet names and raw label/direction context as read-only information, Physical approach (`N/E/S/W`), Movement (`Left/Through/Right/U-turn`), report/Peak inclusion, and inline derived code/status. Advanced remains the technical editing surface. Existing canonical formats, Mapping Excel/Preset compatibility, normalization, WorkflowState revision/invalidation, calculation, QC, Peak, and export paths were preserved.

UX-3, UX-4, UX-5, UX-6, calculation methodology, PCE methodology, Peak methodology, export backend redesign, template redesign, Project Session schema changes, and automatic direction inference were not started.

## Validation evidence

| Gate | Command / Method | Result | Evidence / Artifact |
|---|---|---|---|
| Baseline alignment | `git fetch origin`; `git pull --ff-only origin main`; `git rev-parse HEAD`; `git rev-parse origin/main` | PASS - both resolved to `e70066e651c0c0d182307614a6ea03b7f241a416`; working tree clean before branching | Branch created from refreshed authoritative `origin/main` |
| Baseline regression | `python -m pytest` | PASS - 316 passed, 11 expected openpyxl DrawingML warnings, 224.62s | Accepted baseline before implementation |
| Focused UX-2/state suite | `python -m pytest tests/test_physical_mapping.py tests/test_phase_l0_dual_scheme.py tests/test_workflow_contract_adapter.py -q --disable-warnings` | PASS - 84 passed in 1.91s | Physical derivation, hydration, UI frame, metadata, imports, invalidation, view-only coverage |
| Full regression | `python -m pytest` | PASS - 334 passed, 11 expected openpyxl DrawingML warnings, 180.46s | Final implementation head `c2d3328` |
| Python compile | `python -m compileall -q app.py src` | PASS | No compile errors |
| Diff hygiene | `git diff --check` | PASS | No whitespace errors |
| CI | [GitHub Actions run #76](https://github.com/bokoboss/tmc-processor/actions/runs/33080837854) | PASS - completed/success; Python 3.10 and 3.12 jobs passed | Head `c2d33286de525224f761a4e368d5bb7313baa77a` |

## Implementation and focused test results

- Added pure physical-semantic helpers for all 16 `from_to` combinations and all 16 `approach_movement` combinations.
- `from_to` derivation follows the accepted left-hand-traffic semantics: N Left/Through/Right/U-turn -> `NE/NS/NW/NU`; equivalent E/S/W combinations are covered.
- Valid canonical codes hydrate Basic controls losslessly. Invalid or scheme-incompatible codes remain unchanged and receive `Advanced required` status, including excluded rows.
- Basic frame columns are limited to source sheet, raw direction/label, physical approach/movement, report/Peak inclusion, derived code, and status. Technical fields such as `source_stream`, `movement_code`, aggregation, facility, from/to, turn, and legacy metadata are excluded from Basic.
- Basic semantic edits derive canonical codes and pass through the existing scheme normalizers, preserving `from_leg`, `to_leg`, and `turn_type`.
- Basic edits preserve `source_stream`, `aggregation_method`, `facility_type`, optional scheme-specific data, notes, remarks, and compatible legacy metadata.
- Mapping Excel and Mapping Preset round trips hydrate Basic without changing equivalent canonical mappings.
- WorkflowState tests confirm a physical mapping edit invalidates Analysis, Review, and Export, while Basic/Advanced view changes leave revisions and artifacts unchanged.
- Default mapping tests confirm source labels such as numeric, directional, and Thai labels do not infer physical approach or movement.
- Batch Mapping uses the existing shared preset path and presents source matching/coverage plus a shared Basic-style preset preview; it does not introduce a second mapping model.

## Browser / UX / real-data evidence

The final-head app was run with the Windows Streamlit server at `http://127.0.0.1:8501` and exercised through the in-app Browser.

### Single workflow

- Uploaded protected Kabin Buri: 24 detected source sheets were displayed by actual sheet name (including combined sheets such as `ทิศ 2+3`); the initial default mapping left physical controls blank rather than inferring direction from names.
- Loaded a temporary deterministic qualification-only `from_to` preset: 24/24 sheets matched with 0 missing and 0 extra rows.
- Basic hydrated the first canonical `NE` row as `N + Left` with inline `NE` and `Ready` status.
- Switched Basic -> Advanced -> Basic. The physical selection and derived code persisted; Advanced showed the technical movement-code editor.
- Analyze completed after the deterministic mapping: 13,824 normalized rows, 82,621 vehicles, 87,023.5 PCU, 0 QC errors, 0 QC warnings, and 8 initial QC info findings. Suggested/effective peaks were AM `07:00-08:00` and PM `17:00-18:00`.
- Changed the first Basic row from `N + Left` (`NE`) to `N + Through` (`NS`). Analyze then showed `Re-analysis required`; re-analysis completed successfully with the downstream Review flow available.
- Safe PNG Export Mode completed. Downloaded artifact: `C:\Users\kitti\Downloads\tmc_processor_output_package.zip`, 1,592,061 bytes, 8 entries including report workbook, summary, session, mapping, and three PNG visuals.
- Excel Template Mode completed with Excel COM available at version 16.0. Downloaded artifact: `C:\Users\kitti\Downloads\tmc_processor_output (3).xlsx`, 1,138,581 bytes, 19 sheets including `Diagram_Data`, `Report_Text`, and `Peak_PHF`.
- `Peak_PHF` in the downloaded workbook retained `fixed_hourly`, search windows `07:00-12:00` and `15:00-19:00`, effective peaks `07:00-08:00` and `17:00-18:00`, and `auto_suggested` selection source.

### Mapping imports

- Loaded the existing demo Mapping Excel file `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx`; Basic hydrated the first `NS` row as `N + Through` and retained its source/metadata path.
- Mapping Excel and Mapping Preset hydration plus advanced/legacy metadata preservation are covered directly by `tests/test_physical_mapping.py`.

### Batch workflow

- Uploaded all three protected workbooks: Kabin Buri, Bo Phloi, and Nong Prue. Data retained exactly 3 source files.
- Applied one shared deterministic qualification-only preset containing the union of detected source-sheet labels. Source coverage was 24/24 for Kabin Buri, 15/15 for Bo Phloi, and 12/12 for Nong Prue; the latter two correctly reported 9 and 12 preset-only rows as extras, with 0 missing detected sheets.
- Batch Analyze completed for 3/3 files with 0 failed files. Review showed 3/3 Peak assignments. Selected Kabin Buri matched the Single run at 82,621 vehicles and 87,023.5 PCU, with 0 QC errors, 0 QC warnings, and 8 QC info findings.
- Batch Safe PNG Export completed. Downloaded artifact: `C:\Users\kitti\Downloads\UX-2_Deterministic_Qualification_Only_batch_package.zip`, 3,994,978 bytes, 22 entries including `batch_summary.xlsx`, 3 report workbooks, 3 session files, 3 mapping files, and 9 PNG visuals.
- `batch_summary.xlsx` reported `success` and `Safe PNG Export Mode` for all three files; the package contained no raw source workbooks.

## Protected real-workbook qualification

The protected raw workbooks under `samples/raw/` were used only as local runtime inputs. The approved physical-direction Mapping Presets were not available in the execution environment. The qualification preset was synthetic and deterministic: it assigned valid canonical codes from a stable sheet-inventory list and did not infer physical directions from numeric or textual sheet names. It was removed after UAT and is not part of the branch.

Therefore these runs qualify ingestion, source-sheet presentation, Mapping persistence, analysis, QC, Peak, and export mechanics only. They are not engineering approval of the physical direction mapping. The approved physical Mapping Preset and normal mapping/QC review remain required for release use.

## CI and review status

- PR [#17](https://github.com/bokoboss/tmc-processor/pull/17) is open against `main`, is mergeable, and includes `Closes #5`.
- The PR was not merged.
- GitHub Actions run [#76](https://github.com/bokoboss/tmc-processor/actions/runs/33080837854) completed successfully for the implementation head.
- Status: **READY FOR REVIEW**.
- Release/merge limitation: do not treat the deterministic physical mapping qualification as approval of real-world direction semantics until the approved preset is available and reviewed.

## Changed files and diff summary

- `app.py` - Basic physical-semantics editor, inline derived/status presentation, progressive Advanced controls, shared Batch preview/coverage.
- `src/tmc_processor/mapping.py` - Pure physical-to-canonical derivation and canonical-to-physical hydration helpers.
- `tests/test_physical_mapping.py` - Dedicated UX-2 physical, hydration, compatibility, invalidation, and view-only coverage.
- `tests/test_phase_l0_dual_scheme.py` - Updated Basic visible-column contract.
- `docs/development/EVIDENCE_PACKAGE_UX2.md` - This evidence package.

Implementation commit diff: 4 files changed, 728 insertions, 79 deletions. Evidence documentation is the follow-up artifact commit and is intentionally separate from the implementation commit.

## Known limitations / follow-up

- Approved physical-direction Mapping Presets were unavailable; physical direction decisions remain operator-owned and unapproved in this qualification.
- Basic supports the requested four-leg physical semantics. Non-lossless/custom canonical codes remain Advanced-only and are preserved.
- The existing `approach_movement` v2 pipeline remains intentionally blocked for production processing by its pre-existing contract; UX-2 derives and hydrates its v2 codes without converting schemes.
- Native Excel COM/template behavior was smoke-qualified where available; no COM/chart/template redesign was made.
