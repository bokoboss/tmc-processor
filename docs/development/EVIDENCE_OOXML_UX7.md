# UX-7 OOXML Template Preservation Evidence

> **Chronological investigation record.** Early three-workbook outputs below used a synthetic mapping projection and are not accepted operator Mapping/Review fixtures or backend parity oracles. The accepted checkpoint is deterministic post-Review native-template parity, including manual Excel fidelity and U9:U22 = 1..14. The final UX-7 acceptance decision and exact-branch real-workbook repeat waiver are recorded in `EVIDENCE_PACKAGE_UX7.md`; retain the historical entries here only to explain the path to that result.

Date: 2026-09-27 (Asia/Bangkok)
Checkout: `codex/ux-7-release-qualification`
Manual Excel fidelity: **PENDING — current Excel process/user environment not disturbed.**

This qualification does not change Standard Report routing. It uses the accepted deterministic sheet-inventory mapping as a software qualification fixture. The mapping is synthetic and does not approve physical movement directions.

## COM parity contract

| Behavior | COM data and destination | Sheet/formula behavior | Chart/recalc dependency | Classification |
|---|---|---|---|---|
| Summary metadata | `metadata_cells` in the template map: title B2; project E4; survey point E5; date K5; weather Q5; responsible party E6; survey period K6. Movement title/date C8/C9; direction labels G12/O32/Q17/D26; road labels K11/K32/R19/D24; caption E35. | Writes mapped cells; B2 is the explicitly permitted formula overwrite. Other formula cells are guarded. | None directly; Excel recalculates retained formulas. | Product report content |
| Effective Peak label | A3 includes AM start/end, PM start/end, and `peak_selection_source`. | Direct overwrite is part of the export contract. | Excel recalculates workbook formulas. | Product report content |
| Movement summary | 16 canonical movements use the mapped total/PM/AM cells in the four approach tables on Summary. | Current map sets `formula_write_mode=preserve_template`; native formulas stay intact. | Existing formulas read `Diagram_Data`. | Product report content |
| Summary totals | AM total F30, PM total F31, 12-hour total F32. | Current map sets `summary_formula_write_mode=preserve_template`; formula nodes stay intact. | Existing formulas reference `Peak_PHF` and `Diagram_Data`. | Product report content |
| Hourly movement table | Summary V8:AM22; 12 time rows plus total, movements and total PCU. | Fills mapped headers/data; protected formula cells remain protected. | Native hourly chart reads V10:V21 and AM10:AM21. | Product report content |
| Vehicle-class table | Summary V26:AJ40; 12 time rows plus total, PCU, vehicles, 12 classes. | Fills mapped headers/data; protected formula cells remain protected. | Native composition chart reads AP39:BA39 and AP40:BA40. | Product report content |
| Support report surfaces | Adds every frame in `report_data["sheets"]`: Export_Metadata, Setup, PCE_Factors, Mapping, Movement_Aggregation_Audit, Normalized_Data, QC_Check, Hourly_Summary, Movement_Summary, Vehicle_Composition, Hourly_Movement_PCU, Hourly_Vehicle_Class, Vehicle_Composition_Report, Vehicle_Group_PCE, PHF_15min, Peak_PHF, Report_Text. | Replaces same-named support sheets; header bold, white on blue, columns autofit. New sheets are visible by default. | Hourly/Peak source tables feed formulas and charts. | Product report content |
| Diagram_Data | Appends four columns (movement_code, total_pcu, pm_peak_pcu, am_peak_pcu), one row per standard movement. | PM/AM formulas read Hourly_Movement_PCU and the effective PM/AM start in Peak_PHF H2/D2. | Recalculated by Excel. | Product report content |
| Chart sources/caches | Writes mapped category/value ranges into Summary; COM recalculation refreshes cached chart results. | Existing chart definitions stay in place. | Excel recalculation refreshes formulas and chart caches. | Product report content |
| Visibility | COM-created/replaced support worksheets use Excel's default visible state; it does not hide them. | Existing Summary sheet identity is retained. | None. | Product behavior |
| Calculation | Sets application calculation to automatic; sets ForceFullCalculation; calls CalculateFullRebuild and Calculate. | OOXML sets calcMode=auto, fullCalcOnLoad=1, forceFullCalc=1 and removes the stale calcChain part/relationship/content-type entry. Formula caches are cleared. | Excel recalculates on open. | Product behavior |
| Value coercion | pandas NA -> empty; pandas Timestamp/datetime -> datetime; date -> ISO text; time -> HH:mm. | Summary strings beginning `=` become formulas only when authorized. | Date serials remain numeric OOXML dates when style applies. | Product report content |
| COM process controls | Hidden Excel instance, alerts/events/link prompts, COM teardown. | Not report content; no OOXML representation. | COM runtime detail. | Implementation detail |

## OOXML implementation and package proof

`src/tmc_processor/ooxml_template_export.py` patches ZIP package parts directly and does not use Excel COM or save through openpyxl. It resolves the Summary worksheet from workbook relationships, preserves the template's existing sheet ID/rId, writes only map-authorized Summary cells, adds deterministic support worksheet IDs/relationships/content types, and appends one header style while retaining all existing style records. It patches only chart cache nodes after filling their referenced Summary ranges; it does not recreate charts or drawings.

`scripts/inspect_template_package.py --compare-output` now emits a per-part SHA-256 manifest with all seven requested classifications. The qualification manifests have no `UNEXPECTED_*` parts. For each output: 8 expected existing parts changed, 18 support worksheet parts were added, and `xl/calcChain.xml` was removed. Template parts classified unchanged include styles-independent content, all existing relationship targets except the removed calcChain relationship, drawings, media, theme, and chart style/color parts. The `xl/styles.xml` change appends a support-header font/fill/XF; existing style records remain present.

Native chart references remain `Summary!$V$10:$V$21` / `Summary!$AM$10:$AM$21` and `Summary!$AP$39:$BA$39` / `Summary!$AP$40:$BA$40`. Both chart caches are updated from the current report payload because their formulas point at newly written Summary source values. Chart definitions, drawing relationships, chart style/color parts, and the 33 custom-shape definitions are retained. The output requests recalculation because chart formula caches and Summary formula caches are not Excel-calculated by the OOXML writer.

## Three real-workbook qualification runs

One shared deterministic mapping was built from the stable union of the three detected sheet inventories (24 mapping rows). It assigns valid v1 codes in stable order and is only a software fixture. Application Peak mode and suggestion authority were used; provenance is accurately recorded as `auto_suggested`.

| Sample | Source SHA-256 | Normalized rows | QC | Suggested/effective AM | Suggested/effective PM | Output bytes / SHA-256 |
|---|---|---:|---|---|---|---|
| Kabin Buri | `f68f084c6a1a822714ac2c80051575dbab67531c1107b24e27688ad2d367c94b` | 13,824 | 8 info | 07:00–08:00 (`auto_suggested`) | 17:00–18:00 (`auto_suggested`) | 1,153,080 / `3195018727f6adce96b26108a8acff5650cae0ea4d45c25b11bf356fb9b9a67e` |
| Bo Phloi | `5ca76e635b20ac176c9b41115a9b1a139d23f0ae9fe6237c8b9dd0c41bb2c3ef` | 8,640 | 6 info | 08:00–09:00 (`auto_suggested`) | 17:00–18:00 (`auto_suggested`) | 733,408 / `e2e7768a5a5ced99a581aeca3cfcfd3ec9f67358d122a40d0f953587d47d70d5` |
| Nong Prue | `c6890839f1d4e8c0424fb9bb5740e60c7379e0cc783e69cbec75cbed7140d695` | 6,912 | 4 info | 11:00–12:00 (`auto_suggested`) | 16:00–17:00 (`auto_suggested`) | 597,939 / `e72e78ba5fd57bbc2c75355ce3bcf78c897e3ed330416e0819cbef3dfee27021` |

Each output opened read-only with openpyxl, passed ZIP/XML parsing and `audit_template`, retained Summary, both native charts and 33 shapes, contained all 18 report support sheets plus Diagram_Data, and had no unexpected package changes, missing internal relationship targets, external-link parts, or `#REF!` tokens. Peak start/end agree between Summary A3, Export_Metadata, and Peak_PHF; Diagram_Data formulas reference the same Peak_PHF AM/PM start cells. Per-part manifests and output workbooks are under `%TEMP%\ux7_ooxml_real_qualification\` and are not committed.

## Capability preservation matrix

| Capability | Prior COM behavior | OOXML evidence | Status |
|---|---|---|---|
| Template layout | Writes mapped report values into the Excel-authored Summary layout. | Summary part retained and only mapped cells/chart sources changed; automated package proof. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Formatting | Retains Summary formatting; formats support headers and autofits columns. | Existing cell style IDs and style records remain; one support-header style appended. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Formulas | Guards protected formulas; B2 is an allowed formula overwrite; forces recalculation. | Protected Summary formula XML retained, B2 authorized overwrite verified; caches cleared and calc flags set. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Charts | Fills existing chart source ranges and calculates in Excel. | Both chart parts/references retained; caches updated from report data. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Shapes/drawings | Uses existing template drawings. | Drawing parts/relationships byte-identical; 33 custom shapes retained. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Support sheets | Writes 17 report frames plus Diagram_Data. | All 18 named worksheets created, visible, linked, and readable. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Metadata | Writes mapped metadata and effective Peak provenance. | Mapped fields, Export_Metadata values, A3 label, and Peak_PHF agree. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Movement data | Writes mapped hourly PCU and vehicle-class tables. | Three real output payloads verified against Summary tables and support data. | AUTOMATED-PROVEN / MANUAL-PENDING |
| PHF | Writes Peak_PHF report frame and recalculates formulas. | Peak_PHF fields/provenance retained and internally consistent. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Diagram_Data | Writes formula-driven total, PM, and AM PCU values. | Formula links resolve to hourly movement and effective Peak source cells. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Effective Peak/provenance | Uses app-resolved values/source; writes A3 and support surfaces. | Automatic suggestions, support values, and Diagram_Data references agree across all samples. | AUTOMATED-PROVEN / MANUAL-PENDING |
| Workbook relationships | Excel manages new sheets and package links. | IDs/targets validated; Summary identity unchanged; no collisions/missing targets. | PRESERVED |
| Recalculation request | Automatic mode, full rebuild, calculate. | OOXML automatic/full/force flags; stale calcChain removed. | PRESERVED |
| Safe PNG coexistence | Separate fallback backend remains available. | No Standard routing change; existing Safe PNG/COM tests run in final validation. | PRESERVED |

## Manual gate

The previously accepted one-cell POC remains structurally successful at `C:\Users\kitti\AppData\Local\Temp\tmc-ooxml-poc-a5bc078480d644199bdab587adf97035\template_one_cell.xlsx`. Its manual Excel fidelity is **UNVERIFIED**. No Excel process was inspected, modified, terminated, or relaunched for this checkpoint. Visual rendering of layout, charts, shapes, fonts, merged areas, and formula results still requires a later manual Excel inspection. This evidence does not claim release readiness and does not change Standard Report routing.

## Kabin Buri vehicle table and composition chart remediation

Manual evidence received 2026-09-27: the user opened the earlier Kabin Buri OOXML output and pressed Ctrl+Alt+F9. The Summary values under `ปริมาณจราจรบนทางแยก (คัน)` and the `สัดส่วนประเภทยานพาหนะบริเวณจุดสำรวจ` chart remained blank. This is a real fidelity failure; recalculation is secondary and is not considered the primary or sole cause.

The authoritative map places the vehicle-class table at `Summary!V26:AJ40` (time rows `W28:W39`, direct class values `X28:AJ39`, totals `W40:AJ40`). The hourly vehicle formulas in W28:W39 link to `AM10:AM21`; totals in row 40 sum the table. The native composition chart is `xl/charts/chart2.xml` (drawing object Chart 18, anchor U41:AG61), with categories `Summary!$AP$39:$BA$39` and values `Summary!$AP$40:$BA$40`. AP38:BA38 links to class totals in Y40:AJ40; AP40:BA40 computes each class share against the sum of AP38:BA38.

Root causes found in the failed package and exporter path:

1. The qualification payload stored `Hourly_Vehicle_Class` in `report_data["sheets"]` but did not pass the optional duplicate top-level `hourly_vehicle_class` frame. The OOXML writer read only that optional field and filled the direct Summary class cells `X28:AJ39` with zeros. Those cells are values, so Ctrl+Alt+F9 cannot populate them from the support sheet.
2. The OOXML column resolver used mojibake instead of the actual `Total (คัน)` label, so the total-vehicles source column was not resolved correctly.
3. Native chart source extraction used a mojibake share-column label rather than the actual `สัดส่วน (%)` column in `Vehicle_Composition_Report`, so chart cache values were zeros.
4. As a secondary calculation issue, cache clearing left stale `t="e"` error types on the AP40:BA40 formulas. Cache clearing now removes the stale formula-cell type while retaining each formula.

COM parity was checked in `exporter.py`: the COM path passes both `sheets` and an explicit hourly vehicle-class frame, writes the Summary table from that frame, writes chart source ranges, then asks Excel to recalculate. The OOXML path now accepts the support-sheet frame as its table input when the optional duplicate is absent, writes numeric table values and matching chart sources/caches, and requests automatic full recalculation. Formula caches are also populated deterministically for visible report cells and validated against their sources. Formula nodes and native chart definitions remain intact.

The remediation changed `src/tmc_processor/ooxml_template_export.py`, `src/tmc_processor/exporter.py`, and `tests/test_ooxml_template_export.py`. Regression coverage checks the support-sheet fallback, all 12 numeric hourly rows, totals, retained formulas and non-error formula types, composition percentages, chart references, chart cache categories/values, and nonzero chart payload. Targeted command: `python -m pytest tests/test_ooxml_template_export.py tests/test_template_ooxml_poc.py -q -p no:tmpdir -p pytest_tmp_path_override -p no:cacheprovider`; result: **3 passed, 1 warning**. The warning is openpyxl's read-only DrawingML limitation; the package validator separately confirms the native charts and all 33 custom shapes remain present. A pytest temporary-directory fixture override was used because the sandbox could not enumerate pytest's default per-user temporary directory.

Only Kabin Buri was regenerated. Its original inspected output was preserved; a read-only diagnostic copy was used. No source workbook was modified, and no Excel process was inspected, launched, terminated, or interacted with. No COM, production-routing change, commit, or PR was used. Corrected output:

`C:\Users\kitti\AppData\Local\Temp\UX7_OOXML_KabinBuri_Fixed_Manual_Retest_v2_20260927\03.TMC1 แยกกบินทร์บุรี_พุธ 17 กค67f_OOXML_FIXED_MANUAL_RETEST.xlsx`

Size: 1,153,524 bytes. SHA-256: `128989c777291e48bd8a00185a7c9da27adf5aae224e3eccaf42dd5b72c39414`. Effective AM Peak 07:00-08:00 and PM Peak 17:00-18:00 retain provenance `auto_suggested`.

Structural and semantic validation passed: no unexpected package changes; internal relationships valid; charts, drawings, media, and all 33 shapes preserved; all support sheets present; template audit clean; effective Peak consistent; formulas retained; Summary vehicle-class values equal the support-sheet rows and are numeric; chart ranges and 12 cached category/share values reconcile to report data. The exact formula/chart ranges above remain the template's native sources. Expected manual result is that the Summary vehicle-count table contains hourly values and totals, and the composition chart displays the 12 vehicle classes. This is ready for manual retest only; it is not a visual pass or release-readiness claim. Standard routing remains unchanged.

## Kabin Buri manual pass and remaining real-sample retests

The user subsequently confirmed that the fixed Kabin Buri workbook visibly populated both the Summary vehicle-count table and vehicle-composition chart in real Excel. Kabin Buri manual fidelity is **PASSED** for those reported surfaces. This does not qualify Office startup or unrelated workbook surfaces.

Fresh Bo Phloi and Nong Prue outputs were generated from the exact saved qualification support payloads, mappings, effective Peaks, and provenance. No Excel COM or Excel process interaction was used; no source workbook was modified, and Standard routing is unchanged. Both outputs passed the exact package mutation allowlist (8 expected existing parts changed, 18 support worksheet parts added, `xl/calcChain.xml` removed), ZIP/XML parsing, internal relationship target checks, native chart checks, 33-shape count, support sheet inventory, `audit_template`, formula preservation/cache checks, table-to-support-data equality, chart category/share source/cache equality, and effective Peak consistency.

| Sample | Retest output | Bytes / SHA-256 | Effective AM / PM | Provenance | Result |
|---|---|---|---|---|---|
| Bo Phloi | `C:\Users\kitti\AppData\Local\Temp\UX7_OOXML_Remaining_Real_Sample_Retest_20260927\Bo_Ploi_OOXML_FIXED_MANUAL_RETEST.xlsx` | 733,910 / `4e59fb8091d80a636978efc4f3793a0e91eddc88ebf0e83d0d7b894d8725a8be` | 08:00-09:00 / 17:00-18:00 | `auto_suggested` | Structural, table, chart, formula/cache, and Peak gates **PASS**; manual retest pending |
| Nong Prue | `C:\Users\kitti\AppData\Local\Temp\UX7_OOXML_Remaining_Real_Sample_Retest_20260927\Nong_Prue_OOXML_FIXED_MANUAL_RETEST.xlsx` | 598,430 / `b5ecb3414dc20dfaf5b757e77ca0726eaa3efde291b8a1c7ee5804258f9354f5` | 11:00-12:00 / 16:00-17:00 | `auto_suggested` | Structural, table, chart, formula/cache, and Peak gates **PASS**; manual retest pending |

The Kabin Buri OOXML regression tests were rerun after these exports: `tests/test_ooxml_template_export.py` and `tests/test_template_ooxml_poc.py`, **3 passed, 1 openpyxl DrawingML warning**. No routing change, commit, PR, merge, or release was performed. The remaining two outputs are **OOXML REMAINING REAL-SAMPLE RETEST READY**.


## Movement Completeness Qualification

This qualification traces semantic movements from accepted `Mapping` and `Normalized_Data` through the hourly export payload, `Hourly_Movement_PCU`, Summary hourly and diagram slots, and `Diagram_Data`. AM/PM values aggregate every normalized 15-minute row in the effective 60-minute window. Identity is `(from_leg, to_leg, turn_type, movement_code)`; worksheet/direction labels are source provenance only.

Kabin Buri's earlier manual pass covered its vehicle-count table and composition chart; it did not qualify intersection movement values. Earlier Bo Phloi/Nong Prue retest files are superseded. The qualification found rounded values in saved hourly support while normalized data retained quarter-hour PCU precision; an exact Peak-start match compared one 15-minute row with the full hourly value. The remediation aggregates the full hour, derives HLOOKUP helper indices from effective Peaks, and fills deterministic caches while retaining formulas and native layout. Placement uses mapped semantic identity, with no sample-specific or raw-sheet-order logic. Source worksheet/direction counts are 24 / 15 / 12; all three samples map four legs and use turn types L/R/T/U.

### Kabin Buri

Effective Peak: AM 07:00-08:00; PM 17:00-18:00 (`auto_suggested`). Source rows: 24; mapped legs: 4; expected/exported/report movements: 16/16/16; missing/duplicate/unexpected: 0/0/0; slot collisions: 0. Slots used: EN, ES, EU, EW, NE, NS, NU, NW, SE, SN, SU, SW, WE, WN, WS, WU. Unused valid slots: none.

All value pairs are AM / PM PCU. Summary shows hourly table references and diagram slot references. Diagram_Data uses D for AM, C for PM.

| Movement identity (from?to / turn / code) | Mapping source(s) | Summary total / PM / AM slots | Expected AM / PM | Payload AM / PM | Support AM / PM | Summary hourly and diagram | Diagram_Data AM / PM | Status |
|---|---|---|---:|---:|---:|---|---|---|
| N?E / L / NE | ทิศ 1/dir 1; ทิศ 3/dir 3 | M14 / M15 / M16 | 667.487 / 884.026 | 667.487 / 884.026 | 667.487 / 884.026 | table W10/W20=667.487 / 884.026; slots M16/M15=667.487 / 884.026 | row 2 D/C=667.487 / 884.026 | MATCH |
| N?S / T / NS | ทิศ 4/dir 4; ทิศ 10/dir 10 | L14 / L15 / L16 | 323.707 / 291.201 | 323.707 / 291.201 | 323.707 / 291.201 | table X10/X20=323.707 / 291.201; slots L16/L15=323.707 / 291.201 | row 3 D/C=323.707 / 291.201 | MATCH |
| N?W / R / NW | ทิศ 5/dir 5; ทิศ 11/dir 11 | K14 / K15 / K16 | 246.889 / 198.342 | 246.889 / 198.342 | 246.889 / 198.342 | table Y10/Y20=246.889 / 198.342; slots K16/K15=246.889 / 198.342 | row 4 D/C=246.889 / 198.342 | MATCH |
| N?N / U / NU | ทิศ 6/dir 6; ทิศ 12/dir 12 | J14 / J15 / J16 | 656.144 / 447.793 | 656.144 / 447.793 | 656.144 / 447.793 | table Z10/Z20=656.144 / 447.793; slots J16/J15=656.144 / 447.793 | row 5 D/C=656.144 / 447.793 | MATCH |
| S?W / L / SW | ทิศ 7/dir 7; ทิศ 12+13/dir 12+13 | J29 / J28 / J27 | 979.546 / 985.800 | 979.546 / 985.800 | 979.546 / 985.800 | table AA10/AA20=979.546 / 985.800; slots J27/J28=979.546 / 985.800 | row 6 D/C=979.546 / 985.800 | MATCH |
| S?N / T / SN | ทิศ 7+8/dir 7+8; ทิศ 13/dir 13 | K29 / K28 / K27 | 1394.619 / 1502.456 | 1394.619 / 1502.456 | 1394.619 / 1502.456 | table AB10/AB20=1394.619 / 1502.456; slots K27/K28=1394.619 / 1502.456 | row 7 D/C=1394.619 / 1502.456 | MATCH |
| S?E / R / SE | ทิศ 8/dir 8; ทิศ 14/dir 14 | L29 / L28 / L27 | 842.063 / 972.744 | 842.063 / 972.744 | 842.063 / 972.744 | table AC10/AC20=842.063 / 972.744; slots L27/L28=842.063 / 972.744 | row 8 D/C=842.063 / 972.744 | MATCH |
| S?S / U / SU | ทิศ 9/dir 9; ทิศ 15/dir 15 | M29 / M28 / M27 | 257.130 / 355.834 | 257.130 / 355.834 | 257.130 / 355.834 | table AD10/AD20=257.130 / 355.834; slots M27/M28=257.130 / 355.834 | row 9 D/C=257.130 / 355.834 | MATCH |
| W?N / L / WN | ทิศ 16/dir 16 | E19 / F19 / G19 | 162.649 / 288.467 | 162.649 / 288.467 | 162.649 / 288.467 | table AE10/AE20=162.649 / 288.467; slots G19/F19=162.649 / 288.467 | row 10 D/C=162.649 / 288.467 | MATCH |
| W?E / T / WE | ทิศ 17/dir 17 | E20 / F20 / G20 | 77.978 / 138.617 | 77.978 / 138.617 | 77.978 / 138.617 | table AF10/AF20=77.978 / 138.617; slots G20/F20=77.978 / 138.617 | row 11 D/C=77.978 / 138.617 | MATCH |
| W?S / R / WS | ทิศ 17+18/dir 17+18 | E21 / F21 / G21 | 498.567 / 735.943 | 498.567 / 735.943 | 498.567 / 735.943 | table AG10/AG20=498.567 / 735.943; slots G21/F21=498.567 / 735.943 | row 12 D/C=498.567 / 735.943 | MATCH |
| W?W / U / WU | ทิศ 18/dir 18 | E22 / F22 / G22 | 420.589 / 597.326 | 420.589 / 597.326 | 420.589 / 597.326 | table AH10/AH20=420.589 / 597.326; slots G22/F22=420.589 / 597.326 | row 13 D/C=420.589 / 597.326 | MATCH |
| E?S / L / ES | ทิศ 19/dir 19 | Q24 / P24 / O24 | 310.017 / 349.171 | 310.017 / 349.171 | 310.017 / 349.171 | table AI10/AI20=310.017 / 349.171; slots O24/P24=310.017 / 349.171 | row 14 D/C=310.017 / 349.171 | MATCH |
| E?W / T / EW | ทิศ 2/dir 2 | Q23 / P23 / O23 | 129.350 / 252.025 | 129.350 / 252.025 | 129.350 / 252.025 | table AJ10/AJ20=129.350 / 252.025; slots O23/P23=129.350 / 252.025 | row 15 D/C=129.350 / 252.025 | MATCH |
| E?N / R / EN | ทิศ 2+3/dir 2+3 | Q22 / P22 / O22 | 639.114 / 873.620 | 639.114 / 873.620 | 639.114 / 873.620 | table AK10/AK20=639.114 / 873.620; slots O22/P22=639.114 / 873.620 | row 16 D/C=639.114 / 873.620 | MATCH |
| E?E / U / EU | ทิศ 20/dir 20 | Q21 / P21 / O21 | 14.933 / 46.997 | 14.933 / 46.997 | 14.933 / 46.997 | table AL10/AL20=14.933 / 46.997; slots O21/P21=14.933 / 46.997 | row 17 D/C=14.933 / 46.997 | MATCH |

Approach, turn-type, and overall reconciliations (expected / actual AM / PM):

| Group | Expected AM / PM | Actual AM / PM | Status |
|---|---:|---:|---|
| OVERALL | 7620.782 / 8920.362 | 7620.782 / 8920.362 | MATCH |
| from_leg=E | 1093.414 / 1521.813 | 1093.414 / 1521.813 | MATCH |
| from_leg=N | 1894.227 / 1821.362 | 1894.227 / 1821.362 | MATCH |
| from_leg=S | 3473.358 / 3816.834 | 3473.358 / 3816.834 | MATCH |
| from_leg=W | 1159.783 / 1760.353 | 1159.783 / 1760.353 | MATCH |
| turn_type=L | 2119.699 / 2507.464 | 2119.699 / 2507.464 | MATCH |
| turn_type=R | 2226.633 / 2780.649 | 2226.633 / 2780.649 | MATCH |
| turn_type=T | 1925.654 / 2184.299 | 1925.654 / 2184.299 | MATCH |
| turn_type=U | 1348.796 / 1447.950 | 1348.796 / 1447.950 | MATCH |

Fresh output: `C:\Users\kitti\AppData\Local\Temp\UX7_OOXML_Movement_Completeness_Retest_20260929\Kabin_Buri_OOXML_MOVEMENT_COMPLETENESS_RETEST.xlsx`; 1160693 bytes; SHA-256 `5ed3883e3420af6079c834cf9700833ec70c4104e5e84b9c19eccee987652b6b`. Package validation: zero unexpected changes, XML and relationships valid, charts and 33 shapes preserved, 18 support sheets present, template audit clean, effective Peak consistent. Manual Excel movement inspection remains pending.

### Bo Phloi

Effective Peak: AM 08:00-09:00; PM 17:00-18:00 (`auto_suggested`). Source rows: 15; mapped legs: 4; expected/exported/report movements: 9/9/9; missing/duplicate/unexpected: 0/0/0; slot collisions: 0. Slots used: EW, NE, NS, NU, NW, SE, SN, SU, SW. Unused valid slots: EN, ES, EU, WE, WN, WS, WU.

All value pairs are AM / PM PCU. Summary shows hourly table references and diagram slot references. Diagram_Data uses D for AM, C for PM.

| Movement identity (from?to / turn / code) | Mapping source(s) | Summary total / PM / AM slots | Expected AM / PM | Payload AM / PM | Support AM / PM | Summary hourly and diagram | Diagram_Data AM / PM | Status |
|---|---|---|---:|---:|---:|---|---|---|
| N?E / L / NE | ทิศ 1/dir 1; ทิศ 3/dir 3 | M14 / M15 / M16 | 199.302 / 199.188 | 199.302 / 199.188 | 199.302 / 199.188 | table W11/W20=199.302 / 199.188; slots M16/M15=199.302 / 199.188 | row 2 D/C=199.302 / 199.188 | MATCH |
| N?S / T / NS | ทิศ 4/dir 4; ทิศ 10/dir 10 | L14 / L15 / L16 | 251.277 / 326.545 | 251.277 / 326.545 | 251.277 / 326.545 | table X11/X20=251.277 / 326.545; slots L16/L15=251.277 / 326.545 | row 3 D/C=251.277 / 326.545 | MATCH |
| N?W / R / NW | ทิศ 5/dir 5; ทิศ 11/dir 11 | K14 / K15 / K16 | 132.190 / 114.607 | 132.190 / 114.607 | 132.190 / 114.607 | table Y11/Y20=132.190 / 114.607; slots K16/K15=132.190 / 114.607 | row 4 D/C=132.190 / 114.607 | MATCH |
| N?N / U / NU | ทิศ 6/dir 6; ทิศ 12/dir 12 | J14 / J15 / J16 | 47.992 / 58.988 | 47.992 / 58.988 | 47.992 / 58.988 | table Z11/Z20=47.992 / 58.988; slots J16/J15=47.992 / 58.988 | row 5 D/C=47.992 / 58.988 | MATCH |
| S?W / L / SW | ทิศ 7/dir 7 | J29 / J28 / J27 | 105.458 / 139.119 | 105.458 / 139.119 | 105.458 / 139.119 | table AA11/AA20=105.458 / 139.119; slots J27/J28=105.458 / 139.119 | row 6 D/C=105.458 / 139.119 | MATCH |
| S?N / T / SN | ทิศ 13/dir 13 | K29 / K28 / K27 | 5.666 / 8.997 | 5.666 / 8.997 | 5.666 / 8.997 | table AB11/AB20=5.666 / 8.997; slots K27/K28=5.666 / 8.997 | row 7 D/C=5.666 / 8.997 | MATCH |
| S?E / R / SE | ทิศ 8/dir 8; ทิศ 14/dir 14 | L29 / L28 / L27 | 33.327 / 19.992 | 33.327 / 19.992 | 33.327 / 19.992 | table AC11/AC20=33.327 / 19.992; slots L27/L28=33.327 / 19.992 | row 8 D/C=33.327 / 19.992 | MATCH |
| S?S / U / SU | ทิศ 9/dir 9; ทิศ 15/dir 15 | M29 / M28 / M27 | 234.010 / 205.290 | 234.010 / 205.290 | 234.010 / 205.290 | table AD11/AD20=234.010 / 205.290; slots M27/M28=234.010 / 205.290 | row 9 D/C=234.010 / 205.290 | MATCH |
| E?W / T / EW | ทิศ 2/dir 2 | Q23 / P23 / O23 | 228.855 / 237.806 | 228.855 / 237.806 | 228.855 / 237.806 | table AJ11/AJ20=228.855 / 237.806; slots O23/P23=228.855 / 237.806 | row 15 D/C=228.855 / 237.806 | MATCH |

Approach, turn-type, and overall reconciliations (expected / actual AM / PM):

| Group | Expected AM / PM | Actual AM / PM | Status |
|---|---:|---:|---|
| OVERALL | 1238.077 / 1310.532 | 1238.077 / 1310.532 | MATCH |
| from_leg=E | 228.855 / 237.806 | 228.855 / 237.806 | MATCH |
| from_leg=N | 630.761 / 699.328 | 630.761 / 699.328 | MATCH |
| from_leg=S | 378.461 / 373.398 | 378.461 / 373.398 | MATCH |
| turn_type=L | 304.760 / 338.307 | 304.760 / 338.307 | MATCH |
| turn_type=R | 165.517 / 134.599 | 165.517 / 134.599 | MATCH |
| turn_type=T | 485.798 / 573.348 | 485.798 / 573.348 | MATCH |
| turn_type=U | 282.002 / 264.278 | 282.002 / 264.278 | MATCH |

Fresh output: `C:\Users\kitti\AppData\Local\Temp\UX7_OOXML_Movement_Completeness_Retest_20260929\Bo_Ploi_OOXML_MOVEMENT_COMPLETENESS_RETEST.xlsx`; 739956 bytes; SHA-256 `45bd75011cfcaa0e2194d9b0e62909972735dbff3bab405dac8d509f93663481`. Package validation: zero unexpected changes, XML and relationships valid, charts and 33 shapes preserved, 18 support sheets present, template audit clean, effective Peak consistent. Manual Excel movement inspection remains pending.

### Nong Prue

Effective Peak: AM 11:00-12:00; PM 16:00-17:00 (`auto_suggested`). Source rows: 12; mapped legs: 4; expected/exported/report movements: 8/8/8; missing/duplicate/unexpected: 0/0/0; slot collisions: 0. Slots used: EW, NE, NS, NU, NW, SE, SU, SW. Unused valid slots: EN, ES, EU, SN, WE, WN, WS, WU.

All value pairs are AM / PM PCU. Summary shows hourly table references and diagram slot references. Diagram_Data uses D for AM, C for PM.

| Movement identity (from?to / turn / code) | Mapping source(s) | Summary total / PM / AM slots | Expected AM / PM | Payload AM / PM | Support AM / PM | Summary hourly and diagram | Diagram_Data AM / PM | Status |
|---|---|---|---:|---:|---:|---|---|---|
| N?E / L / NE | ทิศ 1/dir 1; ทิศ 3/dir 3 | M14 / M15 / M16 | 103.421 / 120.850 | 103.421 / 120.850 | 103.421 / 120.850 | table W14/W19=103.421 / 120.850; slots M16/M15=103.421 / 120.850 | row 2 D/C=103.421 / 120.850 | MATCH |
| N?S / T / NS | ทิศ 4/dir 4; ทิศ 10/dir 10 | L14 / L15 / L16 | 80.994 / 67.585 | 80.994 / 67.585 | 80.994 / 67.585 | table X14/X19=80.994 / 67.585; slots L16/L15=80.994 / 67.585 | row 3 D/C=80.994 / 67.585 | MATCH |
| N?W / R / NW | ทิศ 5/dir 5; ทิศ 11/dir 11 | K14 / K15 / K16 | 99.319 / 95.554 | 99.319 / 95.554 | 99.319 / 95.554 | table Y14/Y19=99.319 / 95.554; slots K16/K15=99.319 / 95.554 | row 4 D/C=99.319 / 95.554 | MATCH |
| N?N / U / NU | ทิศ 6/dir 6; ทิศ 12/dir 12 | J14 / J15 / J16 | 77.161 / 94.927 | 77.161 / 94.927 | 77.161 / 94.927 | table Z14/Z19=77.161 / 94.927; slots J16/J15=77.161 / 94.927 | row 5 D/C=77.161 / 94.927 | MATCH |
| S?W / L / SW | ทิศ 7/dir 7 | J29 / J28 / J27 | 35.830 / 34.159 | 35.830 / 34.159 | 35.830 / 34.159 | table AA14/AA19=35.830 / 34.159; slots J27/J28=35.830 / 34.159 | row 6 D/C=35.830 / 34.159 | MATCH |
| S?E / R / SE | ทิศ 8/dir 8 | L29 / L28 / L27 | 71.497 / 78.496 | 71.497 / 78.496 | 71.497 / 78.496 | table AC14/AC19=71.497 / 78.496; slots L27/L28=71.497 / 78.496 | row 8 D/C=71.497 / 78.496 | MATCH |
| S?S / U / SU | ทิศ 9/dir 9 | M29 / M28 / M27 | 36.331 / 34.162 | 36.331 / 34.162 | 36.331 / 34.162 | table AD14/AD19=36.331 / 34.162; slots M27/M28=36.331 / 34.162 | row 9 D/C=36.331 / 34.162 | MATCH |
| E?W / T / EW | ทิศ 2/dir 2 | Q23 / P23 / O23 | 74.329 / 72.830 | 74.329 / 72.830 | 74.329 / 72.830 | table AJ14/AJ19=74.329 / 72.830; slots O23/P23=74.329 / 72.830 | row 15 D/C=74.329 / 72.830 | MATCH |

Approach, turn-type, and overall reconciliations (expected / actual AM / PM):

| Group | Expected AM / PM | Actual AM / PM | Status |
|---|---:|---:|---|
| OVERALL | 578.882 / 598.563 | 578.882 / 598.563 | MATCH |
| from_leg=E | 74.329 / 72.830 | 74.329 / 72.830 | MATCH |
| from_leg=N | 360.895 / 378.916 | 360.895 / 378.916 | MATCH |
| from_leg=S | 143.658 / 146.817 | 143.658 / 146.817 | MATCH |
| turn_type=L | 139.251 / 155.009 | 139.251 / 155.009 | MATCH |
| turn_type=R | 170.816 / 174.050 | 170.816 / 174.050 | MATCH |
| turn_type=T | 155.323 / 140.415 | 155.323 / 140.415 | MATCH |
| turn_type=U | 113.492 / 129.089 | 113.492 / 129.089 | MATCH |

Fresh output: `C:\Users\kitti\AppData\Local\Temp\UX7_OOXML_Movement_Completeness_Retest_20260929\Nong_Prue_OOXML_MOVEMENT_COMPLETENESS_RETEST.xlsx`; 603411 bytes; SHA-256 `d6da0bc0c5f54cb16d854ea40a5e6a0f4cf6d042889a78c46fc4519880a1217f`. Package validation: zero unexpected changes, XML and relationships valid, charts and 33 shapes preserved, 18 support sheets present, template audit clean, effective Peak consistent. Manual Excel movement inspection remains pending.

### Generic root cause and regression gates

The generic data defect was an hourly support value that did not retain normalized quarter-hour PCU precision, combined with a validator that selected only the interval beginning exactly at the Peak start. The corrected path sums the effective 60-minute window and validates semantic identities through Mapping, Normalized_Data, Movement_Summary, Hourly_Movement_PCU, Summary, and Diagram_Data. The exporter derives helper indexes from the effective Peaks and fills formula caches without replacing native formulas. Kabin Buri uses all sixteen template movement slots; Bo Phloi and Nong Prue use nine and eight, leaving legitimate unused slots at zero.

Kabin Buri's prior manual result applies to the vehicle table and composition chart. The new three-sample regression randomizes Mapping row order and exercises 15-minute inputs, checks full identity sets and destinations, AM/PM values, from-leg/turn-type/overall totals, no missing/duplicate/extra movements, and the vehicle table/chart control. No COM, Excel interaction, source edits, Windows/Office changes, routing changes, commit, or PR were used.

## Backend write-contract parity (2026-09-29; corrected scope)

This checkpoint begins with the fixed post-Review export payload. It does not infer physical mappings from raw workbooks or change Data, Mapping, Analyze, or Review. The earlier movement-completeness projection above is historical evidence, not an independently trusted real-sample fixture for this backend checkpoint.

`template_write_plan.resolve_template_write_plan` now resolves the legacy COM destination sequence once. Both `export_with_excel_com` and `export_template_ooxml` consume that plan. The production Standard route remains unchanged. The plan has ordered, typed Summary cell intents; support-sheet frames; Diagram_Data formulas and deterministic caches; the two native chart cache series; preserved native formulas; and attempted writes skipped because their destinations are protected formulas. A table column intent writes its header, clears each mapped destination, then writes each row from the exact post-Review frame and the original template time labels. This preserves the legacy COM write order and the template's time layout.

| Legacy write surface | Destination and post-Review source | Formula / cache rule |
|---|---|---|
| Summary metadata | Mapped `metadata_cells` (B2, E4, E5, K5, Q5, E6, K6) and mapped diagram labels/caption from `metadata`; effective Peak/provenance label A3 | B2 is the map-authorized formula overwrite. Formula-backed diagram title/date and any other protected formula targets are skipped. |
| Hourly movement | `hourly_movement_table`: V9:AM9 headers, V10:AM21 hourly rows, V22:AM22 total row; values from `hourly_movement_pcu` keyed by the template's time labels and column keys | Native Summary movement HLOOKUP formulas remain unchanged. The three effective-Peak row helpers at U11/U20/U22 are direct writes. |
| Vehicle class | `hourly_vehicle_class_table`: V27:AJ27 headers, V28:AJ39 hourly rows, V40:AJ40 total row; values from top-level `hourly_vehicle_class` when present, otherwise the same `sheets['Hourly_Vehicle_Class']` frame | Native total/PCU and vehicle-composition formulas remain unchanged; OOXML seeds their formula caches from populated sources. |
| Native charts | Mapped hourly categories/values V10:V21 / AM10:AM21 and composition categories/values AP39:BA39 / AP40:BA40, from `chart_source_data` | AP40:BA40 are protected formulas, so their attempted direct writes are skipped. COM asks Excel to recalculate chart caches; OOXML updates only the native cache points in chart1/chart2 and seeds the Summary formula caches. Chart definitions and drawings remain unchanged. |
| Support sheets | Every `(name, DataFrame)` in `report_data['sheets']`, including Export_Metadata, Peak_PHF, Hourly_Movement_PCU, Hourly_Vehicle_Class, and all other Review payload sheets | Headers and values are copied, preserving source order and cell types. Provenance and PHF remain payload values. |
| Diagram_Data | A1:D17, 16 `diagram_movement_codes`, formulas against Hourly_Movement_PCU and effective Peak cells in Peak_PHF | COM writes formulas and lets Excel calculate; OOXML retains the same formulas and seeds their total/AM/PM caches. Native Summary movement formulas are unchanged. |

The previous blank vehicle table and composition chart were traced to an OOXML support/table-input mismatch: a qualification payload could provide `Hourly_Vehicle_Class` in `sheets` without the optional duplicate top-level field, and the earlier OOXML table resolver also used a malformed `Total (คัน)` column label. The accepted fix populated numeric Summary vehicle sources and chart caches; the user visually confirmed Kabin Buri's two affected surfaces. The shared plan now gives both transports the same vehicle frame and resolved column. Recalculation remains a secondary mechanism, not the root-cause explanation.

Deterministic parity fixture: all 16 movement slots, 12 hourly rows, 12 vehicle classes, AM/PM, effective Peak helpers, Export_Metadata, Peak_PHF, Diagram_Data, native chart sources/caches, and provenance. It produces 3,196 typed, globally ordered cell intents, including 861 Summary writes across 451 final Summary destinations; 100 attempted formula writes are skipped according to the map. The COM transport's Summary dispatch is captured without launching COM and matches every resolved intent. The OOXML verifier checks each final Summary destination, every support-sheet cell, all 16 Diagram_Data rows/formula caches and visible movement caches, hourly PCU/vehicle totals, chart source caches, native chart XML against cache-only expected edits, all protected formulas, every drawing/media part, XML, internal relationships, sheet order, and the explicit package change allowlist. Negative tests deliberately omit one planned write, corrupt a destination, swap movement source columns, omit a support sheet, stale chart/movement/vehicle caches, change a protected formula, change a chart definition, and change a drawing; each is rejected.

Focused run: `python -m pytest tests/test_ooxml_template_export.py tests/test_template_ooxml_poc.py tests/test_excel_com_export.py tests/test_effective_peak_export.py tests/test_template_integrity.py -q -p no:cacheprovider` passed **39 tests, 6 openpyxl DrawingML read warnings**. The warning concerns openpyxl's in-memory object model; the OOXML package comparison proves that the 33 original shape definitions are unchanged. No Excel process or COM runtime was touched.

Manual inspection artifact (synthetic deterministic payload, not a real sample): `C:\Users\kitti\AppData\Local\Temp\UX7_OOXML_Backend_Parity_20260929\Deterministic_All_16_Slots_OOXML.xlsx`; **56,683 bytes**, SHA-256 `d6da900255996a89b712a037be28922c172357663adc731d958a4200c3a574ff`. Authoritative comparison template: `C:\MyRD\tmc-processor-public\templates\four_leg_tmc_report_template.xlsx`. The saved real-sample movement projection is derived from earlier output and is not accepted as a fixed upstream post-Review payload. A trusted serialized payload/state for Kabin Buri, Bo Phloi, or Nong Prue is still needed before this corrected-scope real-workbook qualification can be claimed.

## Deterministic formula-cache parity (2026-09-30)

The exporter cleared the authoritative template's sample caches when copying the Summary part. The earlier deterministic artifact reseeded the 48 visible movement HLOOKUP results, hourly vehicle PCU/total and composition sources, Diagram_Data results, and both native chart caches, but missed 16 movement total source formulas (`W22:AL22`), the grand hourly total (`AM22`), three displayed movement rollups (`F30:F32`), and the diagram title (`C8`). The template itself has sample data caches in these cells; those values cannot be reused for the new payload. Its shared formula stubs in part of row 22 are native and remain byte-identical at the formula node level.

The following audit compares the **earlier deterministic artifact** with the new artifact. The template's formulas are preserved. `YES` means the cache is required and deterministically resolved from the fixed export payload. `NO` records the explicit text-date exception described below.

| Cell/range | Native formula | Upstream inputs | Expected cache for deterministic payload | Earlier cache | New cache | Downstream / display | Required? |
|---|---|---|---|---|---|---|---|
| Summary!W22:AL22 | Per-column `SUM(column10:column21)`; shared formula nodes retained | 12 direct hourly cells per movement, reconciled against Diagram_Data and Hourly_Movement_PCU | W:AL = `78, 90, 102, 114, 126, 138, 150, 162, 174, 186, 198, 210, 222, 234, 246, 258` | 16 blank | All 16 match | `HLOOKUP(...,$U$22,...)` movement totals, visible rollup | YES |
| Summary!AM22 | `SUM(AM10:AM21)` | 12 direct hourly totals; sum of 16 movement totals | `2688` | Blank | `2688` | Visible hourly total | YES |
| Summary!F30 | `SUM(G19:G22,J16:M16,O21:O24,J27:M27)` | Cached native AM movement results | `136` | Blank | `136` | Visible AM report total | YES |
| Summary!F31 | `SUM(J15:M15,F19:F22,J28:M28,P21:P24)` | Cached native PM movement results | `296` | Blank | `296` | Visible PM report total | YES |
| Summary!F32 | `SUM(J14:M14,E19:E22,J29:M29,Q21:Q24)` | Cached native all-hour movement results | `2688` | Blank | `2688` | Visible 12-hour report total | YES |
| Summary!C8 | `E5` | Written survey point `TMC-01` | `TMC-01` (string cache) | Blank | `TMC-01` | Diagram title display | YES |
| Summary!C9 | `+K5` | K5 contains text `2026-09-27` | Not asserted: Excel's coercion of a text date is not specified by this payload | Blank | Blank | Diagram date display; no formula/chart dependent | NO for deterministic cache contract |
| Summary native movement cells (48) | `HLOOKUP` against `$W$9:$AL$22` | Hourly rows, fixed U11/U20/U22 relative indexes | 16 total, 16 AM, 16 PM values from the fixed plan | Populated | Populated and verified | Diagram movement labels and rollups | YES |
| Summary!W28:W39, W40:AJ40, AP38:BA38, AP40:BA40 | Native linked PCU, vehicle total, numerator, and share formulas | Written movement totals and vehicle-class cells | Values reconciled with Summary and native chart sources | Populated | Populated and verified | Vehicle table and composition chart | YES |
| Diagram_Data!B2:D17 | Native `IFERROR/INDEX` formulas | Hourly_Movement_PCU and effective Peak_PHF AM/PM starts | 16 total, 16 PM, 16 AM caches from the fixed plan | Populated | Populated and verified | Diagram support data | YES |
| Peak_PHF / chart cache points | Peak_PHF has literal payload values; charts retain native series references | Fixed payload and Summary chart source ranges | Effective Peak cells and both 12-point series | Populated | Populated and verified | PHF and native charts | YES |

No general Excel formula evaluator or date parser was added. With text K5, `C9` remains a native formula with an empty cache; the verifier rejects a fabricated/stale nonempty cache. The diagram date's display still requires manual Excel inspection. If K5 is numeric or blank, the bounded cache rule seeds and verifies C9. No other formula cell in the new Summary part has a blank cache.

Fresh deterministic artifact: `C:\Users\kitti\AppData\Local\Temp\UX7_OOXML_Cache_Parity_20260930\Deterministic_All_16_Slots_CACHE_PARITY.xlsx`; **56,741 bytes**; SHA-256 `26598fc7409abf4af2cc9a578855a9c62d32ff47c6ab9032d0f25784d0a2e390`. The 16 row-22 movement formulas remain native; all 16 caches match the hourly input sums exactly. Required missing caches: **0**; required stale/wrong caches: **0**; protected formula changes: **0**. The package verifier and template audit pass: two chart definitions and all drawing/media parts are preserved; XML and relationships are valid; only allowlisted parts change. The new negative cases blank W22, set W22 to 999, blank AM22/F30/C8, insert a fabricated C9 cache, or alter the W22 formula; each is rejected. This artifact is synthetic and is for manual retest, not real-workbook or release acceptance.

Focused run: `python -m pytest tests/test_ooxml_template_export.py tests/test_template_ooxml_poc.py tests/test_excel_com_export.py tests/test_effective_peak_export.py tests/test_template_integrity.py -q --tb=short -p no:cacheprovider` passed **46 tests** with 6 openpyxl DrawingML warnings. `python -m compileall -q app.py src` and `git diff --check` passed. Excel and COM were not used; Standard routing was not changed.

## Native effective-Peak formula binding (2026-09-30)

The existing upstream workflow already resolves the effective/confirmed AM and PM intervals. Earlier tests checked the resulting A3 label, Export_Metadata, Peak_PHF, and Diagram_Data, but did not directly assert that the 32 visible native AM/PM HLOOKUP formulas selected the corresponding hourly rows. The UX-7 interim implementation wrote new row indices into U11/U20, which changed the Excel-authored helper column. The authoritative template has U9:U22 = 1..14, where each value is a row index relative to W9:AL22; U11 is always 3 and U20 is always 12.

The corrected backend-neutral plan reads the actual template V10:V21 time labels, matches the already-resolved AM/PM start-and-end intervals exactly, and records `(effective interval, worksheet row, U helper cell, relative index)`. A missing or ambiguous match fails before transport. It derives all 16 movement cell groups from `movement_diagram_cells.approach_tables`, validates the complete Excel-authored HLOOKUP structure (header reference, W9:AL22 lookup range, U helper reference, `FALSE`), and records 32 narrowly authorized old-formula-to-new-formula bindings. All total formulas keep U22. U9:U22 is removed from the writable map and is neither a payload field nor a Peak destination.

The COM transport and OOXML transport consume the same binding plan. The COM path checks the live cell formula before setting the planned formula. The OOXML path changes only the helper-row number in each explicit formula; Excel-authored shared-formula stubs stay in place and follow their master. The verifier checks U9:U22, every planned formula, all 32 selected-row caches, all 16 total formulas/caches, support sheets, chart caches, and byte-identical drawings. This is a template formula-binding correction only; Mapping, Analyze, Review, Peak calculation, report data, and Standard routing remain unchanged.

The focused OOXML file passed **31 tests**; the affected effective-Peak, explicit review, COM contract, template, metadata, package, and batch suite passed **76 tests**. The matrix covers template-default AM 08:00, non-default AM 07:00, AM 17:00 using a normally PM-time row, PM 16:00/17:00, all 16 total formulas, confirmed AM change, draft-only confirmed-state behavior, and unmatched intervals. Negative tests change one AM helper reference, duplicate U11's value, or change a native HLOOKUP token; qualification rejects each. No Excel or COM runtime was used.

Two fresh deterministic artifacts use 16 movements with hour-distinct values. Both retain U9:U22 = 1..14, two native charts, and all 33 drawing shapes. The package verifier passes its write-plan, support, Diagram_Data, chart cache, protected-formula, drawing/media, XML, and relationship gates; `audit_template` is clean. For every movement, AM and PM formula caches equal the movement's value in the selected Summary hourly row, and the total formula/cache remains bound to U22.

| Artifact | Effective AM / helper | Effective PM / helper | 16 AM / 16 PM / 16 total formula references | File | Size / SHA-256 |
|---|---|---|---|---|---|
| A | 08:00-09:00 / row 11 / U11=3 | 17:00-18:00 / row 20 / U20=12 | `$U$11` / `$U$20` / `$U$22` | `C:\Users\kitti\AppData\Local\Temp\UX7_Native_Peak_Binding_20260930\Deterministic_All_16_Slots_AM08_PM17_NATIVE_BINDING.xlsx` | 56,749 / `a6108d0da7b34e905f939d0a9b3028a5189457951c596ac76cadb30c43b50d39` |
| B | 07:00-08:00 / row 10 / U10=2 | 16:00-17:00 / row 19 / U19=11 | `$U$10` / `$U$19` / `$U$22` | `C:\Users\kitti\AppData\Local\Temp\UX7_Native_Peak_Binding_20260930\Deterministic_All_16_Slots_AM07_PM16_NATIVE_BINDING.xlsx` | 56,732 / `76eaa0ae506b8154cc41e9ae8c5f6c1025d638006fb80298886ac428fe7c64de` |

These are synthetic manual-inspection workbooks. Manual Excel fidelity and real-workbook release acceptance remain separate gates. The earlier cached date-display limitation at Summary!C9 remains documented above.
