# UX-7 release qualification evidence

Finalization date: 2026-10-01 (Asia/Bangkok). Acceptance decision: **ACCEPT WITH DOCUMENTED LIMITATIONS**. The acceptance authority explicitly waived repeating the three protected real-workbook end-to-end runs on this exact UX-7 branch; Issue #10 records the rationale and limitations. The earlier blocked finding below is historical qualification context, not a current gate. Standard routing remains unchanged. No push, PR, merge, tag, or release is authorized by this evidence.

## Repository and baseline

- Authoritative repository: `C:\MyRD\tmc-processor-public`
- Branch: `codex/ux-7-release-qualification`
- Accepted baseline HEAD: `f17d416d46d6b04e1d84add022e93905b933141a` (`main` at the start of UX-7 finalization).
- Historical stable baseline supplied by the user: 257 tests. No corresponding run at this HEAD is asserted.
- Accepted UX-6 baseline: 369 passed, 11 warnings, recorded in `docs/development/EVIDENCE_PACKAGE_UX6.md`.
- Earlier pre-P0 full suite: **401 passed, 13 warnings** in 220.78 seconds; 144 more passing tests than the user-supplied historical count and 32 more than UX-6. The final full-suite result is recorded in the finalization section below. The warnings are openpyxl DrawingML in-memory support warnings, not package failures; OOXML package preservation has separate byte/structure assertions.

Pre-existing UX-7 dirty files were preserved. At evidence creation, the complete dirty inventory is:

| State | Path |
|---|---|
| Modified | `pyproject.toml` |
| Modified | `src/tmc_processor/excel_com_export.py` |
| Modified | `src/tmc_processor/exporter.py` |
| Modified | `src/tmc_processor/report_template.py` |
| Modified | `src/tmc_processor/ui/app_shell.py` |
| Modified | `src/tmc_processor/ui/components/export.py` |
| Modified | `src/tmc_processor/ui/components/peak.py` |
| Modified | `src/tmc_processor/ui/workflows/batch.py` |
| Modified | `src/tmc_processor/ui/workflows/single.py` |
| Modified | `templates/four_leg_tmc_report_template_map.json` |
| Modified | `tests/test_setup_persistence.py` |
| Untracked | `docs/development/EVIDENCE_OOXML_UX7.md` |
| Untracked | `docs/development/EVIDENCE_PACKAGE_UX7.md` (this file) |
| Untracked | `scripts/inspect_template_package.py` |
| Untracked | `src/tmc_processor/ooxml_template_export.py` |
| Untracked | `src/tmc_processor/template_write_plan.py` |
| Untracked | `src/tmc_processor/template_write_verify.py` |
| Untracked | `tests/test_ooxml_template_export.py` |
| Untracked | `tests/test_template_ooxml_poc.py` |

The obsolete synthetic raw-mapping projection fixture and its sole test consumer were removed during finalization. This inventory records the earlier dirty tree, excluding that removed untracked file.

## Automated gates

| Gate | Command / evidence | Result |
|---|---|---|
| Full regression | `.venv\Scripts\python.exe -m pytest -q --tb=short -p no:cacheprovider --basetemp C:\MyRD\tmc-processor-public\pytest-ux7-full` | PASS, 401 passed, 13 DrawingML warnings, 220.78 s. Test temp directory removed after run. |
| Python syntax | `.venv\Scripts\python.exe -m compileall -q app.py src` | PASS |
| Whitespace / patch integrity | `git diff --check` | PASS |
| OOXML deterministic native-template qualification | `tests/test_ooxml_template_export.py` and `tests/test_template_ooxml_poc.py`, included in full suite; detail in `EVIDENCE_OOXML_UX7.md` | PASS for fixed post-Review payload, including backend write-plan parity, zero unexpected package mutations, XML/relationships, support sheets, two native charts, 33 shapes/drawings, protected formulas, required caches, and U9:U22 = 1..14. Negative mutations are rejected. |

The user accepted deterministic OOXML parity, native HLOOKUP effective-Peak binding, the U9:U22 invariant, and manual Excel inspection at this checkpoint. The manual acceptance is recorded as user-provided evidence; it is not treated as a fresh visual inspection of the three protected real samples. The authoritative template is `C:\MyRD\tmc-processor-public\templates\four_leg_tmc_report_template.xlsx`.

The latest deterministic artifacts, each with all 16 movement slots and native charts/shapes, are:

| Effective Peaks | Artifact | Size / SHA-256 |
|---|---|---|
| AM 08:00-09:00; PM 17:00-18:00 | `C:\Users\kitti\AppData\Local\Temp\UX7_Native_Peak_Binding_20260930\Deterministic_All_16_Slots_AM08_PM17_NATIVE_BINDING.xlsx` | 56,749 bytes / `a6108d0da7b34e905f939d0a9b3028a5189457951c596ac76cadb30c43b50d39` |
| AM 07:00-08:00; PM 16:00-17:00 | `C:\Users\kitti\AppData\Local\Temp\UX7_Native_Peak_Binding_20260930\Deterministic_All_16_Slots_AM07_PM16_NATIVE_BINDING.xlsx` | 56,732 bytes / `76eaa0ae506b8154cc41e9ae8c5f6c1025d638006fb80298886ac428fe7c64de` |

### Effective Peak and native formula chain

The full suite retains checks for confirmed/effective AM/PM Peaks flowing through `Export_Metadata`, `Peak_PHF`, `Diagram_Data`, `Summary!A3`, 32 visible native AM/PM HLOOKUP bindings, and formula caches for all 16 movement slots. `tests/test_effective_peak_export.py`, `tests/test_explicit_peak_review.py`, and `tests/test_ooxml_template_export.py` cover the cross-surface values, a changed confirmation, a draft-only edit, non-default Peak hours, and unmatched interval rejection. `Summary!U9:U22` stays `1..14`; the formula binding changes its U reference to the effective hour. Total formulas retain U22. Draft-only changes do not alter confirmed exports; changed confirmation stales Export while retaining Analysis.

### Workflow/state invalidation matrix

The following tests passed as part of the 401-test run. “Stale” means an existing downstream result must be regenerated, without implying that source data or analysis was destroyed.

| Change / gate | Expected and observed contract | Test evidence |
|---|---|---|
| Mapping semantic change | Analysis, Review/confirmation, and Export invalidated; nonsemantic editor/view changes do not churn state | `test_workflow_contract.py`, `test_workflow_contract_adapter.py`, `test_physical_mapping.py`, `test_explicit_peak_review.py` |
| PCE factor change | Analysis, Review, and Export invalidated; Batch Analysis and Export stale | `test_workflow_contract.py`, `test_workflow_contract_adapter.py`, `test_workflow_state.py` |
| Confirmed Peak change | Analysis retained; Review updated; Export stale; native binding and caches change to the confirmed hour | `test_workflow_contract.py`, `test_explicit_peak_review.py`, `test_workflow_state.py`, `test_ooxml_template_export.py` |
| Draft-only Peak edit | Confirmation and effective exported Peaks unchanged; no implicit confirmation | `test_explicit_peak_review.py` |
| Report metadata change | Analysis retained; Export stale; changed metadata reaches generated output | `test_workflow_contract.py`, `test_workflow_contract_adapter.py`, `test_metadata_export.py`, `test_batch.py` |
| Export backend/mode change | Analysis and confirmed Peaks retained; Export invalidated/stale, with requested/used/fallback state tested | `test_workflow_state.py`, `test_ux4_ux5.py` |
| Project Session reload | Confirmed Peak, setup, mapping and movement-code scheme round-trip; unchanged semantic inputs do not clear valid artifacts | `test_explicit_peak_review.py`, `test_setup_persistence.py`, `test_mapping_preset.py`, `test_workflow_contract_adapter.py` |
| Mapping Preset compatibility | Legacy v1 defaults to `from_to`; v2 detection/validation and mapping round-trips pass; Batch preset change stales Analysis | `test_mapping_preset.py`, `test_phase_l0_dual_scheme.py`, `test_workflow_state.py` |

## Protected real-workbook historical qualification and waived repeat

The three named source files exist and passed read-only Data-stage detection and parse. This was not an export run and did not infer directions:

| Sample | Source SHA-256 | Detected/parsed direction sheets | Data stage |
|---|---|---:|---|
| Kabin Buri | `f68f084c6a1a822714ac2c80051575dbab67531c1107b24e27688ad2d367c94b` | 24 | PASS |
| Bo Phloi | `5ca76e635b20ac176c9b1a139d23f0ae9fe6237c8b9dd0c41bb2c3ef` | 15 | PASS |
| Nong Prue | `c6890839f1d4e8c0424fb9bb5740e60c7379e0cc783e69cbec75cbed7140d695` | 12 | PASS |

Their exact source paths are `samples\raw\03.TMC1 แยกกบินทร์บุรี_พุธ 17 กค67f.xlsx`, `samples\raw\03.TMC1 แยกบ่อพลอย_พุธ 28 พค68.xlsx`, and `samples\raw\03.TMC4 แยกหนองปรือ_พุธ 28 พค68.xlsx` beneath the repository.

**Historical gap, now waived for this branch:** Mapping → Analyze → explicit Peak Review → OOXML native-template Export and Safe PNG Export were not repeated for these three protected samples on this exact UX-7 branch. No newly qualified real-sample output is claimed. Earlier UX-1/UX-2 and UX-7 qualification mappings were explicitly synthetic and did not constitute approved physical directions. The available `samples/demo` preset/session apply to demo files, not these sources. Historical real-sample outputs in temporary folders are not a trusted post-Review state. The acceptance authority accepted historical validation and operational use for unchanged Mapping/Analyze behavior and waived the repeat; no raw directions were reinterpreted to fill the fixture gap.

**Documented limitation:** reusable accepted Mapping Presets or saved Project Sessions for these exact three workbook hashes, with their reviewed AM/PM Peaks, were not preserved in the workspace. This is a fixture/provenance limitation for future repeatable qualification. The acceptance decision does not require reconstructing those inputs or reopening the waived gate for this commit.

## Batch workflow qualification using accepted demo fixture

An independent product-path Batch run used the repository's demo Mapping Preset and two demo workbooks plus one intentionally invalid input. Analyze returned `success/success/failed`, with zero QC errors or warnings for successful files, and no implicit confirmed Peaks. Explicit bulk acceptance confirmed the two clean files; the second was then explicitly excluded. Reviewed Safe PNG generation produced `success/excluded/failed`; only the first report was packaged. The success row used AM `08:15-09:15`, PM `17:30-18:30`; `Export_Metadata` and the saved Project Session both record `user_confirmed_batch`. The ZIP has 8 entries, including 3 PNGs. The failed input remained explicit and did not block the reviewed success; the excluded file produced no report. The full suite also covers UX-5 clean eligibility (QC info allowed; warning/error excluded), exclusion/restore, per-file confirmation, output metadata, and backend fallback semantics.

Demo Batch artifact: `C:\Users\kitti\AppData\Local\Temp\UX7_Release_Qualification_20260930\demo_reviewed_batch_safe_png.zip`, 1,150,295 bytes, SHA-256 `e735ab4adfe1705e2e8b52d9ed9971ad9ec48e6dcb67c3b6d673837925c3e3d2`. This qualifies Batch mechanics only. The three protected real-workbook exact-branch repeat lacks preserved accepted Mapping/Review state and was explicitly waived for this commit.

## Limits and next review decision

`Summary!C9` retains a native `+K5` formula. When K5 is a text date, its deterministic cache is intentionally blank because the bounded OOXML exporter does not assert Excel's date coercion; manual Excel rendering is the visual gate. Other required deterministic formula caches are verified. The user accepted the current deterministic manual inspection. No Excel, Excel COM, or Office/Windows configuration was touched in this qualification run. Source workbooks were read only.

Standard routing is unchanged in this commit; any future switch is a separate reviewed decision. Compatible v1 Standard Recommended → OOXML native template, Safe PNG alternative/fallback, and optional COM legacy/Advanced are candidates for that future decision. Preserve requested/used/fallback metadata and export invalidation behavior. The `pyproject.toml` addition defines `pywin32` only in the Windows `excel-com` optional extra, with no change to normal installation. COM remains available as an optional legacy/Advanced capability, so the isolated extra is retained; it can be removed if COM is retired in a separately reviewed decision.

Historical working-tree status at this qualification checkpoint: UX-7 edits and this evidence file were uncommitted. No routing change was made. Finalization gates and commit status are recorded below.

## UX-7 P0 operator UI cleanup (2026-09-30)

Scope is presentation only; the broader operator UI polish remains deferred to UX-8 Issue #22. The UX-7 export changes above, Standard decision/routing, COM and OOXML engines, Safe PNG path, Mapping/Analyze/Peak/QC logic, template, Project Session, and invalidation rules were not changed by this P0 task.

- Normal Single and Batch Export now name the report outcome: **รายงานมาตรฐาน — แนะนำ** for the template report and **รายงานสำรอง** for Safe PNG. The card descriptions describe the Excel report with charts/diagram or the compatible image-based alternative. The underlying radio values and backend selection remain unchanged.
- The normal Export card, Single readiness checklist, Batch readiness checklist, sidebar, and top status presentation no longer show Excel COM availability, Excel version, raw COM reason/detail, a COM test prompt, the COM readiness item, or template/Native Chart transport copy. Internal backend eligibility checks are retained. Normal fallback notice is the concise **ระบบใช้รูปแบบรายงานสำรองสำหรับการส่งออกครั้งนี้**; the original technical notice remains in export provenance.
- The existing sidebar COM probe and status are retained inside a collapsed **Advanced / Diagnostics** expander. Single and Batch Export have collapsed diagnostic panels for requested/used/fallback details and COM status. A runtime COM fallback displays the technical warning only inside an additional collapsed diagnostic expander.
- Review wording now marks suggested Peaks as draft, changes the confirmation button to **ยืนยันช่วง Peak**, marks confirmed selections **ยืนยันแล้ว**, and says **มีการเปลี่ยนช่วง Peak แต่ยังไม่ได้ยืนยัน** for draft edits after confirmation. Confirmation keys and state transitions are unchanged.
- P0 task files: `src/tmc_processor/ui/app_shell.py`, `src/tmc_processor/ui/components/export.py`, `src/tmc_processor/ui/components/peak.py`, `src/tmc_processor/ui/workflows/single.py`, `src/tmc_processor/ui/workflows/batch.py`, `tests/test_setup_persistence.py`, `tests/test_ux7_p0_ui_cleanup.py`, and this evidence section.

Validation sequence: focused headless UI/Review integration tests **5 passed**; related export UI, Standard decision, Safe PNG, explicit Review, workflow state/invalidation, architecture, setup persistence, effective Peak, and OOXML groups **159 passed, 4 existing DrawingML warnings**; full pytest **405 passed, 13 existing DrawingML warnings** in 239.54 seconds. The full count is four above the preceding 401-test UX-7 checkpoint because the new focused file contributes four parametrized cases. The tests assert that ordinary Single and Batch Export text omits COM/version/raw reason and the COM readiness item, diagnostics retain those fields, fallback provenance remains technical internally, Standard routing and Safe PNG selection are unchanged, and the Review button plus explicit confirmation path still work. Existing tests continue to cover draft-only edits, changed confirmation invalidation, and effective Peak native formula binding. No Excel or COM runtime was launched by these tests.

Headless Batch Export presentation testing needed one-render isolation: the pre-existing Standard Batch adapter stores a base export mode while the shell's option list contains a suffixed display label, causing repeated `st.rerun()` in AppTest when entering Export. The P0 test suppresses only that rerun hook to inspect the rendered copy; the production adapter was not altered. This was verified against baseline `HEAD` source and is a separate Batch UI/manual qualification limitation, not a P0 presentation change. Reusable real-workbook Mapping/Review fixtures remain unavailable as described above; the exact-branch repeat is waived.

P0 checkpoint validation: `.venv\Scripts\python.exe -m compileall -q app.py src` **PASS**; `git diff --check` **PASS**. Work remained uncommitted at that checkpoint; Standard routing was unchanged.

## Finalization scope audit (2026-10-01)

Compared with accepted `main` at `f17d416d46d6b04e1d84add022e93905b933141a`, the intended commit contains only these UX-7 files:

| File | Scope and reason |
|---|---|
| `pyproject.toml` | Dependency plumbing: optional Windows `excel-com` extra for retained legacy COM transport; normal dependencies unchanged. Retain. |
| `src/tmc_processor/excel_com_export.py` | Legacy COM transport consumes the shared post-Review write plan and effective-Peak formula bindings. |
| `src/tmc_processor/exporter.py` | Report chart-source extraction uses the payload's vehicle-composition share column. |
| `src/tmc_processor/report_template.py` | Native movement values from the fixed post-Review hourly payload for deterministic formula caches. |
| `src/tmc_processor/ooxml_template_export.py` | Bounded native-template OOXML transport, support sheets, chart caches, and package preservation. |
| `src/tmc_processor/template_write_plan.py` | Backend-neutral destination/value/formula write contract. |
| `src/tmc_processor/template_write_verify.py` | Output verification against that plan, formulas, charts, drawings, relationships, and package scope. |
| `templates/four_leg_tmc_report_template_map.json` | Explicit native Peak helper-column contract (`U`). |
| `src/tmc_processor/ui/app_shell.py` | UX-7 P0 operator presentation: report labels and diagnostics placement. |
| `src/tmc_processor/ui/components/export.py` | Shared operator report/fallback copy. |
| `src/tmc_processor/ui/components/peak.py` | Draft/confirmed Peak labels. |
| `src/tmc_processor/ui/workflows/batch.py` | Batch Export presentation and diagnostics placement. |
| `src/tmc_processor/ui/workflows/single.py` | Single Review/Export presentation and diagnostics placement. |
| `tests/test_setup_persistence.py` | Review UI assertion updated for the accepted operator text. |
| `tests/test_ooxml_template_export.py` | Shared contract, effective Peak/cache, package fidelity, and negative mutation regression tests. |
| `tests/test_template_ooxml_poc.py` | One-cell template package preservation proof. |
| `tests/test_ux7_p0_ui_cleanup.py` | Headless operator UI presentation and unchanged mode-decision checks. |
| `scripts/inspect_template_package.py` | Read-only template/package manifest diagnostic. |
| `docs/development/EVIDENCE_OOXML_UX7.md` | Chronological OOXML fidelity investigation, explicitly marking early synthetic Mapping projection as nonauthoritative. |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | Acceptance, qualification, limitations, finalization, and scope evidence. |

No raw source workbook, physical Mapping, analysis/Peak calculation, template binary, Standard routing, Safe PNG engine, Project Session format, or production dependency set changes are in scope. The obsolete untracked `tests/ux7_movement_completeness.json` synthetic projection and its only test consumer were removed before staging; they are not part of the proposed commit. All 16 native movement slots remain covered by the deterministic post-Review payload fixture.

The accepted deterministic manual Excel fidelity checkpoint, verified 33 shapes/drawings, two native charts, native HLOOKUP binding, and `Summary!U9:U22 = 1..14` remain release evidence. The protected three-real-workbook exact-branch replay is waived by the acceptance decision, with the missing reusable Mapping/Review fixture documented above. The local Excel runtime and `Summary!C9` text-date cache are documented limits; neither is represented as a new failure of the package proof. Batch Standard mode-label rerun remains a known UI limitation. Broader operator UI polish is deferred to UX-8 Issue #22.

Standard routing remains unchanged. Any later routing switch requires a separate review; this commit only qualifies the OOXML transport and operator workflow. No Office or Windows configuration changes were made.

### Final automated gate results

| Gate | Final result |
|---|---|
| Complete `.venv\Scripts\python.exe -m pytest -q --tb=short -p no:cacheprovider --basetemp C:\MyRD\tmc-processor-public\pytest-ux7-final-full` | **405 passed, 13 warnings, 0 failures**, 228.18 seconds. This meets the accepted 405+ suite and exceeds the historical 257 and UX-6 369 pass counts. |
| `.venv\Scripts\python.exe -m compileall -q app.py src` | **PASS**. |
| `git diff --check` | **PASS** before staging. Staged patch integrity is checked separately during finalization. |

The 13 warnings are openpyxl's known DrawingML read warning. The OOXML exporter does not save through openpyxl; independent package/relationship/chart/drawing preservation tests passed. The final suite includes deterministic post-Review write-plan and negative mutation tests, confirmed/effective Peak propagation and native formula caches, workflow invalidation and session/preset compatibility, Safe PNG, Batch mechanics, and the P0 operator UI assertions. Excel COM execution and fresh protected real-workbook replay were not part of this final run; the latter is explicitly waived for this branch.
