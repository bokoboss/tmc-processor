# Fixed-hour Peak contract v1.0.1 — review evidence

Prepared 2026-10-03. This is an uncommitted maintenance candidate, not a published release.

## A. Baseline and change control

- Authoritative repository: `C:\MyRD\tmc-processor-public`.
- Initial branch/status: clean `main`, tracking `origin/main`.
- Initial HEAD and origin/main: `20b166c9d3f9dd9af06830522009f384109bd55d`.
- Initial/final worktrees: only `C:/MyRD/tmc-processor-public`.
- Requested branch: `maintenance/fixed-hour-peak-contract`; HEAD remains the accepted baseline above. Changes are uncommitted and available for review.
- Remote main was reconfirmed at the same SHA after qualification.
- No push, merge, issue mutation, tag creation, release mutation or asset upload performed.
- `main`, `legacy/v0.2.0`, existing tags/releases, local archive, `samples/raw` and original `.venv` preserved.

Classification: STRICT contract/session maintenance. Success gates were fixed-only UI and APIs; whole-hour confirmation; explicit legacy migration; preserved calculation/export/state behavior; focused/full tests; compile/whitespace; synthetic Single/Batch export; isolated Windows bundle qualification. The installed engineering workflow supplied scrutiny and independent review; its legacy `.ai-workflow` directory was not recreated.

## B. Exposure before implementation

See [complete baseline audit](ROLLING_PEAK_AUDIT_v1.0.1.md), including exact reference/line search results.

Single and Batch Analyze offered rolling through `PEAK_MODE_OPTIONS`. Batch had three public rolling defaults. The backend scanned raw quarter-hour starts in rolling mode and accepted arbitrary positive confirmed duration. Session load restored old mode/decisions. Single and Batch Review injected previous/suggested values into actual hourly options. Confirmation checked presence rather than alignment. Application fingerprints included obsolete mode variability. README called native rolling support a current limitation. The committed synthetic demo session already uses `fixed_hourly`.

## C. Resulting product contract

Only `fixed_hourly` is supported. Candidates start at HH:00 inside the unchanged AM/PM search windows. Confirmed Peaks must be valid whole hours lasting exactly 60 minutes; four raw 15-minute intervals compose a Peak Hour. No rounding or coercion occurs. Single/Batch Analyze show fixed-hour guidance and no Peak mode selector. Review options come from actual supported hourly rows.

The rolling constant remains explicitly deprecated for legacy identification. Its active calculation branch is removed. Direct Single/v2/Batch processing and confirmed Peak calculation reject rolling mode. The generic WorkflowState authority and its transition semantics are retained; application analysis fingerprints use the fixed mode.

## D. Changed files and scope

Runtime (11 files):

- `src/tmc_processor/constants.py`: fixed-only product options, deprecated legacy constant.
- `src/tmc_processor/time_utils.py`: shared strict whole-hour validation and label parser.
- `src/tmc_processor/peaks.py`: fixed candidates and confirmation/session/export-boundary validation.
- `src/tmc_processor/pipeline.py`: early unsupported-mode rejection for Single/v2.
- `src/tmc_processor/batch.py`: fixed API defaults, validated labels/readiness, actual hourly options only.
- `src/tmc_processor/session.py`: explicit migration warning, preservation, clearing, authoritative invalidation, validation before serialization.
- `src/tmc_processor/application/state.py`: validate before confirmation mutation.
- `src/tmc_processor/application/workflow.py`: fixed mode in application analysis revisions.
- `src/tmc_processor/ui/app_shell.py`: validated confirmation/readiness and actual hourly options.
- `src/tmc_processor/ui/workflows/single.py`: fixed Analyze, removed selector.
- `src/tmc_processor/ui/workflows/batch.py`: fixed Analyze and actual supported Review choices.

Tests (6 files): `test_fixed_hour_peak_contract.py` (new), `test_export_package.py`, `test_setup_persistence.py`, `test_ux7_standard_ooxml_routing.py`, `test_ux8_p1_ui_hierarchy.py`, `test_workflow_contract_adapter.py`.

Docs/metadata (7 files): `pyproject.toml`, `README.md`, `USER_GUIDE.md`, `PROJECT_PROFILE.md`, `docs/releases/v1.0.1.md` (new), baseline audit (new), this evidence (new).

Final scope: 24 changed/new files. Tracked diff: 20 files, 194 insertions / 147 deletions; the four new test/audit/release/evidence files are additional. Product changes are confined to contract boundaries, adapters and the obsolete candidate branch. Normalization, Mapping/PCE implementation, PHF formula, export implementation, template files, and `workflow_state.py` are unchanged.

## E. Legacy sessions

Example input: `peak_mode=rolling_60min`, confirmed AM `08:15–09:15`, confirmed PM `16:00–17:00`.

Load warns that unsupported legacy selections were not reused, changes mode to fixed, and clears both Peak decisions/sources so the entire old analysis requires fresh Analyze → Review → explicit confirmation. Project/survey metadata, directions, Mapping, PCE and AM/PM windows survive. Applying to populated state clears old analysis/export and invalidates Single WorkflowState analysis/review/export readiness and analysis/review revisions using existing transitions; source/mapping revisions and separate Batch state survive. Incompatible decisions are also cleared when their saved mode is already fixed. Valid fixed-hour round trips retain their values without migration warnings.

The transient re-analysis marker is not a new persisted session schema. New session creation validates before time serialization, preventing truncation of second-bearing invalid values.

## F. Enforcement evidence

`is_supported_peak_period` is shared by label and session validation; domain confirmation validates before calculations, and application confirmation validates before state changes.

| Period | Result |
|---|---|
| 07:00–08:00, 11:00–12:00, 15:00–16:00 | accepted |
| 08:15–09:15, 08:30–09:30, 08:00–09:15 | rejected |
| malformed/out-of-range times or nonzero seconds | rejected |
| 23:00–00:00 | rejected, consistent with existing non-overnight calculation |

Direct tests cover empty-data calculation, state adapters, export resolution, raw/session application, serialization and public processing APIs. Same valid reconfirmation remains no-churn. Changed confirmation follows the existing invalidation contract.

## G–J. Automated validation

Environment: Windows, Python 3.12.10, project `.venv`.

- Focused new contract: `python -m pytest tests/test_fixed_hour_peak_contract.py -q` → **27 passed**.
- Broader focused iterations: **35 passed**, then **30 passed** after replacing obsolete rolling expectations.
- Full `python -m pytest` → **504 passed, 0 failed, 14 warnings**, 281.26 seconds.
- `python -m compileall -q app.py src` → exit 0.
- `git diff --check` → exit 0; also independently checked by reviewer.

The 14 full-suite warnings are openpyxl's known DrawingML limitation. Native template fidelity checks and Standard routing tests pass. Initial red contract evidence: 12 failed / 1 passed. The first full run passed 503 tests and failed only a stale hardcoded 1.0.0 assertion; it was updated for this patch and the complete suite rerun passed. An early AppTest fixture invoked optional COM; the focused fixture was corrected to disable the probe, as existing UI tests do. Product COM implementation was untouched.

Independent fresh-context review: **PASS**, no material unresolved findings. Review examined confirmation/session/API/export paths and scope, and independently probed rejected seconds/midnight values and authoritative migration invalidation. Its optional focused rerun was incomplete when the review returned; the completed automated results above are executor evidence. Four review findings were remedied before green qualification: pre-serialization validation, direct authoritative state invalidation, midnight consistency, and Batch suggestion injection.

## K. Synthetic populated workflows

Synthetic committed demos only: `DEMO_TMC1_FourLeg.xlsx`, `DEMO_TMC1_FourLeg_Day2.xlsx`, and their Mapping Preset. No protected real workbook was used.

Using Streamlit AppTest against the extracted application:

| Workflow | Analyze/Review | Standard | Safe PNG |
|---|---|---|---|
| Single demo | fixed, actual hourly options, explicit confirmation | valid XLSX, native template mode, 2 native charts | valid XLSX, Safe PNG mode, PNG chart artifacts |
| Two-file Batch | both analyzed, actual hourly options, explicit bulk acceptance | valid ZIP, 2 reports, 2 native charts/report, no failure | valid ZIP, 2 Safe PNG reports, no failure |

Recorded Batch confirmed periods and Single export setup satisfy the shared whole-hour rule: AM `08:00–09:00`, PM `17:00–18:00` in all four qualification runs. Output ZIP/XLSX integrity was checked. Mode switches preserve confirmed Peak decisions under the existing export invalidation behavior.

## L. Prepared artifacts

Builder: `python scripts/build_windows_release.py --demo --output dist/v1.0.1-candidate` → exit 0.

| Candidate | Bytes | SHA-256 |
|---|---:|---|
| TMC-Processor-v1.0.1-Windows.zip | 307047 | 5cac112adf71434a86a5af561f4df388167b43befaa14eb79acd53eee3e7ec05 |
| TMC-Processor-v1.0.1-Demo-Files.zip | 55551 | 3c6572505f4e15da4aeb06246372e991ab59253b63f4832ae9a100868d6e41d4 |

Local location: `dist/v1.0.1-candidate/`. Both manifests and per-file hashes are beside the ZIPs. Runtime payload has 61 files. All non-template extracted payload files match current source byte-for-byte; existing builder portable-template checks pass in the full suite. Demo data comes from committed synthetic HEAD content; no raw samples or environment are packaged. Neither archive was published.

## M. Fresh-bundle runtime proof

- Extracted into new `dist/v1.0.1-qualification/` directory.
- Created a new `.venv` with Python 3.12.10; installed editable package/dependencies against extracted `pyproject.toml`. No copied source environment or system site packages.
- Fresh imports, `sys.prefix`, `app.py`, template and template-map paths all resolve inside `TMC-Processor-v1.0.1-Windows`.
- Version resolves to 1.0.1. Fresh dependencies include Streamlit 1.65.0, pandas 3.0.6, openpyxl 3.1.5 and Matplotlib 3.11.2.
- Actual `start_tmc_processor.bat` launched the extracted `.venv` and passed its import check. Streamlit health at `http://127.0.0.1:8511/_stcore/health` returned HTTP 200 / `ok`.
- Playwright with installed Chrome opened that server and verified Single and Batch Analyze show the whole-hour caption, expose no combobox/mode selector, and contain neither `rolling_60min` nor `Peak calculation mode`.
- Populated Analyze → Review → Standard/Safe PNG proof is the fresh-environment AppTest flow in section K. Optional Excel COM was disabled for that qualification; Standard OOXML works without it.

Local reproducibility/evidence artifacts (ignored, not shipped): `dist/verify_fixed_hour_runtime.py`, `dist/v1.0.1-fresh-runtime.json`, `dist/v1.0.1-runtime.log`, `dist/v1.0.1-install.log`, `dist/v1.0.1-launcher.log`, and browser snapshots under `dist/v1.0.1-qualification/.playwright-cli/`.

## N. Proposed Issue #23 closure comment — not posted

> This issue assumed that rolling_60min was an intended product requirement. The product contract has now been clarified as fixed whole-hour Peak reporting: valid HH:00–HH+1:00 periods composed of four 15-minute survey intervals within AM/PM search windows. Non-hour-aligned rolling Peaks are intentionally unsupported, so redesigning the native Excel template for them is unnecessary. No Peak is rounded or coerced. Legacy rolling selections are explicitly discarded with a warning while reusable setup/Mapping/PCE are preserved; fresh analysis and explicit confirmation are required. Following merge/release acceptance of the v1.0.1 maintenance patch, close this issue as NOT PLANNED / NO LONGER APPLICABLE.

GitHub Issue #23 was reconfirmed **open**, with zero comments; no closure/comment was submitted.

## O. Limits and remaining acceptance work

- Candidate remains local and uncommitted. No v1.0.1 remote CI or Python 3.10 run exists for this diff because push/publication are prohibited. Remote Windows Python 3.10/3.12 qualification belongs to the later authorized PR/release gate.
- Runtime qualification uses synthetic demos, as requested; protected real-workbook regression was not invoked. Mapping, normalization, PCE/PHF algorithms and templates were not changed, and their existing regression coverage passed.
- Fresh Streamlit logged a nonfatal Arrow conversion diagnostic for a mixed object/time display table, then applied its conversion fallback. Both workflows and all exports completed with no AppTest exception. This diagnostic was not repaired as unrelated display/dependency scope.
- Native OOXML content/fidelity is covered by existing tests and generated-artifact inspection; no new manual Excel visual fidelity exercise or COM qualification was performed.
- Launcher proof starts with the separately created/installed fresh environment; its first-time bootstrap branch was not rerun. Existing launcher/builder regression coverage passed.
- Existing `v1.0.0` release remains ID 402390636, assets 607507294 (Windows, 305413 bytes) and 607507338 (Demo, 55551 bytes). `v0.2.0` release remains ID 328527203. No existing release assets were modified.

Review disposition: ready for review of the fixed-hour maintenance candidate. Merge, release and Issue #23 closure remain deferred pending acceptance and subsequent authorization.
