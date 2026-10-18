# v1.2.0 release readiness — evidence and handoff

Date: 2026-10-18 (Asia/Bangkok). **GO for independent engineering/release review; NO-GO for publication until independent acceptance and the exceptions below are resolved or explicitly accepted.** No release, tag, merge, deployment, branch deletion or modification of main is authorized or performed.

## Provenance and scope

| Item | Verified state |
|---|---|
| Repository | `C:\MyRD\tmc-processor-public`; origin `https://github.com/bokoboss/tmc-processor.git` |
| Initial checkout | `codex/issue42-v1-precision` / `a3df88956d3922afd2ad2b1ea524a42f2dda7026`; clean tracked tree; one pre-existing untracked investigation |
| Accepted main/base | `3f5a0f8e92893ac5f8b3174d47081d350a636133`, verified through GitHub branch API and fetched origin/main; initial HEAD has the same tracked tree |
| Feature branch | `codex/release-v1.2.0-preparation`, created directly from accepted main |
| Qualified implementation/payload commit | `97e042dd5714ea697e73aa8a2f6e3269b7c75cbf`; subsequent evidence-only commit does not change either ZIP payload |
| Final handoff HEAD | Exact final evidence-commit SHA, PR URL and exact-head CI status are recorded together in the PR description, final chat handoff and local `outputs/release-v1.2.0/handoff.json`. This avoids a self-referential commit hash in its own contents. |
| Local main | Remains `1514b142e566f1edf64ac1ba5aaf6b5a1523541b`; no checkout/update of local main |
| Published release | v1.1.0, published 2026-10-12; annotated tag object `0055c8771a3f1a92fcb99020146619e359fde4c9`, peeled commit `1514b142e566f1edf64ac1ba5aaf6b5a1523541b`. Latest-release API and remote tag read agree. No v1.2.0 remote tag exists at inspection. |
| Accepted-main CI | [39120306254](https://github.com/bokoboss/tmc-processor/actions/runs/39120306254): completed/success on exact base; Windows Python 3.10, Windows Python 3.12 and Linux web runtime Python 3.12 all successful |
| Preserved investigation | `docs/development/ISSUE42_INVESTIGATION.md` remains untracked, unchanged SHA-256 `ac7c5893566cb5f29c15aca642bbd720a36d253b43ff745f1dbdde527060ab29`; never staged |

Read AGENTS.md, PROJECT_PROFILE.md, v1.1.0 release notes, accepted Phase A/B/#42 evidence, installed workflow reference/manifest/templates, and pinned upstream model-routing policy. [Execution contract and pre-implementation gates](EXECUTION_RELEASE_V1_2_0.md) were defined before version edits. Applied Ponytail, workflow scrutiny, Computer Use for visible Excel, and debugging/recovery for qualification-harness/environment failures. STRICT qualification; existing inherited GPT-6 Codex session/effort, one worker, no subagents/model changes. Luna High was considered sufficient for mechanical edits; cross-tool evidence continuity justified keeping the existing session. No model-capability escalation was needed.

The PR contains **8 files**: `pyproject.toml` (one version string), two existing release-version tests (four literal expectations), README, PROJECT_PROFILE, candidate release notes, execution contract and this evidence. No application source, calculations, parser/cache/QC, WorkflowState, mapping, Peak/PHF, Single/Batch logic, exporter, native templates/maps, launcher, builder, dependencies or demo fixture bytes changed. The first seven files total 113 insertions/13 deletions; this evidence document is the only subsequent addition. Final Git diff statistics are included in the handoff.

The existing ignored output directories were preserved. All new qualification scripts, environments, logs, synthetic/private generated files and screenshots are confined to the new ignored `outputs/release-v1.2.0/`. Original raw workbooks and protected template hashes were checked unchanged. No private source/mapping/output file enters public Git history or either release ZIP.

## Accepted changes and evidence reconciliation

- Phase A ([PR #45](https://github.com/bokoboss/tmc-processor/pull/45)): source-integrity diagnostics/provenance, blank info severity, excluded-sheet diagnostic boundary; [implementation evidence](EVIDENCE_DATA_INTEGRITY_PHASE_A.md).
- Phase B ([PR #47](https://github.com/bokoboss/tmc-processor/pull/47)): one-reader loading and bounded, content/parser-revision keyed session-local Single cache, safe eviction/mutation isolation; [implementation evidence](EVIDENCE_EXCEL_IMPORT_PHASE_B.md).
- Issue #42 ([PR #48](https://github.com/bokoboss/tmc-processor/pull/48)): five V1 Native OOXML final totals, precise per-hour reconciliation and duplicate group rejection; [implementation evidence](EVIDENCE_ISSUE42_V1_PRECISION.md).
- The older #42 evidence and PR body still say GUI pending. The later [accepted merge record](https://github.com/bokoboss/tmc-processor/commit/3f5a0f8e92893ac5f8b3174d47081d350a636133), [GUI discussion](https://github.com/bokoboss/tmc-processor/pull/48), and closed issue establish later acceptance of opening/recalculation/charts and Fit Sheet on One Page preview, retaining original-pagination/A4-readability limitations. The merge record is authoritative for that later decision; it does not certify every candidate view/export.

Published v1.1.0 links and its historical limitation remain documented. Version 1.2.0 is explicitly unreleased. Linux CI/local health do not establish production Cloud acceptance; live deployment revision and candidate Cloud workflow remain unverified. No deploy was attempted.

## Environment and command/results ledger

`PY` below means `.venv/Scripts/python.exe -X utf8` from the repository root. `BUNDLE_PY` means `outputs/release-v1.2.0/extracted/TMC-Processor-v1.2.0-Windows/.venv/Scripts/python.exe -X utf8`. Commands use the repository-local Python because unqualified Windows Python aliases are unsuitable in this environment. Each logged invocation redirects stdout/stderr with `*> outputs/release-v1.2.0/<log>` and checks the process exit code. Validation harnesses assert their results and retain machine-readable JSON locally.

Repository environment: Windows 11 build 26300, Python 3.12.10, pytest 9.1.1, pandas 3.0.6, openpyxl 3.1.5, Matplotlib 3.11.2, Altair 6.3.0, Streamlit 1.64.0. Fresh bundle installation selected Streamlit 1.66.0 and pywin32 312 with the same versions of the other listed runtime libraries. Required dependencies remain the existing pyproject constraints; no lockfile or universal future-version guarantee was introduced.

| Command / method | Result / local record |
|---|---|
| `Get-Content AGENTS.md`, `PROJECT_PROFILE.md`, `.engineering-workflow.json`, workflow reference/templates, prior release and Phase A/B/#42 evidence; `rg --files` / targeted `rg -n` / `Get-Content` across builder, tests, launcher, metadata, CI, README, exporter and Batch APIs | Read before edits; stale release/issue text identified. Existing untracked investigation and prior local qualification harnesses read without modification. |
| `git status --short --branch`; `git remote -v`; `git log -1 --format=fuller`; `git branch -vv`; `git tag --list`; `git rev-parse HEAD origin/main main`; `git log --oneline -16 origin/main`; initial `git diff --stat HEAD origin/main` | Initial state above. Cached origin/main was initially `63a10ab`; remote inspection/fetch resolved it to exact accepted base; initial tracked tree equals accepted base. |
| GitHub connector GET branch/main, releases/latest, actions/runs/39120306254 and /jobs, issues/42, pulls/48 and issues/48/comments; pinned upstream MODEL_ROUTING_POLICY.md | Exact baseline and all 3 CI jobs verified; latest published release v1.1.0; Issue #42 closed; later manual GUI acceptance reconciled with older pending records. |
| Initial sandbox `gh --version`, `gh api .../releases/latest`, `gh run view 39120306254 ...`, `git ls-remote ...` | gh 2.102.0 available; sandbox DNS/network requests failed. Connected GitHub reads and bounded outside-sandbox Git requests succeeded; failures were not treated as a baseline pass. |
| Connector `/tags` read; raw web fetch of pinned governance README; `git -C` against the separate local governance checkout | Connector rejected `/tags`; web fetch cache miss; cross-repo Git read denied. Used fetched Git tags, remote tag read and pinned connector policy content. No other repository was changed. |
| `git fetch origin main --tags`; `git switch -c codex/release-v1.2.0-preparation 3f5a0f8e92893ac5f8b3174d47081d350a636133` | PASS outside sandbox. Initial in-sandbox switch could not create index.lock; retry succeeded. Existing untracked file carried through untouched. |
| `New-Item -ItemType Directory -Path outputs/release-v1.2.0 -ErrorAction Stop` | PASS; new output directory, no replacement of old outputs. |
| `PY -m pytest tests/test_windows_release_bundle.py tests/test_export_package.py -q --basetemp=outputs/release-v1.2.0/baseline-tmp` | **36 passed in 6.38 s**, before version edits; `baseline-release-tests.log`. |
| Same focused test command with `--basetemp=outputs/release-v1.2.0/focused-tmp` | **36 passed in 5.26 s** after version/expectation updates; `focused-release-tests.log`. UI version stamp, runtime/export metadata and bundle naming agree on 1.2.0. |
| `PY -m pytest -q --basetemp=outputs/release-v1.2.0/full-tmp --tb=short -o faulthandler_timeout=90` | **636 passed, 43 warnings, no failures/skips, 648.22 s; exit 0**; `full-regression.log`. Outside sandbox to avoid the previously diagnosed socketpair/AppTest sandbox hang. No tests disabled or product workaround added. |
| `PY -m compileall -q app.py src`; `git diff --check`; later `git diff --cached --check` | PASS, exit 0; `compile.log`, `diff-check.log`; final checks repeated after documentation. |
| `PY scripts/build_windows_release.py --demo --output outputs/release-v1.2.0/packages` | PASS, exit 0; `build.log`, both ZIPs and manifests. |
| Same builder with `--output outputs/release-v1.2.0/rebuild` | PASS, exit 0; `rebuild.log`; archive and manifest bytes identical for both packages. |
| `PY outputs/release-v1.2.0/verify_packages.py` | PASS, exit 0; `package-verification.log`, `package_results.json`; CRC, exact allowlists, every manifest entry size/hash, archive hash/size, nested OOXML/privacy checks, Git demo equality and payload equality. |
| `PY outputs/release-v1.2.0/smoke_launcher.py` | PASS, exit 0; `launcher-console.log`, `launcher-smoke.log`, `launcher_result.json`. Existing batch file runs from fresh extracted folder, creates local .venv, installs editable 1.2.0 plus dependencies/optional pywin32, imports extracted package, reaches HTTP 200 `ok`. Owned process tree stopped after smoke. Headless localhost environment variables suppress browser launch; no launcher edits/stubs. |
| Fresh bundle Python `-m pip check`; import/version assertions in launcher harness; repository Python `-m pip check`; `PY -c` platform/dependency inventory | Both environments: no broken requirements. Fresh installed distribution and APP_VERSION both 1.2.0; imports resolve inside extracted bundle; `environment.log`, `repository-pip-check.log`, launcher log. |
| `PY outputs/release-v1.2.0/smoke_extracted.py` (first attempt) | V1 Standard/Safe and reviewed Batch passed, then local harness raised TypeError on nonexistent V2 `use_native_template` argument. Preserved `extracted-demo-first-attempt.log`; inspected existing API and corrected only ignored harness. No product failure/fix is claimed. |
| `BUNDLE_PY outputs/release-v1.2.0/smoke_extracted.py` | **PASS**, exit 0; 6 Single workbooks and 3 reviewed Batch packages, both days, version/effective-Peak/QC/precise-total/final-cache checks; `extracted-demo.log`, `demo_results.json`. V2 uses existing native-only workbook route; no V2 Safe PNG workbook capability invented. |
| `PY scripts/smoke_demo.py` | PASS, exit 0; 4 raw sheets, 2,304 normalized rows, AM 08:00 / PM 17:00; `repository-demo.log`. Safe openpyxl export smoke; no Excel COM required. |
| `BUNDLE_PY outputs/release-v1.2.0/smoke_apptest.py` | PASS, exit 0; initial installed app execution plus 2 reruns, uploader rendered, no exceptions, module types unchanged; `extracted-apptest.log`. Only optional COM availability stubbed unavailable, as in existing runtime regression. One expected bare-mode ScriptRunContext warning. No fresh populated browser walkthrough claimed. |
| Copy/read existing ignored #42 real-workbook harness into new `qualify_real.py`, adapting only its old evidence path; `BUNDLE_PY outputs/release-v1.2.0/qualify_real.py` | PASS, exit 0; all 3 local workbooks, precise totals and five native final cells; raw/template hashes unchanged; `real-parity.log`, `corrected_real_workbook_evidence.json`. An initial shell quoting SyntaxError while constructing the copy was corrected with a literal PowerShell here-string; no original script/output was overwritten. |
| `PY outputs/release-v1.2.0/smoke_excel.py` first two attempts | Both stopped on first Safe PNG workbook: Excel CalculationState remained 2 (pending), including a 45 s wait and an automatic-mode/COM-message-pump diagnostic retry. Neither is a passing full-recalculation gate; `excel-first-attempt.log`, `excel-second-attempt.log`. |
| Final same Excel harness, native files first, collecting pending states instead of aborting subsequent checks | Read-only normal open and formula-error scans pass on all **6** files, hashes unchanged. Full recalculation **passes on 4 native V1/V2 files**; both Safe PNG files remain **UNVERIFIED, state 2**. `excel-smoke.log`, `excel_results.json`. No root cause is claimed for pending Safe calculation. |
| Computer Use `sky.list_apps`, launch Excel, observed open-dialog/keyboard/coordinate inputs, V1 Day 1 **Open Read-Only**, Ctrl+Alt+F9, Ctrl+P, Fit Sheet on One Page, close/Don't Save | V1 candidate opens without a repair/error prompt; diagram/tables/two charts render. Original A4 landscape Custom Scaling = 2 pages; one-page setting = 1 page with small content. `###` visible in final-total area at 55% worksheet view. Temporary preview changes discarded; source hash unchanged. Local `gui/` screenshots. Cached accessibility-element errors were recovered using refreshed screenshot-backed controls; no unobserved UI input or security bypass. |
| `PY -c` SHA-256 comparison against `pre-excel-hashes.json`; `Get-FileHash` investigation/templates/manifests | All 6 source exports unchanged after read-only GUI/COM checks; investigation hash unchanged. |
| `git diff --name-only 3f5a0f8 -- src templates scripts start_tmc_processor.bat assets samples requirements.txt` | Empty. Includes accepted numerical, Phase A QC, Phase B cache, WorkflowState, Peak, V1/V2 export/template/mapping and launcher/builder implementations. |
| `git ls-remote origin refs/heads/main refs/tags/v1.1.0 refs/tags/v1.2.0`; `git rev-parse v1.1.0 'v1.1.0^{}'`; final branch/release API reads | main still exact accepted SHA; published tag unchanged; no v1.2.0 tag. |
| Exact-file `git add`, staged diff/check, `git commit -m "Prepare v1.2.0 release candidate metadata and notes"` | PASS; implementation commit above. Only 7 authorized files staged; investigation stays untracked. Evidence-only commit/push/PR identifiers are recorded in final handoff. |

The 43 suite warnings comprise existing openpyxl DrawingML-read limitations and one deliberately invalid-infinity Peak/PHF RuntimeWarning. Native workbooks are inspected read-only and never resaved with openpyxl. The full regression includes source-integrity, import-reuse, confirmed/effective Peak, WorkflowState/session, Single/Batch, Standard/Safe routing, V1/V2 precision and native-template fidelity tests. No protected behavior is inferred solely from a version test.

## Package identities and inspection

Local artifacts: `outputs/release-v1.2.0/packages/`. These are **source-based distributions requiring Python >=3.10**, not standalone Windows executables. Both were built with the unchanged existing release builder. They have no GitHub Release download URL because publication is explicitly excluded.

| Archive | Bytes / files | SHA-256 |
|---|---|---|
| `TMC-Processor-v1.2.0-Windows.zip` | 1,222,201 / 64 | `6f7220309cef49e7d4837e8284a383d0104a37e6b74e0b047b337e6a8afed04c` |
| `TMC-Processor-v1.2.0-Demo-Files.zip` | 55,551 / 8 | `ea010c9fe310d08fb96327b8d34b50c24a9c7ddd438e7a4b4ff1abd10d3c58f7` |

Manifest SHA-256: Windows `404440d69d9c2f7a2b4d339dc27651a662686514954d0374f474acfb9f142394`; Demo `f9884c77cd1cbc2911a2d485c8d58e2e65d8bc43ca0ad75d5946cd56d01e9159`. Manifests contain ordered member paths, byte sizes and individual SHA-256 values. Extracted payload sizes are 2,296,037 and 66,715 bytes respectively. Both archives and manifests are byte-identical across two builds on this Windows checkout. Cross-platform line-ending/zlib identity and timestamped Excel-export byte reproducibility are not claimed.

Runtime allowlist equals the 11 existing required files plus 53 tracked Python sources. Required imports and declared dependencies are available: Altair, Matplotlib, openpyxl, pandas, Streamlit, and conditional tomli on Python <3.11. Optional pywin32 installed through the existing launcher. No dependency-policy change or standalone runtime embedding.

Support QR matches the committed image. Both native V1/V2 template binaries/maps are present; bundled copies equal the existing builder's portable-template output, which removes only saved-location metadata. Existing template-part regression confirms every other OOXML part is preserved. Demo files exactly equal the eight allowlisted committed Git blobs, including both input days, V1/V2 Excel/JSON mappings and Single session fixture. Runtime ZIP contains no demos/private sources; Demo ZIP contains only synthetic fixtures.

Every archive/member CRC, manifest hash/size, name uniqueness and allowlist passed. Nested XLSX contents were scanned for local absolute paths and common credential/private-key patterns; no findings. No `.git`, `.venv`, private raw data, outputs, secrets/config credentials, tests, governance docs or untracked investigation is shipped. The allowlist and committed/source byte comparison provide stronger provenance than pattern scans alone; no claim of an exhaustive universal secret detector.

## Numerical, workflow and Excel qualification

| Scenario | Precise PCU | AM / PM final PCU | F32 / W40 / AM22 |
|---|---:|---|---|
| Synthetic day 1, V1 and V2 native | 23,959.041 | 2,981 / 3,153 | 23,959 each |
| Synthetic day 2, V1 and V2 native | 23,590.054 | 2,872 / 3,036 | 23,590 each |
| Protected workbook 1 (Kabin Buri software fixture) | 87,023.550 | 6,433 / 8,920 | 87,024 each |
| Protected workbook 2 (Bo Phloi software fixture) | 14,032.315 | 1,238 / 1,311 | 14,032 each |
| Protected workbook 3 (Nong Prue software fixture) | 6,811.479 | 571 / 578 | 6,811 each |

Both synthetic days retain AM 08:00–09:00 and PM 17:00–18:00, confirmed before generation; exported metadata version is 1.2.0. Mapping Excel and JSON fixtures load for both schemes. V1 Standard native and explicit Safe PNG export work for both days. V2 retains its existing native workbook route. Reviewed Batch V1 Standard/Safe and V2 native each process both files, retain precise per-file totals, and produce valid ZIPs with batch_summary.xlsx/Batch_QC. No unsupported V2 Safe workbook or Batch Project Session is invented. Full tests qualify custom confirmed Peaks, excluded/failed Batch outcomes and invalidation separately; the fresh package smoke uses successful demo items.

The three protected mappings are existing deterministic **software parity fixtures**, not engineering-approved physical intersection mappings. Source files remain private and unmodified; V1/V2 original template hashes remain `1b9b8ddffe1a44799b5ba9012b627a8c0f88750bcf8cc2937674d51d9fc2c66b` / `409965c21f9ea4872af9d3e3ca416592bf31b2f9821b14305a76334e847df5d9`. Real-data reports/mappings are local only. No new blanket engineering-direction acceptance is claimed.

Fresh Excel COM uses Excel **16.0 build 20430**. All six synthetic workbooks open normally read-only with zero formula-result errors and unchanged file hashes. All four native files complete full recalculation with the five values above, 33 Summary shapes and two native charts. The two Safe PNG workbooks retain application calculation state 2; their complete COM recalculation remains unverified despite successful generation, opening and error scans.

Fresh visible GUI smoke covers **V1 day 1 only**. No repair/error prompt appeared. The effective-Peak note, intersection diagram, tables and both chart visuals were present. Ctrl+Alt+F9 was invoked; numerical stability is established separately by COM. Final-total `###` at the observed 55% worksheet view prevents an unconditional readable-display sign-off. A4 landscape original settings preview two pages; manual Fit Sheet on One Page previews a complete one-page arrangement with small text/charts. Neither proves readable physical output. Close/Don't Save plus SHA-256 verifies the preview did not modify the export. Candidate V2 GUI and other candidate views/printers remain unverified; accepted #42 manual V1/V2 baseline evidence is explicitly separate.

## Limitations, unverified gates and decision

The [candidate notes](../releases/v1.2.0.md) document original print pagination, A4 readability, candidate `###` display observation, Safe pending recalculation, Batch Project Session, historical #40 compatibility boundaries, QC/caching/Peak limits, independently rounded detailed subtotals, legacy COM-generation scope, and GUI qualification limits. No behavior was altered to conceal a limitation.

| Gate | Status / required independent action |
|---|---|
| Metadata/docs/package/focused/full/compile/whitespace | **PASS** on qualified payload; final evidence-only diff must remain documentation only. |
| First-run Windows launcher, installed runtime, extracted synthetic exports | **PASS**, with fresh dependencies selected by existing constraints. Only tested Windows/Python environment is certified locally; CI covers additional runtimes. |
| Native V1/V2 Excel opening/recalculation/objects/hashes | **PASS** for both synthetic days on recorded M365 build. No universal Excel or legacy COM-generation qualification. |
| Safe PNG Excel full recalculation | **UNVERIFIED** (CalculationState 2); independent reviewer must investigate or explicitly accept this scoped exception. Opening and formula-error scans pass. |
| GUI/print readability | **LIMITED**: fresh V1 smoke only; `###`, original two-page pagination and small A4 one-page output remain. Candidate V2 GUI, broader views/builds/printers and readable physical/PDF output not newly accepted. Requires explicit exception acceptance or separately authorized follow-up; no template/UI fix in this PR. |
| Production Cloud v1.2.0 | **UNVERIFIED**. No deploy or live end-to-end acceptance performed; v1.1.0 acceptance does not transfer automatically. |
| Exact final-head hosted CI | Verify PR checks for the final evidence commit; status/link captured at handoff, not inferred from accepted-main CI. Pending CI is not a pass. |
| Independent engineering/release acceptance | **PENDING**. Executor evidence and self-scrutiny do not substitute for independent acceptance, including numerical interpretation and known-limit decisions. |

Recommendation: **GO for independent review of this release-preparation PR. NO-GO for publication now.** The implementation/package regression gates pass, and no feature/calculation changes were needed. Publication requires independent review of the exact final head and CI plus explicit resolution/acceptance of the Safe recalculation and display/print exceptions and the intended Cloud acceptance scope. Any later fix needs separate authorization and qualification. This task ends with a feature-branch PR; it grants no merge/tag/release/deployment authority.

Local evidence (`outputs/release-v1.2.0/`) includes scripts, logs, manifests, package/demo/Excel/launcher JSON, original/fitted GUI captures, source-export hashes and the final handoff. These local artifacts are not uploaded to public history; a remote reviewer can rerun the existing public tests/builder and compare package manifests, or review the retained local evidence. Dependency resolution is not hermetic; future installs may select different versions. Rebuild/recheck the candidate payload from the accepted checkout before a separately authorized publication.
