# Excel import performance — Phase B evidence

Issue: [#46](https://github.com/bokoboss/tmc-processor/issues/46). Date: 2026-10-13, Asia/Bangkok.
Execution contract: [EXECUTION_EXCEL_IMPORT_PHASE_B.md](EXECUTION_EXCEL_IMPORT_PHASE_B.md).

## Scope and provenance

- Root: `C:\MyRD\tmc-processor-public`; origin: `https://github.com/bokoboss/tmc-processor.git`.
- Initial branch/HEAD: `codex/data-integrity-qc-phase-a` / `460b6544f514876c67a9d29b51cf8944ff4698df`; working tree clean. Initial remote-tracking `origin/main` was `29bcfc00d05de1a1316278a939d608d0ee4ea457`.
- Read AGENTS.md, PROJECT_PROFILE.md, Issue #46 (open, no comments), installed workflow manifest/reference/templates and upstream governance at installed revision `3547ae260feacf8fc9a102b2abfdb13881e36dab` (v1.4.1). The older published-release paragraph in PROJECT_PROFILE.md does not supersede the explicit Phase B starting SHA.
- Fetched `origin/main`, verified `3758330c15e648fa07c0a95cfe35bc2bd8b20d4b`, and created `codex/excel-import-phase-b` directly from that SHA. The Phase B implementation was committed and pushed on this feature branch and opened as PR #47. This final cache-version remediation started from PR HEAD `039681f86930a6e14c3732b2e39830fc4e8e3d86`.
- Local `main` remains `1514b142e566f1edf64ac1ba5aaf6b5a1523541b`; `origin/main` remains the required base. No main modification, merge, release, deployment, branch deletion or history rewrite.
- Existing chat, one worker on the inherited session model/effort (exact runtime selection not exposed); no subagents/model changes. Applied Luna-first consideration and bounded execution routing as recorded in the contract. Applied Ponytail, TDD, Streamlit, debugging, and the installed workflow's scrutiny/checkpoint guidance. No legacy `.ai-workflow/` recreated.

## Root cause and intervention

Single's app-shell rerun path called detection, preview and details independently. Preview itself called details; each details load opened once for detection and again for `pd.read_excel`. Result: **5 workbook opens, 2 read calls and 2 parses per source sheet on every rerun**, even without upload changes. All collapsed preview containers still incurred that work.

The shared importer now owns one `pd.ExcelFile` context for both sheet-name detection and reading, and closes it deterministically. Parsing/provenance logic and pandas NA options are unchanged. Single loads details once; detected names are dictionary keys and preview is derived from those same details. Preview frames are independent copies.

`application/workbook_import.py` retains one immutable serialized details payload in the caller's session mapping. It hashes the actual source bytes with SHA-256 on every request; filename and size alone cannot produce a cache hit. The cache identity also includes `SINGLE_WORKBOOK_PARSER_REVISION`, stored with the digest and payload. Revision-less legacy entries, malformed envelopes, and entries from another revision are evicted as misses. If a matching entry fails deserialization or does not contain `ParsedSheet` details, it is evicted and reparsed from the original upload bytes; errors from that source parse propagate normally. Bump the parser revision whenever parsed output/provenance semantics or the serialized `ParsedSheet` contract changes, including importer fixes that should affect already-open Streamlit sessions. Serialization occurs only on a miss, deserialization returns independent DataFrames/debug lists on hits, and the initial returned objects cannot mutate the serialized entry. It unpickles only its own internal serialized result, never upload or Project Session content.

Replacement evicts the previous entry **before** parsing, including failed replacements. Upload clear removes the entry immediately, and a run with no retained upload also removes it. Reupload after clear parses again. Renamed identical bytes reuse details while the existing filename/UI/workflow adapters still update source identity and invalidate obsolete output. Existing nonempty title/survey metadata remains preserved, as at baseline.

The cache adds no engineering readiness rules: WorkflowState and its existing revision/invalidation adapters are unchanged. Metadata/mapping/PCE/Peak changes do not change the parsed source, and still follow existing downstream invalidation. An inactive Single upload may remain retained when switching to Batch, matching existing upload persistence; its single cache entry remains bounded and is removed with source clear/session disposal. Project Session JSON excludes the cache and raw workbook data.

## Environment and benchmark method

Windows 11 `10.0.26300`, Python **3.12.10**, pandas **3.0.6**, openpyxl **3.1.5**, Streamlit **1.64.0**, pytest **9.1.1**, existing `.venv`. Both reported comparisons used this same interpreter/runtime and the same source SHA-256 values. Originals were read without modification.

Five timed repetitions per scenario, reporting median and min/max; tracing memory in a **separate untimed** repetition. GC runs before each sample. Counters wrap actual `openpyxl.load_workbook`, `pd.read_excel`, `ExcelFile.parse`, `OpenpyxlReader.get_sheet_data` and `parse_raw_count_sheet`. Counts were identical in all five repetitions for every reported fixture/scenario. Counter wrappers remain benchmark/test-only.

“Cold” means a new upload/session cache, **not** a flushed Windows filesystem cache or a fresh Python interpreter. “Warm” means the same already-loaded content/session. Import-only times include detection, preview and details; actual Streamlit times cover the entire `app.py` Data rerun with COM availability stubbed identically. These are local measurements, not network upload timings or population-wide speed guarantees. Normal scheduling/background-load noise remains; sample ranges are reported rather than estimated speedups.

| Fixture | Upload bytes | Total/source sheets | Parsed frame deep bytes | Retained serialized bytes |
|---|---:|---:|---:|---:|
| Synthetic demo | 20,664 | 4 / 4 | 531,754 | 250,347 |
| Kabin Buri | 3,978,591 | 25 / 24 | 4,085,828 | 1,763,269 |
| Bo Phloi | 6,677,459 | 16 / 15 | 1,945,672 | 1,003,607 |
| Nong Prue | 6,856,192 | 13 / 12 | 1,579,553 | 800,891 |

### Import-only medians (seconds)

| Fixture | Single cold before → after | Single warm before → after | Batch source-load before → after |
|---|---:|---:|---:|
| Demo | 0.0988 → 0.0475 | 0.0980 → 0.0015 | 0.0479 → 0.0449 |
| Kabin Buri | 0.9684 → 0.4387 | 0.9711 → 0.0116 | 0.4688 → 0.4322 |
| Bo Phloi | 0.5499 → 0.2410 | 0.5653 → 0.0068 | 0.2618 → 0.2343 |
| Nong Prue | 0.4111 → 0.1830 | 0.4162 → 0.0060 | 0.1976 → 0.1784 |

| Fixture | Single cold range before → after | Single warm range before → after | Batch source-load range before → after |
|---|---|---|---|
| Demo | 0.098341–0.105337 → 0.044088–0.064318 | 0.097069–0.098013 → 0.001407–0.001588 | 0.047354–0.049079 → 0.043047–0.046414 |
| Kabin Buri | 0.954387–0.983093 → 0.434832–0.458874 | 0.964909–0.984875 → 0.011340–0.012025 | 0.452841–0.478442 → 0.423076–0.432440 |
| Bo Phloi | 0.545298–0.561880 → 0.238231–0.249753 | 0.555819–0.569332 → 0.006636–0.007098 | 0.254142–0.271388 → 0.232252–0.234855 |
| Nong Prue | 0.410308–0.418427 → 0.182268–0.186794 | 0.407446–0.442075 → 0.005507–0.007191 | 0.192434–0.200098 → 0.176291–0.181480 |

Actual demo Streamlit cold median: **0.224356 → 0.168860 s** (ranges 0.217367–0.434577 → 0.153090–0.368317). Actual unchanged warm median: **0.119679 → 0.015889 s** (ranges 0.117316–0.128953 → 0.014505–0.037486).

### Exact open/read/parse counts per run

`N` is the number of detected source sheets (4/24/15/12 respectively).

| Path | Workbook opens | read_excel / ExcelFile.parse calls | Worksheet reads | Source parses |
|---|---:|---:|---:|---:|
| Single cold or warm, baseline | 5 | 2 / 2 | 2N | 2N |
| Single cold, after | 1 | 1 / 1 | N | N |
| Single unchanged warm, after | **0** | **0 / 0** | **0** | **0** |
| Batch one source load, baseline | 2 | 1 / 1 | N | N |
| Batch one source load, after | 1 | 1 / 1 | N | N |

Actual demo Streamlit reruns reproduce these Single counts, including zero warm parsing, across all five repetitions. Integration regressions also instrument navigation through Data, Mapping, Analyze, Review and Export on unchanged content.

Batch was evaluated separately: Mapping preset application detects workbook sheets; `analyze_batch_files` parses each input once; reviewed export `_process_one_file` / `_process_one_file_v2` loads the input again for confirmed-peak processing. The shared loader removes one redundant open from those loads. **No Batch cache or Batch stage logic changed.** Batch timing above measures its source-import component, not total analysis/export latency. Retaining parsed frames for every Batch input would add privacy/memory policy and cache ownership beyond the demonstrated Single fix; it is deferred.

## Memory, isolation and invalidation

One entry maximum; serialized payload cap **33,554,432 bytes (32 MiB)** per session, plus small key/digest/tuple overhead. No global decorator/module cache, disk cache or retained Excel reader/handle. Measured entries were all below 1.8 MB. Inputs whose parsed serialization exceeds the cap still load correctly but are not cached and will parse again on rerun. The limit bounds **retained cache payload**, not the existing source upload or transient DataFrames/serialization during a load.

| Fixture | Cold traced peak before → after (bytes) | Warm traced peak after (bytes) | Warm allocations left after releasing result (bytes) |
|---|---:|---:|---:|
| Demo | 3,259,527 → 1,874,967 | 594,410 | 2,419 |
| Kabin Buri | 7,673,362 → 8,579,220 | 4,962,258 | 5,815 |
| Bo Phloi | 8,102,654 → 4,613,336 | 2,427,374 | 4,198 |
| Nong Prue | 7,177,656 → 4,154,201 | 1,968,650 | 4,784 |

Cold memory is not uniformly reduced: Kabin Buri temporarily adds serialization to the live frames. Warm tracing starts after cache initialization, so the existing cache is excluded from traced new allocations; the separately measured serialized-byte column above accounts for that retention. These are Python tracemalloc measurements, **not process RSS or allocator-release guarantees**. Explicit clear/replacement/session isolation and forced oversized-cache bypass are tested deterministically, without relying on noisy memory thresholds. Clearing one session leaves another session's cache intact.

## Validation gates

Logs and private snapshots under ignored `outputs/` remain local and are not part of the public diff.

| Gate | Command / method | Result |
|---|---|---|
| Baseline runtime | `.venv/Scripts/python.exe -m pytest tests/test_app_runtime.py -q -o faulthandler_timeout=45` | 11 passed in 1.98 s outside sandbox; `outputs/phase_b_runtime_baseline.log`. |
| RED | `.venv/Scripts/python.exe -X utf8 -m pytest tests/test_excel_import_reuse.py -q -o faulthandler_timeout=45` before runtime edits | 14 failed in 1.79 s; `outputs/phase_b_red.log`. Cold open count and actual Single cold parsing reproduced duplication; absent reuse API accounts for the other failures. |
| Initial GREEN | Same focused command after implementation/test-contract corrections | 14 passed in 2.22 s; `outputs/phase_b_green.log`. |
| Final focused | `.venv/Scripts/python.exe -X utf8 -m pytest tests/test_excel_import_reuse.py tests/test_upload_persistence.py tests/test_workflow_state.py tests/test_workflow_contract_adapter.py tests/test_web_session_safety.py -q -o faulthandler_timeout=45` | **75 passed in 2.79 s**; includes 16 new cases and Project Session cache exclusion/round-trip. `outputs/phase_b_focused.log`. |
| Phase B full regression | `.venv/Scripts/python.exe -X utf8 -m pytest --basetemp=outputs/phase_b_cache_revision_full_tmp -q -o faulthandler_timeout=90` | **620 passed, 0 failed/skipped, 31 warnings in 521.86 s (8:41)**, completed with exit 0. |
| Cache revision RED/GREEN | `.venv/Scripts/python.exe -X utf8 -m pytest tests/test_excel_import_reuse.py -q -k 'parser_revision_change or revisionless_legacy or corrupt_cached_payload or corrupt_cache_does_not_hide'`, then full import file | RED: all 4 regressions failed on starting HEAD (stale reuse/revision ignored, stale legacy payload, unpickle error, and source parse error masked). GREEN: **20 passed in 2.40 s**. |
| Cache revision focused compatibility | Existing cache/upload/WorkflowState/Project Session command plus four new regressions | **79 passed in 2.91 s**; includes unchanged revision/bytes hit, changed revision exactly one parse, corrupt and legacy entry eviction, error propagation. |
| Single/Batch/Project Session lifecycle | Full-suite AppTest and contract tests, including populated Single analyze/review/export, Single/Batch stage round trips, explicit clear and session round trips | PASS. Includes `test_streamlit_apptest_setup_values_survive_processing_and_export`, Batch analyze/setup navigation, source/preset clear, reviewed Batch confirmation/export and Project Session mapping/setup/Peak contracts. New import tests instrument unchanged stage reruns and real uploader replacement. |
| Three real workbooks | `.venv/Scripts/python.exe -X utf8 scripts/validate_excel_import_parity.py baseline`, then `green` | **PASS**: exact parsed/provenance/normalized/QC/Peak/PHF/summary/export-content parity, all 3 Single and reviewed Batch exports, and original source hashes unchanged. Logs: `outputs/phase_b_parity_baseline.log`, `outputs/phase_b_parity_green.log`; ignored snapshot `outputs/phase_b_real_baseline.pkl`. |
| Benchmarks | `.venv/Scripts/python.exe -X utf8 scripts/benchmark_excel_import.py --output outputs/phase_b_before.json --streamlit`, then add `--optimized` and use `outputs/phase_b_after.json` | Completed both with five runs. Environment/fixture hashes match and all repetition counts are stable. |
| Compile | `.venv/Scripts/python.exe -X utf8 -m compileall -q app.py src scripts/benchmark_excel_import.py scripts/validate_excel_import_parity.py tests/test_excel_import_reuse.py` | PASS. |
| Whitespace | `git diff --check`; new files checked with `git diff --no-index --check -- NUL <file>` | PASS: no whitespace errors. No-index returns 1 for added-file differences and prints only normal LF/CRLF conversion warnings. |
| Server health | `.venv/Scripts/python.exe -X utf8 scripts/smoke_streamlit.py` | PASS; `outputs/phase_b_streamlit_health.log`. |
| Protected scope | Diff against baseline for PCU/Peak/summaries/QC/WorkflowState/Session/exporter/Batch/templates | Empty. Both templates byte-identical to baseline Git blobs; all 3 source hashes unchanged. |
| CI | GitHub Actions on the cache-version remediation head | Pending until the focused fix is pushed and the PR checks complete. Do not infer hosted CI success from earlier Phase B runs. |

The first **baseline** runtime test hung in the sandbox before app execution: faulthandler located `socket.accept` inside Python's `_fallback_socketpair` while AppTest constructed its asyncio event loop. That invocation was interrupted and **is not a pass**. The identical unmodified tests passed outside the sandbox. All reported final AppTest/full-suite and before/after benchmarks use the same outside-sandbox condition; no app/test workaround, disabled test or extended-loop shortcut was introduced. Subsequent temporary failures were test assumptions/API usage, corrected after inspecting baseline metadata persistence, inactive-mode upload retention and AppTest's public mapping interface; the runtime contracts were preserved.

During the remediation full suite, one existing Streamlit privacy-disclosure test printed a Windows fatal COM exception (`0x80010108`) from the background Excel availability probe. The test still passed, the full suite exited 0, and the same test reproduced the printed exception when run alone while passing. This code path is outside the cache remediation and was not changed.

## Real-workbook parity method and limitations

The qualification script captures exact frames for parsed data/raw/debug provenance, automatic and confirmed Peak/PHF, normalized records, QC and all summaries. It compares SHA-256 of **all exported worksheet cell contents including QC and formula strings**, with fixed generation metadata, for direct Safe exports, requested native-template exports and all reviewed Batch workbooks including batch summary. Source hashes before/after must match. Batch totals are checked against corresponding Single totals.

The identical qualification mapping cycles existing V1 canonical movements across source-sheet names. It demonstrates software parity; it is **not physical movement engineering acceptance**. On all three baseline sources the requested native route falls back to Safe PNG with `Hourly movement total disagrees with support-sheet total for Total.` The final run preserved that result; successful native-template export on these artificial mappings is not claimed. Existing synthetic native/V2/template regressions passed in the full suite and qualify those unchanged routes separately. No Excel GUI recalculation or manual browser acceptance is claimed.

Final comparison passed for all three sources. Kabin Buri: 13,824 normalized rows, 82,621 vehicles, 87,023.550 PCU; Bo Phloi: 8,640 rows, 16,905 vehicles, 14,032.315 PCU; Nong Prue: 6,912 rows, 7,908 vehicles, 6,811.479 PCU. No QC cells or other workbook sheets were excluded from export-content comparison.

Protected template hashes match the baseline byte-for-byte:

- V1: `1b9b8ddffe1a44799b5ba9012b627a8c0f88750bcf8cc2937674d51d9fc2c66b`.
- V2: `409965c21f9ea4872af9d3e3ca416592bf31b2f9821b14305a76334e847df5d9`.

## Changed files and review

| File | Purpose |
|---|---|
| `src/tmc_processor/importer.py` | One workbook context; preview from existing details; close readers. |
| `src/tmc_processor/application/workbook_import.py` | One session-local content cache with parser revision, safe incompatible/corrupt-entry fallback, 32 MiB retained-payload cap and mutation isolation. |
| `src/tmc_processor/ui/app_shell.py` | Consume one details result; clear cache/detected names with source clear. |
| `tests/test_excel_import_reuse.py` | Parsing counts, stale-result/replacement/clear, filename behavior, parser revision/legacy/corrupt payload, provenance, empty source sheet, mutation/privacy/bounds, actual reruns/uploader and Project Session. |
| `scripts/benchmark_excel_import.py` | Reproducible timings, resource counters, memory and actual AppTest reruns. |
| `scripts/validate_excel_import_parity.py` | Local source/frame/export/Batch parity qualification; private snapshots ignored. |
| `docs/development/EXECUTION_EXCEL_IMPORT_PHASE_B.md` | Bounded contract, routing, scrutiny and success gates. |
| `docs/development/EVIDENCE_EXCEL_IMPORT_PHASE_B.md` | Review/handoff evidence. |

The baseline benchmark was run **before runtime edits** with the benchmark script copied into the exact baseline checkout. Reproduction requires the same approach in a separate clean checkout at the baseline SHA; running the default legacy three-call path on the final checkout benefits from the updated shared importer and does not recreate the original five opens. No existing checkout must be reset to reproduce it. Real parity baseline snapshots likewise require the baseline source tree.

Final diff summary: **8 files**; runtime intervention confined to **3 files, 56 insertions and 16 deletions** (including the 34-line new helper). The remaining files are regression tests, reproducible qualification scripts, contract and evidence. No unrelated user changes existed or were discarded; all final untracked files are the six named additions above, and generated exporter temporary directories cleaned up normally. Local logs, snapshots and real workbooks remain ignored.

Final local scrutiny: **GO for independent acceptance review**, with all local implementation gates passing. The diff was inspected against the contract: one session cache, actual content hashing, eviction before errors, mutation isolation, byte/entry bounds, filename adapters preserved, no cache-defined readiness, shared loader improvement only for Batch, no formula/template changes. Hosted CI was queued after the authorized push/PR; its result remains an external acceptance gate. This package does not grant engineering acceptance or merge approval.
