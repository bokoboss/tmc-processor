# Project Profile

## Identity
- Project name: TMC Processor
- Repository URL: https://github.com/bokoboss/tmc-processor
- Authoritative local path: `C:\MyRD\tmc-processor-public`
- Primary branch: `main`
- Package/application version: `1.1.0` — published Stable/Latest release [`v1.1.0`](https://github.com/bokoboss/tmc-processor/releases/tag/v1.1.0)
- Legacy public beta: `v0.2.0`, preserved at branch `legacy/v0.2.0`
- Installed workflow source: https://github.com/bokoboss/engineering-development-workflow
- Installed workflow revision: `3547ae260feacf8fc9a102b2abfdb13881e36dab`
- Installed workflow version: `v1.4.1`

## Current accepted baseline and release state
- Current published and accepted release baseline: `main@1514b142e566f1edf64ac1ba5aaf6b5a1523541b` (`v1.1.0`, Stable/Latest, published 2026-10-12).
- v1.1.0 GitHub Actions CI run [#38263923359](https://github.com/bokoboss/tmc-processor/actions/runs/38263923359) passed all 3 jobs; Streamlit Community Cloud acceptance passed at https://tmc-process.streamlit.app/.
- Issue #36 release publication, Issue #38 Cloud rerun fix, and Issue #40 Excel interoperability investigation are closed. Issue #42 remains open as a documented V1 Summary PCU discrepancy.
- Historical v1.0.1 baseline before Issue #30: `main@39c1ba906dcc284b714e5980b9ef9e23dfe47827`; its full suite had 504 passed and CI #113 passed on Windows / Python 3.10 and 3.12.
- Before the separately approved Issue #30 web-app enhancement, the project was in maintenance mode.
- Historical Issue #36 preparation baseline: `main@203bc5f3f584d0245f46c27d58da74c3d973fd65` (PRs #31, #33, and #35 had merged at that point).
- v1.1.0 release includes version alignment, Windows Support QR packaging, hosted Streamlit readiness, Native V2 Summary and round-once V2 Summary PCU. Calculation, Mapping, QC, Peak and V1 behavior remain protected.
- Accepted UX/product baseline: `main@ff26e48a904d4dae53c9d0f92ecf853f86ea93a9` from PR #25 (UX-8).
- PR #25 qualification: full automated suite 451 passed, 0 failed; CI passed on Windows / Python 3.10 and 3.12.
- Repository housekeeping baseline before v1.0.0 preparation: `main@6ac5914574a66f881b3cad0671caedd5014d94da` after PR #26 preserved the legacy version and refreshed README legacy links.
- GitHub Actions CI #106 passed on Windows / Python 3.10 and 3.12 for that housekeeping baseline.
- Current published release: `v1.1.0` is Stable/Latest at `main@1514b142e566f1edf64ac1ba5aaf6b5a1523541b`.
- Streamlit Community Cloud is deployed at https://tmc-process.streamlit.app/; the v1.1.0 Cloud operator acceptance passed.

## Product workflow
Canonical operator flow for Single and Batch:

`Data -> Mapping -> Analyze -> Review -> Export`

Core behavior:
- Data owns source and project/survey setup.
- Mapping owns movement semantics and Mapping Preset reuse.
- Analyze owns Peak search/PCE processing and analysis execution.
- Review owns QC and explicit Peak confirmation.
- Export owns Standard/Safe PNG report generation and package/ZIP output.

## Technology stack
- Language: Python
- UI: Streamlit
- Core libraries: pandas, openpyxl, Altair, Matplotlib
- Optional legacy/diagnostic integration: Microsoft Excel COM via pywin32
- Package manager: pip / setuptools via `pyproject.toml`
- Supported runtime: Windows, Python `>=3.10`
- GitHub Actions qualification: Windows / Python 3.10 and 3.12

## Standard commands
### Install/bootstrap
```text
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```

### Focused validation
```text
python -m pytest <relevant test files> -q
```

### Full validation
```text
python -m pytest
python -m compileall -q app.py src
git diff --check
```

### Local run
```text
start_tmc_processor.bat
```

Equivalent developer command:

```text
python -m streamlit run app.py
```

## Architecture / invariants
- `app.py` is the Streamlit entrypoint/compatibility facade.
- UI composition is under `src/tmc_processor/ui/`.
- Application boundaries are under `src/tmc_processor/application/`.
- Domain/calculation/export code remains under `src/tmc_processor/`.
- `src/tmc_processor/workflow_state.py` is the authority for workflow revisions, readiness and invalidation.
- UI/application code adapts to WorkflowState; it must not create a second workflow state machine.
- Suggested Peak, draft Peak and explicitly confirmed Peak are distinct states.
- Confirmed/effective Peak is authoritative for Peak-dependent export artifacts.
- Supported operator Peak periods are fixed whole hours, composed of four 15-minute intervals inside AM/PM search windows. Legacy rolling selections require re-analysis and explicit confirmation.
- Standard report prefers package-preserving OOXML native-template output when the exact confirmed/effective Peak is representable by the current template.
- Safe PNG is an explicit alternative/fallback and must preserve the exact confirmed/effective Peak.
- Excel COM is optional legacy/diagnostic capability, not a requirement for Standard report.

## Protected behavior
Changes must not alter the following without explicit approval and qualification:
- source parsing and TMC calculation methodology
- PCU/PCE processing
- Peak detection/calculation semantics
- suggested/draft/confirmed/effective Peak semantics
- QC rules
- movement aggregation and movement-code derivation
- Mapping domain semantics
- WorkflowState/readiness/invalidation rules
- Mapping Excel and Mapping Preset contracts
- Project Session schema and supported Single round-trip behavior
- Batch processing and per-file Review/Export semantics
- Standard OOXML / Safe PNG routing and effective-Peak consistency
- native Excel template fidelity
- protected real-workbook behavior under local `samples/raw/`

## Important paths
- Source: `src/tmc_processor/`
- UI: `src/tmc_processor/ui/`
- Application: `src/tmc_processor/application/`
- Entry point: `app.py`
- Tests: `tests/`
- Documentation: `docs/`
- Synthetic demos: `samples/demo/`
- Generated output: `outputs/`
- Real local samples: `C:\MyRD\tmc-processor-public\samples\raw\` (not part of public Git history)

## Validation matrix
| Gate | Command / Method | Required |
|---|---|---|
| Focused tests | `python -m pytest <relevant tests> -q` | Yes for affected areas |
| Full regression | `python -m pytest` | Yes |
| Compile | `python -m compileall -q app.py src` | Yes for implementation/runtime changes |
| Whitespace | `git diff --check` | Yes |
| Browser/UI | Manual populated Single/Batch workflow smoke | When UI/workflow presentation changes materially |
| Real-data/reference | Protected local workbooks in `samples/raw/` | Required when changes affect relevant workflow/mapping/Peak/export/state behavior |
| Native template/Excel | Package/formula/cache/fidelity checks and Excel/manual checks where applicable | Required for affected export changes |
| CI | GitHub Actions Windows Python 3.10 / 3.12 | Yes |

## Release history
### v1.1.0
Published Stable/Latest on 2026-10-12 from accepted `main@1514b142e566f1edf64ac1ba5aaf6b5a1523541b`. [GitHub Release](https://github.com/bokoboss/tmc-processor/releases/tag/v1.1.0). Windows and Demo ZIP assets were published. All 3 CI jobs and Streamlit Cloud acceptance passed.

Release highlights: Native V2 Summary, the supplied Support QR in the Windows bundle, and round-once PCU totals in the V2 Summary. The synthetic demo V2 workbook has 24 worksheets and two charts. V1 Summary PCU total discrepancy remains tracked by open Issue #42.

### v1.0.1
Historical published stable maintenance release from `main@39c1ba906dcc284b714e5980b9ef9e23dfe47827`.

### v1.0.0
Stable closure release after UX-0 through UX-8 modernization and qualification.

Key capabilities:
- canonical five-stage Single and Batch workflow
- sheet-centric physical Mapping
- explicit Peak Review confirmation
- exception-first Batch review
- application/UI architecture split
- package-preserving OOXML native-template Standard export
- Safe PNG fallback without Peak coercion
- operator UI cleanup, stage guidance and global Single Project Session controls

### v0.2.0 Public Beta
Original public release from 24 May 2026.

Preserved as:
- tag/release `v0.2.0`
- branch `legacy/v0.2.0`

The legacy snapshot is historical and must not be moved or rewritten.

## Known limitations / backlog
- Issue #23 was closed as `not_planned` after the fixed-hour Peak contract was accepted. Non-hour-aligned rolling Peaks are outside the supported contract.
- Issue [#42](https://github.com/bokoboss/tmc-processor/issues/42) remains open: on the synthetic demo, V1 Summary F32 displays 23,956 PCU while supporting total cells W40/AM22 display 23,959. V2 Summary uses its separately accepted round-once totals. V1 formulas and templates remain unchanged pending separate investigation and approval.
- The Excel GUI opening rejection reported during Issue #40 investigation was not reproduced in the later current-M365 acceptance; the V2 Cloud workbook was accepted. The original cause was not established, and universal Excel compatibility is not claimed.
- Batch Project Session is intentionally disabled because the current session format cannot round-trip Batch uploads and Batch Mapping Presets.
- UX-8 live Batch qualification used successful demo items; failed/excluded outcomes were covered by existing automated behavior but not recreated in that final UX browser smoke.
- `WorkflowContext` / `WorkflowOperations` retain a relatively broad dependency surface from UX-6; avoid expanding it without architectural review.

## Git / maintenance policy
- `main` is the maintained current line.
- `legacy/v0.2.0` is the preserved historical Public Beta snapshot.
- Use short-lived task branches for future fixes/enhancements.
- Merged task branches should normally be deleted after remote acceptance.
- Preserve unrelated local dirty changes.
- Do not rewrite release tags/history.
- Implementation changes require evidence appropriate to affected risk areas before merge/release.

## Current next objective
Issue #36 v1.1.0 publication is complete. The next project priority is the separately scoped investigation in open Issue #42; preserve the current V1/V2 report contracts until an explicit change is accepted and qualified.
