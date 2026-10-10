# Project Profile

## Identity
- Project name: TMC Processor
- Repository URL: https://github.com/bokoboss/tmc-processor
- Authoritative local path: `C:\MyRD\tmc-processor-public`
- Primary branch: `main`
- Package/application version: `1.1.0` (release-preparation branch; `v1.0.1` remains the latest published stable release until approval and publication)
- Legacy public beta: `v0.2.0`, preserved at branch `legacy/v0.2.0`
- Installed workflow source: https://github.com/bokoboss/engineering-development-workflow
- Installed workflow revision: `3547ae260feacf8fc9a102b2abfdb13881e36dab`
- Installed workflow version: `v1.4.1`

## Current accepted baseline and release state
- Accepted released baseline before Issue #30: `main@39c1ba906dcc284b714e5980b9ef9e23dfe47827` (`v1.0.1`).
- v1.0.1 qualification: full automated suite 504 passed; GitHub Actions CI #113 passed on Windows / Python 3.10 and 3.12.
- Before the separately approved Issue #30 web-app enhancement, the project was in maintenance mode.
- Accepted Issue #36 release baseline: `main@203bc5f3f584d0245f46c27d58da74c3d973fd65` (PRs #31, #33, and #35 merged; v1.0.1 remains the latest published release).
- Current task: Issue #36 v1.1.0 release preparation, with calculation, mapping, Peak, export, template, drawing, and chart behavior protected.
- Accepted UX/product baseline: `main@ff26e48a904d4dae53c9d0f92ecf853f86ea93a9` from PR #25 (UX-8).
- PR #25 qualification: full automated suite 451 passed, 0 failed; CI passed on Windows / Python 3.10 and 3.12.
- Repository housekeeping baseline before v1.0.0 preparation: `main@6ac5914574a66f881b3cad0671caedd5014d94da` after PR #26 preserved the legacy version and refreshed README legacy links.
- GitHub Actions CI #106 passed on Windows / Python 3.10 and 3.12 for that housekeeping baseline.
- Current published release: `v1.0.1` is stable; `v1.1.0` is not tagged or published during preparation.
- Streamlit Community Cloud is deployed at https://tmc-process.streamlit.app/ from `main`; the accepted v1.1.0 branch must pass review and CI before deployment changes.

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
### v1.0.1
Published stable maintenance release from `main@39c1ba906dcc284b714e5980b9ef9e23dfe47827`.

### v1.1.0 (preparation)
Release preparation is based on accepted `main@203bc5f3f584d0245f46c27d58da74c3d973fd65`. It aligns package/runtime version, packages the supplied Support QR in the Windows bundle, and records the hosted readiness, V2 Summary and round-once PCU work. It remains unpublished until PR review, CI, merge, and separate release approval.

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
Complete Issue #36 release preparation from the accepted `203bc5f3f584d0245f46c27d58da74c3d973fd65` baseline without changing validated engineering or export behavior. Merge, tag, and publish remain separate approval gates.
