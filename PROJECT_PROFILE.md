# Project Profile

## Identity
- Project name: TMC Processor
- Repository URL: https://github.com/bokoboss/tmc-processor
- Authoritative local path: `C:\MyRD\tmc-processor-public`
- Primary branch: `main`
- Package/application version: `1.0.0` (from `pyproject.toml` after release preparation)
- Legacy public beta: `v0.2.0`, preserved at branch `legacy/v0.2.0`
- Installed workflow source: https://github.com/bokoboss/engineering-development-workflow
- Installed workflow revision: `3547ae260feacf8fc9a102b2abfdb13881e36dab`
- Installed workflow version: `v1.4.1`

## Current accepted baseline and release state
- Accepted UX/product baseline: `main@ff26e48a904d4dae53c9d0f92ecf853f86ea93a9` from PR #25 (UX-8).
- PR #25 qualification: full automated suite 451 passed, 0 failed; CI passed on Windows / Python 3.10 and 3.12.
- Repository housekeeping baseline before v1.0.0 preparation: `main@6ac5914574a66f881b3cad0671caedd5014d94da` after PR #26 preserved the legacy version and refreshed README legacy links.
- GitHub Actions CI #106 passed on Windows / Python 3.10 and 3.12 for that housekeeping baseline.
- Current phase: `v1.0.0` release preparation and project closure.
- After the `v1.0.0` tag is created, that tag is the authoritative stable release baseline.
- Post-release operating mode: maintenance mode.

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
- Issue #23: native-template support for non-hour-aligned rolling 60-minute Peaks (for example `08:15–09:15`). Until implemented, Standard falls back to Safe PNG while preserving the exact confirmed Peak.
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
No active modernization milestone is scheduled after `v1.0.0`.

Operate in maintenance mode:
- bug fixes through focused issues/PRs
- optional enhancement #23 when prioritized
- preserve validated behavior unless a separately approved change requires modification
