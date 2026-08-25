# Project Profile

## Identity
- Project name: TMC Processor
- Repository URL: https://github.com/bokoboss/tmc-processor
- Authoritative local path: `C:\MyRD\tmc-processor-public`
- Primary branch: `main`
- Package/application version: `0.2.0` (from `pyproject.toml`)
- Installed workflow source: https://github.com/bokoboss/engineering-development-workflow
- Installed workflow revision: `3547ae260feacf8fc9a102b2abfdb13881e36dab`
- Installed workflow version: `v1.4.1`

## Current accepted baseline
- Accepted branch: `main`
- Accepted HEAD SHA: `e840e09db25fce15ab2e66e938713165382e3ca8`
- Accepted date: 2026-08-25
- Current phase/milestone: UX-0 workflow/state contract merged; UX-1 is next.
- Last accepted PR / CI run: Not recorded in the local repository metadata.

## Technology stack
- Languages: Python
- Frameworks/libraries: Streamlit, pandas, openpyxl, Altair, Matplotlib; optional Microsoft Excel COM via pywin32 for native-template export
- Package manager: pip / setuptools via `pyproject.toml`
- Supported OS/runtime: Windows is the authoritative runtime; Python `>=3.10`. GitHub Actions currently qualifies Windows with Python 3.10 and 3.12.

## Standard commands
### Install/bootstrap
```text
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```
### Fast validation
```text
python -m pytest <relevant test files> -q
```
### Full validation
```text
python -m pytest
python -m compileall -q app.py src
git diff --check
```
### Build/package
```text
No separate release build or package command is defined.
```
### Local run
```text
start_tmc_processor.bat
```

When the environment is already prepared, the equivalent developer command is:

```text
python -m streamlit run app.py
```

## Architecture / invariants
- `app.py` is the Streamlit application shell; the validated calculation/domain engine is under `src/tmc_processor/`.
- `src/tmc_processor/workflow_state.py` is pure Python with no Streamlit dependency. `WorkflowState`, semantic revisions/fingerprints, readiness, and `transition_workflow` define the dependency chain and invalidation rules; `app.py` adapts these rules to legacy session-state keys.
- The intended operator flow for Single and Batch is `Data -> Mapping -> Analyze -> Review -> Export`.
- `from_to` remains the compatibility movement-code scheme; `approach_movement` is also supported with its documented export limitations.
- Effective Peak decisions must remain consistent in every Peak-dependent export artifact.

## Protected behavior
Changes must not alter the following unless explicitly approved:
- Calculation methodology, PCU/PCE processing, Peak calculations, QC logic, movement aggregation, movement-code derivation, and existing `from_to` / `approach_movement` behavior.
- Mapping Excel, Mapping Preset, Project Session, Single, Batch, effective-Peak propagation, export metadata, and movement-diagram contracts.
- Excel Template Mode behavior when Microsoft Excel/COM is available and Safe PNG/generated export fallback behavior when it is not.
- Real-workbook regression behavior for the local files in `samples/raw/`.

## Important paths
- Source: `src/tmc_processor/`
- Application shell: `app.py`
- Tests: `tests/`
- Documentation: `docs/`
- Generated output: `outputs/`
- Synthetic demos: `samples/demo/`
- Real local samples: `C:\MyRD\tmc-processor-public\samples\raw\` (not part of public Git history)

## Validation matrix
| Gate | Command / Method | Required |
|---|---|---|
| Unit / targeted | `python -m pytest <relevant test files> -q` | Yes for affected areas |
| Integration / regression | `python -m pytest` | Yes |
| Browser/UI | Manual Streamlit workflow/UAT; no dedicated browser automation gate is defined | When UI behavior changes materially |
| Build/package/runtime | `python -m compileall -q app.py src`; Windows Streamlit launch | Yes for implementation/runtime changes |
| Real-data/reference | Local Windows validation of the three workbooks in `samples/raw/`, including required Excel/export checks | Yes for release qualification and changes affecting workflow, mapping, Peak, export, or state |
| CI | GitHub Actions on `windows-latest`, Python 3.10 and 3.12 | Yes |

## Execution characteristics
- Typical task ambiguity: Preserve established engineering outputs while evolving the Streamlit workflow.
- High-risk areas: `app.py` session-state orchestration, mapping, Peak review, Excel export, and real-workbook compatibility.
- Modules safe to parallelize: Pure calculation and test-only changes when their contracts are independent.
- Modules tightly coupled / single-owner: Streamlit state transitions and export/readiness integration in `app.py`.
- Preferred local execution constraints: Windows; use the repository virtual environment where available; Excel COM qualification requires desktop Microsoft Excel and pywin32.

## Git / release policy
- Branch naming: isolated task branches; use `codex/<scope>-<description>` or the task-specified branch.
- Commit policy: focused commits scoped to the approved task; preserve unrelated changes; no destructive history rewriting without authorization.
- PR policy: implementation changes require a PR against `main` with baseline SHA, changed files, validation evidence, scope/non-scope, and limitations.
- Merge policy: do not merge automatically; require review of the diff and applicable regression, CI, runtime, and real-data evidence.
- Release policy: qualify applicable automated tests, Windows runtime, real-workbook behavior, Excel COM/native-template behavior, Safe PNG fallback, effective-Peak consistency, and `git diff --check`.

## Current known limitations / risks
- `app.py` remains a large Streamlit orchestration surface; stale-result regressions remain a risk while UX modernization continues.
- Peak suggestion versus explicit human confirmation is not yet fully separated; this is planned for UX-3.
- Batch review is still primarily file-by-file.
- Excel Template Mode depends on Windows Microsoft Excel/Excel COM; Safe PNG/generated export is the fallback. Batch `approach_movement` Excel Template Mode remains blocked.
- No dedicated browser automation, wheel/package qualification, or separate release-build gate is defined.
- The older project-owned `.ai-workflow/` directory overlaps conceptually with this workflow and is intentionally preserved for a separate follow-up.

## Current next objective
- UX-1: implement the canonical five-stage navigation shell, `Data -> Mapping -> Analyze -> Review -> Export`, for Single and Batch while preserving the UX-0 state/invalidation contract and validated engineering outputs.
