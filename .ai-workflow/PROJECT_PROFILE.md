# Project Profile

## Identity
- Project name: TMC Processor
- Repository URL: https://github.com/bokoboss/tmc-processor
- Authoritative local path: `C:\MyRD\tmc-processor-public`
- Primary branch: `main`
- Package / application version: `0.2.0` public beta

## Current accepted baseline
- Accepted branch: `main`
- Accepted HEAD SHA: `a1cd6b49bc1b7cc8bfbc0591590d62965a6ab108`
- Accepted date: 2026-08-24
- Current phase / milestone: UX/UAT modernization — UX-0 state/invalidation architecture is the next implementation phase
- Last accepted PR / CI run: PR #2 — “Document TMC Processor UX/UAT modernization architecture”; CI #54 / run `32710586892` passed on Windows with Python 3.10 and 3.12

## Technology stack
- Language(s): Python
- Framework(s): Streamlit; pandas; openpyxl; Altair; Matplotlib; Microsoft Excel COM for native-template export when available
- Package manager: pip / setuptools via `pyproject.toml`
- Supported OS / runtime: Windows is the authoritative runtime; Python `>=3.10`; CI currently qualifies Python 3.10 and 3.12 on `windows-latest`; Microsoft Excel desktop is required only for Excel Template / COM export

## Standard commands
### Install / bootstrap
```text
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```

For the normal Windows user workflow, `start_tmc_processor.bat` creates/uses the local virtual environment and installs required packages as needed.

### Fast validation
```text
python -m pytest <relevant test files> -q
```

Use targeted pytest for the area changed during development. There is currently no separate repository-defined fast-validation script.

### Full validation
```text
python -m pytest
git diff --check
```

### Build / package
```text
No separate release-build/package command is currently defined.
The project is installed from source with:
python -m pip install -e ".[dev]"
```

### Local run
```text
start_tmc_processor.bat
```

Equivalent developer launch when the environment is already prepared:
```text
python -m streamlit run app.py
```

## Architecture / invariants
- The validated TMC calculation/domain engine lives under `src/tmc_processor`; Streamlit is the operator UI/application shell and must not become the source of calculation methodology.
- The authoritative operator workflow direction is `Data → Mapping → Analyze → Review → Export` for both Single and Batch modes; implementation is staged through GitHub issues UX-0 to UX-7.
- Raw workbook sheet/source identity is the primary mapping anchor. Basic mapping should let the operator assign physical traffic semantics and derive canonical movement fields automatically where deterministic, without requiring Advanced mode.
- Workflow state must follow explicit dependency/invalidation contracts: Source → Mapping → Analysis configuration → Analysis result → Review/effective Peak → Export configuration → Export artifact.
- Effective Peak must propagate consistently to every exported artifact that reports Peak-dependent results.
- Windows + real-workbook + Excel export qualification remains an acceptance gate for release-level changes; automated unit tests alone are not sufficient for final release acceptance.

## Protected behavior
Changes must not alter the following unless explicitly approved:
- Validated calculation methodology, PCU/PCE processing, Peak calculations, QC logic, movement aggregation, movement-code derivation, and V1/V2 movement-scheme behavior.
- Existing Mapping Excel, Mapping Preset, Project Session, Single processing, Batch processing, effective-Peak propagation, Excel Template export, Safe PNG export, export metadata, and movement-diagram contracts.
- Previously validated behavior that changing a confirmed Peak clears stale export output while preserving analysis when recalculation is not required.
- Real sample compatibility for Kabin Buri, Bo Phloi, and Nong Prue workbooks.
- Existing engineering outputs must not be changed as a side effect of UX modernization unless a separate defect is demonstrated and explicitly approved.

## Important paths
- Source: `src/tmc_processor/`
- Application shell / Streamlit UI: `app.py`
- Tests: `tests/`
- Documentation: `docs/`
- Generated output: `outputs/`
- Demo / synthetic samples: `samples/demo/`
- Real local samples: `C:\MyRD\tmc-processor-public\samples\raw`
- Local-only / sensitive data: real survey workbooks and project/customer data under local sample/work directories; do not add real or sensitive raw data to public Git history unless explicitly approved

## Validation matrix
| Gate | Command / Method | Required |
|---|---|---|
| Unit tests | `python -m pytest` | Yes |
| Integration tests | Covered within pytest suite for pipeline, mapping, batch, export, sessions, and effective-Peak behavior | Yes |
| Browser/UI tests | No dedicated browser automation gate currently defined; Streamlit workflow/UAT remains manual plus helper/state tests | No for every PR; Yes where UI behavior changes materially |
| Build | No separate release-build gate currently defined | No |
| Package / wheel | No wheel qualification gate currently defined | No |
| Runtime smoke | Launch Streamlit on Windows and complete the affected operator workflow | Yes for workflow/UI changes |
| Real-data validation | Kabin Buri, Bo Phloi, Nong Prue end-to-end; verify required Excel/export artifacts | Yes for release qualification / changes affecting workflow, mapping, Peak, export, or state |
| CI | GitHub Actions on Windows / Python 3.10 and 3.12 | Yes |

## Git / release policy
- Branch naming: use isolated task branches, preferably `codex/<phase-or-issue>-<short-description>`; do not implement directly on `main`
- Commit policy: focused commits scoped to the approved issue; preserve unrelated dirty user changes; no destructive history rewriting unless explicitly authorized
- PR policy: implementation changes require a PR against `main` with issue reference, starting baseline SHA, changed files, validation evidence, known limitations, and explicit scope/non-scope notes
- Merge policy: do not merge merely because implementation is “complete” or tests passed; require independent review of diff, regression evidence, CI, and applicable runtime/real-data gates; prefer squash merge after acceptance
- Release policy: release qualification must include applicable automated tests, Windows runtime, real-workbook validation, Excel COM/native-template validation when relevant, Safe PNG fallback, effective-Peak consistency, `git diff --check`, and documented residual limitations

## Current known limitations / risks
- `app.py` is a large Streamlit orchestration surface and currently owns substantial workflow/session-state behavior; stale-result and hidden dependency regressions remain a key risk until UX-0/UX-6 are completed.
- Single-file Peak selection currently blurs system-suggested values and explicit human confirmation; explicit Peak Review provenance is intentionally deferred to UX-3.
- Batch review is still primarily file-by-file and is not yet exception-first/bulk-review optimized.
- Batch `approach_movement` currently supports Safe PNG/generated ZIP; Batch Excel Template Mode is not yet supported for that path.
- Excel Template Mode depends on Windows Microsoft Excel / Excel COM; Safe PNG Export Mode is the validated fallback.
- No dedicated browser automation, wheel/package qualification, or separate release-build gate is currently defined.
- UX modernization must not be combined with a frontend-framework migration or MCP implementation until the application/state contracts are stabilized.

## Current next objective
- Implement GitHub Issue #3 / UX-0: centralize workflow state and invalidation contracts in a small pure-Python application/state layer, add deterministic state-transition tests, route the highest-risk stale-state paths through the new contract, preserve current visible behavior and validated engineering outputs, and open a PR for independent review before starting UX-1.
