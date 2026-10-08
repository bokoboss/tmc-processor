# Community Cloud preparation evidence — Issue #30 / PR #31

Date: 2026-10-08 (Asia/Bangkok). Scope and gates:
[execution contract](COMMUNITY_CLOUD_EXECUTION.md).

## Provenance and changes

Repository: `C:\MyRD\tmc-processor-public` / `bokoboss/tmc-processor`.
Branch: `codex/issue-30-web-readiness-support`.
Baseline local/remote HEAD: `46901e719a9b4938b568ad9de3c393dba2ecdc76`;
initial working tree clean. Final commit identifier and current CI results are
reported in [PR #31](https://github.com/bokoboss/tmc-processor/pull/31).

Changed files in this increment:
- `src/tmc_processor/ui/components/support.py`: final Issue #30 copy.
- `src/tmc_processor/ui/app_shell.py`: Privacy moved from header to collapsed
  Data upload disclosure shared by Single and Batch.
- `tests/test_support_view.py`: independently recorded approved text and
  runtime Data disclosure coverage.
- `requirements.txt`: `-e .`; pyproject remains authoritative.
- `.github/workflows/ci.yml`: Linux Python 3.12 uv + pip deployment install,
  dependency consistency/no COM checks, app.py execution and startup smoke.
- `README.md`: brief deployment/remaining qualification instructions;
  correct stale published-version wording.
- This evidence file and `COMMUNITY_CLOUD_EXECUTION.md`.

No core processing, Mapping, Peak, QC, WorkflowState, export, theme, QR,
package-version or Windows launcher changes in this increment.

## Local validation

| Command / check | Result |
| --- | --- |
| `.venv/Scripts/python.exe -m pytest tests/test_support_view.py -q --basetemp=.pytest_cache/support-baseline` before edits | 7 passed |
| Updated tests against original code, before production edits | Header test reported failure; sandbox AppTest did not complete normally |
| AST comparison of baseline SUPPORT_BODY against independently recorded final Issue copy | Expected AssertionError (RED); obsolete wording reproduced |
| `.venv/Scripts/python.exe -m pytest tests/test_support_view.py -q --basetemp=.pytest_cache/support-elevated` | 8 passed in 2.08s |
| `.venv/Scripts/python.exe -m compileall -q app.py src` | Passed |
| `git diff --check` | Passed |
| QR SHA-256 before and after edits | Identical: `07cb6abdc58680ecf333ec4295f9c5ec9a1dd66b68268647ebbc267b02d52e31` |

Sandbox temp/runtime limitations caused initial test errors/hangs; use of a
repository-local pytest basetemp and an approved unsandboxed focused run resolved
execution. No application changes were made to compensate for the sandbox.

## Linux / CI evidence

Local host is Windows. The updated PR's Ubuntu/Python 3.12 job is the clean Linux
qualification environment. It installs the exact root requirements with uv in a
fresh venv and separately with pip, checks dependencies/package import and absence
of win32com, executes app.py with Streamlit AppTest, runs the existing regression
suite and starts `streamlit run app.py` through the headless health helper.
Existing Windows Python 3.10/3.12 jobs are preserved. CI results are pending at
commit creation; the final report/PR description must supply the completed run URL
and results before this preparation can be accepted. No Linux PASS is inferred
from local Windows results.

## Limitations / remaining review

No local full-suite, real-workbook or COM integration rerun: user requested
risk-based testing and protected paths are untouched. No physical final-dialog
QR scan or desktop/mobile browser qualification performed in this increment;
Issue #30 records the user's source-asset scan PASS. QR bytes/render size/layout
are unchanged. AppTest checks runtime elements, not visual browser fidelity.
Dependencies retain existing lower bounds and are not reproducibly locked.

No production deployment was performed. After approved merge, deploy main/app.py
with Python 3.12 through Community Cloud sign-in/GitHub authorization, then run
the non-sensitive demo upload-to-export and Support smoke described in README.
Public deployment/resource qualification and final evidence review remain
required before tagging/releasing v1.1.0 or closing Issue #30. Version remains
1.0.1 until the approved release step.
