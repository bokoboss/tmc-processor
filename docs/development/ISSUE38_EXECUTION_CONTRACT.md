# Issue #38 runtime hotfix execution contract

Date: 2026-10-11 (Asia/Bangkok).

## Objective and baseline

Fix compatibility-facade recursion during repeated module execution, preserving
legacy `import app` monkeypatches and the Streamlit entrypoint.

- Repository: `C:\MyRD\tmc-processor-public`, `bokoboss/tmc-processor`.
- Issue: https://github.com/bokoboss/tmc-processor/issues/38.
- Verified clean local HEAD and fetched `origin/main`:
  `b4360c2dc5732fba75d1e9fd691a8ef2a653528f`.
- Branch: `codex/issue-38-runtime-hotfix`.

## Execution route

Use installed engineering-development-workflow v1.4.1, revision
`3547ae260feacf8fc9a102b2abfdb13881e36dab`, including its model-routing and
systematic-debug policies. Upstream now reports v1.7.4; upgrading governance is
outside this hotfix. Luna High is sufficient for bounded implementation after
the module-lifecycle reproducer is established. Continue in the existing
user-selected chat with its configured model and one worker, as in the project's
Community Cloud execution contract; no additional model or agent is required.
Escalate only for a capability limitation after diagnosing environment and
reproducer problems.

## Scope and invariants

Change `app.py`, runtime regression tests, Linux rerun CI qualification, and
instantiated evidence documents only. No new dependencies or feature changes.
Compatibility facade installation must be restricted to ordinary `import app`.
Module metadata (especially `__class__`) must stay local. Existing helper
delegation and monkeypatch restoration must work. Streamlit `__main__` must still
call `_run_streamlit_app()` followed by `st.stop()`.

Do not change Calculation, Mapping, PCE, QC, Peak, WorkflowState, Export,
templates, drawings, charts, or Support QR. Preserve local real workbooks and
user changes. Do not merge, tag, publish, close issues, build release ZIPs, or
reuse the ZIPs from `b4360c2`.

## Success gates (defined before implementation)

| Gate | Required proof |
| --- | --- |
| Baseline | Clean worktree; local and remote main match accepted SHA; issue inspected |
| RED | Regression fails against unmodified app.py on repeated module class assignment; record exact output |
| Facade | At least three executions of imported app; shell class/metadata unchanged; delegated helpers and patch/restore work |
| Entrypoint | At least three executions of __main__; original module class retained; run then stop order retained |
| Streamlit | AppTest.from_file actual runs and reruns at least three times; no exceptions; shell and __main__ class isolation |
| Regression | Focused tests and full pytest pass |
| Hygiene | compileall app.py/src and git diff --check pass; actual diff limited to scope |
| CI | GitHub Actions Linux Python 3.12 and Windows Python 3.10/3.12 pass on hotfix head |
| Delivery | Hotfix PR and evidence package with exact results, changed files, limitations, and review handoff |

Tests exercise the module import/assignment, script entrypoint, and AppTest seams
explicitly specified by the user and Issue #38. Restore process-global module
state even when a regression fails.

## Acceptance boundary

Stop for scope review if the fix requires protected logic changes. Human hotfix
review precedes merge. After separately approved merge, production Cloud reboot,
reload/rerun, full demo workflow, version/banner, V2 Summary/totals, Support QR,
and workflow-state smoke remain mandatory before release. Rebuild release ZIPs
from the latest accepted hotfix merge commit only.
