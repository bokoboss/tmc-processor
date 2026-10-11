# Issue #38 — production runtime hotfix evidence

## Provenance and scope

- Date: 2026-10-11 (Asia/Bangkok).
- Repository: `bokoboss/tmc-processor`, `C:\MyRD\tmc-processor-public`.
- Issue: https://github.com/bokoboss/tmc-processor/issues/38.
- Related release task: https://github.com/bokoboss/tmc-processor/issues/36.
- Authoritative baseline: `main@b4360c2dc5732fba75d1e9fd691a8ef2a653528f`.
- Local HEAD and fetched `origin/main` matched that SHA; initial worktree clean.
- Branch: `codex/issue-38-runtime-hotfix`.
- Execution contract: [ISSUE38_EXECUTION_CONTRACT.md](ISSUE38_EXECUTION_CONTRACT.md).
- PR/head/CI identifiers: recorded in the hotfix PR's validation section after
  creation; evidence document content is versioned with that PR.

Scope is module-facade assignment, entrypoint regression tests, Linux runtime CI
reruns, and evidence. No Calculation, Mapping, PCE, QC, Peak, WorkflowState,
Export, Excel template, drawing, chart, Support QR, or package version changes.
No local real workbook or release ZIP was added to Git.

## Execution route and governance

Reviewed the installed workflow v1.4.1 model-routing and systematic-debug skills
at pinned revision `3547ae260feacf8fc9a102b2abfdb13881e36dab`. Current upstream
reports v1.7.4; the hotfix preserves the installed project governance. Luna High
is sufficient for this bounded execution; continued with the configured model
in the existing user-selected chat, one worker, no model escalation. The explicit
user requirements establish module import/assignment, __main__, and AppTest as
the test boundaries. No legacy `.ai-workflow/` directory was recreated.

## Root cause and RED proof

Before modifying `app.py`, added
`test_repeated_import_execution_does_not_recurse_or_change_shell_class`, which
executes the real entrypoint source three times in one registered `app` module.
The test restores both the original `sys.modules['app']` entry and shell class
even when the baseline recurses.

Command (using repository `.venv` Python 3.12.10):

```text
python -m pytest tests/test_app_runtime.py -q --tb=short
```

Recorded baseline result: **1 failed**, exit 1, in 5.84 seconds:

```text
app.py:31: in <module>
    sys.modules[__name__].__class__ = _FacadeModule
app.py:28: in __setattr__
    setattr(_app_shell, name, value)
app.py:28: in __setattr__
    setattr(_app_shell, name, value)
E   RecursionError: maximum recursion depth exceeded
```

Execution 1 installs the facade on the entry module. Execution 2 invokes the
existing facade setter while assigning its replacement class; because
`app_shell` has `__class__`, the setter also installs that facade on `app_shell`.
Execution 3 invokes the shell's facade setter, which forwards the assignment
back to the same shell indefinitely. This directly reproduces the production
trace's failing assignment and recursive call using actual baseline code.

The installed Streamlit 1.64.0 runner creates/registers a `__main__` ModuleType
and executes script code in its dictionary. Exact production lifecycle timing
that first contaminated the Cloud shell is not independently reconstructed.
The repeated-module defect and its corresponding production traceback are
proven; the fix removes both the entrypoint mutation and propagation paths.

## Correction

- Install `_FacadeModule` only when `__name__ == 'app'`.
- Keep module dunder metadata, including `__class__`, local in the setter.
- Retain `_app_shell` assignment isolation and ordinary helper forwarding.
- Preserve delegated lookup/dir and legacy monkeypatch/restore behavior.
- Leave `__main__` execution's `_run_streamlit_app()` then `st.stop()` unchanged.
- Strengthen the Linux no-COM CI script from one execution to three, checking
  exceptions, actual rendered upload widgets, and both module classes each run.

## Validation evidence

All local Python commands use `./.venv/Scripts/python.exe` (Python 3.12.10,
Streamlit 1.64.0, Windows).

| Gate | Command / method | Result |
| --- | --- | --- |
| RED | `python -m pytest tests/test_app_runtime.py -q --tb=short` before app.py edit | 1 failed with baseline RecursionError, exit 1 |
| First GREEN | Same single regression immediately after fix | 1 passed in 1.75s |
| Runtime regressions | `python -u -m pytest tests/test_app_runtime.py -q --tb=short` | 11 passed in 1.93s |
| Repeated import | Actual app.py source, 3 executions in one imported module | PASS; no recursion; shell class unchanged |
| Metadata isolation | __class__, __name__, __file__, __package__, __loader__, __spec__, __doc__ writes | PASS; assignments stay local |
| Legacy compatibility | Helper lookup, dir, monkeypatch and restoration | PASS; shell receives helper patches and original is restored |
| Entrypoint | Actual app.py as __main__, 3 executions with recorded calls | PASS; run then stop on every execution; main/shell classes unchanged |
| Actual Streamlit | AppTest.from_file, 3 fresh sessions x 3 executions | PASS; 9 executions, no exceptions, Data upload UI rendered, main/shell classes unchanged |
| Full regression | `python -u -m pytest --basetemp outputs/issue38/pytest-temp` | 551 passed, 30 OpenPyXL DrawingML warnings, exit 0, in 432.34s |
| Compile | `python -m compileall -q app.py src` | PASS |
| Whitespace | `git diff --check` | PASS |
| CI | Linux Python 3.12; Windows Python 3.10/3.12 | Head-specific final results and run/job links are recorded in the accompanying PR validation section |

Local raw logs are ignored artifacts under `outputs/issue38/`:
`red-repeated-import.log`, `green-repeated-import.log`, `focused-runtime.log`,
`full-pytest.log`, and `baseline-lifecycle.log`. The essential RED and GREEN results are preserved above
so reviewers do not need private workspace artifacts.

Initial sandbox AppTest attempts were interrupted. Diagnostic pytest with
`-o faulthandler_timeout=25` located the hang in Windows
`socket._fallback_socketpair -> asyncio.new_event_loop`, before app execution.
The same tests passed outside the sandbox in 1.93s. Full pytest also runs
outside the sandbox. This is an environment limitation, not a skipped test or
an application workaround. The new AppTest regression explicitly sets Excel
COM unavailable to model Cloud; the Linux CI entrypoint step runs with real
no-COM dependencies and no probe monkeypatch.

## Changed files and audit

- `app.py`: facade metadata guard and import-only installation.
- `tests/test_app_runtime.py`: 11 runtime/facade regression cases.
- `.github/workflows/ci.yml`: three actual Linux entrypoint executions.
- `docs/development/ISSUE38_EXECUTION_CONTRACT.md`: bounded execution and gates.
- `docs/development/EVIDENCE_PACKAGE_ISSUE38.md`: evidence and review handoff.

Reviewed actual diff against baseline for scope. Protected source, templates,
assets, sample workbooks, and release packaging files are unchanged. Existing
regressions remain required in the full suite and platform CI.

## Limitations and acceptance handoff

This PR prepares the hotfix for review. It does not claim production Cloud
acceptance. No production/browser smoke against the patched Cloud deployment
is possible until the separately approved merge/deployment. No real-workbook
or Excel GUI recalculation qualification was added for this facade-only fix.

After reviewer acceptance and separately authorized merge:

1. Verify latest main CI and reboot/reload/rerun Streamlit Cloud repeatedly.
2. Confirm v1.1.0 banner and full demo Data -> Mapping -> Analyze -> Review ->
   Export workflow, V2 Summary's 24 worksheets, round-once totals
   F30=2981, F31=3153, F32/W40/AM22=23959, Support QR, and state preservation.
3. Rebuild both release ZIPs from the latest accepted hotfix merge commit;
   verify QR bytes, templates, version, archive provenance and SHA.
4. Keep Issue #36 release blocked until those gates pass. Never publish or reuse
   the pre-hotfix ZIPs from `b4360c2`.

No merge, tag, release publication, production deployment change, or issue
closure is authorized/performed in this executor task.
