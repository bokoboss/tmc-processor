"""Issue #38: module facade isolation and Streamlit execution regressions."""

from __future__ import annotations

import sys
import types
from pathlib import Path

import pytest

from tmc_processor.ui import app_shell


ENTRYPOINT = Path(__file__).resolve().parents[1] / "app.py"


@pytest.fixture
def imported_app(monkeypatch):
    """Exercise a fresh import namespace and restore even a contaminated shell."""
    shell_class = type(app_shell)
    module = types.ModuleType("app")
    module.__file__ = str(ENTRYPOINT)
    monkeypatch.setitem(sys.modules, "app", module)
    try:
        yield module
    finally:
        # Bypass a broken facade setter when the RED test changes shell's class.
        types.ModuleType.__setattr__(app_shell, "__class__", shell_class)


def test_repeated_import_execution_does_not_recurse_or_change_shell_class(imported_app):
    shell_class = type(app_shell)
    code = compile(ENTRYPOINT.read_text(encoding="utf-8"), str(ENTRYPOINT), "exec")
    for _ in range(3):
        exec(code, imported_app.__dict__)
    assert type(app_shell) is shell_class
    assert imported_app._run_streamlit_app is app_shell._run_streamlit_app


@pytest.mark.parametrize("name", ["__name__", "__file__", "__package__", "__loader__", "__spec__", "__doc__"])
def test_import_metadata_assignments_stay_local(imported_app, name):
    exec(compile(ENTRYPOINT.read_text(encoding="utf-8"), str(ENTRYPOINT), "exec"), imported_app.__dict__)
    original = getattr(app_shell, name)
    replacement = object()
    try:
        setattr(imported_app, name, replacement)
        assert getattr(imported_app, name) is replacement
        assert getattr(app_shell, name) is original
    finally:
        types.ModuleType.__setattr__(app_shell, name, original)


def test_import_class_assignment_stays_local(imported_app):
    exec(compile(ENTRYPOINT.read_text(encoding="utf-8"), str(ENTRYPOINT), "exec"), imported_app.__dict__)
    shell_class = type(app_shell)

    class ImportOnlyModule(types.ModuleType):
        pass

    imported_app.__class__ = ImportOnlyModule
    assert type(imported_app) is ImportOnlyModule
    assert type(app_shell) is shell_class


def test_legacy_helper_lookup_and_monkeypatch_restore(imported_app, monkeypatch):
    exec(compile(ENTRYPOINT.read_text(encoding="utf-8"), str(ENTRYPOINT), "exec"), imported_app.__dict__)
    original = app_shell._run_streamlit_app
    replacement = lambda: None
    assert imported_app._run_streamlit_app is original
    assert "_run_streamlit_app" in dir(imported_app)
    with monkeypatch.context() as patch:
        patch.setattr(imported_app, "_run_streamlit_app", replacement)
        assert app_shell._run_streamlit_app is replacement
        assert imported_app._run_streamlit_app is replacement
    assert imported_app._run_streamlit_app is original
    assert app_shell._run_streamlit_app is original


def test_main_executes_run_then_stop_three_times_without_installing_facade(monkeypatch):
    module = types.ModuleType("__main__")
    monkeypatch.setitem(sys.modules, "__main__", module)
    calls = []
    shell_class = type(app_shell)
    monkeypatch.setattr(app_shell, "_run_streamlit_app", lambda: calls.append("run"))
    monkeypatch.setattr(app_shell, "st", types.SimpleNamespace(stop=lambda: calls.append("stop")))
    code = compile(ENTRYPOINT.read_text(encoding="utf-8"), str(ENTRYPOINT), "exec")
    try:
        for _ in range(3):
            exec(code, module.__dict__)
            assert type(module) is types.ModuleType
            assert type(app_shell) is shell_class
        assert calls == ["run", "stop"] * 3
    finally:
        types.ModuleType.__setattr__(app_shell, "__class__", shell_class)


def test_apptest_fresh_loads_and_reruns_preserve_module_classes(monkeypatch):
    from streamlit.testing.v1 import AppTest
    from tmc_processor.excel_com_export import ExcelComStatus

    # AppTest replaces __main__ in the process; restore it after this test.
    monkeypatch.setitem(sys.modules, "__main__", sys.modules["__main__"])
    monkeypatch.setattr(
        app_shell,
        "_probe_excel_com_for_ui",
        lambda force=False: ExcelComStatus(available=False, reason="Cloud runtime test", detail="", version=""),
    )
    shell_class = type(app_shell)
    try:
        for _ in range(3):
            at = AppTest.from_file(str(ENTRYPOINT), default_timeout=60)
            for _ in range(3):
                at.run()
                assert not at.exception, [error.message for error in at.exception]
                assert at.file_uploader
                assert type(sys.modules["__main__"]) is types.ModuleType
                assert type(app_shell) is shell_class
    finally:
        types.ModuleType.__setattr__(app_shell, "__class__", shell_class)
