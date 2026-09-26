"""Executable Streamlit entrypoint for TMC Processor.

The implementation lives under :mod:`tmc_processor.ui`; this module remains a
small compatibility facade for established tests and internal imports.
"""

from tmc_processor.ui import app_shell as _app_shell
import sys
import types


def __getattr__(name: str):
    """Preserve legacy ``app.<helper>`` access without duplicating logic."""

    return getattr(_app_shell, name)


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(dir(_app_shell)))


class _FacadeModule(types.ModuleType):
    """Keep legacy monkeypatch seams pointed at the shell implementation."""

    def __setattr__(self, name: str, value: object) -> None:
        super().__setattr__(name, value)
        if name != "_app_shell" and hasattr(_app_shell, name):
            setattr(_app_shell, name, value)


sys.modules[__name__].__class__ = _FacadeModule


if __name__ == "__main__":
    _app_shell._run_streamlit_app()
    _app_shell.st.stop()
