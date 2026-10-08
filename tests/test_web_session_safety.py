from __future__ import annotations

from threading import Thread

from tmc_processor.excel_com_export import get_last_export_diagnostics


def test_excel_com_diagnostics_are_isolated_between_execution_threads() -> None:
    current = get_last_export_diagnostics()
    current.formula_cells_preserved[:] = ["session-a-only"]
    observed: list[list[str]] = []

    try:
        worker = Thread(
            target=lambda: observed.append(
                list(get_last_export_diagnostics().formula_cells_preserved)
            )
        )
        worker.start()
        worker.join(timeout=5)

        assert not worker.is_alive()
        assert observed == [[]]
    finally:
        current.formula_cells_preserved.clear()
