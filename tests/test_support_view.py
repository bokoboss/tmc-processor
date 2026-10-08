from __future__ import annotations

from contextlib import nullcontext
from pathlib import Path

from tmc_processor.ui.components import support
from tmc_processor.ui.app_shell import (
    APP_SUPPORT_VIEW,
    APP_VIEW_STATE_KEY,
    get_app_view,
    set_app_view,
)
from tmc_processor.ui.components.support import resolve_support_qr_path


def test_support_navigation_preserves_the_active_workflow_and_session_state() -> None:
    state: dict[str, object] = {
        "active_workflow_tab": "Review",
        "work_mode": "ประมวลผลหลายไฟล์",
        "mapping_table": [{"movement_code": "N1"}],
        "tmc_batch_analysis_result": object(),
        "tmc_batch_export_result": object(),
    }
    original_state = dict(state)

    assert set_app_view(APP_SUPPORT_VIEW, state=state) == APP_SUPPORT_VIEW
    assert get_app_view(state=state) == APP_SUPPORT_VIEW
    assert state[APP_VIEW_STATE_KEY] == APP_SUPPORT_VIEW

    set_app_view("TMC Processor", state=state)

    assert state["active_workflow_tab"] == original_state["active_workflow_tab"]
    assert state["work_mode"] == original_state["work_mode"]
    assert state["mapping_table"] == original_state["mapping_table"]
    assert state["tmc_batch_analysis_result"] is original_state["tmc_batch_analysis_result"]
    assert state["tmc_batch_export_result"] is original_state["tmc_batch_export_result"]


def test_support_qr_resolver_uses_the_supplied_asset_and_handles_missing_files(tmp_path: Path) -> None:
    supplied_qr = tmp_path / "support_qr.png"
    supplied_qr.write_bytes(b"authoritative qr bytes")

    assert resolve_support_qr_path(supplied_qr) == supplied_qr
    assert resolve_support_qr_path(tmp_path / "missing.png") is None


def test_support_qr_resolver_finds_the_repository_asset() -> None:
    expected = Path(__file__).parents[1] / "assets" / "support_qr.png"

    assert expected.is_file()
    assert resolve_support_qr_path() == expected


def test_support_view_shows_a_clean_fallback_without_a_qr(tmp_path: Path, monkeypatch) -> None:
    class RecordingView:
        def __init__(self) -> None:
            self.captions: list[str] = []
            self.images: list[str] = []

        def title(self, value: str) -> None:
            pass

        def markdown(self, value: str) -> None:
            pass

        def caption(self, value: str) -> None:
            self.captions.append(value)

        def image(self, value: str, *, width: int) -> None:
            self.images.append(value)

        def container(self, *, border: bool):
            return nullcontext()

        def columns(self, ratios: list[float], *, gap: str):
            return (nullcontext(), nullcontext())

    view = RecordingView()
    monkeypatch.setattr(support, "st", view)
    monkeypatch.setattr(support, "SUPPORT_QR_PATH", tmp_path / "missing.png")

    support.render_support_view()

    assert view.images == []
    assert any("ไม่มีไฟล์ QR" in caption for caption in view.captions)
