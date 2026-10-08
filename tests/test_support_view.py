from __future__ import annotations

from contextlib import nullcontext
from pathlib import Path

from tmc_processor.ui import app_shell
from tmc_processor.ui.components import support
from tmc_processor.ui.components.support import resolve_support_qr_path


# Issue #30, final copy override: issuecomment-6061137433.
APPROVED_SUPPORT_TEXT = """TMC Processor เปิดให้ใช้งานฟรี

ถ้าโปรแกรมนี้ช่วยลดเวลาจัดการข้อมูล TMC และงาน Excel ของคุณได้ และอยากช่วยสนับสนุนการพัฒนาต่อ จะเลี้ยงชาเย็นคนทำโปรแกรมสักแก้วก็ได้

**สแกน QR ได้ตามสะดวกเลย**

สนับสนุนหรือไม่ก็ใช้งานได้ครบทุกฟังก์ชันเหมือนเดิม"""


class RecordingHeader:
    def __init__(self, *, clicked: bool) -> None:
        self.clicked = clicked
        self.buttons: list[tuple[str, dict[str, object]]] = []
        self.markdown_values: list[str] = []
        self.captions: list[str] = []
        self.session_state: dict[str, object] = {}

    def columns(self, ratios: list[float], **_kwargs):
        return (nullcontext(), nullcontext())

    def markdown(self, value: str, **_kwargs) -> None:
        self.markdown_values.append(value)

    def caption(self, value: str) -> None:
        self.captions.append(value)

    def button(self, label: str, **kwargs) -> bool:
        self.buttons.append((label, kwargs))
        return self.clicked


def test_header_support_button_opens_dialog_without_changing_workflow_state(monkeypatch) -> None:
    state = {
        "active_workflow_tab": "Review",
        "work_mode": "ประมวลผลหลายไฟล์",
        "tmc_batch_source_uploads": [object(), object()],
        "mapping_table": [{"movement_code": "N1"}],
        "tmc_batch_confirmed_peaks": {"file-1": {"am": "07:00-08:00"}},
        "tmc_batch_analysis_result": object(),
        "tmc_batch_export_result": object(),
    }
    header = RecordingHeader(clicked=True)
    header.session_state = state
    original_state = dict(state)
    opened: list[bool] = []
    monkeypatch.setattr(app_shell, "st", header)
    monkeypatch.setattr(support, "open_support_dialog", lambda: opened.append(True))

    app_shell._render_app_header()

    assert header.buttons == [("เลี้ยงชาเย็น", {"key": "open_support_dialog", "type": "secondary"})]
    assert opened == [True]
    assert header.session_state == original_state
    assert header.session_state["tmc_batch_analysis_result"] is original_state["tmc_batch_analysis_result"]
    assert header.session_state["tmc_batch_export_result"] is original_state["tmc_batch_export_result"]


def test_header_keeps_privacy_disclosure_in_data_context(monkeypatch) -> None:
    header = RecordingHeader(clicked=False)
    monkeypatch.setattr(app_shell, "st", header)
    monkeypatch.setattr(support, "open_support_dialog", lambda: None)

    app_shell._render_app_header()

    assert header.captions == []


def test_data_uploads_include_privacy_disclosure_in_single_and_batch() -> None:
    from streamlit.testing.v1 import AppTest

    entrypoint = Path(__file__).resolve().parents[1] / "app.py"
    at = AppTest.from_file(str(entrypoint), default_timeout=60).run()
    assert not at.exception
    for batch in (False, True):
        if batch:
            mode = next(widget for widget in at.radio if widget.key == "work_mode")
            mode.set_value(mode.options[1]).run()
            assert not at.exception
        privacy = next(panel for panel in at.get("expander") if panel.label == "Privacy / ความเป็นส่วนตัว")
        assert [caption.value for caption in privacy.caption] == [support.PRIVACY_DISCLOSURE]
        assert at.file_uploader


def test_support_dialog_renders_exact_approved_text_and_authoritative_qr(monkeypatch) -> None:
    class RecordingView:
        def __init__(self) -> None:
            self.markdown_values: list[str] = []
            self.captions: list[str] = []
            self.images: list[tuple[str, int]] = []

        def markdown(self, value: str) -> None:
            self.markdown_values.append(value)

        def caption(self, value: str) -> None:
            self.captions.append(value)

        def image(self, value: str, *, width: int) -> None:
            self.images.append((value, width))

        def columns(self, _ratios: list[int], *, gap: str):
            assert gap == "small"
            return (nullcontext(), nullcontext(), nullcontext())

    view = RecordingView()
    monkeypatch.setattr(support, "st", view)
    expected_qr = Path("assets/support_qr.png")
    monkeypatch.setattr(support, "SUPPORT_QR_PATH", expected_qr)

    support.render_support_view()

    assert view.markdown_values == [APPROVED_SUPPORT_TEXT]
    assert support.SUPPORT_TITLE == "เลี้ยงชาเย็นคนทำโปรแกรม"
    assert view.images == [(str(expected_qr), 220)]
    assert view.captions == []


def test_support_qr_resolver_uses_supplied_asset_and_cleanly_handles_missing_files(tmp_path: Path) -> None:
    supplied_qr = tmp_path / "support_qr.png"
    supplied_qr.write_bytes(b"authoritative qr bytes")

    assert resolve_support_qr_path(supplied_qr) == supplied_qr
    assert resolve_support_qr_path(tmp_path / "missing.png") is None


def test_support_qr_resolver_finds_the_repository_asset() -> None:
    expected = Path(__file__).parents[1] / "assets" / "support_qr.png"

    assert expected.is_file()
    assert resolve_support_qr_path() == expected


def test_support_dialog_shows_a_clean_fallback_without_a_qr(tmp_path: Path, monkeypatch) -> None:
    class RecordingView:
        def __init__(self) -> None:
            self.captions: list[str] = []
            self.images: list[str] = []

        def markdown(self, _value: str) -> None:
            pass

        def caption(self, value: str) -> None:
            self.captions.append(value)

        def columns(self, *_args, **_kwargs):
            raise AssertionError("the QR layout should not render without an asset")

    view = RecordingView()
    monkeypatch.setattr(support, "st", view)
    monkeypatch.setattr(support, "SUPPORT_QR_PATH", tmp_path / "missing.png")

    support.render_support_view()

    assert view.images == []
    assert view.captions == ["ไม่มีไฟล์ QR สำหรับแสดงในสภาพแวดล้อมนี้"]


def test_support_dialog_uses_streamlit_compact_dialog(monkeypatch) -> None:
    opened: list[bool] = []
    monkeypatch.setattr(support, "_support_dialog", lambda: opened.append(True))

    support.open_support_dialog()

    assert opened == [True]
