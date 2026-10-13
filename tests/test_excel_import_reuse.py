"""Phase B: measure import work and guard session reuse/upload lifecycle."""
from datetime import time
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZIP_STORED

import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

from tmc_processor import importer
from tmc_processor.ui import app_shell
from tmc_processor.excel_com_export import ExcelComStatus


def workbook(count=10, sheet="ทิศ 1"):
    table = pd.DataFrame([
        ["start", "end", "รถยนต์นั่งส่วนบุคคลไม่เกิน 7 คน"],
        *[[time(7, minute), time(8) if minute == 45 else time(7, minute + 15), count]
          for minute in (0, 15, 30, 45)],
    ])
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        table.to_excel(writer, sheet_name=sheet, index=False, header=False)
        pd.DataFrame([["ignored"]]).to_excel(writer, sheet_name="Summary", index=False, header=False)
    # Same-width changed values produce exactly equal upload lengths.
    stored = BytesIO()
    with ZipFile(BytesIO(buffer.getvalue())) as source, ZipFile(stored, "w", ZIP_STORED) as target:
        for name in source.namelist():
            target.writestr(name, source.read(name))
    return stored.getvalue()


@pytest.fixture
def io_counts(monkeypatch):
    import openpyxl
    from pandas.io.excel._openpyxl import OpenpyxlReader
    counts = dict(opens=0, reads=0, parses=0, closes=0)
    def counted(key, function):
        def run(*args, **kwargs):
            counts[key] += 1
            return function(*args, **kwargs)
        return run
    monkeypatch.setattr(openpyxl, "load_workbook", counted("opens", openpyxl.load_workbook))
    monkeypatch.setattr(OpenpyxlReader, "get_sheet_data", counted("reads", OpenpyxlReader.get_sheet_data))
    monkeypatch.setattr(importer, "parse_raw_count_sheet", counted("parses", importer.parse_raw_count_sheet))
    monkeypatch.setattr(pd.ExcelFile, "close", counted("closes", pd.ExcelFile.close))
    return counts


def reuse(payload, state):
    from tmc_processor.application.workbook_import import single_workbook_details
    return single_workbook_details(payload, state)


def test_cold_details_open_and_read_each_source_once(io_counts):
    details = importer.load_detected_sheet_details(BytesIO(workbook()))
    assert list(details) == ["ทิศ 1"]
    assert io_counts["opens"] == 1
    assert io_counts["reads"] == io_counts["parses"] == 1
    assert io_counts["closes"] >= 1


def test_unchanged_content_reuses_parse_and_isolates_mutations(io_counts):
    payload, state = workbook(), {}
    details = reuse(payload, state)
    expected = details["ทิศ 1"].data.copy(deep=True)
    details["ทิศ 1"].data.loc[:, "count"] = 999
    details["ทิศ 1"].raw.iloc[1, 2] = 999
    details["ทิศ 1"].debug.vehicle_class_columns[0]["header"] = "corrupted"
    details.clear()
    second = reuse(payload, state)
    pd.testing.assert_frame_equal(second["ทิศ 1"].data, expected)
    assert second["ทิศ 1"].raw.iloc[1, 2] == 10
    assert second["ทิศ 1"].debug.vehicle_class_columns[0]["header"] != "corrupted"
    assert io_counts["opens"] == io_counts["reads"] == io_counts["parses"] == 1


def test_same_length_changed_bytes_evict_previous_source(io_counts):
    first, second, state = workbook(10), workbook(20), {}
    assert len(first) == len(second) and first != second
    reuse(first, state)
    assert reuse(second, state)["ทิศ 1"].data["count"].tolist() == [20] * 4
    reuse(first, state)
    assert len(state) == 1
    assert io_counts["parses"] == 3  # one current entry, no retention of old uploads


def test_clear_and_failed_replacement_do_not_retain_stale_data(io_counts):
    payload, state = workbook(), {}
    reuse(payload, state)
    assert reuse(None, state) == {} and not state
    reuse(payload, state)
    with pytest.raises(ValueError):
        reuse(b"invalid workbook", state)
    assert not state
    reuse(payload, state)
    assert io_counts["parses"] == 3


def test_separate_sessions_never_share_cached_data(io_counts):
    payload, first, second = workbook(), {}, {}
    reuse(payload, first)
    reuse(payload, second)
    assert io_counts["parses"] == 2
    reuse(None, first)
    assert second and not first
    assert reuse(payload, second)["ทิศ 1"].data["count"].sum() == 40
    assert io_counts["parses"] == 2


def test_preview_uses_existing_details_without_io_and_preserves_raw(io_counts):
    details = reuse(workbook(), {})
    summary, previews = importer.preview_from_details(details, preview_rows=2)
    assert summary.to_dict("records") == [dict(raw_sheet="ทิศ 1", raw_direction="1", rows=5, columns=3)]
    pd.testing.assert_frame_equal(previews["ทิศ 1"], details["ทิศ 1"].raw.head(2))
    previews["ทิศ 1"].iloc[1, 2] = 999
    assert details["ทิศ 1"].raw.iloc[1, 2] == 10
    assert io_counts["opens"] == io_counts["parses"] == 1


def test_cache_byte_budget_bypasses_oversized_results(monkeypatch, io_counts):
    from tmc_processor.application import workbook_import
    monkeypatch.setattr(workbook_import, "MAX_SINGLE_WORKBOOK_CACHE_BYTES", 1)
    payload, state = workbook(), {}
    assert reuse(payload, state)["ทิศ 1"].data["count"].sum() == 40
    assert not state
    reuse(payload, state)
    assert io_counts["parses"] == 2


@pytest.mark.parametrize("value,status", [("N/A", "invalid"), ("bad", "invalid"), (None, "blank"), (0, "zero")])
def test_cached_details_preserve_phase_a_count_provenance(value, status):
    payload, state = workbook(value), {}
    first, second = reuse(payload, state), reuse(payload, state)
    pd.testing.assert_frame_equal(first["ทิศ 1"].data, second["ทิศ 1"].data)
    assert set(second["ทิศ 1"].data.source_count_status) == {status}
    assert second["ทิศ 1"].data.source_row.tolist() == [2, 3, 4, 5]
    assert set(second["ทิศ 1"].data.source_column) == {3}


def test_empty_source_sheet_is_cached(io_counts):
    buffer, state = BytesIO(), {}
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame([["no count table"]]).to_excel(writer, sheet_name="ทิศ 1", index=False, header=False)
    payload = buffer.getvalue()
    assert reuse(payload, state)["ทิศ 1"].data.empty
    assert reuse(payload, state)["ทิศ 1"].data.empty
    assert io_counts["opens"] == io_counts["parses"] == 1


@pytest.fixture
def app_test(monkeypatch):
    import sys
    monkeypatch.setitem(sys.modules, "__main__", sys.modules["__main__"])
    monkeypatch.setattr(app_shell, "_probe_excel_com_for_ui",
                        lambda force=False: ExcelComStatus(False, "test", "", ""))
    return AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py"), default_timeout=60)


def upload(at, name, payload):
    at.session_state[app_shell.SINGLE_SOURCE_UPLOAD_STATE_KEY] = app_shell._upload_record(app_shell._StoredUpload(name, payload))
    at.run()
    assert not at.exception, [error.message for error in at.exception]


def test_actual_streamlit_cold_upload_and_warm_reruns_parse_once(app_test, io_counts):
    upload(app_test, "first.xlsx", workbook())
    assert io_counts["opens"] == io_counts["reads"] == io_counts["parses"] == 1
    for stage in ("Mapping", "Analyze", "Review", "Export", "Data"):
        app_test.session_state["active_workflow_tab"] = stage
        app_test.run()
        assert not app_test.exception
    assert io_counts["opens"] == io_counts["reads"] == io_counts["parses"] == 1


def test_actual_streamlit_filename_content_and_clear_lifecycle(app_test, io_counts):
    from tmc_processor.application.workbook_import import SINGLE_WORKBOOK_CACHE_KEY
    first, second = workbook(10), workbook(20)
    assert len(first) == len(second)
    upload(app_test, "first.xlsx", first)
    app_test.session_state["tmc_output"] = {"workbook_bytes": b"stale export"}
    upload(app_test, "renamed.xlsx", first)
    assert app_test.session_state["tmc_uploaded_identity"][0] == "renamed.xlsx"
    # Existing nonempty metadata survives rename; source/download identity changes.
    assert app_test.session_state["survey_point_input"] == "first"
    assert any(item.value == "Loaded source: renamed.xlsx" for item in app_test.caption)
    assert "tmc_output" not in app_test.session_state
    assert io_counts["parses"] == 1
    upload(app_test, "renamed.xlsx", second)
    assert io_counts["parses"] == 2
    cache = app_test.session_state[SINGLE_WORKBOOK_CACHE_KEY]
    app_test.button(key="clear_single_source_upload").click().run()
    assert not app_test.exception
    assert SINGLE_WORKBOOK_CACHE_KEY not in app_test.session_state
    assert app_test.session_state["tmc_detected_sheet_names"] == []
    upload(app_test, "renamed.xlsx", second)
    assert io_counts["parses"] == 3
    assert app_test.session_state[SINGLE_WORKBOOK_CACHE_KEY] == cache


def test_real_uploader_replacement_invalidates_analysis_review_export(app_test, io_counts):
    from tmc_processor.application.workbook_import import SINGLE_WORKBOOK_CACHE_KEY
    app_test.run()
    uploader = next(widget for widget in app_test.file_uploader if widget.key == "raw_tmc_upload")
    uploader.set_value(("source.xlsx", workbook(10), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")).run()
    assert not app_test.exception
    app_test.session_state["tmc_processed"] = {"obsolete": True}
    app_test.session_state["tmc_output"] = {"workbook_bytes": b"obsolete"}
    app_test.session_state["tmc_confirmed_am_peak_start"] = "07:00"
    uploader = next(widget for widget in app_test.file_uploader if widget.key == "raw_tmc_upload")
    uploader.set_value(("source.xlsx", workbook(20), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")).run()
    assert not app_test.exception
    assert "tmc_processed" not in app_test.session_state
    assert "tmc_output" not in app_test.session_state
    assert "tmc_confirmed_am_peak_start" not in app_test.session_state
    assert io_counts["parses"] == 2
    cached = app_test.session_state[SINGLE_WORKBOOK_CACHE_KEY]
    assert cached
    app_test.radio(key="work_mode").set_value("ประมวลผลหลายไฟล์").run()
    assert not app_test.exception
    assert app_test.session_state[SINGLE_WORKBOOK_CACHE_KEY] == cached
    assert io_counts["parses"] == 2


def test_project_session_round_trip_excludes_cache_and_preserves_reuse(app_test, monkeypatch, io_counts):
    from tmc_processor.application.workbook_import import SINGLE_WORKBOOK_CACHE_KEY
    from tmc_processor.session import apply_session_to_state, session_from_json, session_to_json_bytes
    payload = workbook()
    upload(app_test, "source.xlsx", payload)
    state = dict(app_test.session_state)
    cached = state[SINGLE_WORKBOOK_CACHE_KEY]
    with monkeypatch.context() as local:
        local.setattr(app_shell.st, "session_state", state)
        session = app_shell._build_session_from_state("source.xlsx", len(payload))
    serialized = session_to_json_bytes(session)
    assert SINGLE_WORKBOOK_CACHE_KEY.encode() not in serialized
    assert b"source_count_raw" not in serialized
    loaded = session_from_json(serialized).session
    assert loaded["schema_version"] == 1 and loaded["source_file_name"] == "source.xlsx"
    apply_session_to_state(loaded, state)
    assert state[SINGLE_WORKBOOK_CACHE_KEY] == cached
    assert reuse(payload, state)["ทิศ 1"].data["count"].sum() == 40
    assert io_counts["parses"] == 1


def test_parser_revision_change_reparses_once_then_reuses(io_counts, monkeypatch):
    import hashlib
    from tmc_processor.application import workbook_import
    payload, state = workbook(), {}
    monkeypatch.setattr(workbook_import, "SINGLE_WORKBOOK_PARSER_REVISION", 1, raising=False)
    reuse(payload, state)
    reuse(payload, state)
    assert io_counts["parses"] == 1

    monkeypatch.setattr(workbook_import, "SINGLE_WORKBOOK_PARSER_REVISION", 2, raising=False)
    refreshed = reuse(payload, state)
    assert refreshed["ทิศ 1"].data["count"].tolist() == [10] * 4
    assert io_counts["parses"] == 2
    reuse(payload, state)
    assert io_counts["parses"] == 2
    assert state[workbook_import.SINGLE_WORKBOOK_CACHE_KEY][0] == 2
    assert state[workbook_import.SINGLE_WORKBOOK_CACHE_KEY][1] == hashlib.sha256(payload).hexdigest()


def test_revisionless_legacy_entry_with_stale_details_is_evicted(io_counts):
    import hashlib
    import pickle
    from tmc_processor.application.workbook_import import SINGLE_WORKBOOK_CACHE_KEY
    payload, state = workbook(10), {}
    stale = reuse(workbook(99), {})
    # Legacy tuple shape: (source digest, serialized details), without revision.
    state[SINGLE_WORKBOOK_CACHE_KEY] = (hashlib.sha256(payload).hexdigest(), pickle.dumps(stale))

    current = reuse(payload, state)

    assert current["ทิศ 1"].data["count"].tolist() == [10] * 4
    assert io_counts["parses"] == 2  # stale fixture plus exactly one parse of current bytes
    assert len(state[SINGLE_WORKBOOK_CACHE_KEY]) == 3


def test_corrupt_cached_payload_is_evicted_and_source_is_reparsed(io_counts):
    import hashlib
    from tmc_processor.application import workbook_import
    payload, state = workbook(10), {}
    revision = getattr(workbook_import, "SINGLE_WORKBOOK_PARSER_REVISION", 1)
    state[workbook_import.SINGLE_WORKBOOK_CACHE_KEY] = (
        revision, hashlib.sha256(payload).hexdigest(), b"not a valid pickle",
    )

    current = reuse(payload, state)

    assert current["ทิศ 1"].data["count"].tolist() == [10] * 4
    assert io_counts["parses"] == 1
    assert len(state[workbook_import.SINGLE_WORKBOOK_CACHE_KEY]) == 3


def test_corrupt_cache_does_not_hide_invalid_original_workbook_error():
    import hashlib
    import pytest
    from tmc_processor.application.workbook_import import SINGLE_WORKBOOK_CACHE_KEY
    invalid = b"invalid original workbook"
    state = {SINGLE_WORKBOOK_CACHE_KEY: (hashlib.sha256(invalid).hexdigest(), b"corrupt cache payload")}

    with pytest.raises(ValueError):
        reuse(invalid, state)

    assert SINGLE_WORKBOOK_CACHE_KEY not in state
