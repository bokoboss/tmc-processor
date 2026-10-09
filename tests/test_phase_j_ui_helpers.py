from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook
from PIL import Image
import pandas as pd
import pytest

import app
from tmc_processor.batch import batch_inputs_ready
from tmc_processor.constants import AM_WINDOW, DEFAULT_PEAK_MODE, PM_WINDOW
from tmc_processor.importer import load_detected_sheets
from tmc_processor.mapping import apply_saved_mapping_to_sheets, default_mapping_for_sheets
from tmc_processor.mapping_preset import apply_mapping_preset_to_detected_sheets, load_mapping_preset
from tmc_processor.movement_scheme import MOVEMENT_SCHEME_V1, MOVEMENT_SCHEME_V2
from tmc_processor.pipeline import ProcessingResult, V2DryRunResult


ROOT = Path(__file__).resolve().parents[1]
DEMO_DIR = ROOT / "samples" / "demo"
RAW_WORKBOOK = DEMO_DIR / "DEMO_TMC1_FourLeg.xlsx"
V1_MAPPING_XLSX = DEMO_DIR / "DEMO_TMC1_FourLeg_mapping.xlsx"
V2_PRESET = DEMO_DIR / "DEMO_TMC1_FourLeg_approach_v2.mapping.json"


def _setup(scheme: str) -> dict[str, object]:
    return {
        "project_name": "Phase J UI Helper",
        "tmc_id": "PHASE-J",
        "tmc_name": "Phase J Demo",
        "survey_date_text": "2026-05-26",
        "movement_code_scheme": scheme,
    }


def _peak_windows() -> dict[str, tuple[str, str]]:
    return {"AM": AM_WINDOW, "PM": PM_WINDOW}


def _raw_sheets() -> dict[str, pd.DataFrame]:
    return load_detected_sheets(RAW_WORKBOOK)


def _v2_mapping(raw_sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    preset = load_mapping_preset(V2_PRESET.read_bytes()).preset
    return apply_mapping_preset_to_detected_sheets(preset, list(raw_sheets)).mapping


def _v1_mapping(raw_sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    return apply_saved_mapping_to_sheets(list(raw_sheets), pd.read_excel(V1_MAPPING_XLSX, sheet_name="Mapping"))


def _v2_result() -> tuple[V2DryRunResult, pd.DataFrame]:
    raw_sheets = _raw_sheets()
    mapping = _v2_mapping(raw_sheets)
    result = app._process_single_file_for_ui(
        raw_sheets=raw_sheets,
        mapping=mapping,
        setup=_setup(MOVEMENT_SCHEME_V2),
        detected_sheets=list(raw_sheets),
        peak_mode=DEFAULT_PEAK_MODE,
        peak_windows=_peak_windows(),
        pce_factors={},
    )
    assert isinstance(result, V2DryRunResult)
    return result, mapping


def test_single_file_ui_helper_processes_v2_through_dry_run_path() -> None:
    result, _ = _v2_result()

    assert result.movement_code_scheme == MOVEMENT_SCHEME_V2
    assert not result.normalized.empty
    assert set(result.normalized["movement_code_scheme"]) == {MOVEMENT_SCHEME_V2}


def test_single_file_ui_helper_derives_v2_processor_legs_from_basic_mapping() -> None:
    raw_sheets = {
        "Sheet 1": pd.DataFrame(
            [
                {
                    "raw_sheet": "Sheet 1",
                    "raw_direction": "West",
                    "time_start": "07:00",
                    "time_end": "07:15",
                    "vehicle_class": "PC<7",
                    "count": 3,
                }
            ]
        )
    }
    mapping = pd.DataFrame(
        [
            {
                "raw_sheet": "Sheet 1",
                "raw_direction": "West",
                "movement_code": "WL",
                "source_stream": "mainline",
                "raw_movement_label": "West left",
                "from_leg": "",
                "to_leg": "",
                "turn_type": "",
                "facility_type": "at_grade",
                "include_in_peak": True,
                "include_in_report": True,
                "aggregation_method": "sum",
            }
        ]
    )

    result = app._process_single_file_for_ui(
        raw_sheets=raw_sheets,
        mapping=mapping,
        setup=_setup(MOVEMENT_SCHEME_V2),
        detected_sheets=list(raw_sheets),
        peak_mode=DEFAULT_PEAK_MODE,
        peak_windows=_peak_windows(),
        pce_factors={},
    )

    assert isinstance(result, V2DryRunResult)
    assert result.normalized.loc[0, "from_leg"] == "W"
    assert result.normalized.loc[0, "to_leg"] == "N"


def test_v2_result_feeds_peak_review_hourly_data_helper() -> None:
    result, mapping = _v2_result()

    hourly = app._hourly_movement_for_ui(result, mapping)

    assert not hourly.empty
    assert "NT" in hourly.columns
    assert "Total" in hourly.columns


def test_v2_single_ui_export_uses_authored_template_without_excel_com(monkeypatch: pytest.MonkeyPatch) -> None:
    result, mapping = _v2_result()
    export_payload = app._workflow_export_payload(
        _setup(MOVEMENT_SCHEME_V1),
        movement_code_scheme=MOVEMENT_SCHEME_V2,
        export_mode=app.EXCEL_TEMPLATE_EXPORT_MODE,
    )
    assert export_payload["template_version"] == "four_leg_approach_movement_v2"
    assert Path(str(export_payload["template_path"])).name == "four_leg_tmc_report_template_approach_v2.xlsx"

    def fail_if_called(*args: object, **kwargs: object) -> bytes:
        raise AssertionError("V2 Single export must use direct OOXML")

    monkeypatch.setattr(app, "export_v2_template_workbook_com", fail_if_called)

    workbook_bytes = app._export_single_file_for_ui(
        result=result,
        mapping=mapping,
        setup=_setup(MOVEMENT_SCHEME_V2),
        export_mode=app.SAFE_PNG_EXPORT_MODE,
        use_template_report_layout=False,
        use_excel_com_native_charts=False,
        source_file_name=RAW_WORKBOOK.name,
        generated_at="2026-05-26 12:00:00",
    )
    workbook = load_workbook(BytesIO(workbook_bytes), read_only=True)

    assert workbook.sheetnames[0] == "Summary"
    assert "Hourly_Movement_PCU" in workbook.sheetnames
    assert "Movement_Diagram_Data" in workbook.sheetnames
    assert "Movement_Source_Stream_Audit" in workbook.sheetnames
    with ZipFile(BytesIO(workbook_bytes)) as exported, ZipFile(ROOT / "templates" / "four_leg_tmc_report_template_approach_v2.xlsx") as template:
        assert exported.testzip() is None
        assert exported.read("xl/drawings/drawing1.xml") == template.read("xl/drawings/drawing1.xml")
        assert exported.read("xl/charts/chart1.xml")
        assert exported.read("xl/charts/chart2.xml")
    package_bytes = app.create_v2_generated_export_package_zip(workbook_bytes=workbook_bytes)
    with ZipFile(BytesIO(package_bytes)) as package:
        assert package.testzip() is None
        assert package.read("approach_movement_v2_template_workbook.xlsx") == workbook_bytes
        with Image.open(BytesIO(package.read("diagram/movement_diagram.png"))) as image:
            assert image.format == "PNG"
            image.verify()


def test_v2_single_template_mode_does_not_call_excel_com(monkeypatch: pytest.MonkeyPatch) -> None:
    result, mapping = _v2_result()

    def fail_if_called(*args: object, **kwargs: object) -> bytes:
        raise AssertionError("V2 template export must not require Excel COM")

    monkeypatch.setattr(app, "export_v2_template_workbook_com", fail_if_called)

    workbook_bytes = app._export_single_file_for_ui(
        result=result,
        mapping=mapping,
        setup=_setup(MOVEMENT_SCHEME_V2),
        export_mode=app.EXCEL_TEMPLATE_EXPORT_MODE,
        use_template_report_layout=True,
        use_excel_com_native_charts=False,
        source_file_name=RAW_WORKBOOK.name,
        generated_at="2026-05-26 12:00:00",
    )
    with ZipFile(BytesIO(workbook_bytes)) as package:
        assert package.testzip() is None
        assert "xl/drawings/drawing1.xml" in package.namelist()


def test_v2_single_export_never_selects_excel_com_backend(monkeypatch: pytest.MonkeyPatch) -> None:
    result, mapping = _v2_result()
    calls: list[dict[str, object]] = []

    def fake_com_export(*args: object, **kwargs: object) -> bytes:
        calls.append({"args": args, "kwargs": kwargs})
        raise AssertionError("V2 Single export must not call Excel COM")

    monkeypatch.setattr(app, "export_v2_template_workbook_com", fake_com_export)

    workbook_bytes = app._export_single_file_for_ui(
        result=result,
        mapping=mapping,
        setup=_setup(MOVEMENT_SCHEME_V2),
        export_mode=app.EXCEL_TEMPLATE_EXPORT_MODE,
        use_template_report_layout=True,
        use_excel_com_native_charts=False,
        source_file_name=RAW_WORKBOOK.name,
        generated_at="2026-05-26 12:00:00",
    )

    assert workbook_bytes.startswith(b"PK")
    assert calls == []
    with ZipFile(BytesIO(workbook_bytes)) as package:
        assert package.testzip() is None
        assert "xl/drawings/drawing1.xml" in package.namelist()


def test_v2_batch_analysis_is_ready_for_phase_k() -> None:
    assert batch_inputs_ready(
        uploaded_workbook_count=1,
        mapping_available=True,
        pce_factors_ready=True,
        movement_code_scheme=MOVEMENT_SCHEME_V2,
    )


def test_v1_single_file_ui_helper_still_uses_processing_result_path() -> None:
    raw_sheets = _raw_sheets()
    mapping = _v1_mapping(raw_sheets)

    result = app._process_single_file_for_ui(
        raw_sheets=raw_sheets,
        mapping=mapping,
        setup=_setup(MOVEMENT_SCHEME_V1),
        detected_sheets=list(raw_sheets),
        peak_mode=DEFAULT_PEAK_MODE,
        peak_windows=_peak_windows(),
        pce_factors={},
    )

    assert isinstance(result, ProcessingResult)
    assert not isinstance(result, V2DryRunResult)
    assert not result.normalized.empty


def test_v1_single_file_ui_helper_processes_movement_code_only_mapping() -> None:
    raw_sheets = {
        "ทิศ 1": pd.DataFrame(
            [
                {
                    "raw_sheet": "ทิศ 1",
                    "raw_direction": "North",
                    "time_start": "07:00",
                    "time_end": "07:15",
                    "vehicle_class": "PC<7",
                    "count": 3,
                }
            ]
        )
    }
    mapping = pd.DataFrame(
        [
            {
                "raw_sheet": "ทิศ 1",
                "movement_code": "NE",
                "from_leg": None,
                "to_leg": None,
                "turn_type": None,
            }
        ]
    )

    result = app._process_single_file_for_ui(
        raw_sheets=raw_sheets,
        mapping=mapping,
        setup=_setup(MOVEMENT_SCHEME_V1),
        detected_sheets=list(raw_sheets),
        peak_mode=DEFAULT_PEAK_MODE,
        peak_windows=_peak_windows(),
        pce_factors={},
    )

    assert isinstance(result, ProcessingResult)
    assert result.normalized.loc[0, ["from_leg", "to_leg", "turn_type"]].to_dict() == {
        "from_leg": "N",
        "to_leg": "E",
        "turn_type": "L",
    }


def test_v1_batch_readiness_unchanged() -> None:
    assert batch_inputs_ready(
        uploaded_workbook_count=1,
        mapping_available=True,
        pce_factors_ready=True,
        movement_code_scheme=MOVEMENT_SCHEME_V1,
    )


def test_project_session_preserves_v2_mapping_scheme() -> None:
    raw_sheets = _raw_sheets()
    mapping = _v2_mapping(raw_sheets)

    session = app.build_project_session(
        metadata={},
        directions={},
        mapping=mapping,
        movement_code_scheme=MOVEMENT_SCHEME_V2,
        detected_sheet_names=list(raw_sheets),
        pce_factors={},
    )

    assert session["mapping"]["movement_code_scheme"] == MOVEMENT_SCHEME_V2
