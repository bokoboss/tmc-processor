from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile

import pandas as pd

from tmc_processor.export_package import build_export_summary_text, create_export_package_zip
from tmc_processor.metadata import APP_VERSION, TEMPLATE_VERSION, get_app_version


def test_app_version_is_available_from_central_metadata() -> None:
    assert get_app_version() == APP_VERSION
    assert APP_VERSION
    assert TEMPLATE_VERSION == "four_leg_v1"


def _mapping(aggregate: bool = False) -> pd.DataFrame:
    rows = [
        {
            "raw_sheet": "North mainline",
            "raw_direction": "North",
            "movement_code": "NS",
            "source_stream": "mainline",
            "raw_movement_label": "North through",
            "from_leg": "N",
            "to_leg": "S",
            "turn_type": "through",
            "facility_type": "at_grade",
            "include_in_peak": True,
            "include_in_report": True,
            "aggregation_method": "sum",
        }
    ]
    if aggregate:
        rows.append(
            {
                "raw_sheet": "North frontage",
                "raw_direction": "North",
                "movement_code": "NS",
                "source_stream": "frontage",
                "raw_movement_label": "North frontage through",
                "from_leg": "N",
                "to_leg": "S",
                "turn_type": "through",
                "facility_type": "frontage",
                "include_in_peak": True,
                "include_in_report": True,
                "aggregation_method": "sum",
            }
        )
    return pd.DataFrame(rows)


def _qc() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"check": "aggregated_output_movement", "severity": "info", "message": "NS is aggregated."},
            {"check": "mapped_movement_zero_total", "severity": "warning", "message": "Zero total."},
        ]
    )


def test_export_summary_content_includes_traceability_fields() -> None:
    summary = build_export_summary_text(
        setup={
            "survey_point": "Main & 1st",
            "tmc_title": "TMC-01",
            "am_peak_start": "07:00",
            "am_peak_end": "08:00",
            "pm_peak_start": "17:00",
            "pm_peak_end": "18:00",
            "peak_selection_source": "user_confirmed",
        },
        source_file_name=r"C:\private\raw_input.xlsx",
        export_mode="Safe PNG Export Mode",
        mapping=_mapping(),
        qc=_qc(),
        workbook_filename="TMC01_output.xlsx",
        pce_factors={"MC": 0.75},
        export_settings={"template_name": "TMC_Template.xlsx"},
        generated_at="2026-05-19T10:00:00Z",
    )

    assert f"App version: {APP_VERSION}" in summary
    assert f"Template version: {TEMPLATE_VERSION}" in summary
    assert "Generated at: 2026-05-19T10:00:00Z" in summary
    assert "Source file name: raw_input.xlsx" in summary
    assert "Survey point / TMC title: Main & 1st" in summary
    assert "Export mode: Safe PNG Export Mode" in summary
    assert "QC warnings / info count: total=2, error=0, warning=1, info=1" in summary
    assert "Output workbook name: TMC01_output.xlsx" in summary


def test_export_summary_includes_pce_factors_and_confirmed_peaks() -> None:
    summary = build_export_summary_text(
        setup={
            "am_peak_start": "07:00",
            "am_peak_end": "08:00",
            "pm_peak_start": "16:15",
            "pm_peak_end": "17:15",
            "peak_selection_source": "user_confirmed",
        },
        mapping=_mapping(),
        pce_factors={"MC": 0.75},
        generated_at="2026-05-19T10:00:00Z",
    )

    assert "AM peak period: 07:00-08:00" in summary
    assert "PM peak period: 16:15-17:15" in summary
    assert "Peak selection source: user_confirmed" in summary
    assert "- MC: 0.75 (user_override)" in summary
    assert "- PC<7: 1 (default)" in summary


def test_export_summary_includes_aggregation_summary_when_present() -> None:
    summary = build_export_summary_text(
        mapping=_mapping(aggregate=True),
        generated_at="2026-05-19T10:00:00Z",
    )

    assert "Mapping aggregation summary:" in summary
    assert "Mapping rows: 2" in summary
    assert "- NS is aggregated from 2 source streams." in summary


def test_export_package_zip_contains_expected_files() -> None:
    summary_text = build_export_summary_text(generated_at="2026-05-19T10:00:00Z")
    package = create_export_package_zip(
        workbook_bytes=b"workbook",
        workbook_filename="TMC01_output.xlsx",
        export_summary_text=summary_text,
        project_session_bytes=b'{"schema_version":1}',
        project_session_filename="TMC01_session.tmcproj.json",
        mapping=_mapping(),
        chart_pngs={"hourly_pcu": b"png1", "vehicle_composition": b"png2"},
        diagram_png=b"diagram",
    )

    with ZipFile(BytesIO(package)) as archive:
        names = set(archive.namelist())
        summary = archive.read("export_summary.txt").decode("utf-8")

    assert "TMC01_output.xlsx" in names
    assert "export_summary.txt" in names
    assert "TMC01_session.tmcproj.json" in names
    assert "mapping_table.xlsx" in names
    assert "charts/hourly_pcu.png" in names
    assert "charts/vehicle_composition.png" in names
    assert "charts/tmc_movement_diagram.png" in names
    assert f"App version: {APP_VERSION}" in summary
    assert f"Template version: {TEMPLATE_VERSION}" in summary


def test_export_package_zip_does_not_contain_raw_excel_input() -> None:
    package = create_export_package_zip(
        workbook_bytes=b"workbook",
        workbook_filename="TMC01_output.xlsx",
        export_summary_text="summary",
        project_session_bytes=b'{"source_file_name":"raw_input.xlsx"}',
        project_session_filename="TMC01_session.tmcproj.json",
        mapping=_mapping(),
    )

    with ZipFile(BytesIO(package)) as archive:
        names = set(archive.namelist())

    assert "raw_input.xlsx" not in names
    assert all("outputs/" not in name and ".venv/" not in name and "egg-info/" not in name for name in names)
