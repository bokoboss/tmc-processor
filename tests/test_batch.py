from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

from tmc_processor.batch import BatchItem, process_batch_files
from tmc_processor.mapping_preset import load_mapping_preset


ROOT = Path(__file__).resolve().parents[1]
DEMO_DIR = ROOT / "samples" / "demo"
DAY1 = DEMO_DIR / "DEMO_TMC1_FourLeg.xlsx"
DAY2 = DEMO_DIR / "DEMO_TMC1_FourLeg_Day2.xlsx"
PRESET = DEMO_DIR / "DEMO_TMC1_FourLeg.mapping.json"


def _demo_items() -> list[BatchItem]:
    return [
        BatchItem(file_name=DAY1.name, workbook_bytes=DAY1.read_bytes()),
        BatchItem(file_name=DAY2.name, workbook_bytes=DAY2.read_bytes()),
    ]


def _preset() -> dict[str, object]:
    return load_mapping_preset(PRESET.read_bytes()).preset


def _setup() -> dict[str, object]:
    return {
        "project_name": "Synthetic Demo Project",
        "tmc_id": "DEMO-TMC1",
        "tmc_title": "Synthetic Four-Leg Demo Intersection",
        "survey_point": "Synthetic Four-Leg Demo Point",
        "survey_date_text": "2026-01-01",
        "survey_period": "07:00-19:00",
        "north_label": "Demo North",
        "south_label": "Demo South",
        "east_label": "Demo East",
        "west_label": "Demo West",
    }


def test_processes_two_demo_workbooks_with_one_mapping_preset() -> None:
    result = process_batch_files(
        _demo_items(),
        mapping_preset=_preset(),
        setup=_setup(),
        generated_at="2026-05-19T10:00:00Z",
    )

    assert [row.status for row in result.summary_rows] == ["success", "success"]
    assert all(row.AM_peak for row in result.summary_rows)
    assert all(row.PM_peak for row in result.summary_rows)
    assert all(row.total_vehicles > 0 for row in result.summary_rows)
    assert all(row.total_PCU > 0 for row in result.summary_rows)
    assert all(row.QC_errors >= 0 and row.QC_warnings >= 0 and row.QC_info >= 0 for row in result.summary_rows)


def test_batch_zip_contains_summary_and_one_folder_per_success_without_raw_inputs() -> None:
    result = process_batch_files(
        _demo_items(),
        mapping_preset=_preset(),
        setup=_setup(),
        generated_at="2026-05-19T10:00:00Z",
    )

    with ZipFile(BytesIO(result.package_bytes)) as archive:
        names = set(archive.namelist())

    assert "batch_summary.xlsx" in names
    success_folders = {row.folder_name for row in result.summary_rows if row.status == "success"}
    for folder in success_folders:
        assert f"{folder}/report.xlsx" in names
        assert f"{folder}/export_summary.txt" in names
        assert f"{folder}/session.tmcproj.json" in names
        assert f"{folder}/mapping_preset.mapping.json" in names
    assert len(success_folders) == 2
    assert DAY1.name not in names
    assert DAY2.name not in names
    assert all(not name.endswith(".xlsm") and name not in {DAY1.name, DAY2.name} for name in names)


def test_failed_file_does_not_stop_batch_and_summary_contains_success_and_failure_rows() -> None:
    items = [
        BatchItem(file_name=DAY1.name, workbook_bytes=DAY1.read_bytes()),
        BatchItem(file_name="broken_input.xlsx", workbook_bytes=b"not an excel workbook"),
    ]
    result = process_batch_files(
        items,
        mapping_preset=_preset(),
        setup=_setup(),
        generated_at="2026-05-19T10:00:00Z",
    )

    assert [row.status for row in result.summary_rows] == ["success", "failed"]
    assert result.summary_rows[1].notes

    with ZipFile(BytesIO(result.package_bytes)) as archive:
        names = set(archive.namelist())
        summary_bytes = archive.read("batch_summary.xlsx")

    assert "file_01_DEMO_TMC1_FourLeg/report.xlsx" in names
    assert "file_02_broken_input/report.xlsx" not in names

    workbook = load_workbook(BytesIO(summary_bytes), read_only=True, data_only=True)
    rows = list(workbook["batch_summary"].iter_rows(values_only=True))
    headers = list(rows[0])
    records = [dict(zip(headers, row)) for row in rows[1:]]

    assert [record["status"] for record in records] == ["success", "failed"]
    assert {"QC_errors", "QC_warnings", "QC_info"}.issubset(headers)
    assert records[1]["notes"]
