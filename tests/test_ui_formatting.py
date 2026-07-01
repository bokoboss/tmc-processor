from __future__ import annotations

import pandas as pd

import app


def test_numeric_format_helpers_are_report_friendly() -> None:
    assert app.format_count(12345.67) == "12,346"
    assert app.format_pcu(12345) == "12,345"
    assert app.format_pcu(12345.25) == "12,345.2"
    assert app.format_percent(0.1234) == "12.3%"
    assert app.format_phf(0.876) == "0.88"


def test_time_range_helper_normalizes_peak_labels() -> None:
    assert app.format_time_range("07.00-08.00") == "07:00–08:00"
    assert app.format_time_range("07:00-08:00") == "07:00–08:00"
    assert app.format_time_range("") == "-"


def test_batch_peak_review_display_keeps_status_and_qc_near_front() -> None:
    analysis = type(
        "Analysis",
        (),
        {
            "items": [
                type(
                    "Item",
                    (),
                    {
                        "file_name": "long_demo_file.xlsx",
                        "survey_date_text": "2026-05-24",
                        "output_stem": "demo",
                        "folder_name": "file_01_demo",
                        "status": "success",
                        "mapping_status": "ready",
                        "suggested_AM_peak": "07:00-08:00",
                        "suggested_PM_peak": "17:00-18:00",
                        "confirmed_AM_peak": "07:00-08:00",
                        "confirmed_PM_peak": "17:00-18:00",
                        "total_PCU": 1234.5,
                        "QC_errors": 0,
                        "QC_warnings": 2,
                        "QC_info": 1,
                        "notes": "",
                    },
                )()
            ]
        },
    )()

    display = app._batch_peak_review_display_frame(analysis)

    assert list(display.columns[:4]) == ["ชื่อไฟล์", "วันที่สำรวจ", "สถานะ", "AM แนะนำ"]
    assert display.loc[0, "สถานะ"] == "สำเร็จ"
    assert display.loc[0, "QC เตือน"] == "2"


def test_qc_display_frame_uses_thai_labels_without_mutating_source() -> None:
    source = pd.DataFrame(
        [
            {
                "severity": "warning",
                "category": "peak",
                "check": "low_peak_phf",
                "message": "Low PHF",
                "detail": "phf=0.72",
                "movement_code": "NB",
                "raw_sheet": "North",
            }
        ]
    )

    display = app._qc_display_frame(source, thai_labels=True)

    assert "ระดับ" in display.columns
    assert display.loc[0, "ระดับ"] == "เตือน"
    assert source.loc[0, "severity"] == "warning"


def test_movement_label_th_mapping() -> None:
    assert app._movement_label_th("L") == "ซ้าย"
    assert app._movement_label_th("T") == "ตรง"
    assert app._movement_label_th("R") == "ขวา"
    assert app._movement_label_th("U") == "กลับรถ"


def test_basic_mapping_derived_summary_uses_normalized_from_to_mapping() -> None:
    source = pd.DataFrame(
        [
            {"raw_sheet": "ทิศ 1", "movement_code": "NE"},
            {"raw_sheet": "ทิศ 2", "movement_code": "NS"},
            {"raw_sheet": "ทิศ 3", "movement_code": "NW"},
            {"raw_sheet": "ทิศ 10", "movement_code": "WU"},
        ]
    )

    summary = app._basic_mapping_derived_summary(source)

    assert summary[["Code", "Derived movement", "Type", "Thai label", "Status"]].to_dict("records") == [
        {"Code": "NE", "Derived movement": "N → E", "Type": "L", "Thai label": "ซ้าย", "Status": "Valid"},
        {"Code": "NS", "Derived movement": "N → S", "Type": "T", "Thai label": "ตรง", "Status": "Valid"},
        {"Code": "NW", "Derived movement": "N → W", "Type": "R", "Thai label": "ขวา", "Status": "Valid"},
        {"Code": "WU", "Derived movement": "W → W", "Type": "U", "Thai label": "กลับรถ", "Status": "Valid"},
    ]


def test_basic_mapping_derived_summary_reports_invalid_code_clearly() -> None:
    summary = app._basic_mapping_derived_summary(pd.DataFrame([{"raw_sheet": "ทิศ 1", "movement_code": "XX"}]))

    assert summary.loc[0, "Status"].startswith("รหัสทิศทางไม่ถูกต้อง")
    assert "NE, NS, NW, NU" in summary.loc[0, "Status"]
