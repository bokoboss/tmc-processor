"""Report-ready PNG chart generation."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Any, Mapping

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd

from .time_utils import hourly_interval_rows


CHART_FILENAMES = {
    "hourly_pcu": "hourly_pcu_chart.png",
    "vehicle_composition": "vehicle_composition_chart.png",
}


def _configure_fonts() -> None:
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = [
        "Leelawadee UI",
        "Tahoma",
        "Noto Sans Thai",
        "Arial Unicode MS",
        "DejaVu Sans",
    ]
    plt.rcParams["axes.unicode_minus"] = False


def _write_png(fig, output_path: str | Path | None = None) -> bytes:
    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    png = buffer.getvalue()
    if output_path is not None:
        Path(output_path).write_bytes(png)
    return png


def _first_column(frame: pd.DataFrame) -> str | None:
    return str(frame.columns[0]) if len(frame.columns) else None


def _without_total_rows(frame: pd.DataFrame) -> pd.DataFrame:
    return hourly_interval_rows(frame)


def _coerce_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def _chart_title(base_title: str, setup: Mapping[str, Any] | None = None) -> str:
    setup = setup or {}
    context = str(setup.get("tmc_title") or setup.get("tmc_name") or setup.get("survey_point") or "").strip()
    return f"{base_title}\n{context}" if context else base_title


def hourly_pcu_chart_png(
    hourly_movement_pcu: pd.DataFrame,
    output_path: str | Path | None = None,
    setup: Mapping[str, Any] | None = None,
) -> bytes:
    """Create an hourly PCU line chart from the Hourly_Movement_PCU table."""

    _configure_fonts()
    chart_data = _without_total_rows(hourly_movement_pcu)
    time_column = _first_column(chart_data)

    if chart_data.empty or time_column is None:
        chart_data = pd.DataFrame({"เวลา": [], "Total": []})
        time_column = "เวลา"

    if "Total" not in chart_data.columns:
        value_columns = [column for column in chart_data.columns if column != time_column]
        chart_data["Total"] = chart_data[value_columns].apply(_coerce_number).sum(axis=1) if value_columns else 0

    x_labels = chart_data[time_column].astype(str).tolist()
    y_values = _coerce_number(chart_data["Total"])

    fig, ax = plt.subplots(figsize=(10, 5.2))
    ax.plot(x_labels, y_values, color="#1F6F8B", linewidth=2.4, marker="o", markersize=4.8)
    ax.fill_between(range(len(x_labels)), y_values.to_numpy(), color="#C7E3EA", alpha=0.45)
    ax.set_title("ปริมาณจราจรรายชั่วโมง (PCU)", fontsize=16, fontweight="bold", pad=14)
    ax.set_xlabel("เวลา", fontsize=12)
    ax.set_ylabel("ปริมาณจราจร (PCU/ชั่วโมง)", fontsize=12)
    ax.grid(axis="y", color="#D9E2E8", linewidth=0.8)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="x", rotation=35)
    ax.set_ylim(bottom=0)
    ax.set_title(_chart_title(ax.get_title(), setup), fontsize=16, fontweight="bold", pad=14)

    for index, value in enumerate(y_values):
        if value:
            ax.annotate(f"{value:,.0f}", (index, value), textcoords="offset points", xytext=(0, 7), ha="center", fontsize=9)

    fig.tight_layout()
    return _write_png(fig, output_path)


def _vehicle_columns(frame: pd.DataFrame) -> tuple[str, str, str, str | None]:
    label_column = "ประเภทยานพาหนะ" if "ประเภทยานพาหนะ" in frame.columns else str(frame.columns[0])
    count_column = "จำนวนคัน" if "จำนวนคัน" in frame.columns else str(frame.columns[2] if len(frame.columns) > 2 else frame.columns[-1])
    percent_column = "สัดส่วน (%)" if "สัดส่วน (%)" in frame.columns else str(frame.columns[-1])
    class_column = "vehicle_class" if "vehicle_class" in frame.columns else None
    return label_column, count_column, percent_column, class_column


def vehicle_composition_chart_png(
    vehicle_composition_report: pd.DataFrame,
    output_path: str | Path | None = None,
    setup: Mapping[str, Any] | None = None,
) -> bytes:
    """Create a horizontal bar chart from the Vehicle_Composition_Report table."""

    _configure_fonts()
    chart_data = vehicle_composition_report.copy()

    if chart_data.empty:
        chart_data = pd.DataFrame({"ประเภทยานพาหนะ": [], "จำนวนคัน": [], "vehicle_class": []})

    label_column, count_column, percent_column, class_column = _vehicle_columns(chart_data)
    labels = chart_data[label_column].astype(str).str.strip().str.lower()
    total_mask = labels.isin({"รวม", "total"})
    if class_column:
        total_mask = total_mask | chart_data[class_column].astype(str).str.strip().str.lower().eq("total")

    chart_data = chart_data[~total_mask].copy()
    chart_data[count_column] = _coerce_number(chart_data[count_column])
    chart_data[percent_column] = _coerce_number(chart_data[percent_column])
    chart_data = chart_data.sort_values(percent_column, ascending=True)

    height = max(4.8, 0.38 * len(chart_data) + 1.8)
    fig, ax = plt.subplots(figsize=(10, height))
    colors = ["#A23E48" if index % 2 else "#2E6F95" for index in range(len(chart_data))]
    bars = ax.barh(chart_data[label_column].astype(str), chart_data[percent_column], color=colors, alpha=0.88)
    ax.set_title("สัดส่วนประเภทยานพาหนะ", fontsize=16, fontweight="bold", pad=14)
    ax.set_xlabel("สัดส่วน (%)", fontsize=12)
    ax.set_ylabel("ประเภทยานพาหนะ", fontsize=12)
    ax.grid(axis="x", color="#D9E2E8", linewidth=0.8)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_xlim(left=0)
    ax.set_title(_chart_title(ax.get_title(), setup), fontsize=16, fontweight="bold", pad=14)
    ax.xaxis.set_major_formatter(lambda value, _position: f"{value * 100:.0f}%")

    for bar in bars:
        value = bar.get_width()
        if value:
            ax.annotate(
                f"{value * 100:.2f}%",
                (value, bar.get_y() + bar.get_height() / 2),
                textcoords="offset points",
                xytext=(5, 0),
                va="center",
                fontsize=9,
            )

    fig.tight_layout()
    return _write_png(fig, output_path)


def report_chart_pngs(
    hourly_movement_pcu: pd.DataFrame,
    vehicle_composition_report: pd.DataFrame,
    output_dir: str | Path | None = None,
    setup: Mapping[str, Any] | None = None,
) -> Mapping[str, bytes]:
    """Generate both report charts and optionally write them to an output directory."""

    output_root = Path(output_dir) if output_dir is not None else None
    if output_root is not None:
        output_root.mkdir(parents=True, exist_ok=True)

    return {
        "hourly_pcu": hourly_pcu_chart_png(
            hourly_movement_pcu,
            output_root / CHART_FILENAMES["hourly_pcu"] if output_root is not None else None,
            setup=setup,
        ),
        "vehicle_composition": vehicle_composition_chart_png(
            vehicle_composition_report,
            output_root / CHART_FILENAMES["vehicle_composition"] if output_root is not None else None,
            setup=setup,
        ),
    }
