"""Static four-leg TMC movement diagram generation."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from io import BytesIO
from pathlib import Path
from typing import Any, Mapping

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, Polygon, Rectangle
import pandas as pd


MOVEMENT_CODES = [
    "NE",
    "NS",
    "NW",
    "NU",
    "SW",
    "SN",
    "SE",
    "SU",
    "WN",
    "WE",
    "WS",
    "WU",
    "ES",
    "EW",
    "EN",
    "EU",
]

APPROACH_TABLES = {
    "N": ["NU", "NW", "NS", "NE"],
    "S": ["SW", "SN", "SE", "SU"],
    "W": ["WN", "WE", "WS"],
    "E": ["EN", "EW", "ES"],
}


@dataclass(frozen=True)
class DiagramConfig:
    tmc_id: str = ""
    tmc_name: str = ""
    survey_date_text: str = ""
    north_label: str = ""
    south_label: str = ""
    east_label: str = ""
    west_label: str = ""
    north_road: str = ""
    south_road: str = ""
    east_road: str = ""
    west_road: str = ""
    survey_period_text: str = "07.00-19.00 น."
    caption_text: str = ""
    show_u_turn: bool = True


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


def _time_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        value = value.time()
    if isinstance(value, time):
        return value.strftime("%H:%M")
    text = str(value).strip()
    return text[:5] if len(text) >= 5 and text[2:3] == ":" else text


def _coerce_int(value: Any) -> int:
    return int(round(pd.to_numeric(pd.Series([value]), errors="coerce").fillna(0).iloc[0]))


def _total_row(hourly_movement_pcu: pd.DataFrame) -> pd.Series | None:
    if hourly_movement_pcu.empty:
        return None
    label_column = _first_column(hourly_movement_pcu)
    if label_column is None:
        return None
    labels = hourly_movement_pcu[label_column].astype(str).str.strip().str.lower()
    matches = hourly_movement_pcu[labels.isin({"รวม", "total"})]
    return matches.iloc[-1] if not matches.empty else None


def _hour_row(hourly_movement_pcu: pd.DataFrame, start_text: str) -> pd.Series | None:
    if hourly_movement_pcu.empty or not start_text:
        return None
    label_column = _first_column(hourly_movement_pcu)
    if label_column is None:
        return None
    labels = hourly_movement_pcu[label_column].astype(str).str.strip()
    matches = hourly_movement_pcu[labels.str.startswith(start_text)]
    return matches.iloc[0] if not matches.empty else None


def _peak_start(peaks: pd.DataFrame, period: str) -> str:
    if peaks.empty or "period" not in peaks.columns or "peak_start" not in peaks.columns:
        return ""
    matches = peaks[peaks["period"].astype(str).str.upper() == period]
    return _time_text(matches.iloc[0]["peak_start"]) if not matches.empty else ""


def _peak_total(peaks: pd.DataFrame, period: str) -> int:
    if peaks.empty or "period" not in peaks.columns or "hourly_pcu" not in peaks.columns:
        return 0
    matches = peaks[peaks["period"].astype(str).str.upper() == period]
    return _coerce_int(matches.iloc[0]["hourly_pcu"]) if not matches.empty else 0


def _movement_value(row: pd.Series | None, movement: str) -> int:
    if row is None or movement not in row:
        return 0
    return _coerce_int(row[movement])


def movement_diagram_values(hourly_movement_pcu: pd.DataFrame, peaks: pd.DataFrame) -> dict[str, dict[str, int]]:
    """Return Total/PM/AM PCU values for every standard movement code."""

    total = _total_row(hourly_movement_pcu)
    am = _hour_row(hourly_movement_pcu, _peak_start(peaks, "AM"))
    pm = _hour_row(hourly_movement_pcu, _peak_start(peaks, "PM"))
    return {
        movement: {
            "total": _movement_value(total, movement),
            "pm": _movement_value(pm, movement),
            "am": _movement_value(am, movement),
        }
        for movement in MOVEMENT_CODES
    }


def _summary_values(hourly_movement_pcu: pd.DataFrame, peaks: pd.DataFrame) -> dict[str, int]:
    total = _total_row(hourly_movement_pcu)
    return {
        "am": _peak_total(peaks, "AM"),
        "pm": _peak_total(peaks, "PM"),
        "total": _movement_value(total, "Total"),
    }


def _config_from_mapping(config: DiagramConfig | Mapping[str, Any] | None) -> DiagramConfig:
    if config is None:
        return DiagramConfig()
    if isinstance(config, DiagramConfig):
        return config
    fields = DiagramConfig.__dataclass_fields__
    return DiagramConfig(**{key: value for key, value in config.items() if key in fields})


def _draw_road(ax) -> None:
    ax.add_patch(Rectangle((42, 0), 16, 100, facecolor="white", edgecolor="none", zorder=0))
    ax.add_patch(Rectangle((0, 42), 100, 16, facecolor="white", edgecolor="none", zorder=0))
    for x in (42, 58):
        ax.plot([x, x], [0, 42], color="black", linewidth=4, solid_capstyle="butt")
        ax.plot([x, x], [58, 100], color="black", linewidth=4, solid_capstyle="butt")
    for y in (42, 58):
        ax.plot([0, 42], [y, y], color="black", linewidth=4, solid_capstyle="butt")
        ax.plot([58, 100], [y, y], color="black", linewidth=4, solid_capstyle="butt")
    ax.plot([42, 58, 58, 42, 42], [42, 42, 58, 58, 42], color="black", linewidth=3)
    ax.plot([50, 50], [6, 38], color="#777777", linewidth=1.2, linestyle="--")
    ax.plot([50, 50], [62, 94], color="#777777", linewidth=1.2, linestyle="--")
    ax.plot([6, 38], [50, 50], color="#777777", linewidth=1.2, linestyle="--")
    ax.plot([62, 94], [50, 50], color="#777777", linewidth=1.2, linestyle="--")


def _draw_arrow(ax, start: tuple[float, float], end: tuple[float, float], color: str, width: float = 2.0) -> None:
    ax.add_patch(
        FancyArrowPatch(
            start,
            end,
            arrowstyle="-|>",
            mutation_scale=14,
            linewidth=width,
            color=color,
            zorder=5,
        )
    )


def _draw_direction_labels(ax, config: DiagramConfig) -> None:
    yellow = "#F0C419"
    _draw_arrow(ax, (52, 93), (52, 82), yellow, 3.0)
    _draw_arrow(ax, (48, 7), (48, 18), yellow, 3.0)
    _draw_arrow(ax, (7, 52), (18, 52), yellow, 3.0)
    _draw_arrow(ax, (93, 48), (82, 48), yellow, 3.0)
    ax.text(54, 87, config.north_label, fontsize=10, va="center", ha="left", color="#333333")
    ax.text(46, 13, config.south_label, fontsize=10, va="center", ha="right", color="#333333")
    ax.text(13, 54, config.west_label, fontsize=10, va="bottom", ha="center", color="#333333")
    ax.text(87, 46, config.east_label, fontsize=10, va="top", ha="center", color="#333333")
    ax.text(60, 75, config.north_road, fontsize=10, va="center", ha="left", rotation=90, color="#333333")
    ax.text(40, 25, config.south_road, fontsize=10, va="center", ha="right", rotation=90, color="#333333")
    ax.text(25, 39, config.west_road, fontsize=10, va="top", ha="center", color="#333333")
    ax.text(75, 61, config.east_road, fontsize=10, va="bottom", ha="center", color="#333333")


def _draw_compass(ax) -> None:
    ax.text(88, 91, "N", ha="center", va="bottom", fontsize=13, fontweight="bold")
    ax.add_patch(Polygon([(88, 89), (85.8, 83), (88, 84.3), (90.2, 83)], closed=True, facecolor="black"))
    ax.plot([88, 88], [78, 83.8], color="black", linewidth=1.5)


def _draw_title_and_summary(
    ax,
    config: DiagramConfig,
    summary: Mapping[str, int],
) -> None:
    title = " ".join(part for part in [config.tmc_id, config.tmc_name] if str(part).strip())
    ax.add_patch(Rectangle((2, 88), 35, 10, facecolor="white", edgecolor="black", linewidth=1.2))
    ax.text(3.5, 95.5, title or "TMC movement diagram", fontsize=12, fontweight="bold", ha="left", va="center")
    ax.text(3.5, 91.5, f"Survey date: {config.survey_date_text}", fontsize=10, ha="left", va="center")

    ax.add_patch(Rectangle((2, 2), 26, 12, facecolor="#F8FAFC", edgecolor="black", linewidth=1.0))
    rows = [
        ("AM peak PCU", summary["am"]),
        ("PM peak PCU", summary["pm"]),
        ("Total PCU", summary["total"]),
    ]
    for index, (label, value) in enumerate(rows):
        y = 11 - index * 3.5
        ax.text(4, y, label, fontsize=9.2, ha="left", va="center")
        ax.text(26, y, f"{value:,.0f}", fontsize=9.2, ha="right", va="center", fontweight="bold")


def _draw_movement_table(
    ax,
    x: float,
    y: float,
    movements: list[str],
    values: Mapping[str, Mapping[str, int]],
    title: str,
) -> None:
    cell_w = 7.0
    row_h = 3.6
    label_w = 16.0
    width = label_w + cell_w * len(movements)
    height = row_h * 4 + 4.0
    ax.add_patch(Rectangle((x, y), width, height, facecolor="white", edgecolor="black", linewidth=1.1, zorder=4))
    ax.add_patch(Rectangle((x, y + height - 4.0), width, 4.0, facecolor="#E9EEF3", edgecolor="black", linewidth=1.0, zorder=4))
    ax.text(x + width / 2, y + height - 2.0, title, fontsize=9, fontweight="bold", ha="center", va="center", zorder=6)

    top = y + height - 4.0
    headers = ["", *movements]
    row_labels = [("Total 12-hour", "total"), ("PM peak hour", "pm"), ("AM peak hour", "am")]
    for column_index, header in enumerate(headers):
        left = x if column_index == 0 else x + label_w + (column_index - 1) * cell_w
        width_here = label_w if column_index == 0 else cell_w
        ax.add_patch(Rectangle((left, top - row_h), width_here, row_h, facecolor="#F8FAFC", edgecolor="#9AA8B4", linewidth=0.7, zorder=4))
        ax.text(left + width_here / 2, top - row_h / 2, header, fontsize=8.2, fontweight="bold", ha="center", va="center", zorder=6)
    for row_index, (label, key) in enumerate(row_labels):
        bottom = top - row_h * (row_index + 2)
        ax.add_patch(Rectangle((x, bottom), label_w, row_h, facecolor="#F8FAFC", edgecolor="#9AA8B4", linewidth=0.7, zorder=4))
        ax.text(x + 0.7, bottom + row_h / 2, label, fontsize=7.8, ha="left", va="center", zorder=6)
        for column_index, movement in enumerate(movements):
            left = x + label_w + column_index * cell_w
            ax.add_patch(Rectangle((left, bottom), cell_w, row_h, facecolor="white", edgecolor="#9AA8B4", linewidth=0.7, zorder=4))
            ax.text(
                left + cell_w - 0.5,
                bottom + row_h / 2,
                f"{values[movement][key]:,.0f}",
                fontsize=7.6,
                ha="right",
                va="center",
                zorder=6,
            )


def _draw_red_movement_arrows(ax) -> None:
    red = "#D62828"
    for start, end in [
        ((43, 71), (49, 61)),
        ((49, 72), (49, 61)),
        ((55, 71), (51, 61)),
        ((57, 29), (51, 39)),
        ((51, 28), (51, 39)),
        ((45, 29), (49, 39)),
        ((29, 57), (39, 51)),
        ((28, 51), (39, 51)),
        ((29, 45), (39, 49)),
        ((71, 43), (61, 49)),
        ((72, 49), (61, 49)),
        ((71, 55), (61, 51)),
    ]:
        _draw_arrow(ax, start, end, red, 1.8)


def generate_four_leg_tmc_diagram(
    hourly_movement_pcu: pd.DataFrame,
    peaks: pd.DataFrame,
    config: DiagramConfig | Mapping[str, Any] | None = None,
    output_path: str | Path | None = None,
) -> bytes:
    """Generate a report-ready standard four-leg TMC movement diagram PNG."""

    _configure_fonts()
    cfg = _config_from_mapping(config)
    values = movement_diagram_values(hourly_movement_pcu, peaks)
    summary = _summary_values(hourly_movement_pcu, peaks)

    fig, ax = plt.subplots(figsize=(11, 8.5))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.set_aspect("equal")
    ax.axis("off")

    _draw_road(ax)
    _draw_direction_labels(ax, cfg)
    _draw_compass(ax)
    _draw_title_and_summary(ax, cfg, summary)

    tables = {
        "North approach": (28, 73, APPROACH_TABLES["N"]),
        "South approach": (28, 7, APPROACH_TABLES["S"]),
        "West approach": (2, 61, APPROACH_TABLES["W"]),
        "East approach": (62, 61, APPROACH_TABLES["E"]),
    }
    for title, (x, y, movements) in tables.items():
        if not cfg.show_u_turn:
            movements = [movement for movement in movements if not movement.endswith("U")]
        _draw_movement_table(ax, x, y, movements, values, title)

    _draw_red_movement_arrows(ax)

    caption = cfg.caption_text or (
        f"ปริมาณจราจรบนทางแยก {cfg.survey_date_text} PCU/12 ชม. ({cfg.survey_period_text})"
    )
    ax.text(50, -2.5, caption, ha="center", va="top", fontsize=12)

    fig.tight_layout(pad=0.6)
    return _write_png(fig, output_path)
