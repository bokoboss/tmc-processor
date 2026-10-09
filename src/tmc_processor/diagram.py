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

from .movement_scheme import (
    APPROACH_MOVEMENT_CODES,
    approach_direction_label,
    approach_movement_display_label,
    movement_type_label,
    parse_approach_movement_code,
)


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
# V1 from-to diagram table groups: the first character is the source/from leg
# and the second character is the destination/to leg. This is intentionally
# separate from v2 notation, where the first character is travel direction.


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


def _numeric_value(value: Any) -> float:
    return float(pd.to_numeric(pd.Series([value]), errors="coerce").fillna(0).iloc[0])


def _movement_total_lookup(movement_summary: pd.DataFrame | None) -> dict[str, dict[str, Any]]:
    if movement_summary is None or movement_summary.empty or "movement_code" not in movement_summary.columns:
        return {}
    frame = movement_summary.copy()
    frame["movement_code"] = frame["movement_code"].fillna("").astype(str).str.strip()
    aggregations = {}
    if "count" in frame.columns:
        aggregations["count"] = ("count", "sum")
    if "pcu" in frame.columns:
        aggregations["pcu"] = ("pcu", "sum")
    if aggregations:
        frame = frame.groupby("movement_code", dropna=False, as_index=False).agg(**aggregations)
    lookup: dict[str, dict[str, Any]] = {}
    for _, row in frame.iterrows():
        code = str(row.get("movement_code") or "").strip()
        if code:
            lookup[code] = row.to_dict()
    return lookup


def build_v2_movement_diagram_data(
    movement_summary: pd.DataFrame | None = None,
    hourly_movement_pcu: pd.DataFrame | None = None,
    peaks: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Return table-based v2 movement visualization data.

    V2 approach-movement codes use the first letter as travel direction:
    ``NL`` means northbound left, not "from north leg to left". Phase H keeps
    the visualization table-based so it does not borrow v1 from-to arrows or
    physical-approach assumptions.
    """

    total_lookup = _movement_total_lookup(movement_summary)
    hourly = hourly_movement_pcu if hourly_movement_pcu is not None else pd.DataFrame()
    peak_frame = peaks if peaks is not None else pd.DataFrame()
    am_row = _hour_row(hourly, _peak_start(peak_frame, "AM"))
    pm_row = _hour_row(hourly, _peak_start(peak_frame, "PM"))

    rows = []
    for order, code in enumerate(APPROACH_MOVEMENT_CODES, start=1):
        parsed = parse_approach_movement_code(code)
        totals = total_lookup.get(code, {})
        rows.append(
            {
                "movement_code": code,
                "approach_direction": parsed.approach_direction,
                "approach_direction_label": approach_direction_label(parsed.approach_direction),
                "movement_type": parsed.movement_type,
                "movement_type_label": movement_type_label(parsed.movement_type),
                "display_label": approach_movement_display_label(code),
                "diagram_group": approach_direction_label(parsed.approach_direction),
                "position_order": order,
                "total_count": _numeric_value(totals.get("count", 0)),
                "total_pcu": _numeric_value(totals.get("pcu", 0)),
                "am_peak_count": None,
                "am_peak_pcu": _numeric_value(am_row[code]) if am_row is not None and code in am_row else 0,
                "pm_peak_count": None,
                "pm_peak_pcu": _numeric_value(pm_row[code]) if pm_row is not None and code in pm_row else 0,
                "diagram_order": order,
            }
        )
    return pd.DataFrame(
        rows,
        columns=[
            "movement_code",
            "approach_direction",
            "approach_direction_label",
            "movement_type",
            "movement_type_label",
            "display_label",
            "diagram_group",
            "position_order",
            "total_count",
            "total_pcu",
            "am_peak_count",
            "am_peak_pcu",
            "pm_peak_count",
            "pm_peak_pcu",
            "diagram_order",
        ],
    )


def _v2_diagram_frame(diagram_data: pd.DataFrame) -> pd.DataFrame:
    frame = diagram_data.copy() if diagram_data is not None else pd.DataFrame()
    if "movement_code" not in frame.columns:
        frame["movement_code"] = []
    frame["movement_code"] = frame["movement_code"].fillna("").astype(str).str.strip()
    rows = []
    by_code = {str(row.get("movement_code") or "").strip(): row for _, row in frame.iterrows()}
    for order, code in enumerate(APPROACH_MOVEMENT_CODES, start=1):
        source = by_code.get(code, {})
        parsed = parse_approach_movement_code(code)
        rows.append(
            {
                "movement_code": code,
                "approach_direction_label": source.get(
                    "approach_direction_label",
                    approach_direction_label(parsed.approach_direction),
                ),
                "movement_type_label": source.get("movement_type_label", movement_type_label(parsed.movement_type)),
                "total_count": _numeric_value(source.get("total_count", 0)),
                "total_pcu": _numeric_value(source.get("total_pcu", 0)),
                "diagram_order": order,
            }
        )
    return pd.DataFrame(rows)


def _format_v2_metric(value: Any) -> str:
    numeric = _numeric_value(value)
    return f"{numeric:,.0f}"


def render_v2_movement_diagram_png(diagram_data: pd.DataFrame) -> bytes:
    """Render a deterministic table-style PNG for v2 approach-movement data.

    The visual groups movements by travel direction and movement type only. It
    intentionally avoids v1-style turn arrows because v2 codes are not from-to
    geometry and left/right placement needs a separately tested coordinate model.
    """

    _configure_fonts()
    frame = _v2_diagram_frame(diagram_data)
    lookup = {row.movement_code: row for row in frame.itertuples(index=False)}
    groups = [
        ("Northbound", ["NL", "NT", "NR", "NU"]),
        ("Southbound", ["SL", "ST", "SR", "SU"]),
        ("Eastbound", ["EL", "ET", "ER", "EU"]),
        ("Westbound", ["WL", "WT", "WR", "WU"]),
    ]
    movement_short = {
        "Left turn": "Left",
        "Through": "Through",
        "Right turn": "Right",
        "U-turn": "U-turn",
    }

    fig, axes = plt.subplots(2, 2, figsize=(12, 8))
    fig.patch.set_facecolor("white")
    for ax, (title, codes) in zip(axes.flat, groups):
        ax.set_xlim(0, 4)
        ax.set_ylim(0, 1)
        ax.axis("off")
        ax.set_title(title, fontsize=15, fontweight="bold", color="#1F4E78", pad=12)
        for index, code in enumerate(codes):
            row = lookup[code]
            x = index + 0.05
            width = 0.9
            ax.add_patch(
                Rectangle(
                    (x, 0.08),
                    width,
                    0.78,
                    facecolor="#F8FAFC",
                    edgecolor="#8EAADB",
                    linewidth=1.2,
                )
            )
            ax.add_patch(
                Rectangle(
                    (x, 0.70),
                    width,
                    0.16,
                    facecolor="#D9EAF7",
                    edgecolor="#8EAADB",
                    linewidth=1.0,
                )
            )
            label = movement_short.get(str(row.movement_type_label), str(row.movement_type_label))
            ax.text(x + width / 2, 0.78, code, ha="center", va="center", fontsize=13, fontweight="bold", color="#1F4E78")
            ax.text(x + width / 2, 0.58, label, ha="center", va="center", fontsize=10, color="#333333")
            ax.text(
                x + width / 2,
                0.40,
                f"PCU {_format_v2_metric(row.total_pcu)}",
                ha="center",
                va="center",
                fontsize=10,
                fontweight="bold",
                color="#111827",
            )
            ax.text(
                x + width / 2,
                0.24,
                f"Count {_format_v2_metric(row.total_count)}",
                ha="center",
                va="center",
                fontsize=9,
                color="#4B5563",
            )

    fig.suptitle("Approach-Movement Diagram Data", fontsize=18, fontweight="bold", color="#111827", y=0.98)
    fig.text(
        0.5,
        0.025,
        "V2 semantics: N/S/E/W are travel directions. This visual does not use v1 from-to arrows.",
        ha="center",
        va="center",
        fontsize=10,
        color="#4B5563",
    )
    fig.tight_layout(rect=(0.02, 0.05, 0.98, 0.94), h_pad=2.0, w_pad=1.2)
    return _write_png(fig)


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
