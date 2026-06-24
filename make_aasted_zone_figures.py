from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from analyze_aasted_runs_zone_patterns import (
    COMPARISON_CSV,
    COMPARISON_RUN_NAME,
    REFERENCE_CSV,
    REFERENCE_RUN_NAME,
    OUTPUT_DIR,
    TEMP_SENSORS,
    build_seconds,
    detect_pattern_zones,
    read_run,
)


COLORS = {
    "T1": "#7F1D1D",
    "T2": "#DC2626",
    "T3": "#2563EB",
    "T4": "#16A34A",
    "T5": "#F97316",
    "T6": "#7C3AED",
    "T7": "#0891B2",
    "T8": "#111827",
    "T9": "#A16207",
    "acc z": "#000000",
}

ZONE_COLORS = [
    "#EFF6FF",
    "#F0FDF4",
    "#FFF7ED",
    "#FEF2F2",
    "#F8FAFC",
    "#ECFEFF",
    "#F0FDFA",
    "#F7FEE7",
    "#FDF2F8",
    "#F5F3FF",
    "#FAFAFA",
    "#F1F5F9",
]


def nice_ticks(vmin: float, vmax: float, n: int = 6) -> list[float]:
    if not np.isfinite(vmin) or not np.isfinite(vmax) or vmin == vmax:
        return [vmin]
    raw_step = (vmax - vmin) / max(n - 1, 1)
    mag = 10 ** np.floor(np.log10(abs(raw_step)))
    nice = np.array([1, 2, 2.5, 5, 10]) * mag
    step = float(nice[np.argmin(np.abs(nice - raw_step))])
    start = np.floor(vmin / step) * step
    end = np.ceil(vmax / step) * step
    ticks = []
    val = start
    while val <= end + step * 0.5:
        ticks.append(round(float(val), 6))
        val += step
    return ticks


def polyline(points: list[tuple[float, float]], color: str, width: float = 1.4, opacity: float = 1.0) -> str:
    if not points:
        return ""
    d = " ".join(f"{x:.1f},{y:.1f}" for x, y in points if np.isfinite(x) and np.isfinite(y))
    return f'<polyline points="{d}" fill="none" stroke="{color}" stroke-width="{width}" stroke-opacity="{opacity}" />'


def downsample(df: pd.DataFrame, max_points: int = 1800) -> pd.DataFrame:
    if len(df) <= max_points:
        return df
    step = int(np.ceil(len(df) / max_points))
    return df.iloc[::step].copy()


def make_figure(path: Path, run_name: str, csv_path: Path) -> Path:
    df = read_run(csv_path)
    seconds = build_seconds(df)
    zones = detect_pattern_zones(seconds, run_name)

    temp = df.dropna(subset=TEMP_SENSORS, how="all")[["time", "elapsed_s", *TEMP_SENSORS]].copy()
    temp = downsample(temp, 1800)
    motion = seconds[["elapsed_sec", "accz_mean"]].copy()
    motion["time"] = df["time"].min() + motion["elapsed_sec"]
    motion = downsample(motion, 1800)

    x_min = float(df["time"].min())
    x_max = float(df["time"].max())
    temp_min = float(np.nanmin(temp[TEMP_SENSORS].to_numpy()))
    temp_max = float(np.nanmax(temp[TEMP_SENSORS].to_numpy()))
    temp_pad = (temp_max - temp_min) * 0.06
    y_temp_min = temp_min - temp_pad
    y_temp_max = temp_max + temp_pad
    acc_min = float(np.nanquantile(motion["accz_mean"], 0.005))
    acc_max = float(np.nanquantile(motion["accz_mean"], 0.995))
    acc_pad = max((acc_max - acc_min) * 0.10, 0.05)
    y_acc_min = acc_min - acc_pad
    y_acc_max = acc_max + acc_pad

    width, height = 1680, 880
    left, right, top, bottom = 96, 104, 74, 170
    plot_w = width - left - right
    plot_h = height - top - bottom

    def sx(t: float) -> float:
        return left + (t - x_min) / (x_max - x_min) * plot_w

    def sy_temp(v: float) -> float:
        return top + (y_temp_max - v) / (y_temp_max - y_temp_min) * plot_h

    def sy_acc(v: float) -> float:
        return top + (y_acc_max - v) / (y_acc_max - y_acc_min) * plot_h

    parts: list[str] = []
    parts.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">')
    parts.append('<rect width="100%" height="100%" fill="white"/>')
    parts.append(
        f'<text x="{left}" y="34" font-family="Arial" font-size="24" font-weight="700">{run_name}: temperatures and acc z with detected process zones</text>'
    )
    parts.append(
        f'<text x="{left}" y="58" font-family="Arial" font-size="14" fill="#475569">x-axis uses absolute CSV time. Left axis: temperature sensors T1-T9. Right axis: per-second mean acc z.</text>'
    )

    # Zone background.
    for idx, row in enumerate(zones.itertuples(index=False)):
        start_abs = x_min + float(row.detected_start_s)
        end_abs = x_min + float(row.detected_end_s)
        x0 = max(left, sx(start_abs))
        x1 = min(left + plot_w, sx(end_abs))
        color = ZONE_COLORS[idx % len(ZONE_COLORS)]
        parts.append(f'<rect x="{x0:.1f}" y="{top}" width="{max(0, x1 - x0):.1f}" height="{plot_h}" fill="{color}" opacity="0.92"/>')
        if x1 - x0 > 42:
            label = str(row.zone).replace("_", " ")
            parts.append(
                f'<text transform="translate({(x0+x1)/2:.1f},{top+18}) rotate(-28)" text-anchor="middle" font-family="Arial" font-size="12" fill="#334155">{label}</text>'
            )

    # Grid and axes.
    temp_ticks = nice_ticks(y_temp_min, y_temp_max, 7)
    for tick in temp_ticks:
        y = sy_temp(tick)
        parts.append(f'<line x1="{left}" x2="{left+plot_w}" y1="{y:.1f}" y2="{y:.1f}" stroke="#E2E8F0" stroke-width="1"/>')
        parts.append(f'<text x="{left-10}" y="{y+4:.1f}" text-anchor="end" font-family="Arial" font-size="12" fill="#334155">{tick:g}</text>')
    x_ticks = nice_ticks(x_min, x_max, 9)
    for tick in x_ticks:
        x = sx(tick)
        if left <= x <= left + plot_w:
            parts.append(f'<line x1="{x:.1f}" x2="{x:.1f}" y1="{top}" y2="{top+plot_h}" stroke="#F1F5F9" stroke-width="1"/>')
            parts.append(f'<text x="{x:.1f}" y="{top+plot_h+28}" text-anchor="middle" font-family="Arial" font-size="12" fill="#334155">{tick:g}</text>')
    acc_ticks = nice_ticks(y_acc_min, y_acc_max, 6)
    for tick in acc_ticks:
        y = sy_acc(tick)
        parts.append(f'<text x="{left+plot_w+10}" y="{y+4:.1f}" font-family="Arial" font-size="12" fill="#111827">{tick:g}</text>')

    parts.append(f'<rect x="{left}" y="{top}" width="{plot_w}" height="{plot_h}" fill="none" stroke="#334155" stroke-width="1.2"/>')
    parts.append(f'<text x="{left+plot_w/2}" y="{height-82}" text-anchor="middle" font-family="Arial" font-size="15" fill="#111827">absolute time [s]</text>')
    parts.append(f'<text transform="translate(28,{top+plot_h/2}) rotate(-90)" text-anchor="middle" font-family="Arial" font-size="15" fill="#111827">temperature [deg C]</text>')
    parts.append(f'<text transform="translate({width-26},{top+plot_h/2}) rotate(90)" text-anchor="middle" font-family="Arial" font-size="15" fill="#111827">acc z [g]</text>')

    # Lines.
    for sensor in TEMP_SENSORS:
        points = [(sx(float(t)), sy_temp(float(v))) for t, v in zip(temp["time"], temp[sensor]) if pd.notna(v)]
        parts.append(polyline(points, COLORS[sensor], 1.5 if sensor != "T8" else 2.2, 0.92))
    acc_points = [(sx(float(t)), sy_acc(float(v))) for t, v in zip(motion["time"], motion["accz_mean"]) if pd.notna(v)]
    parts.append(polyline(acc_points, COLORS["acc z"], 1.8, 0.72))

    # Legend.
    lx, ly = left, height - 52
    items = [*TEMP_SENSORS, "acc z"]
    for idx, item in enumerate(items):
        x = lx + (idx % 10) * 132
        y = ly + (idx // 10) * 22
        parts.append(f'<line x1="{x}" x2="{x+26}" y1="{y}" y2="{y}" stroke="{COLORS[item]}" stroke-width="{2.4 if item in ["T8", "acc z"] else 1.8}"/>')
        parts.append(f'<text x="{x+34}" y="{y+4}" font-family="Arial" font-size="13" fill="#111827">{item}</text>')

    # Zone boundary table summary.
    table_x, table_y = left + plot_w - 430, top + plot_h + 48
    parts.append(f'<text x="{table_x}" y="{table_y}" font-family="Arial" font-size="13" font-weight="700" fill="#111827">Detected zone starts (absolute s)</text>')
    for i, row in enumerate(zones.itertuples(index=False)):
        if i >= 12:
            break
        label = str(row.zone).replace("_", " ")
        abs_start = x_min + float(row.detected_start_s)
        parts.append(f'<text x="{table_x}" y="{table_y + 18 + i*14}" font-family="Arial" font-size="11" fill="#334155">{abs_start:.0f}: {label}</text>')

    parts.append("</svg>")
    path.write_text("\n".join(parts), encoding="utf-8")
    return path


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ref_path = OUTPUT_DIR / "reference_2291_4160_temperature_accz_zones.svg"
    comparison_path = OUTPUT_DIR / "old_configuration_temperature_accz_zones.svg"
    make_figure(ref_path, REFERENCE_RUN_NAME, REFERENCE_CSV)
    make_figure(comparison_path, COMPARISON_RUN_NAME, COMPARISON_CSV)
    print(ref_path)
    print(comparison_path)


if __name__ == "__main__":
    main()
