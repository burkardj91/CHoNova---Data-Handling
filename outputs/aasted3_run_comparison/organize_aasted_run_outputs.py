from __future__ import annotations

import shutil
from pathlib import Path

from analyze_aasted_runs_zone_patterns import (
    COMPARISON_CSV,
    COMPARISON_FIGURE_PREFIX,
    COMPARISON_RUN_NAME,
    FINAL_REPORT_WORKBOOK,
    OUTPUT_DIR,
    REFERENCE_FIGURE_PREFIX,
    RUN_OUTPUT_DIR,
    TEMPERATURE_CORRECTION_CSV,
)
from build_aasted_summary_outputs import (
    CONTOUR_DATA_WORKBOOK,
    RAW_TABLES_WORKBOOK,
    SUMMARY_WORKBOOK,
    ZONE_FINDINGS_WORKBOOK,
)


def copy_if_exists(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def copy_many(pattern: str, dst_dir: Path) -> None:
    dst_dir.mkdir(parents=True, exist_ok=True)
    for src in sorted(OUTPUT_DIR.glob(pattern)):
        if src.is_file():
            shutil.copy2(src, dst_dir / src.name)


def organize_outputs() -> Path:
    reports = RUN_OUTPUT_DIR / "reports"
    figures = RUN_OUTPUT_DIR / "figures"
    contours = figures / "contours"
    inputs = RUN_OUTPUT_DIR / "inputs"
    corrected = RUN_OUTPUT_DIR / "temperature_corrected_raw"

    copy_if_exists(FINAL_REPORT_WORKBOOK, reports / FINAL_REPORT_WORKBOOK.name)
    copy_if_exists(SUMMARY_WORKBOOK, reports / SUMMARY_WORKBOOK.name)
    copy_if_exists(ZONE_FINDINGS_WORKBOOK, reports / ZONE_FINDINGS_WORKBOOK.name)
    copy_if_exists(CONTOUR_DATA_WORKBOOK, reports / CONTOUR_DATA_WORKBOOK.name)
    copy_if_exists(RAW_TABLES_WORKBOOK, reports / RAW_TABLES_WORKBOOK.name)

    copy_many(f"{REFERENCE_FIGURE_PREFIX}_temperature_accz_zones.*", figures / "zone_overview")
    copy_many(f"{COMPARISON_FIGURE_PREFIX}_temperature_accz_zones.*", figures / "zone_overview")

    us_dir = OUTPUT_DIR / "us_change_point_figures"
    if us_dir.exists():
        for src in sorted(us_dir.glob("*.png")):
            copy_if_exists(src, figures / "us_change_points" / src.name)

    contour_dir = OUTPUT_DIR / "contour_images"
    if contour_dir.exists():
        for src in sorted(contour_dir.glob(f"{COMPARISON_RUN_NAME}_*.png")):
            copy_if_exists(src, contours / "comparison" / src.name)
        for src in sorted(contour_dir.glob("reference_2291-4160_*.png")):
            copy_if_exists(src, contours / "reference" / src.name)

    copy_if_exists(COMPARISON_CSV, inputs / COMPARISON_CSV.name)
    copy_if_exists(TEMPERATURE_CORRECTION_CSV, inputs / TEMPERATURE_CORRECTION_CSV.name)
    for name in ["aa3_trials_experimental_setup.xlsx", "aa3_trials_experimental_summary.xlsx"]:
        copy_if_exists(COMPARISON_CSV.parent / name, inputs / name)
    copy_if_exists(COMPARISON_CSV.parent / "aasted3_update_request_checklist.md", inputs / "aasted3_update_request_checklist.md")

    corrected_dir = COMPARISON_CSV.parent / "temperature_corrected_raw"
    if corrected_dir.exists():
        for src in sorted(corrected_dir.glob("*.csv")):
            copy_if_exists(src, corrected / src.name)

    readme = RUN_OUTPUT_DIR / "README.txt"
    readme.write_text(
        "\n".join(
            [
                f"Aasted-3 consolidated output folder for comparison run: {COMPARISON_RUN_NAME}",
                "",
                "reports/",
                "  Customer summary, technical comparison workbook, zone findings, contour data, and raw detail tables.",
                "figures/zone_overview/",
                "  Reference and comparison overview figures with zones, landmarks, and T-corrected US.",
                "figures/us_change_points/",
                "  Per-run/per-channel change-point decision figures. Search domain is detachment onset to run end.",
                "figures/contours/",
                "  Product temperature contour figures by detected zone.",
                "inputs/",
                "  Input protocol, experimental summary, comparison raw CSV, and temperature-correction coefficients.",
                "temperature_corrected_raw/",
                "  Raw files with T7-normalized ultrasound columns (*_tc).",
            ]
        ),
        encoding="utf-8",
    )
    return RUN_OUTPUT_DIR


def main() -> None:
    print(organize_outputs())


if __name__ == "__main__":
    main()
