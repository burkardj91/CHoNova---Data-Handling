# CHoNova Data Understanding

This workspace separates reusable analysis profiles by measurement setup. The two current workflows are intentionally different because the available signals, mould geometry, and zone definitions are different.

## Folder Layout

| Folder | Purpose |
|---|---|
| `inputs/lab_trials` | Source workbooks for laboratory trials. Raw repetitions are stored as Excel sheets; summary landmarks and experimental setup metadata are stored in separate Excel files. |
| `inputs/AAK_cooling_tunnel` | Filtered subgroup containing only the `a_test` / cooling-tunnel lab trials. Generated from `inputs/lab_trials`. |
| `inputs/aasted2` | Source CSV files, setup/summary workbooks, mould profile, checklist, and mould sketch for Aasted 2 line runs. |
| `inputs/aasted3` | Source CSV files, setup files, summary landmarks, temperature-correction coefficients, and the reusable Aasted-3 update checklist. |
| `configs` | Reusable YAML profiles describing mould geometry, sensors, coordinates, and default analysis logic. |
| `outputs/lab_trials` | Lab-trial repeatability and factor-screening reports. |
| `outputs/AAK_cooling_tunnel` | Cooling-tunnel subgroup reports, including pattern-derived zones, product temperature deltas/hotspots, and ultrasound detachment pattern analysis. |
| `outputs/aasted2_comparison` | Aasted 2 reports, including first Aasted-2 vs Aasted-3 reference comparison and packaged input archive. |
| `outputs/aasted3_run_comparison` | Aasted 3 run-comparison reports with zone figures, contour plots, and irregularity analysis. |

## Lab Trials Profile

Use this profile for laboratory mould trials stored in Excel workbooks.

Expected inputs:

- `lab_trials_raw.xlsx`
- `lab_trials_summary.xlsx`
- `lab_trials_experimental setup.xlsx`

Main logic:

- Uses lab mould geometry from `configs/lab_trials_mould_profile.yaml`.
- Product temperature sensors: `T1`, `T3`, `T4`, `T5`.
- Mould temperature sensor: `T2`.
- Ambient temperature sensor: `T6`.
- Primary ultrasound channels: `Rx1Tx1`, `Rx2Tx2`.
- Summary landmarks such as crystallization onset, detachment onset, and detachment completion come from the summary workbook.

Typical output:

- Screens whether extracted summary parameters vary with main experimental factors such as chocolate recipe, cooling configuration, cooling speed, tempering, and trial group.
- Reports repeatability using CV.
- Only investigates temperature explanations for low-repeatability cases when requested.
- Treats ultrasound sensor codes `11` and `22` separately when pattern repeatability is analyzed.

## AAK Cooling Tunnel Subgroup

Use this subgroup for the `a_test` lab trials from the Marcos/AAK cooling tunnel setup.

Generated inputs:

- `AAK_cooling_tunnel_raw.xlsx`
- `AAK_cooling_tunnel_summary.xlsx`
- `AAK_cooling_tunnel_experimental setup.xlsx`

Current subgroup repetitions:

- `at2455_a_test`
- `at2455_a_test3`
- `at2455_a_test4`

Main logic:

- Uses the same lab mould geometry as `lab_trials`.
- Applies a pattern-based cooling-zone split from the user-defined example `at2455_a_test`.
- Uses `T6` ambient temperature and product/mould temperature behavior to identify cooling-zone transitions.
- Calculates product temperature delta and hotspot summaries by detected zone.
- Analyzes ultrasound detachment behavior inside the cooling/detachment window, not the demoulding zone.

Typical output:

- Detected zone boundaries for each repetition.
- Per-zone product temperature distribution summaries.
- Hotspot sensor summaries using only product sensors.
- Ultrasound change-point/detachment checks for `Rx1Tx1` and `Rx2Tx2`.
- Figures with temperature, ambient, ultrasound, and labelled change points.

## Aasted 3 Profile

Use this profile for Aasted 3 line trials stored as CSV files.

Expected inputs:

- Two or more CSV files from the Aasted 3 line.
- A reference run can be supplied to define normal zone timing and pattern behavior.
- Use `inputs/aasted3/aasted3_update_request_checklist.md` when requesting an update and you want to tick which inputs/outputs should be refreshed.

Main logic:

- Uses Aasted 3 mould geometry from `configs/aasted3_mould_profile.yaml`.
- Treats `T1` as RH, not product temperature.
- Product contour sensors: `T2`, `T3`, `T4`, `T5`, `T7`.
- Uses `T8`, `acc z`, and `gyro y` to derive or align process zones.
- Includes ultrasound overlays for comparison, but ultrasound coordinates do not influence temperature contours.

Typical output:

- Zone-labelled reference and irregular-run figures.
- Stop/irregularity detection from IMU and timing behavior.
- Product-only contour plots with shared temperature scales per zone.
- Product delta and hotspot summaries by zone.
- Product Temp By Zone also includes deposition temperature pseudo-zone rows for mould temperature at deposition and chocolate temperature after deposition when deposition landmarks are available.

## Aasted 2 Profile

Use this profile for Aasted 2 line trials stored as CSV files.

Expected inputs:

- `260604_aasted2_Run 1_1.csv` or another chocolate-filled Aasted-2 raw CSV.
- Optional empty-run CSV for context.
- `aa2_trials_experimental_setup.xlsx`.
- `aa2_trials_experimental_summary.xlsx`.
- `aasted2_mould_profile.yaml`.
- `Form_Aa2.png` or another mould configuration sketch.

Main logic:

- Uses Aasted 2 mould geometry and sensor coordinates from `configs/aasted2_mould_profile.yaml`.
- Treats `T1` as humidity, not product temperature; AA2 product temperature uses `T3`, `T4`, and `T6`.
- Uses the current user-defined AA2 zones until pattern-derived AA2 zone detection is added.
- Uses manual detachment onset/offset from the AA2 summary until ultrasound temperature correction is available.
- Compares AA2 against the Aasted-3 reference with emphasis on cooling temperature variation, product/mould/ambient spread, vibration intensity, twisting intensity, viscosity proxy, and detachment timing.

Typical output:

- One consolidated run folder under `outputs/aasted2_comparison/runs/...`.
- A technical comparison report workbook.
- A zone overview figure.
- An archived copy of the AA2 inputs plus the Aasted-3 reference inputs used for comparison.

## Practical Use

When adding new data, first state which profile applies:

- `lab_trials`
- `AAK_cooling_tunnel`
- `Aasted 2`
- `Aasted 3`

Then place or upload raw data, summary data, and experimental setup metadata into the matching input folder. The report should be generated into the matching `outputs` folder.
