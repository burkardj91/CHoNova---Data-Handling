# Aasted-3 Update Request Checklist

Use this checklist when asking Codex to update or compare Aasted-3 runs. Fill in the labels and tick the outputs you want.

## 1. Run Setup

- [ ] This is an `Aasted-3` analysis.
- [ ] Reference run label: `...`
- [ ] New/comparison run label(s): `...`
- [ ] Should the new run replace the current comparison run?
  - [ ] Yes
  - [ ] No, add it as an additional comparison/repeatability run
- [ ] Should the output be pushed to GitHub?
  - [ ] Yes
  - [ ] No

## 2. Files I Am Uploading

- [ ] Raw CSV file(s) for all new Aasted-3 runs.
- [ ] Reference raw CSV, if the reference changed.
- [ ] `aa3_trials_experimental_setup.xlsx` with run metadata, remarks, and quality scoring.
- [ ] `aa3_trials_experimental_summary.xlsx` or `parameter_summary.xlsx` with:
  - deposition time
  - crystallization onset
  - detachment onset for Rx1Tx1 and/or Rx2Tx2
- [ ] Updated ultrasound temperature-correction coefficients, if T7 correction changed.
- [ ] Extra notes about conveyor speed, cooling configuration, demoulding behavior, or known process anomalies.

## 3. Analyses To Run

- [ ] Reference-based zone comparison.
- [ ] Repeatability analysis across multiple runs.
- [ ] Product temperature by zone.
- [ ] Deposition temperature rows in Product Temp By Zone:
  - `mould_temperature_deposition` from T7, deposition-5 s to deposition
  - `chocolate_temperature_deposition` from T2/T3/T4/T5, deposition+20 s to deposition+25 s
- [ ] Product contour plots and hotspot summary.
- [ ] Mechanical/IMU zone comparison.
- [ ] T7-corrected ultrasound visualization.
- [ ] Detachment change-point and offset analysis.
- [ ] Quality comparison against experimental setup score.

## 4. Outputs To Update

- [ ] Consolidated run folder under `outputs/aasted3_run_comparison/runs/...`.
- [ ] Customer-facing summary workbook.
- [ ] Technical comparison report workbook.
- [ ] Zone visualization figures.
- [ ] US change-point decision figures.
- [ ] Contour figures and contour data workbook.
- [ ] Raw/detail tables workbook.
- [ ] Input archive inside the run folder.
- [ ] Temperature-corrected raw CSV exports.

## 5. Prompt To Paste

Please update the Aasted-3 analysis using the uploaded files. Use the checked options in this checklist, keep the Aasted-3 profile logic, regenerate the requested outputs, organize them into one run-specific folder, and push the result to GitHub if selected.

