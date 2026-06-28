# Aasted-2 Update Request Checklist

Use this checklist when asking Codex to update Aasted-2 line analyses.

## 1. Run Setup

- [ ] This is an `Aasted-2` analysis.
- [ ] Reference run label, if any: `...`
- [ ] New/comparison run label(s): `...`
- [ ] Chocolate-filled run(s) uploaded.
- [ ] Empty run(s) uploaded, if needed for baseline/context.
- [ ] Should the result be pushed to GitHub?
  - [ ] Yes
  - [ ] No

## 2. Files I Am Uploading

- [ ] Raw Aasted-2 CSV file(s).
- [ ] `aa2_trials_experimental_setup.xlsx` with protocol metadata and quality information.
- [ ] `aa2_trials_experimental_summary.xlsx` with deposition, crystallization onset, detachment onset, and detachment offset.
- [ ] Mould sketch/configuration image, if changed.
- [ ] Updated sensor coordinates or sensor group assignments, if changed.
- [ ] Temperature-correction coefficients for ultrasound, if available.
- [ ] Notes on zone definitions, conveyor speed, cooling configuration, or known anomalies.

## 3. Analyses To Run

- [ ] Aasted-2 zone visualization.
- [ ] Aasted-2 vs Aasted-3 reference comparison.
- [ ] Product temperature spread until detachment onset.
- [ ] Cooling temperature variation summary.
- [ ] Vibration intensity comparison.
- [ ] Twisting intensity comparison.
- [ ] Raw-ultrasound viscosity proxy.
- [ ] Manual detachment onset/offset summary.
- [ ] Future T-corrected US change-point detachment analysis, once coefficients are available.

## 4. Outputs To Update

- [ ] Consolidated run folder under `outputs/aasted2_comparison/runs/...`.
- [ ] Technical comparison report workbook.
- [ ] Zone overview figures.
- [ ] Input archive inside the run folder.
- [ ] Raw/detail tables.
- [ ] README/checklist copy.

## 5. Prompt To Paste

Please update the Aasted-2 analysis using the uploaded files. Use the checked options in this checklist, keep the Aasted-2 profile logic, regenerate the requested outputs, organize them into one run-specific folder, and push the result to GitHub if selected.

