---
name: clinical-patient-profile-html
description: Trigger this skill when the user says “制作patient profile”, “patient profile”, “制作受试者画像”, “制作个例profile”, or asks to build subject-level or center-level clinical patient profile HTML from uploaded listings, protocol files, and finding trackers. The skill should automatically inspect the uploaded files, determine whether they are sufficient, ask only for blocking or ambiguous mappings, and otherwise proceed directly to generate an offline interactive Chinese HTML patient profile plus supporting CSV outputs.
---

# Clinical Patient Profile HTML

Use this skill for clinical-study patient profile HTML work where the output must stay source-grounded and inspection-ready.

This skill is best when the user provides subject listings, finding trackers, and optionally protocol/amendment files, and wants an offline HTML that combines subject profile, efficacy trends, lab trends, vital signs, and finding review.

If the user explicitly says `制作patient profile` or an equivalent phrase after uploading files, invoke this skill immediately and start with file inspection rather than waiting for extra instructions.
Do not require an additional “开始构建” instruction.

## Workflow

1. Run the bundled precheck as soon as the uploaded files are available:

```bash
python3 scripts/build_patient_profile_html.py \
  --work-dir "<project-folder>" \
  --output-dir "<project-folder>/patient_profile_output" \
  --precheck-only
```

2. Read `input_precheck.md`.
Also inspect `suggested_project_config.json`.

3. If precheck shows only `提示`, continue directly to full build.

4. If precheck shows `阻断`, ask only the smallest necessary question.
Do not guess missing centers, finding sheet mappings, subject ID columns, or substitute sheet names.
If the user confirms that some metrics, sheets, or fields do not need to be included, treat that as authorization to proceed without them.

5. If precheck shows `需确认` but the user has already made a scope decision such as “该字段不纳入” or “该数据不体现”, update the config accordingly and continue directly.

6. Once inputs are sufficient or the user has accepted the reduced scope, run the full build immediately. Prefer using the generated config file so the build uses the same detected mappings:

```bash
python3 scripts/build_patient_profile_html.py \
  --work-dir "<project-folder>" \
  --config-json "<project-folder>/patient_profile_output/suggested_project_config.json" \
  --output-dir "<project-folder>/patient_profile_output"
```

7. Validate the result against the source tables before presenting it.
Use `references/validation_checklist.md`.

## Required Inputs

- A main listing workbook with subject-level tabs for demographics, randomization, visits, efficacy, labs, AE, and vital signs.
- A finding workbook or self-inspection/audit tracker.

## Optional Inputs

- Protocol, amendment, or errata files for visit-window and endpoint-definition confirmation.
- A subject report workbook for cross-checking sex, age group, and treatment assignment.
- Reference HTML files if the project already has an approved patient-profile style to follow.

## When To Ask The User

Ask the user instead of proceeding when any of these happen:

- More than one center is detected and the target center scope is not explicitly stated.
- The finding workbook exists but the relevant sheet or subject column cannot be confirmed.
- Key efficacy sheets are missing or likely renamed.
- Lab reference ranges are absent and the user must decide whether to keep those rows with blank normal-range display.
- A sheet appears to contain needed data but the field mapping is ambiguous.

Do not ask the user to “confirm starting the build” when uploaded files and scope are already sufficient.

Use the prompt patterns in `references/mapping_and_decisions.md`.

## Adaptation Rules

This skill ships with a working builder that already reproduces the current RUX patient-profile style and logic closely. For a structurally similar project, prefer changing configuration and mapping before changing rendering.

Prefer updating `suggested_project_config.json` before editing Python.

Adjust the script only in these places when the new project differs beyond config:

- `EFFICACY_CONFIG` for efficacy sheet names and value columns
- `LAB_CONFIG` for lab workbook tabs
- `CORE_SHEET_CATALOG` for nonstandard main listing tabs
- `FINDING_SHEET_SPECS` or finding-sheet auto-detection behavior
- group and center normalization helpers

Do not fabricate subject dates, ranges, response flags, or normal limits.

## Outputs

The builder writes at least these files to the chosen output directory:

- `patient_profile.html`
- `cleaned_subject_profile_dataset.csv`
- `efficacy_longitudinal_dataset.csv`
- `lab_longitudinal_dataset.csv`
- `vital_signs_longitudinal_dataset.csv`
- `finding_subject_level_dataset.csv`
- `data_source_field_mapping.csv`
- `data_qc_summary.csv`
- `data_issue_log.csv`
- `subject_data_completeness.csv`
- `unresolved_data_questions.md`
- `input_precheck.md`
- `input_precheck.json`

The output HTML must stay Chinese-first in title, labels, filters, tables, and review text, consistent with the prior RUX patient-profile style.

## References

- For missing-field questions and exclusion decisions, read `references/mapping_and_decisions.md`.
- For final source-consistency checks, read `references/validation_checklist.md`.
- For a copy-paste quick start, read `references/quickstart_template.md`.
