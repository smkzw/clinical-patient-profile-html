# Validation Checklist

Run this after building and before delivering the HTML.

## Input consistency

- Confirm the selected center list matches the intended scope.
- Confirm the main listing workbook is the intended source workbook.
- Confirm the finding workbook and sheet mappings match the intended centers.

## Row-count checks

Compare output CSV counts against the source-driven expectation:

- subjects
- efficacy rows
- lab rows
- vital-sign rows
- finding rows

Large unexpected drops usually mean a center filter, sheet-name mismatch, or value-column mismatch.

## Subject-level spot checks

Check at least a few subjects across different scenarios:

- randomized active subject
- randomized control subject
- screen-failure subject
- subject with findings
- subject with abnormal labs

For each spot check, compare HTML against the source workbook on:

- sex
- age group
- treatment group
- visit labels
- efficacy values
- lab values and reference ranges
- vital-sign values
- linked finding IDs

## Response logic checks

Verify that response flags are only assigned where the project definition supports them.

Examples in the current RUX implementation:

- EASI
- IGA
- Itch NRS
- PROMIS 8a / 8b

Do not assume the same response rules for a new project without confirming the endpoint definitions.

## Final decision rule

If the source workbook and HTML disagree, treat the source workbook as the current truth until the mapping or build logic is corrected.
