# Mapping And Decision Rules

Use this file when precheck shows `阻断` or `需确认`.

## Non-negotiable rules

- Use subject-level listings and finding trackers as the primary source.
- Use protocol/amendment files only for design, visit, endpoint, or window confirmation.
- If a field cannot be mapped confidently, mark it as unresolved and ask the user.
- If a sheet should be excluded, get explicit user confirmation.
- Keep the generated `suggested_project_config.json` as the first place to revise mappings. Do not edit Python unless the project structure falls outside the config model.

## Questions to ask the user

Use short, concrete questions.

### Center scope

`识别到多个中心。请确认是纳入全部中心，还是仅纳入以下中心：<center list>？`

### Main listing workbook

`当前目录识别到多个可能的主listing文件。请确认以哪一个作为主数据源：<file list>？`

### Finding sheet mapping

`Finding工作簿已识别到 <sheet name>，但无法确认受试者字段/中心归属。请确认该sheet对应中心，以及应使用哪一列作为受试者编号。`

### Missing efficacy sheet

`未识别到 <metric> 的明确sheet。请确认该项目没有该指标，还是sheet名称不同；如名称不同，请提供正确sheet名。`

### Optional exclusion

`以下sheet存在但字段不足以稳定映射：<sheet list>。请确认是补充字段说明后纳入，还是本轮先排除。`

## Configuration blocks to review first

If the project structure differs, inspect these blocks before changing rendering logic:

- `EFFICACY_CONFIG`
- `LAB_CONFIG`
- `FINDING_SHEET_SPECS`
- `FIELD_ALIASES`
- `normalize_study_group`
- `standardize_center_name`

## Typical exclusion decisions

These can be excluded only after user confirmation:

- empty duplicate exports
- site-level admin sheets unrelated to subject data
- malformed finding tabs without subject linkage
- efficacy/pro tables that only duplicate already preferred total-score tabs
