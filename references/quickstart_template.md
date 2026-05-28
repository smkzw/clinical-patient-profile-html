# Quickstart Template

给其他同事直接复制用，先做预检查，再正式构建。

## 1. 预检查

把下面命令里的 `<项目文件夹>` 替换成当前项目资料所在目录：

```bash
python3 ~/.codex/skills/clinical-patient-profile-html/scripts/build_patient_profile_html.py \
  --work-dir "<项目文件夹>" \
  --output-dir "<项目文件夹>/patient_profile_output" \
  --precheck-only
```

看输出目录里的：

- `input_precheck.md`
- `suggested_project_config.json`

如果里面有 `阻断` 或 `需确认`，先补文件、确认字段映射，或明确哪些内容本轮不纳入。

## 2. 正式构建

确认完预检查结果后，运行：

```bash
python3 ~/.codex/skills/clinical-patient-profile-html/scripts/build_patient_profile_html.py \
  --work-dir "<项目文件夹>" \
  --config-json "<项目文件夹>/patient_profile_output/suggested_project_config.json" \
  --output-dir "<项目文件夹>/patient_profile_output"
```

主要输出文件：

- `patient_profile.html`
- `cleaned_subject_profile_dataset.csv`
- `efficacy_longitudinal_dataset.csv`
- `lab_longitudinal_dataset.csv`
- `vital_signs_longitudinal_dataset.csv`
- `finding_subject_level_dataset.csv`

## 常见情况

- 缺少研究方案：可以先做 profile，但访视定义、终点定义、窗期解释要后补核对。
- 自动识别到多个中心：先明确本轮纳入哪些中心，再继续构建。
- 某个指标的结果列或日期列没识别准：优先修改 `suggested_project_config.json`，不要先改 Python。
- Finding sheet 找不到受试者字段：先让项目组确认该列，或决定该 sheet 暂不纳入。
