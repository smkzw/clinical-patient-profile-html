from __future__ import annotations

import csv
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path("/Users/smkzw/Documents/康哲项目资料/Ruxolitinib-AD/CFDI Inspection/准备阶段")
SCRIPT_PATH = PROJECT_ROOT / "skills" / "clinical-patient-profile-html" / "scripts" / "build_patient_profile_html.py"
LISTING_XLSX = PROJECT_ROOT / "RUX-03-002_列表_数据集_Excel_20250612_处理后_南方医科大学皮肤病医院_杭州市第一人民医院.xlsx"
FINDING_XLSX = PROJECT_ROOT / "自查稽查问题汇总-260511.xlsx"
PD_DEF_XLSX = PROJECT_ROOT / "附录2：RUX-03-002_方案偏离分类界定表_V1.0_20240528_clean.xlsx"
SUBJECT_REPORT_XLS = PROJECT_ROOT / "RUX-03-002_受试者报表_20250613.xls"
CENTERS = "南方医科大学皮肤病医院；杭州市第一人民医院"


class BuildPatientProfileHtmlTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = Path(tempfile.mkdtemp(prefix="patient-profile-skill-", dir="/private/tmp"))

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def run_builder(self, *extra_args: str) -> None:
        cmd = [
            "python3",
            str(SCRIPT_PATH),
            "--work-dir",
            str(PROJECT_ROOT),
            "--listing-xlsx",
            str(LISTING_XLSX),
            "--finding-xlsx",
            str(FINDING_XLSX),
            "--pd-def-xlsx",
            str(PD_DEF_XLSX),
            "--subject-report-xls",
            str(SUBJECT_REPORT_XLS),
            "--centers",
            CENTERS,
            "--output-dir",
            str(self.temp_dir),
            *extra_args,
        ]
        subprocess.run(cmd, check=True)

    def count_rows(self, file_name: str) -> int:
        with (self.temp_dir / file_name).open(encoding="utf-8-sig", newline="") as handle:
            return max(sum(1 for _ in csv.reader(handle)) - 1, 0)

    def test_precheck_only_writes_reports(self) -> None:
        self.run_builder("--precheck-only")
        self.assertTrue((self.temp_dir / "input_precheck.md").exists())
        self.assertTrue((self.temp_dir / "input_precheck.json").exists())
        self.assertTrue((self.temp_dir / "suggested_project_config.json").exists())
        precheck = (self.temp_dir / "input_precheck.md").read_text(encoding="utf-8")
        self.assertIn("可继续构建：是", precheck)
        self.assertIn("南方医科大学皮肤病医院；杭州市第一人民医院", precheck)

    def test_full_build_writes_expected_outputs(self) -> None:
        self.run_builder()
        for name in [
            "patient_profile.html",
            "cleaned_subject_profile_dataset.csv",
            "efficacy_longitudinal_dataset.csv",
            "lab_longitudinal_dataset.csv",
            "vital_signs_longitudinal_dataset.csv",
            "finding_subject_level_dataset.csv",
        ]:
            self.assertTrue((self.temp_dir / name).exists(), name)
        self.assertGreater(self.count_rows("cleaned_subject_profile_dataset.csv"), 0)
        self.assertGreater(self.count_rows("efficacy_longitudinal_dataset.csv"), 0)
        self.assertGreater(self.count_rows("lab_longitudinal_dataset.csv"), 0)
        html = (self.temp_dir / "patient_profile.html").read_text(encoding="utf-8")
        self.assertIn("南方医科大学皮肤病医院", html)
        self.assertIn("杭州市第一人民医院", html)

    def test_build_with_generated_config(self) -> None:
        self.run_builder("--precheck-only")
        config_path = self.temp_dir / "suggested_project_config.json"
        second_dir = self.temp_dir / "from-config"
        cmd = [
            "python3",
            str(SCRIPT_PATH),
            "--work-dir",
            str(PROJECT_ROOT),
            "--listing-xlsx",
            str(LISTING_XLSX),
            "--finding-xlsx",
            str(FINDING_XLSX),
            "--pd-def-xlsx",
            str(PD_DEF_XLSX),
            "--subject-report-xls",
            str(SUBJECT_REPORT_XLS),
            "--config-json",
            str(config_path),
            "--output-dir",
            str(second_dir),
        ]
        subprocess.run(cmd, check=True)
        self.assertTrue((second_dir / "patient_profile.html").exists())
        self.assertGreater(self.count_rows_from_dir(second_dir, "efficacy_longitudinal_dataset.csv"), 0)

    def count_rows_from_dir(self, directory: Path, file_name: str) -> int:
        with (directory / file_name).open(encoding="utf-8-sig", newline="") as handle:
            return max(sum(1 for _ in csv.reader(handle)) - 1, 0)


if __name__ == "__main__":
    unittest.main()
