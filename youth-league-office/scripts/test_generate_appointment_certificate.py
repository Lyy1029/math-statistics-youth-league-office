#!/usr/bin/env python3

from __future__ import annotations

import csv
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import generate_appointment_certificate as appointment
from generate_certificate import InputValidationError, qn


HEADERS = [
    "内部人员键",
    "姓名",
    "班级",
    "学号",
    "组织",
    "部门",
    "职务",
    "任职证明类型",
    "任期开始",
    "任期结束",
    "批准状态",
]


def write_csv(path: Path, rows: list[list[str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(HEADERS)
        writer.writerows(rows)


def document_xml(path: Path) -> tuple[ET.Element, str]:
    with zipfile.ZipFile(path) as package:
        raw = package.read("word/document.xml")
    return ET.fromstring(raw), raw.decode("utf-8")


class AppointmentCertificateTests(unittest.TestCase):
    def test_staff_only_diagnostic_generates_editable_docx(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            roster = root / "roster.csv"
            write_csv(
                roster,
                [
                    ["S001", "张三", "统计24", "2400000001", "团委", "办公室", "干事", "干事", "", "", "已批准"],
                    ["S002", "李四", "数学24", "2400000002", "团委", "办公室", "干事", "干事", "", "", "已批准"],
                ],
            )
            result = appointment.generate_appointments(
                list_path=roster,
                academic_year="2025-2026",
                issue_date="2025-09-01",
                output_dir=root / "output",
                make_pdf=False,
            )
            self.assertEqual(result.person_count, 2)
            self.assertEqual(result.staff_proof_count, 1)
            self.assertEqual(result.leader_certificate_count, 0)
            generated = sorted((root / "output").glob("*.docx"))
            self.assertEqual(len(generated), 1)
            payload = "\n".join(document_xml(path)[1] for path in generated)
            self.assertNotIn("{{", payload)
            self.assertIn("2025至2026学年", payload)
            self.assertIn("2400000001", payload)
            for path in generated:
                root_xml, _ = document_xml(path)
                visible_runs = [
                    run
                    for run in root_xml.iter(qn("r"))
                    if any((node.text or "") for node in run.iter(qn("t")))
                ]
                self.assertTrue(visible_runs)
                for run in visible_runs:
                    fonts = run.find(qn("rPr") + "/" + qn("rFonts"))
                    self.assertIsNotNone(fonts)
                    self.assertEqual(
                        {fonts.attrib.get(qn(slot)) for slot in ("ascii", "hAnsi", "eastAsia", "cs")},
                        {"Noto Serif SC"},
                    )
                    self.assertFalse(
                        any(qn(theme) in fonts.attrib for theme in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"))
                    )

    def test_leader_svg_uses_reference_background_and_expected_text(self) -> None:
        record = appointment.AppointmentRecord(
            source_row=2,
            name="合成甲",
            class_name="",
            student_id="",
            organization="团委",
            department="",
            position="副书记",
            kind="leader",
            term_start="2025年9月",
            term_end="2026年9月",
        )
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "leader.svg"
            appointment.build_leader_svg(record, "2025年9月1日", output)
            payload = output.read_text(encoding="utf-8")
        self.assertIn("leader-appointment-background", appointment.LEADER_BACKGROUND.name)
        self.assertIn("data:image/png;base64,", payload)
        self.assertIn('font-family="Source Han Serif SC"', payload)
        self.assertIn('x="243.653" y="342.295" font-size="17.474">担</text>', payload)
        self.assertIn('x="109.218" y="342.270" font-size="17.474">聘</text>', payload)
        self.assertIn('x="379.762" y="675.729" font-size="14.585">2</text>', payload)
        self.assertIn('x="115.556" y="380.914"', payload)
        self.assertIn('font-size="14.585">2</text>', payload)
        self.assertNotIn("蔡欣辰", payload)

    def test_leader_certificate_is_pdf_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            roster = root / "leader.csv"
            write_csv(
                roster,
                [["S003", "王五", "", "", "团委", "", "副书记", "部长及以上", "2025-09", "2026-09", "已批准"]],
            )
            with self.assertRaisesRegex(appointment.CertificateError, "不支持 --no-pdf"):
                appointment.generate_appointments(
                    list_path=roster,
                    academic_year="2025-2026",
                    issue_date="2025-09-01",
                    output_dir=root / "output",
                    make_pdf=False,
                )

    def test_unapproved_row_stops_the_whole_batch(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            roster = Path(temp) / "roster.csv"
            write_csv(
                roster,
                [["S001", "张三", "统计24", "2400000001", "团委", "办公室", "干事", "干事", "", "", "待审核"]],
            )
            with self.assertRaisesRegex(InputValidationError, "未明确批准"):
                appointment.read_appointment_records(roster)

    def test_leader_term_and_staff_position_are_validated(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            invalid_staff = root / "staff.csv"
            write_csv(
                invalid_staff,
                [["S001", "张三", "统计24", "2400000001", "团委", "办公室", "部长", "干事", "", "", "已批准"]],
            )
            with self.assertRaisesRegex(InputValidationError, "非干事职务"):
                appointment.read_appointment_records(invalid_staff)
            invalid_leader = root / "leader.csv"
            write_csv(
                invalid_leader,
                [["S002", "李四", "", "", "学生会", "", "主席", "部长及以上", "2026-09", "2025-09", "已批准"]],
            )
            with self.assertRaisesRegex(InputValidationError, "结束时间必须晚于"):
                appointment.read_appointment_records(invalid_leader)

    def test_blank_template_contains_only_expected_headers(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "template.xlsx"
            appointment.create_list_template(path)
            rows = appointment._xlsx_rows(path, "任职证明批准名单")
            self.assertEqual(rows[0][1], HEADERS)
            self.assertEqual(len(rows), 1)

    def test_pdf_target_planning_does_not_extend_while_iterating(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            roster = root / "roster.csv"
            write_csv(
                roster,
                [["S001", "张三", "统计24", "2400000001", "团委", "办公室", "干事", "干事", "", "", "已批准"]],
            )
            with (
                mock.patch.object(appointment, "_find_soffice", return_value=Path("/tmp/soffice")),
                mock.patch.object(
                    appointment,
                    "convert_sources_to_pdf",
                    side_effect=lambda pairs, _: [destination for _, destination in pairs],
                ),
            ):
                result = appointment.generate_appointments(
                    list_path=roster,
                    academic_year="2025-2026",
                    issue_date="2025-09-01",
                    output_dir=root / "output",
                    make_pdf=True,
                )
            self.assertEqual(result.pdf_count, 1)


if __name__ == "__main__":
    unittest.main()
