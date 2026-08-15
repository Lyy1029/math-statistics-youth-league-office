#!/usr/bin/env python3
"""Synthetic-data tests for generate_certificate.py.

No name in this test file is sourced from an office document or a real roster.
"""

from __future__ import annotations

import csv
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest import mock
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generate_certificate as cert


W = "{" + cert.W_NS + "}"


def docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    return "".join(node.text or "" for node in root.iter(W + "t"))


def docx_font_families(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    families: set[str] = set()
    for run in root.iter(W + "r"):
        if not any((node.text or "") for node in run.iter(W + "t")):
            continue
        properties = run.find(W + "rPr")
        fonts = properties.find(W + "rFonts") if properties is not None else None
        if fonts is None:
            continue
        for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
            value = fonts.attrib.get(W + slot)
            if value:
                families.add(value)
        for slot in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            value = fonts.attrib.get(W + slot)
            if value:
                families.add("theme:" + value)
    return families


def paragraph_indent_attributes(path: Path, needle: str) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    for paragraph in root.iter(W + "p"):
        text = "".join(node.text or "" for node in paragraph.iter(W + "t"))
        if needle not in text:
            continue
        ppr = paragraph.find(W + "pPr")
        indent = ppr.find(W + "ind") if ppr is not None else None
        return dict(indent.attrib) if indent is not None else {}
    raise AssertionError(f"未找到段落：{needle}")


def name_table_geometry(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    table = next(root.iter(W + "tbl"))
    properties = table.find(W + "tblPr")
    layout = properties.find(W + "tblLayout")
    alignment = properties.find(W + "jc")
    table_width = int(properties.find(W + "tblW").attrib[W + "w"])
    grid = table.find(W + "tblGrid")
    grid_widths = [int(node.attrib[W + "w"]) for node in grid.findall(W + "gridCol")]
    rows = table.findall(W + "tr")
    cells = [cell for row in rows for cell in row.findall(W + "tc")]
    cell_widths = [
        int(cell.find(W + "tcPr").find(W + "tcW").attrib[W + "w"])
        for cell in cells
    ]
    cell_alignments = [
        cell.find(W + "p").find(W + "pPr").find(W + "jc").attrib[W + "val"]
        for cell in cells
    ]
    return {
        "layout": layout.attrib[W + "type"],
        "alignment": alignment.attrib[W + "val"],
        "table_width": table_width,
        "grid_widths": grid_widths,
        "row_cell_counts": [len(row.findall(W + "tc")) for row in rows],
        "cell_widths": cell_widths,
        "cell_alignments": cell_alignments,
    }


def outstanding_name_layout(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    body = root.find(".//" + W + "body")
    if body is None:
        raise AssertionError("文档缺少正文。")
    children = list(body)
    statement_index = next(
        index
        for index, child in enumerate(children)
        if child.tag == W + "p"
        and "兹证明" in "".join(node.text or "" for node in child.iter(W + "t"))
    )
    table_index = next(index for index, child in enumerate(children) if child.tag == W + "tbl")
    between = children[statement_index + 1 : table_index]
    blank_lines = sum(
        child.tag == W + "p"
        and not "".join(node.text or "" for node in child.iter(W + "t")).strip()
        for child in between
    )
    table_properties = children[table_index].find(W + "tblPr")
    position = table_properties.find(W + "tblpPr") if table_properties is not None else None
    if position is None:
        raise AssertionError("突出贡献名单表缺少浮动位置。")
    return {
        "blank_lines": blank_lines,
        "table_y": int(position.attrib[W + "tblpY"]),
    }


def activity_spacer_lines(path: Path) -> int:
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    body = root.find(".//" + W + "body")
    if body is None:
        raise AssertionError("文档缺少正文。")
    children = list(body)
    table_index = next(index for index, child in enumerate(children) if child.tag == W + "tbl")
    blanks = 0
    for child in children[table_index + 1 :]:
        if child.tag != W + "p" or "".join(
            node.text or "" for node in child.iter(W + "t")
        ).strip():
            break
        blanks += 1
    return blanks


def build_synthetic_xlsx(path: Path, rows: list[list[str]]) -> None:
    cert.create_approved_list_template(path)
    with zipfile.ZipFile(path) as package:
        infos = package.infolist()
        parts = {info.filename: package.read(info.filename) for info in infos}
    parts["xl/worksheets/sheet1.xml"] = cert._xlsx_sheet_xml(
        rows, [16, 14, 15, 30, 18, 15]
    )
    with zipfile.ZipFile(path, "w") as package:
        for info in infos:
            package.writestr(info, parts[info.filename])


class CertificateGeneratorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="certificate-test-")
        self.work = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_bundled_font_is_valid_and_pdf_conversion_prefers_it(self) -> None:
        font_path = cert._validate_bundled_font_asset()
        self.assertEqual(font_path, cert.BUNDLED_CJK_FONT_PATH)
        self.assertEqual(font_path.stat().st_size, cert.BUNDLED_CJK_FONT_SIZE)

        docx = self.work / "font-check.docx"
        docx.write_bytes(b"synthetic docx placeholder")
        pdf = self.work / "font-check.pdf"

        def fake_run(command, **kwargs):
            config = Path(kwargs["env"]["FONTCONFIG_FILE"])
            root = ET.fromstring(config.read_bytes())
            directories = [node.text for node in root.findall("dir")]
            self.assertEqual(directories[0], str(cert.BUNDLED_FONT_DIR))
            output_dir = Path(command[command.index("--outdir") + 1])
            (output_dir / "font-check.pdf").write_bytes(
                b"%PDF-1.7 /FontName /ABCDEF+NotoSerifSC-Regular synthetic"
            )
            return SimpleNamespace(returncode=0)

        with mock.patch.object(cert.subprocess, "run", side_effect=fake_run):
            cert.convert_to_pdf(docx, pdf, Path("/synthetic/soffice"))
        self.assertTrue(pdf.is_file())

    def test_activity_txt_sorting_and_identity_only(self) -> None:
        roster = self.work / "approved.txt"
        roster.write_text("姓名\n甲乙丙\n丁戊\n欧阳甲乙\n己庚辛\n", encoding="utf-8")
        result = cert.generate_certificate(
            kind="activity",
            names_path=roster,
            academic_year="2025-2026",
            activity_name="新生见面活动",
            role="参与人员",
            issue_date="2026-01-09",
            output_dir=self.work,
            make_pdf=False,
        )
        text = docx_text(result.docx_path)
        self.assertEqual(docx_font_families(result.docx_path), {cert.FALLBACK_CJK_FONT})
        self.assertEqual(result.columns, 8)
        self.assertNotIn("活动活动", text)
        self.assertNotRegex(text, cert.SCORE_LANGUAGE_RE)
        self.assertIn("参与人员", text)
        self.assertLess(text.index("甲乙丙"), text.index("己庚辛"))
        self.assertLess(text.index("己庚辛"), text.index("丁　戊"))
        self.assertLess(text.index("丁　戊"), text.index("欧阳甲乙"))
        geometry = name_table_geometry(result.docx_path)
        self.assertEqual(geometry["layout"], "fixed")
        self.assertEqual(geometry["alignment"], "center")
        self.assertEqual(geometry["row_cell_counts"], [8])
        self.assertEqual(len(set(geometry["grid_widths"])), 1)
        self.assertEqual(len(set(geometry["cell_widths"])), 1)
        self.assertEqual(geometry["grid_widths"][0], geometry["cell_widths"][0])
        self.assertEqual(geometry["table_width"], geometry["grid_widths"][0] * 8)
        self.assertEqual(set(geometry["cell_alignments"]), {"center"})
        self.assertEqual(
            activity_spacer_lines(result.docx_path),
            cert.ACTIVITY_SPACER_LINE_BUDGET - result.rows,
        )

    def test_activity_spacer_keeps_large_roster_within_budget(self) -> None:
        roster = self.work / "approved-large.txt"
        roster.write_text(
            "姓名\n"
            + "\n".join("合成人员" + chr(0x4E00 + index) for index in range(80))
            + "\n",
            encoding="utf-8",
        )
        result = cert.generate_certificate(
            kind="activity",
            names_path=roster,
            academic_year="2025-2026",
            activity_name="志愿服务活动",
            role="参与人员",
            issue_date="2026-06-30",
            output_dir=self.work,
            make_pdf=False,
        )
        self.assertEqual(result.rows, 10)
        self.assertEqual(activity_spacer_lines(result.docx_path), 1)

    def test_outstanding_csv_allowed_headers(self) -> None:
        roster = self.work / "approved.csv"
        with roster.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["内部人员键", "姓名", "证明类型", "活动", "角色", "批准状态"])
            writer.writerow(["SYN-001", "甲乙丙", "突出贡献", "迎新晚会活动", "突出贡献", "已批准"])
            writer.writerow(["SYN-002", "丁戊", "突出贡献", "迎新晚会活动", "突出贡献", "已批准"])
        result = cert.generate_certificate(
            kind="outstanding",
            names_path=roster,
            academic_year="2025至2026学年",
            activity_name="迎新晚会活动",
            issue_date="2026-01",
            output_dir=self.work,
            make_pdf=False,
        )
        text = docx_text(result.docx_path)
        self.assertEqual(result.columns, 10)
        self.assertNotIn("活动活动", text)
        self.assertNotRegex(text, cert.SCORE_LANGUAGE_RE)
        self.assertIn("突出贡献证明", text)
        geometry = name_table_geometry(result.docx_path)
        self.assertEqual(geometry["layout"], "fixed")
        self.assertEqual(geometry["row_cell_counts"], [10])
        self.assertEqual(len(set(geometry["grid_widths"])), 1)
        self.assertEqual(len(set(geometry["cell_widths"])), 1)
        self.assertEqual(geometry["table_width"], geometry["grid_widths"][0] * 10)
        layout = outstanding_name_layout(result.docx_path)
        self.assertEqual(layout["blank_lines"], 1)

        long_activity = "示范青春歌会暨共青团表彰大会优秀志愿者交流展示与总结活动成果分享会"
        long_roster = self.work / "approved-long.csv"
        with long_roster.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["内部人员键", "姓名", "证明类型", "活动", "角色", "批准状态"])
            writer.writerow(
                ["SYN-003", "甲乙丙", "突出贡献", long_activity, "突出贡献", "已批准"]
            )
        long_result = cert.generate_certificate(
            kind="outstanding",
            names_path=long_roster,
            academic_year="2025至2026学年",
            activity_name=long_activity,
            issue_date="2026-01",
            output_dir=self.work,
            make_pdf=False,
        )
        long_layout = outstanding_name_layout(long_result.docx_path)
        self.assertEqual(long_layout["blank_lines"], 1)
        self.assertEqual(
            long_layout["table_y"] - layout["table_y"],
            cert.OUTSTANDING_FALLBACK_LINE_HEIGHT_TWIPS,
        )

    def test_outstanding_official_event_name_is_not_rewritten(self) -> None:
        statement = cert._outstanding_statement(
            "2025至2026学年", "示范青春歌会暨共青团表彰大会"
        )
        self.assertIn("“示范青春歌会暨共青团表彰大会”中表现突出", statement)
        self.assertNotIn("表彰大会”活动中", statement)

    def test_xlsx_allowed_headers_and_approval(self) -> None:
        roster = self.work / "approved.xlsx"
        build_synthetic_xlsx(
            roster,
            [
                ["内部人员键", "姓名", "证明类型", "活动", "角色", "批准状态"],
                ["SYN-101", "甲乙丙", "活动证明", "志愿服务", "参与人员", "已批准"],
                ["SYN-102", "丁戊", "活动证明", "志愿服务", "参与人员", "已批准"],
            ],
        )
        result = cert.generate_certificate(
            kind="activity",
            names_path=roster,
            academic_year="2025-2026",
            activity_name="志愿服务",
            issue_date="2026-02-03",
            output_dir=self.work,
            make_pdf=False,
        )
        self.assertEqual(result.name_count, 2)
        text = docx_text(result.docx_path)
        self.assertIn("“志愿服务”参与人员：", text)
        self.assertNotIn("主题活动参与人员", text)
        indent = paragraph_indent_attributes(result.docx_path, "兹证明")
        self.assertNotIn(W + "firstLine", indent)
        self.assertNotIn(W + "firstLineChars", indent)
        geometry = name_table_geometry(result.docx_path)
        self.assertEqual(geometry["row_cell_counts"], [10])
        self.assertEqual(len(set(geometry["grid_widths"])), 1)
        self.assertEqual(len(set(geometry["cell_widths"])), 1)
        self.assertEqual(geometry["table_width"], geometry["grid_widths"][0] * 10)

    def test_cli_defaults_to_pdf_and_supports_diagnostic_opt_out(self) -> None:
        parser = cert._build_parser()
        common = [
            "generate", "--type", "activity", "--names", "synthetic.csv",
            "--academic-year", "2025-2026", "--activity", "示范活动",
            "--issue-date", "2026-06-30",
        ]
        self.assertTrue(parser.parse_args(common).make_pdf)
        self.assertFalse(parser.parse_args(common + ["--no-pdf"]).make_pdf)

    def test_sensitive_header_is_rejected_without_name_echo(self) -> None:
        roster = self.work / "unsafe.csv"
        roster.write_text("姓名,加分\n甲乙丙,20\n", encoding="utf-8")
        with self.assertRaises(cert.InputValidationError) as context:
            cert.read_name_records(roster)
        self.assertIn("证明批准名单", str(context.exception))
        self.assertNotIn("甲乙丙", str(context.exception))

    def test_structured_input_requires_name_header(self) -> None:
        roster = self.work / "no-header.csv"
        roster.write_text("甲乙丙\n丁戊\n", encoding="utf-8")
        with self.assertRaises(cert.InputValidationError) as context:
            cert.read_name_records(roster)
        self.assertIn("姓名", str(context.exception))
        self.assertNotIn("甲乙丙", str(context.exception))

    def test_formula_or_unsupported_header_is_rejected(self) -> None:
        formula = self.work / "formula.csv"
        formula.write_text("姓名,批准状态\n甲乙丙,=IF(1=1,\"已批准\",\"待审核\")\n", encoding="utf-8")
        with self.assertRaises(cert.InputValidationError) as formula_context:
            cert.read_name_records(formula)
        self.assertIn("公式", str(formula_context.exception))
        self.assertNotIn("甲乙丙", str(formula_context.exception))

        unsupported = self.work / "unsupported.csv"
        unsupported.write_text("姓名,手机号\n丁戊,00000000000\n", encoding="utf-8")
        with self.assertRaises(cert.InputValidationError) as unsupported_context:
            cert.read_name_records(unsupported)
        self.assertIn("不允许", str(unsupported_context.exception))
        self.assertNotIn("丁戊", str(unsupported_context.exception))

    def test_structured_context_mismatch_is_rejected_without_name_echo(self) -> None:
        roster = self.work / "mismatch.csv"
        roster.write_text(
            "内部人员键,姓名,证明类型,活动,角色,批准状态\n"
            "SYN-301,甲乙丙,突出贡献证明,其他活动,观众,已批准\n",
            encoding="utf-8",
        )
        with self.assertRaises(cert.InputValidationError) as context:
            cert.generate_certificate(
                kind="activity",
                names_path=roster,
                academic_year="2025-2026",
                activity_name="示范活动",
                role="工作人员",
                issue_date="2026-06-30",
                output_dir=self.work,
                make_pdf=False,
            )
        message = str(context.exception)
        self.assertIn("证明类型不一致", message)
        self.assertIn("活动不一致", message)
        self.assertIn("角色不一致", message)
        self.assertNotIn("甲乙丙", message)

    def test_unapproved_row_stops_instead_of_filtering(self) -> None:
        roster = self.work / "unapproved.csv"
        roster.write_text(
            "姓名,证明类型,活动,角色,批准状态\n"
            "甲乙丙,活动证明,示范活动,参与人员,待审核\n",
            encoding="utf-8",
        )
        with self.assertRaises(cert.InputValidationError) as context:
            cert.read_name_records(roster)
        self.assertIn("未明确批准", str(context.exception))
        self.assertNotIn("甲乙丙", str(context.exception))

    def test_duplicate_and_blank_rows_stop_without_name_echo(self) -> None:
        duplicate = self.work / "duplicate.txt"
        duplicate.write_text("甲乙丙\n甲乙丙\n", encoding="utf-8")
        with self.assertRaises(cert.InputValidationError) as duplicate_context:
            cert.validate_and_sort_names(cert.read_name_records(duplicate))
        self.assertIn("重复", str(duplicate_context.exception))
        self.assertNotIn("甲乙丙", str(duplicate_context.exception))

        blank = self.work / "blank.csv"
        blank.write_text(
            "姓名,证明类型,活动,角色,批准状态\n"
            "丁戊,活动证明,示范活动,参与人员,已批准\n"
            ",活动证明,示范活动,参与人员,已批准\n",
            encoding="utf-8",
        )
        with self.assertRaises(cert.InputValidationError) as blank_context:
            cert.read_name_records(blank)
        self.assertIn("空白姓名", str(blank_context.exception))
        self.assertNotIn("丁戊", str(blank_context.exception))


if __name__ == "__main__":
    unittest.main()
