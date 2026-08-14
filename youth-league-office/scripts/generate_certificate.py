#!/usr/bin/env python3
"""Generate Youth League Office activity and outstanding-contribution certificates.

The implementation intentionally uses only Python's standard library.  DOCX and
XLSX files are OOXML zip packages, so no third-party dependency is required.

Privacy properties:
* Console messages never echo a person's name.
* Duplicate/blank diagnostics report source row numbers only.
* Generated DOCX metadata is scrubbed of template authorship information.
* Certificate prose proves participation/contribution only; scoring language is
  rejected and never generated.
"""

from __future__ import annotations

import argparse
import copy
import csv
import datetime as dt
import io
import math
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
FALLBACK_CJK_FONT = "Songti SC"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

W = "{" + W_NS + "}"
R = "{" + R_NS + "}"
S = "{" + SHEET_NS + "}"
PR = "{" + PKG_REL_NS + "}"

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
ASSET_DIR = SKILL_DIR / "assets"
DEFAULT_TEMPLATES = {
    "activity": ASSET_DIR / "activity-certificate-template.docx",
    "outstanding": ASSET_DIR / "outstanding-contribution-template.docx",
}

HEADER_ALIASES = {
    "姓名",
    "名字",
    "学生姓名",
    "同学姓名",
    "成员姓名",
    "参与人姓名",
    "name",
    "studentname",
}
ALLOWED_STRUCTURED_HEADERS = {
    "内部人员键",
    "姓名",
    "证明类型",
    "活动",
    "角色",
    "批准状态",
}
FORBIDDEN_INTERNAL_HEADER_TERMS = {
    "分值",
    "加分",
    "请假",
    "缺勤",
    "出勤次数",
    "奖励方案",
    "公式",
    "评分",
}
APPROVED_STATUS_VALUES = {"已批准", "批准", "通过", "是", "yes", "approved"}
MAX_NAME_ROWS = 12
OUTSTANDING_REFERENCE_TEXT_LINES = 2
OUTSTANDING_BASE_TABLE_Y_OFFSET_TWIPS = 1440
OUTSTANDING_FALLBACK_LINE_HEIGHT_TWIPS = 500
SCORE_LANGUAGE_RE = re.compile(
    r"(?:加分|分值|综测分|综合测评分|德育成绩|专项教育|"
    r"(?:加|计|奖励|获)\s*\d+(?:\.\d+)?\s*分)"
)


class CertificateError(Exception):
    """A safe, user-facing certificate generation error."""


class InputValidationError(CertificateError):
    """The supplied names or certificate fields are invalid."""


class TemplateError(CertificateError):
    """A template is missing required anchors or is malformed."""


@dataclass(frozen=True)
class NameRecord:
    value: str
    source_row: int
    certificate_type: str | None = None
    activity_name: str | None = None
    role: str | None = None


@dataclass(frozen=True)
class GenerationResult:
    docx_path: Path
    pdf_path: Path | None
    name_count: int
    columns: int
    rows: int


def qn(local_name: str) -> str:
    return W + local_name


def _safe_text(value: object) -> str:
    return "" if value is None else str(value)


def _decode_text_file(path: Path) -> str:
    data = path.read_bytes()
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16")
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise InputValidationError("名单文本编码无法识别；请另存为 UTF-8、UTF-16 或 GB18030。")


def _trim_outer_blank_rows(rows: list[tuple[int, list[str]]]) -> list[tuple[int, list[str]]]:
    def blank(row: tuple[int, list[str]]) -> bool:
        return not any(cell.strip() for cell in row[1])

    start = 0
    end = len(rows)
    while start < end and blank(rows[start]):
        start += 1
    while end > start and blank(rows[end - 1]):
        end -= 1
    return rows[start:end]


def _normalize_header(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).strip().casefold()
    return re.sub(r"[\s_\-—–]+", "", value)


def _letters_to_index(value: str) -> int:
    result = 0
    for char in value.upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(value)
        result = result * 26 + ord(char) - ord("A") + 1
    return result - 1


def _resolve_name_column(
    first_row: Sequence[str], selector: str
) -> tuple[int, bool]:
    normalized = [_normalize_header(cell) for cell in first_row]
    selector_norm = _normalize_header(selector)

    if selector_norm == "auto":
        for index, value in enumerate(normalized):
            if value in HEADER_ALIASES:
                return index, True
        nonempty = [index for index, cell in enumerate(first_row) if cell.strip()]
        if len(nonempty) > 1:
            raise InputValidationError(
                "名单含多列但未识别到“姓名”表头；请用 --name-column 指定列名、列号或列字母。"
            )
        return (nonempty[0] if nonempty else 0), False

    if selector_norm.isdigit():
        index = int(selector_norm) - 1
        if index < 0:
            raise InputValidationError("--name-column 的数字必须从 1 开始。")
        is_header = index < len(normalized) and normalized[index] in HEADER_ALIASES
        return index, is_header

    if re.fullmatch(r"[a-z]+", selector_norm):
        index = _letters_to_index(selector_norm)
        is_header = index < len(normalized) and normalized[index] in HEADER_ALIASES
        return index, is_header

    for index, value in enumerate(normalized):
        if value == selector_norm:
            return index, True
    raise InputValidationError("未找到 --name-column 指定的表头；请核对列名。")


def _records_from_tabular_rows(
    rows: list[tuple[int, list[str]]], selector: str, *, isolate_headers: bool = False
) -> list[NameRecord]:
    rows = _trim_outer_blank_rows(rows)
    if not rows:
        raise InputValidationError("名单为空。")

    first_number, first_values = rows[0]
    del first_number
    if isolate_headers:
        normalized_first_row = [_normalize_header(cell) for cell in first_values]
        if any(
            any(term in value for term in FORBIDDEN_INTERNAL_HEADER_TERMS)
            for value in normalized_first_row
            if value
        ):
            raise InputValidationError(
                "名单含分值、考勤、奖励方案、公式或评分字段；请先导出“证明批准名单”。"
            )
    column, has_header = _resolve_name_column(first_values, selector)
    approval_column: int | None = None
    type_column: int | None = None
    activity_column: int | None = None
    role_column: int | None = None
    if isolate_headers:
        if any(
            cell.lstrip().startswith(("=", "+", "@"))
            for _, values in rows
            for cell in values
            if cell.strip()
        ):
            raise InputValidationError(
                "结构化名单含公式式内容；请先导出只含身份字段的“证明批准名单”。"
            )
        if not has_header:
            raise InputValidationError(
                "CSV/XLSX 名单必须有“姓名”表头；无表头的单列名单请另存为 TXT。"
            )
        normalized_headers = [_normalize_header(cell) for cell in first_values]
        forbidden = [
            value
            for value in normalized_headers
            if any(term in value for term in FORBIDDEN_INTERNAL_HEADER_TERMS)
        ]
        if forbidden:
            raise InputValidationError(
                "名单含分值、考勤、奖励方案、公式或评分字段；请先导出“证明批准名单”。"
            )
        allowed_normalized = {_normalize_header(value) for value in ALLOWED_STRUCTURED_HEADERS}
        unsupported = [value for value in normalized_headers if value and value not in allowed_normalized]
        if unsupported:
            raise InputValidationError(
                "结构化名单含不允许的扩展字段；仅允许内部人员键、姓名、证明类型、活动、角色、批准状态。"
            )
        nonempty_headers = [value for value in normalized_headers if value]
        if len(nonempty_headers) != len(set(nonempty_headers)):
            raise InputValidationError("结构化名单含重复表头；请整理后重试。")
        required_headers = ["姓名", "证明类型", "活动", "角色", "批准状态"]
        missing_headers = [
            header
            for header in required_headers
            if _normalize_header(header) not in normalized_headers
        ]
        if missing_headers:
            raise InputValidationError(
                "结构化证明批准名单缺少必需表头："
                + "、".join(missing_headers)
                + "。"
            )
        name_header = _normalize_header("姓名")
        column = normalized_headers.index(name_header)
        approval_column = normalized_headers.index(_normalize_header("批准状态"))
        type_column = normalized_headers.index(_normalize_header("证明类型"))
        activity_column = normalized_headers.index(_normalize_header("活动"))
        role_column = normalized_headers.index(_normalize_header("角色"))
    data_rows = rows[1:] if has_header else rows
    if not data_rows:
        raise InputValidationError("姓名表头下没有数据。")

    empty_rows: list[int] = []
    unapproved_rows: list[int] = []
    incomplete_context_rows: list[int] = []
    records: list[NameRecord] = []
    for row_number, values in data_rows:
        if not any(cell.strip() for cell in values):
            empty_rows.append(row_number)
            continue
        value = values[column] if column < len(values) else ""
        if not value.strip():
            empty_rows.append(row_number)
            continue
        if approval_column is not None:
            approval = values[approval_column] if approval_column < len(values) else ""
            if _normalize_header(approval) not in APPROVED_STATUS_VALUES:
                unapproved_rows.append(row_number)
                continue
        certificate_type = values[type_column].strip() if type_column is not None and type_column < len(values) else None
        activity_name = values[activity_column].strip() if activity_column is not None and activity_column < len(values) else None
        role = values[role_column].strip() if role_column is not None and role_column < len(values) else None
        if isolate_headers and not all((certificate_type, activity_name, role)):
            incomplete_context_rows.append(row_number)
            continue
        records.append(
            NameRecord(
                value=value,
                source_row=row_number,
                certificate_type=certificate_type,
                activity_name=activity_name,
                role=role,
            )
        )

    if empty_rows:
        locations = _format_row_numbers(empty_rows)
        raise InputValidationError(
            f"发现空白姓名项，位于行 {locations}；为避免漏开，已停止生成。"
        )
    if unapproved_rows:
        raise InputValidationError(
            "发现未明确批准的条目，位于行 "
            + _format_row_numbers(unapproved_rows)
            + "；请完成核实后再生成，不会自动筛掉。"
        )
    if incomplete_context_rows:
        raise InputValidationError(
            "证明批准名单的证明类型、活动或角色为空，位于行 "
            + _format_row_numbers(incomplete_context_rows)
            + "；为防止错用名单，已停止生成。"
        )
    return records


def read_txt_names(path: Path) -> list[NameRecord]:
    text = _decode_text_file(path)
    raw_lines = text.splitlines()
    indexed = [(index, [line]) for index, line in enumerate(raw_lines, start=1)]
    return _records_from_tabular_rows(indexed, "1")


def read_csv_names(path: Path, selector: str) -> list[NameRecord]:
    text = _decode_text_file(path)
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;，")
    except csv.Error:
        dialect = csv.excel
    rows = [
        (index, [_safe_text(cell) for cell in row])
        for index, row in enumerate(csv.reader(io.StringIO(text), dialect), start=1)
    ]
    return _records_from_tabular_rows(rows, selector, isolate_headers=True)


def _xlsx_shared_strings(package: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in package.namelist():
        return []
    root = ET.fromstring(package.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in item.iter(S + "t")) for item in root.findall(S + "si")]


def _xlsx_sheet_path(package: zipfile.ZipFile, requested_sheet: str | None) -> str:
    workbook = ET.fromstring(package.read("xl/workbook.xml"))
    relationships = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships.findall(PR + "Relationship")
    }
    sheets = workbook.find(S + "sheets")
    if sheets is None:
        raise InputValidationError("XLSX 中没有工作表。")

    available = list(sheets.findall(S + "sheet"))
    if requested_sheet is None:
        candidates = [s for s in available if s.attrib.get("state", "visible") == "visible"]
        chosen = (candidates or available)[0] if available else None
    else:
        chosen = next((s for s in available if s.attrib.get("name") == requested_sheet), None)
    if chosen is None:
        if requested_sheet is None:
            raise InputValidationError("XLSX 中没有可读取的工作表。")
        raise InputValidationError("未找到 --sheet 指定的工作表。")

    relationship_id = chosen.attrib.get(R + "id")
    target = targets.get(relationship_id or "")
    if not target:
        raise InputValidationError("XLSX 工作表关系损坏。")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def _xlsx_cell_value(cell: ET.Element, shared: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(t.text or "" for t in cell.iter(S + "t"))
    value_node = cell.find(S + "v")
    value = "" if value_node is None or value_node.text is None else value_node.text
    if cell_type == "s" and value:
        try:
            return shared[int(value)]
        except (ValueError, IndexError) as exc:
            raise InputValidationError("XLSX 共享字符串索引损坏。") from exc
    return value


def read_xlsx_names(path: Path, selector: str, sheet_name: str | None) -> list[NameRecord]:
    try:
        with zipfile.ZipFile(path) as package:
            shared = _xlsx_shared_strings(package)
            sheet_path = _xlsx_sheet_path(package, sheet_name)
            sheet = ET.fromstring(package.read(sheet_path))
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        raise InputValidationError("XLSX 文件损坏或不是有效的 OOXML 工作簿。") from exc

    tabular: list[tuple[int, list[str]]] = []
    if any(True for _ in sheet.iter(S + "f")):
        raise InputValidationError(
            "XLSX 名单工作表含公式；请先导出只含身份字段的“证明批准名单”。"
        )
    sheet_data = sheet.find(S + "sheetData")
    if sheet_data is None:
        raise InputValidationError("XLSX 工作表为空。")
    for ordinal, row in enumerate(sheet_data.findall(S + "row"), start=1):
        try:
            row_number = int(row.attrib.get("r", ordinal))
        except ValueError:
            row_number = ordinal
        values_by_column: dict[int, str] = {}
        for cell in row.findall(S + "c"):
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Za-z]+)", ref)
            column = _letters_to_index(match.group(1)) if match else len(values_by_column)
            values_by_column[column] = _xlsx_cell_value(cell, shared)
        if values_by_column:
            width = max(values_by_column) + 1
            values = [values_by_column.get(index, "") for index in range(width)]
        else:
            values = []
        tabular.append((row_number, values))
    return _records_from_tabular_rows(tabular, selector, isolate_headers=True)


def read_name_records(path: Path, selector: str = "auto", sheet_name: str | None = None) -> list[NameRecord]:
    if not path.is_file():
        raise InputValidationError("名单文件不存在或不是普通文件。")
    suffix = path.suffix.casefold()
    if suffix == ".txt":
        return read_txt_names(path)
    if suffix in {".csv", ".tsv"}:
        return read_csv_names(path, selector)
    if suffix == ".xlsx":
        return read_xlsx_names(path, selector, sheet_name)
    raise InputValidationError("名单仅支持 .txt、.csv、.tsv 或 .xlsx；不支持旧版 .xls。")


def _canonical_name(value: str) -> str:
    value = unicodedata.normalize("NFKC", value)
    return re.sub(r"\s+", "", value).strip()


def _format_row_numbers(numbers: Iterable[int], limit: int = 20) -> str:
    values = list(numbers)
    shown = "、".join(str(number) for number in values[:limit])
    if len(values) > limit:
        shown += f" 等 {len(values)} 行"
    return shown


def validate_and_sort_names(records: Sequence[NameRecord]) -> list[NameRecord]:
    if not records:
        raise InputValidationError("名单为空。")

    normalized: list[NameRecord] = []
    invalid_rows: list[int] = []
    for record in records:
        value = _canonical_name(record.value)
        allowed = all(char.isalpha() or char in {"·", "•", "-", "'", "’"} for char in value)
        if len(value) < 2 or len(value) > 30 or not allowed:
            invalid_rows.append(record.source_row)
            continue
        normalized.append(
            NameRecord(
                value=value,
                source_row=record.source_row,
                certificate_type=record.certificate_type,
                activity_name=record.activity_name,
                role=record.role,
            )
        )
    if invalid_rows:
        raise InputValidationError(
            "姓名含空白、数字、异常符号或长度不合理，位于行 "
            + _format_row_numbers(invalid_rows)
            + "；为保护隐私不在日志中显示原值。"
        )

    first_rows: dict[str, int] = {}
    duplicate_rows: list[int] = []
    for record in normalized:
        key = unicodedata.normalize("NFKC", record.value).casefold()
        if key in first_rows:
            duplicate_rows.extend([first_rows[key], record.source_row])
        else:
            first_rows[key] = record.source_row
    if duplicate_rows:
        rows = sorted(set(duplicate_rows))
        raise InputValidationError(
            "发现重复姓名项，涉及行 "
            + _format_row_numbers(rows)
            + "；为避免重复开具，已停止生成。"
        )

    def bucket(record: NameRecord) -> int:
        length = len(record.value)
        if length == 3:
            return 0
        if length == 2:
            return 1
        return 2

    # Python's sort is stable, so original order is retained inside each length group.
    return sorted(normalized, key=bucket)


def normalize_academic_year(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value).strip()
    normalized = re.sub(r"学年$", "", normalized).strip()
    match = re.fullmatch(r"(20\d{2})\s*(?:-|—|–|至|~|～|/|\\)\s*(20\d{2})", normalized)
    if not match:
        raise InputValidationError("学年格式应为 2025-2026、2025至2026 或带“学年”的同义形式。")
    start, end = map(int, match.groups())
    if end != start + 1:
        raise InputValidationError("学年起止年份必须相邻，例如 2025-2026。")
    return f"{start}至{end}学年"


def normalize_issue_date(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value).strip()
    patterns = [
        (r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", True),
        (r"(20\d{2})年\s*(\d{1,2})月\s*(\d{1,2})日", True),
        (r"(20\d{2})[-/.](\d{1,2})", False),
        (r"(20\d{2})年\s*(\d{1,2})月", False),
    ]
    for pattern, has_day in patterns:
        match = re.fullmatch(pattern, normalized)
        if not match:
            continue
        year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3)) if has_day else 1
        try:
            dt.date(year, month, day)
        except ValueError as exc:
            raise InputValidationError("出具日期不是有效日期。") from exc
        return f"{year}年{month}月{day}日" if has_day else f"{year}年{month}月"
    raise InputValidationError("出具日期格式应为 YYYY-MM-DD、YYYY年M月D日、YYYY-MM 或 YYYY年M月。")


def _normalize_plain_field(value: str, label: str, max_length: int) -> str:
    normalized = unicodedata.normalize("NFKC", value).strip()
    if not normalized or len(normalized) > max_length or any(char in normalized for char in "\r\n\t"):
        raise InputValidationError(f"{label}为空、过长或包含换行/制表符。")
    if SCORE_LANGUAGE_RE.search(normalized):
        raise InputValidationError(f"{label}含加分或分值表述；证明正文只允许证明参与/贡献身份。")
    return normalized


def normalize_activity_name(value: str) -> str:
    normalized = _normalize_plain_field(value, "活动名称", 100)
    pairs = [("“", "”"), ('"', '"'), ("‘", "’"), ("'", "'")]
    for left, right in pairs:
        if len(normalized) >= 2 and normalized.startswith(left) and normalized.endswith(right):
            normalized = normalized[len(left) : -len(right)].strip()
            break
    if not normalized:
        raise InputValidationError("活动名称不能只包含引号。")
    return normalized


def normalize_role(value: str) -> str:
    return _normalize_plain_field(value, "参与身份", 30)


def normalize_certificate_type(value: str) -> str:
    normalized = _normalize_header(value)
    activity_aliases = {_normalize_header(item) for item in ("activity", "活动证明", "普通活动证明")}
    outstanding_aliases = {
        _normalize_header(item) for item in ("outstanding", "突出贡献", "突出贡献证明")
    }
    if normalized in activity_aliases:
        return "activity"
    if normalized in outstanding_aliases:
        return "outstanding"
    raise InputValidationError("证明批准名单含无法识别的证明类型；请使用“活动证明”或“突出贡献证明”。")


def validate_roster_context(
    records: Sequence[NameRecord], *, kind: str, activity_name: str, role: str
) -> None:
    type_rows: list[int] = []
    activity_rows: list[int] = []
    role_rows: list[int] = []
    for record in records:
        # TXT is supported as a minimal one-column input. Structured CSV/XLSX
        # records always carry all three approval-context fields.
        if record.certificate_type is None:
            continue
        try:
            record_kind = normalize_certificate_type(record.certificate_type)
            record_activity = normalize_activity_name(record.activity_name or "")
            record_role = normalize_role(record.role or "")
        except InputValidationError:
            type_rows.append(record.source_row)
            continue
        if record_kind != kind:
            type_rows.append(record.source_row)
        if record_activity != activity_name:
            activity_rows.append(record.source_row)
        if record_role != role:
            role_rows.append(record.source_row)
    problems = []
    if type_rows:
        problems.append("证明类型不一致行 " + _format_row_numbers(type_rows))
    if activity_rows:
        problems.append("活动不一致行 " + _format_row_numbers(activity_rows))
    if role_rows:
        problems.append("角色不一致行 " + _format_row_numbers(role_rows))
    if problems:
        raise InputValidationError(
            "证明批准名单与本次生成参数不一致（"
            + "；".join(problems)
            + "）；为防止错用批准名单，已停止生成。"
        )


def _activity_statement(academic_year: str, activity_name: str, role: str) -> str:
    # 用户输入的是正式全称；普通活动证明不得擅自补“主题活动”等词。
    return f"兹证明以下同学为数统学院{academic_year}“{activity_name}”{role}："


def _outstanding_statement(academic_year: str, activity_name: str) -> str:
    return (
        "兹证明以下同学在北京工商大学数学与统计学院"
        f"{academic_year}“{activity_name}”中表现突出，特此授予突出贡献证明："
    )


def _register_namespaces(raw_xml: bytes) -> str:
    text = raw_xml.decode("utf-8")
    match = re.search(r"<[^!?][^>]*>", text)
    if not match:
        raise TemplateError("OOXML 主文档缺少根元素。")
    start_tag = match.group(0)
    # Drawing namespaces are often declared on descendant elements rather than
    # the document root. Register every declaration so ElementTree does not emit
    # undeclared nsN prefixes when it rebuilds the XML.
    for prefix, uri in re.findall(r'xmlns(?::([\w.-]+))?="([^"]+)"', text):
        if prefix == "xml":
            continue
        try:
            ET.register_namespace(prefix or "", uri)
        except ValueError:
            # Reserved prefixes are already known to ElementTree.
            pass
    return start_tag


def _parse_document(raw_xml: bytes) -> tuple[ET.Element, str]:
    start_tag = _register_namespaces(raw_xml)
    try:
        return ET.fromstring(raw_xml), start_tag
    except ET.ParseError as exc:
        raise TemplateError("模板的 word/document.xml 无法解析。") from exc


def _serialize_document(root: ET.Element, original_start_tag: str) -> bytes:
    del original_start_tag
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qn("t")))


def _paragraphs(root: ET.Element) -> list[ET.Element]:
    body = root.find(".//" + qn("body"))
    if body is None:
        raise TemplateError("模板缺少正文。")
    return [child for child in body if child.tag == qn("p")]


def _find_paragraph(root: ET.Element, predicate, description: str) -> ET.Element:
    matches = [paragraph for paragraph in _paragraphs(root) if predicate(paragraph_text(paragraph))]
    if len(matches) != 1:
        raise TemplateError(f"模板中{description}锚点数量应为 1，实际为 {len(matches)}。")
    return matches[0]


def set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    ppr = paragraph.find(qn("pPr"))
    first_rpr = paragraph.find(".//" + qn("rPr"))
    rpr_copy = copy.deepcopy(first_rpr) if first_rpr is not None else None
    for child in list(paragraph):
        if child is not ppr:
            paragraph.remove(child)
    run = ET.SubElement(paragraph, qn("r"))
    if rpr_copy is not None:
        run.append(rpr_copy)
    run_properties = run.find(qn("rPr"))
    if run_properties is None:
        run_properties = ET.Element(qn("rPr"))
        run.insert(0, run_properties)
    _set_child(
        run_properties,
        "rFonts",
        ascii=FALLBACK_CJK_FONT,
        hAnsi=FALLBACK_CJK_FONT,
        eastAsia=FALLBACK_CJK_FONT,
        cs=FALLBACK_CJK_FONT,
        hint="eastAsia",
    )
    text_node = ET.SubElement(run, qn("t"))
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set("{" + XML_NS + "}space", "preserve")
    text_node.text = text


def clear_first_line_indent(paragraph: ET.Element) -> None:
    """Remove template first-line indentation so the certificate statement does not orphan a role."""
    ppr = paragraph.find(qn("pPr"))
    if ppr is None:
        return
    indent = ppr.find(qn("ind"))
    if indent is None:
        return
    indent.attrib.pop(qn("firstLine"), None)
    indent.attrib.pop(qn("firstLineChars"), None)


def _set_child(parent: ET.Element, local_name: str, **attributes: str) -> ET.Element:
    child = parent.find(qn(local_name))
    if child is None:
        child = ET.SubElement(parent, qn(local_name))
    for key, value in attributes.items():
        child.set(qn(key), value)
    return child


def _apply_cjk_font_fallback(root: ET.Element) -> None:
    """Replace unavailable/template-subset CJK fonts while preserving other styling."""
    for run in root.iter(qn("r")):
        text = "".join(node.text or "" for node in run.iter(qn("t")))
        if not re.search(r"[\u3400-\u9fff\uf900-\ufaff]", text):
            continue
        run_properties = run.find(qn("rPr"))
        if run_properties is None:
            run_properties = ET.Element(qn("rPr"))
            run.insert(0, run_properties)
        _set_child(
            run_properties,
            "rFonts",
            ascii=FALLBACK_CJK_FONT,
            hAnsi=FALLBACK_CJK_FONT,
            eastAsia=FALLBACK_CJK_FONT,
            cs=FALLBACK_CJK_FONT,
            hint="eastAsia",
        )


def _apply_global_document_font(root: ET.Element) -> None:
    """Use a portable CJK font for every Word text run while retaining size/style.

    LibreOffice can select western/theme fonts for mixed Chinese runs even when
    only eastAsia is set, producing tofu boxes. Applying the same family to all
    script slots is deterministic and keeps the source template's layout.
    """
    for run in root.iter(qn("r")):
        if not any((node.text or "") for node in run.iter(qn("t"))):
            continue
        run_properties = run.find(qn("rPr"))
        if run_properties is None:
            run_properties = ET.Element(qn("rPr"))
            run.insert(0, run_properties)
        _set_child(
            run_properties,
            "rFonts",
            ascii=FALLBACK_CJK_FONT,
            hAnsi=FALLBACK_CJK_FONT,
            eastAsia=FALLBACK_CJK_FONT,
            cs=FALLBACK_CJK_FONT,
            hint="eastAsia",
        )


def _display_name(value: str) -> str:
    if len(value) == 2 and all(char.isalpha() for char in value):
        return value[0] + "\u3000" + value[1]
    return value


def _grid_dimensions(names: Sequence[NameRecord]) -> tuple[int, int]:
    columns = 8 if any(len(record.value) >= 4 for record in names) else 10
    rows = math.ceil(len(names) / columns)
    if rows > MAX_NAME_ROWS:
        capacity = columns * MAX_NAME_ROWS
        raise InputValidationError(
            f"名单共 {len(names)} 人，超过单页版式容量 {capacity} 人；请拆分为多份证明后分别生成。"
        )
    return columns, rows


def _page_content_width(root: ET.Element) -> int:
    section = root.find(".//" + qn("sectPr"))
    if section is None:
        return 8300
    page = section.find(qn("pgSz"))
    margins = section.find(qn("pgMar"))
    try:
        page_width = int(page.attrib[qn("w")]) if page is not None else 11906
        left = int(margins.attrib.get(qn("left"), "1800")) if margins is not None else 1800
        right = int(margins.attrib.get(qn("right"), "1800")) if margins is not None else 1800
        return max(5000, page_width - left - right)
    except (KeyError, ValueError):
        return 8300


def make_borderless_name_table(
    names: Sequence[NameRecord], columns: int, width_twips: int
) -> ET.Element:
    cell_width = max(1, width_twips // columns)
    table_width = cell_width * columns
    table = ET.Element(qn("tbl"))
    table_properties = ET.SubElement(table, qn("tblPr"))
    _set_child(table_properties, "tblW", w=str(table_width), type="dxa")
    _set_child(table_properties, "tblLayout", type="fixed")
    _set_child(table_properties, "jc", val="center")
    _set_child(table_properties, "tblCellSpacing", w="0", type="dxa")
    borders = ET.SubElement(table_properties, qn("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _set_child(borders, side, val="nil")
    margins = ET.SubElement(table_properties, qn("tblCellMar"))
    for side, amount in (("top", "30"), ("left", "20"), ("bottom", "30"), ("right", "20")):
        _set_child(margins, side, w=amount, type="dxa")

    grid = ET.SubElement(table, qn("tblGrid"))
    for _ in range(columns):
        grid_column = ET.SubElement(grid, qn("gridCol"))
        grid_column.set(qn("w"), str(cell_width))

    values = [_display_name(record.value) for record in names]
    total_rows = math.ceil(len(values) / columns)
    for row_index in range(total_rows):
        row = ET.SubElement(table, qn("tr"))
        row_properties = ET.SubElement(row, qn("trPr"))
        _set_child(row_properties, "cantSplit")
        _set_child(row_properties, "trHeight", val="330", hRule="atLeast")
        for column_index in range(columns):
            position = row_index * columns + column_index
            value = values[position] if position < len(values) else ""
            cell = ET.SubElement(row, qn("tc"))
            cell_properties = ET.SubElement(cell, qn("tcPr"))
            _set_child(cell_properties, "tcW", w=str(cell_width), type="dxa")
            _set_child(cell_properties, "vAlign", val="center")
            paragraph = ET.SubElement(cell, qn("p"))
            paragraph_properties = ET.SubElement(paragraph, qn("pPr"))
            _set_child(paragraph_properties, "jc", val="center")
            _set_child(paragraph_properties, "spacing", before="0", after="0", line="300", lineRule="exact")
            run = ET.SubElement(paragraph, qn("r"))
            run_properties = ET.SubElement(run, qn("rPr"))
            _set_child(
                run_properties,
                "rFonts",
                ascii=FALLBACK_CJK_FONT,
                hAnsi=FALLBACK_CJK_FONT,
                eastAsia=FALLBACK_CJK_FONT,
                cs=FALLBACK_CJK_FONT,
                hint="eastAsia",
            )
            _set_child(run_properties, "sz", val="21")
            _set_child(run_properties, "szCs", val="21")
            text_node = ET.SubElement(run, qn("t"))
            text_node.text = value
    return table


def validate_equal_width_name_table(table: ET.Element, columns: int) -> None:
    """Fail closed unless the name table is fixed-layout, equal-width and centered."""
    properties = table.find(qn("tblPr"))
    layout = properties.find(qn("tblLayout")) if properties is not None else None
    table_width_node = properties.find(qn("tblW")) if properties is not None else None
    grid = table.find(qn("tblGrid"))
    grid_widths = (
        [int(node.attrib.get(qn("w"), "0")) for node in grid.findall(qn("gridCol"))]
        if grid is not None
        else []
    )
    rows = table.findall(qn("tr"))
    cells = [cell for row in rows for cell in row.findall(qn("tc"))]
    cell_widths: list[int] = []
    centered = True
    for cell in cells:
        cell_properties = cell.find(qn("tcPr"))
        width = cell_properties.find(qn("tcW")) if cell_properties is not None else None
        cell_widths.append(int(width.attrib.get(qn("w"), "0")) if width is not None else 0)
        paragraph = cell.find(qn("p"))
        paragraph_properties = paragraph.find(qn("pPr")) if paragraph is not None else None
        alignment = paragraph_properties.find(qn("jc")) if paragraph_properties is not None else None
        centered = centered and alignment is not None and alignment.attrib.get(qn("val")) == "center"

    table_width = (
        int(table_width_node.attrib.get(qn("w"), "0")) if table_width_node is not None else 0
    )
    equal_grid = len(grid_widths) == columns and len(set(grid_widths)) == 1 and grid_widths[0] > 0
    equal_cells = False
    if equal_grid and cells and all(len(row.findall(qn("tc"))) == columns for row in rows):
        equal_cells = len(set(cell_widths)) == 1 and cell_widths[0] == grid_widths[0]
    fixed = layout is not None and layout.attrib.get(qn("type")) == "fixed"
    width_matches = equal_grid and table_width == grid_widths[0] * columns
    if not (fixed and equal_grid and equal_cells and width_matches and centered):
        raise TemplateError("名单表格未满足固定布局、均分列宽和单元格居中要求。")


def _frame_width(paragraph: ET.Element, fallback: int) -> int:
    properties = paragraph.find(qn("pPr"))
    frame = properties.find(qn("framePr")) if properties is not None else None
    try:
        return int(frame.attrib.get(qn("w"), str(fallback))) if frame is not None else fallback
    except ValueError:
        return fallback


def _paragraph_line_height(paragraph: ET.Element, fallback: int) -> int:
    properties = paragraph.find(qn("pPr"))
    spacing = properties.find(qn("spacing")) if properties is not None else None
    try:
        return int(spacing.attrib.get(qn("line"), str(fallback))) if spacing is not None else fallback
    except ValueError:
        return fallback


def _estimated_paragraph_lines(paragraph: ET.Element, text: str) -> int:
    """Estimate wrapped lines well enough to keep the floating names below the prose.

    The honor-certificate template uses a fixed-width frame and Songti text.
    Full-width characters occupy roughly one font em and Latin/digits half an
    em in both Word and LibreOffice. The estimate only controls vertical
    placement; final delivery still requires one visual check of the PDF.
    """
    properties = paragraph.find(qn("pPr"))
    frame_width = _frame_width(paragraph, 8464)
    indent = properties.find(qn("ind")) if properties is not None else None
    try:
        first_line_indent = int(indent.attrib.get(qn("firstLine"), "0")) if indent is not None else 0
    except ValueError:
        first_line_indent = 0

    run_properties = paragraph.find(".//" + qn("rPr"))
    font_size = run_properties.find(qn("sz")) if run_properties is not None else None
    try:
        # w:sz is measured in half-points; one point is 20 twips.
        full_width = int(font_size.attrib.get(qn("val"), "28")) * 10 if font_size is not None else 280
    except ValueError:
        full_width = 280
    half_width = max(1, full_width // 2)

    lines = 1
    used = 0
    capacity = max(full_width, frame_width - first_line_indent)
    for character in text:
        advance = (
            full_width
            if unicodedata.east_asian_width(character) in {"W", "F", "A"}
            else half_width
        )
        if used and used + advance > capacity:
            lines += 1
            used = 0
            capacity = frame_width
        used += advance
    return lines


def _position_table_from_frame(
    table: ET.Element, paragraph: ET.Element, *, y_offset_twips: int = 0
) -> None:
    """Keep a template's floating name area while replacing its tab grid with a real table."""
    paragraph_properties = paragraph.find(qn("pPr"))
    frame = paragraph_properties.find(qn("framePr")) if paragraph_properties is not None else None
    table_properties = table.find(qn("tblPr"))
    if frame is None or table_properties is None:
        return
    position = _set_child(
        table_properties,
        "tblpPr",
        leftFromText="0",
        rightFromText="0",
        topFromText="0",
        bottomFromText="0",
        vertAnchor=frame.attrib.get(qn("vAnchor"), "page"),
        horzAnchor=frame.attrib.get(qn("hAnchor"), "page"),
    )
    if qn("x") in frame.attrib:
        position.set(qn("tblpX"), frame.attrib[qn("x")])
    if qn("y") in frame.attrib:
        try:
            position.set(qn("tblpY"), str(int(frame.attrib[qn("y")]) + y_offset_twips))
        except ValueError:
            position.set(qn("tblpY"), frame.attrib[qn("y")])


def _replace_paragraph_with_table(
    root: ET.Element, paragraph: ET.Element, table: ET.Element
) -> tuple[ET.Element, int]:
    body = root.find(".//" + qn("body"))
    if body is None:
        raise TemplateError("模板缺少正文。")
    children = list(body)
    index = children.index(paragraph)
    body.remove(paragraph)
    body.insert(index, table)
    return body, index


def _adjust_activity_spacer(body: ET.Element, table_index: int, name_rows: int) -> None:
    children = list(body)
    blanks: list[ET.Element] = []
    for child in children[table_index + 1 :]:
        if child.tag != qn("p") or paragraph_text(child).strip():
            break
        blanks.append(child)
    if not blanks:
        return
    # The original template reserves fourteen text lines between the marker and
    # the closing paragraph.  Reduce blank paragraphs as the name grid grows so
    # the issuer/date remain near the reference position on a one-page document.
    keep = max(1, 14 - name_rows)
    for paragraph in blanks[keep:]:
        body.remove(paragraph)


def _visible_document_text(raw_xml: bytes) -> str:
    root = ET.fromstring(raw_xml)
    return "".join(node.text or "" for node in root.iter(qn("t")))


def _validate_identity_only(raw_xml: bytes) -> None:
    visible = _visible_document_text(raw_xml)
    if "{{" in visible or "}}" in visible:
        raise TemplateError("生成结果仍含未替换的模板字段。")
    if SCORE_LANGUAGE_RE.search(visible):
        raise InputValidationError("生成结果意外包含加分或分值表述，已停止写入。")


def _sanitize_core_xml(raw: bytes) -> bytes:
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return raw
    ET.register_namespace("cp", CP_NS)
    ET.register_namespace("dc", DC_NS)
    ET.register_namespace("dcterms", DCTERMS_NS)
    removable = {
        "{" + DC_NS + "}creator",
        "{" + CP_NS + "}lastModifiedBy",
        "{" + DC_NS + "}title",
        "{" + DC_NS + "}subject",
        "{" + DC_NS + "}description",
        "{" + CP_NS + "}keywords",
        "{" + CP_NS + "}lastPrinted",
        "{" + DCTERMS_NS + "}created",
        "{" + DCTERMS_NS + "}modified",
    }
    for child in list(root):
        if child.tag in removable:
            root.remove(child)
    revision = root.find("{" + CP_NS + "}revision")
    if revision is not None:
        revision.text = "1"
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_custom_xml(raw: bytes) -> bytes:
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return raw
    ET.register_namespace("", CUSTOM_NS)
    ET.register_namespace("vt", VT_NS)
    for child in list(root):
        root.remove(child)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_app_xml(raw: bytes) -> bytes:
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return raw
    ET.register_namespace("", APP_NS)
    for local_name in ("Manager", "Company"):
        node = root.find("{" + APP_NS + "}" + local_name)
        if node is not None:
            node.text = ""
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _write_docx_package(
    source: Path, destination: Path, replacements: dict[str, bytes]
) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(source, "r") as package:
            infos = package.infolist()
            contents = {info.filename: package.read(info.filename) for info in infos}
            comment = package.comment
    except (zipfile.BadZipFile, KeyError) as exc:
        raise TemplateError("DOCX 模板损坏或缺少必要文件。") from exc

    for name, data in replacements.items():
        if name in contents:
            contents[name] = data
    if "docProps/core.xml" in contents:
        contents["docProps/core.xml"] = _sanitize_core_xml(contents["docProps/core.xml"])
    if "docProps/custom.xml" in contents:
        contents["docProps/custom.xml"] = _sanitize_custom_xml(contents["docProps/custom.xml"])
    if "docProps/app.xml" in contents:
        contents["docProps/app.xml"] = _sanitize_app_xml(contents["docProps/app.xml"])

    temp_handle = tempfile.NamedTemporaryFile(
        prefix=destination.stem + "-", suffix=".docx", dir=destination.parent, delete=False
    )
    temp_path = Path(temp_handle.name)
    temp_handle.close()
    try:
        with zipfile.ZipFile(temp_path, "w") as output:
            output.comment = comment
            for info in infos:
                output.writestr(info, contents[info.filename])
        os.replace(temp_path, destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def prepare_template(source: Path, destination: Path, kind: str) -> None:
    try:
        with zipfile.ZipFile(source) as package:
            raw = package.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        raise TemplateError("原始 DOCX 模板损坏。") from exc
    root, start_tag = _parse_document(raw)

    body = _find_paragraph(root, lambda text: "兹证明" in text, "证明正文")
    names = _find_paragraph(root, lambda text: "从此插入" in text, "名单位置")
    if kind == "activity":
        date = _find_paragraph(
            root,
            lambda text: bool(re.search(r"20\d{2}\s*年", unicodedata.normalize("NFKC", text)))
            and "月" in unicodedata.normalize("NFKC", text),
            "出具日期",
        )
        set_paragraph_text(body, "兹证明以下同学为数统学院{{academic_year}}“{{activity_name}}”{{activity_suffix}}{{role}}：")
        set_paragraph_text(names, "{{names_table}}")
        set_paragraph_text(date, "{{issue_date}}")
        closing = _find_paragraph(
            root, lambda text: "感谢以上同学对学院活动的支持与付出" in text, "活动证明结语"
        )
        set_paragraph_text(closing, "感谢以上同学对学院活动的支持与付出，特此证明。")
    elif kind == "outstanding":
        date = _find_paragraph(
            root,
            lambda text: bool(re.search(r"20\d{2}\s*年", unicodedata.normalize("NFKC", text)))
            and "月" in unicodedata.normalize("NFKC", text),
            "出具日期",
        )
        set_paragraph_text(body, "兹证明以下同学在北京工商大学数学与统计学院{{academic_year}}“{{activity_name}}”{{activity_suffix}}表现突出，特此授予突出贡献证明：")
        set_paragraph_text(names, "{{names_table}}")
        set_paragraph_text(date, "{{issue_date}}")
        issuer_second = _find_paragraph(root, lambda text: "数学与统计学院委员会" in text, "落款第二行")
        body_element = root.find(".//" + qn("body"))
        if body_element is None:
            raise TemplateError("模板缺少正文。")
        body_children = list(body_element)
        second_index = body_children.index(issuer_second)
        issuer_first = next(
            (child for child in reversed(body_children[:second_index]) if child.tag == qn("p") and paragraph_text(child).strip()),
            None,
        )
        if issuer_first is None:
            raise TemplateError("模板中落款第一行锚点缺失。")
        set_paragraph_text(issuer_first, "共青团北京工商大学")
        set_paragraph_text(issuer_second, "数学与统计学院委员会")
        closing = _find_paragraph(root, lambda text: "特颁此证" in text, "荣誉证书结语")
        closing_properties = closing.find(qn("pPr"))
        closing_frame = (
            closing_properties.find(qn("framePr"))
            if closing_properties is not None
            else None
        )
        if closing_frame is not None:
            # The source frame is only 548 twips high. Songti SC and common
            # Office substitutes need more line box height or the glyphs clip.
            closing_frame.set(qn("h"), "800")
            closing_frame.set(qn("hRule"), "exact")
    else:
        raise ValueError(kind)

    _apply_global_document_font(root)
    document_xml = _serialize_document(root, start_tag)
    _write_docx_package(source, destination, {"word/document.xml": document_xml})


def prepare_assets(activity_source: Path, outstanding_source: Path, asset_dir: Path) -> None:
    prepare_template(activity_source, asset_dir / "activity-certificate-template.docx", "activity")
    prepare_template(
        outstanding_source,
        asset_dir / "outstanding-contribution-template.docx",
        "outstanding",
    )


def _xlsx_inline_cell(reference: str, value: str, style: int = 0) -> ET.Element:
    cell = ET.Element(S + "c", {"r": reference, "t": "inlineStr"})
    if style:
        cell.set("s", str(style))
    inline = ET.SubElement(cell, S + "is")
    text_node = ET.SubElement(inline, S + "t")
    text_node.text = value
    return cell


def _xlsx_sheet_xml(rows: Sequence[Sequence[str]], widths: Sequence[float]) -> bytes:
    ET.register_namespace("", SHEET_NS)
    worksheet = ET.Element(S + "worksheet")
    views = ET.SubElement(worksheet, S + "sheetViews")
    view = ET.SubElement(views, S + "sheetView", {"workbookViewId": "0"})
    ET.SubElement(
        view,
        S + "pane",
        {"ySplit": "1", "topLeftCell": "A2", "activePane": "bottomLeft", "state": "frozen"},
    )
    columns = ET.SubElement(worksheet, S + "cols")
    for index, width in enumerate(widths, start=1):
        ET.SubElement(
            columns,
            S + "col",
            {"min": str(index), "max": str(index), "width": str(width), "customWidth": "1"},
        )
    sheet_data = ET.SubElement(worksheet, S + "sheetData")
    for row_number, values in enumerate(rows, start=1):
        row = ET.SubElement(sheet_data, S + "row", {"r": str(row_number)})
        for column_index, value in enumerate(values):
            reference = _column_index_to_letters(column_index) + str(row_number)
            row.append(_xlsx_inline_cell(reference, value, style=1 if row_number == 1 else 0))
    if rows and rows[0]:
        ET.SubElement(
            worksheet,
            S + "autoFilter",
            {"ref": f"A1:{_column_index_to_letters(len(rows[0]) - 1)}1"},
        )
    ET.SubElement(
        worksheet,
        S + "pageMargins",
        {"left": "0.3", "right": "0.3", "top": "0.5", "bottom": "0.5", "header": "0.2", "footer": "0.2"},
    )
    return ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def _column_index_to_letters(index: int) -> str:
    result = ""
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def create_approved_list_template(destination: Path) -> None:
    """Create a blank, privacy-minimized XLSX intake template."""

    headers = ["内部人员键", "姓名", "证明类型", "活动", "角色", "批准状态"]
    instructions = [
        ["证明批准名单填写说明"],
        ["仅填写已逐级审核、可以制证的人员；生成器只会把姓名写入证明。"],
        ["允许列", "内部人员键、姓名、证明类型、活动、角色、批准状态"],
        ["批准状态", "统一填写“已批准”；存在其他状态时生成器会停止，不会自动筛掉。"],
        ["禁止列", "分值、加分、请假、缺勤、出勤次数、奖励方案、公式、评分等内部认定字段。"],
        ["名单校验", "空姓名、重复姓名、异常字符、未批准项均会阻止生成。"],
        ["证明正文", "只证明参与或突出贡献身份，不写具体加分数。"],
    ]

    content_types = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''.encode("utf-8")
    root_rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''.encode("utf-8")
    workbook = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{SHEET_NS}" xmlns:r="{R_NS}">
  <sheets>
    <sheet name="证明批准名单" sheetId="1" r:id="rId1"/>
    <sheet name="填写说明" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>'''.encode("utf-8")
    workbook_rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''.encode("utf-8")
    styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{SHEET_NS}">
  <fonts count="2"><font><sz val="11"/><name val="等线"/></font><font><b/><sz val="11"/><name val="等线"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''.encode("utf-8")
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="{CP_NS}" xmlns:dc="{DC_NS}" xmlns:dcterms="{DCTERMS_NS}"/>'''.encode("utf-8")
    app = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="{APP_NS}"><Application>团委办公室证书工具</Application></Properties>'''.encode("utf-8")

    parts = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": root_rels,
        "docProps/core.xml": core,
        "docProps/app.xml": app,
        "xl/workbook.xml": workbook,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/worksheets/sheet1.xml": _xlsx_sheet_xml([headers], [16, 14, 15, 30, 18, 15]),
        "xl/worksheets/sheet2.xml": _xlsx_sheet_xml(instructions, [20, 80]),
        "xl/styles.xml": styles,
    }
    destination.parent.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(
        prefix=destination.stem + "-", suffix=".xlsx", dir=destination.parent, delete=False
    )
    temp_path = Path(handle.name)
    handle.close()
    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as package:
            for name, data in parts.items():
                package.writestr(name, data)
        os.replace(temp_path, destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def _load_template_document(template: Path) -> tuple[ET.Element, str]:
    if not template.is_file():
        raise TemplateError("证书模板不存在。")
    try:
        with zipfile.ZipFile(template) as package:
            raw = package.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        raise TemplateError("证书模板损坏。") from exc
    return _parse_document(raw)


def build_certificate_docx(
    *,
    kind: str,
    names: Sequence[NameRecord],
    academic_year: str,
    activity_name: str,
    role: str,
    issue_date: str,
    template: Path,
    output: Path,
) -> tuple[int, int]:
    root, start_tag = _load_template_document(template)
    columns, rows = _grid_dimensions(names)
    body_paragraph = _find_paragraph(root, lambda text: "{{academic_year}}" in text, "证明正文")
    names_paragraph = _find_paragraph(root, lambda text: "{{names_table}}" in text, "名单位置")
    date_paragraph = _find_paragraph(root, lambda text: "{{issue_date}}" in text, "出具日期")

    if kind == "activity":
        statement = _activity_statement(academic_year, activity_name, role)
        set_paragraph_text(body_paragraph, statement)
        clear_first_line_indent(body_paragraph)
        table = make_borderless_name_table(names, columns, _page_content_width(root))
        validate_equal_width_name_table(table, columns)
        body, table_index = _replace_paragraph_with_table(root, names_paragraph, table)
        _adjust_activity_spacer(body, table_index, rows)
        closing = _find_paragraph(
            root, lambda text: "感谢以上同学对学院活动的支持与付出" in text, "活动证明结语"
        )
        set_paragraph_text(closing, "感谢以上同学对学院活动的支持与付出，特此证明。")
    elif kind == "outstanding":
        statement = _outstanding_statement(academic_year, activity_name)
        set_paragraph_text(body_paragraph, statement)
        table = make_borderless_name_table(
            names, columns, _frame_width(names_paragraph, _page_content_width(root))
        )
        # Keep one complete blank line between the statement and the names.
        # The template's base offset is correct for two prose lines; each extra
        # wrapped line moves the floating table down by the exact line height.
        prose_lines = _estimated_paragraph_lines(body_paragraph, statement)
        extra_lines = max(0, prose_lines - OUTSTANDING_REFERENCE_TEXT_LINES)
        table_offset = OUTSTANDING_BASE_TABLE_Y_OFFSET_TWIPS + extra_lines * _paragraph_line_height(
            body_paragraph, OUTSTANDING_FALLBACK_LINE_HEIGHT_TWIPS
        )
        _position_table_from_frame(table, names_paragraph, y_offset_twips=table_offset)
        validate_equal_width_name_table(table, columns)
        _replace_paragraph_with_table(root, names_paragraph, table)
    else:
        raise InputValidationError("证书类型只能是 activity 或 outstanding。")
    set_paragraph_text(date_paragraph, issue_date)

    _apply_global_document_font(root)
    document_xml = _serialize_document(root, start_tag)
    _validate_identity_only(document_xml)
    _write_docx_package(template, output, {"word/document.xml": document_xml})
    return columns, rows


def _safe_filename_component(value: str) -> str:
    value = re.sub(r"[\\/:*?\"<>|\x00-\x1f]", "-", value).strip(" .")
    value = re.sub(r"\s+", "", value)
    return (value[:80] or "证书")


def _find_soffice() -> Path | None:
    located = shutil.which("soffice") or shutil.which("libreoffice")
    candidates = [
        Path(located) if located else None,
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ]
    return next((candidate for candidate in candidates if candidate and candidate.is_file()), None)


def convert_to_pdf(docx_path: Path, pdf_path: Path, soffice: Path) -> None:
    with (
        tempfile.TemporaryDirectory(prefix="certificate-pdf-") as output_dir,
        tempfile.TemporaryDirectory(prefix="certificate-lo-profile-") as profile_dir,
        tempfile.TemporaryDirectory(prefix="certificate-fontconfig-") as fontconfig_dir,
    ):
        fontconfig_root = ET.Element("fontconfig")
        font_directories = [
            Path("/System/Library/Fonts"),
            Path("/System/Library/Fonts/Supplemental"),
            Path("/Library/Fonts"),
            Path.home() / "Library" / "Fonts",
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
        ]
        for directory in font_directories:
            if directory.is_dir():
                node = ET.SubElement(fontconfig_root, "dir")
                node.text = str(directory)
        cache_dir = Path(fontconfig_dir) / "cache"
        cache_dir.mkdir()
        cache_node = ET.SubElement(fontconfig_root, "cachedir")
        cache_node.text = str(cache_dir)
        fontconfig_path = Path(fontconfig_dir) / "fonts.conf"
        fontconfig_path.write_bytes(
            ET.tostring(fontconfig_root, encoding="utf-8", xml_declaration=True)
        )
        environment = os.environ.copy()
        environment["FONTCONFIG_FILE"] = str(fontconfig_path)
        command = [
            str(soffice),
            f"-env:UserInstallation={Path(profile_dir).as_uri()}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            output_dir,
            str(docx_path),
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
                env=environment,
            )
        except subprocess.TimeoutExpired as exc:
            raise CertificateError("PDF 转换超时；DOCX 已保留，请手动导出 PDF。") from exc
        generated = Path(output_dir) / (docx_path.stem + ".pdf")
        if completed.returncode != 0 or not generated.is_file():
            raise CertificateError("LibreOffice 未能导出 PDF；DOCX 已保留，请手动导出。")
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        os.replace(generated, pdf_path)


def generate_certificate(
    *,
    kind: str,
    names_path: Path,
    academic_year: str,
    activity_name: str,
    issue_date: str,
    output_dir: Path,
    role: str = "参与人员",
    name_column: str = "auto",
    sheet_name: str | None = None,
    template: Path | None = None,
    filename: str | None = None,
    make_pdf: bool = True,
    overwrite: bool = False,
) -> GenerationResult:
    if kind not in DEFAULT_TEMPLATES:
        raise InputValidationError("证书类型只能是 activity 或 outstanding。")
    records = read_name_records(names_path, name_column, sheet_name)
    normalized_year = normalize_academic_year(academic_year)
    normalized_activity = normalize_activity_name(activity_name)
    normalized_date = normalize_issue_date(issue_date)
    normalized_role = normalize_role(role) if kind == "activity" else ""
    context_role = normalized_role if kind == "activity" else "突出贡献"
    validate_roster_context(
        records,
        kind=kind,
        activity_name=normalized_activity,
        role=context_role,
    )
    names = validate_and_sort_names(records)

    chosen_template = template or DEFAULT_TEMPLATES[kind]
    label = "活动证明" if kind == "activity" else "突出贡献证明"
    if filename:
        filename_path = Path(filename)
        if filename_path.name != filename:
            raise InputValidationError("--filename 只能是文件名，不能包含目录。")
        docx_name = filename if filename.casefold().endswith(".docx") else filename + ".docx"
    else:
        docx_name = _safe_filename_component(normalized_activity) + label + ".docx"
    output_dir.mkdir(parents=True, exist_ok=True)
    docx_path = output_dir / docx_name
    pdf_path = docx_path.with_suffix(".pdf") if make_pdf else None

    existing = [path for path in (docx_path, pdf_path) if path is not None and path.exists()]
    if existing and not overwrite:
        raise CertificateError("目标文件已存在；请更换 --filename，或明确使用 --overwrite。")
    soffice = _find_soffice() if make_pdf else None
    if make_pdf and soffice is None:
        raise CertificateError(
            "未找到 LibreOffice/soffice；为避免只生成单一格式，已停止制证。"
            "请配置 PDF 导出环境，或仅在诊断时使用 --no-pdf。"
        )

    columns, rows = build_certificate_docx(
        kind=kind,
        names=names,
        academic_year=normalized_year,
        activity_name=normalized_activity,
        role=normalized_role,
        issue_date=normalized_date,
        template=chosen_template,
        output=docx_path,
    )
    if make_pdf and pdf_path is not None and soffice is not None:
        convert_to_pdf(docx_path, pdf_path, soffice)
    return GenerationResult(docx_path, pdf_path, len(names), columns, rows)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="按办公室模板生成活动证明或突出贡献证明；正文不写任何加分数。"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    generate = subparsers.add_parser("generate", help="从名单生成证书")
    generate.add_argument("--type", required=True, choices=("activity", "outstanding"), dest="kind")
    generate.add_argument("--names", required=True, type=Path, help="TXT/CSV/TSV/XLSX 名单")
    generate.add_argument("--name-column", default="auto", help="姓名列：auto、表头、1-based 列号或 A/B/C")
    generate.add_argument("--sheet", help="XLSX 工作表名；默认读取第一张可见表")
    generate.add_argument("--academic-year", required=True, help="例如 2025-2026")
    generate.add_argument("--activity", required=True, help="活动名称；不要写加分信息")
    generate.add_argument("--role", default="参与人员", help="活动证明中的身份，默认“参与人员”")
    generate.add_argument("--issue-date", required=True, help="YYYY-MM-DD 或 YYYY-MM")
    generate.add_argument("--output-dir", type=Path, default=Path.cwd())
    generate.add_argument("--filename", help="输出 DOCX 文件名，不含目录")
    generate.add_argument("--template", type=Path, help="覆盖内置模板")
    pdf_options = generate.add_mutually_exclusive_group()
    pdf_options.add_argument(
        "--pdf",
        dest="make_pdf",
        action="store_true",
        help="同时导出 PDF（默认）；必须逐页检查中文字体、均分列宽与版式",
    )
    pdf_options.add_argument(
        "--no-pdf",
        dest="make_pdf",
        action="store_false",
        help="仅用于诊断，只生成 DOCX；正常证明任务不得以此作为完整交付",
    )
    generate.set_defaults(make_pdf=True)
    generate.add_argument("--overwrite", action="store_true", help="允许覆盖同名输出")

    prepare = subparsers.add_parser("prepare-assets", help="将办公室原模板清理并变量化为 skill assets")
    prepare.add_argument("--activity-source", required=True, type=Path)
    prepare.add_argument("--outstanding-source", required=True, type=Path)
    prepare.add_argument("--asset-dir", type=Path, default=ASSET_DIR)

    approved = subparsers.add_parser(
        "create-approved-list-template", help="生成空白的“证明批准名单”XLSX 模板"
    )
    approved.add_argument(
        "--output",
        type=Path,
        default=ASSET_DIR / "certificate-approved-list-template.xlsx",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    known_commands = {
        "generate",
        "prepare-assets",
        "create-approved-list-template",
        "-h",
        "--help",
    }
    if arguments and arguments[0] not in known_commands:
        # Convenience form: allow `generate_certificate.py --type ...`.
        arguments.insert(0, "generate")
    parser = _build_parser()
    args = parser.parse_args(arguments)
    try:
        if args.command == "prepare-assets":
            prepare_assets(args.activity_source, args.outstanding_source, args.asset_dir)
            print("模板已变量化并清理元数据：2 份（未读取或输出名单）。")
            return 0
        if args.command == "create-approved-list-template":
            create_approved_list_template(args.output)
            print("已生成空白“证明批准名单”模板；不含真实姓名或内部认定字段。")
            return 0

        result = generate_certificate(
            kind=args.kind,
            names_path=args.names,
            academic_year=args.academic_year,
            activity_name=args.activity,
            issue_date=args.issue_date,
            output_dir=args.output_dir,
            role=args.role,
            name_column=args.name_column,
            sheet_name=args.sheet,
            template=args.template,
            filename=args.filename,
            make_pdf=args.make_pdf,
            overwrite=args.overwrite,
        )
        print(
            f"生成完成：{result.name_count} 人，{result.columns} 列 × {result.rows} 行；"
            "名单表格为固定布局、均分列宽；日志未显示姓名，证明正文未写加分数。"
        )
        print(f"DOCX：{result.docx_path.resolve()}")
        if result.pdf_path is not None:
            print(f"PDF：{result.pdf_path.resolve()}")
            print("PDF 已同步导出；盖章或发群前必须逐页检查中文字体、均分列宽、分页和版式。")
        return 0
    except CertificateError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"文件操作失败：{exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
