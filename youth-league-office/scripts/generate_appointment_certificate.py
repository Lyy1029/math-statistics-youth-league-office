#!/usr/bin/env python3
"""Batch-generate staff appointment proofs and individual leader certificates.

Staff proofs retain an editable DOCX and a PDF.  Leader certificates are drawn
on the fixed artwork recovered from the approved PDF reference and are
delivered directly as PDFs, avoiding Word layout drift.
"""

from __future__ import annotations

import argparse
import base64
import copy
import csv
import io
import os
import re
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from generate_certificate import (
    APP_NS,
    APPROVED_STATUS_VALUES,
    ASSET_DIR,
    BUNDLED_FONT_DIR,
    CONTENT_TYPES_NS,
    CP_NS,
    DC_NS,
    DCTERMS_NS,
    R_NS,
    SHEET_NS,
    S,
    W,
    XML_NS,
    CertificateError,
    InputValidationError,
    TemplateError,
    _apply_global_document_font,
    _column_index_to_letters,
    _decode_text_file,
    _find_paragraph,
    _find_soffice,
    _format_row_numbers,
    _letters_to_index,
    _load_template_document,
    _normalize_header,
    _safe_filename_component,
    _safe_text,
    _serialize_document,
    _trim_outer_blank_rows,
    _validate_bundled_font_asset,
    _write_docx_package,
    _xlsx_cell_value,
    _xlsx_inline_cell,
    _xlsx_shared_strings,
    _xlsx_sheet_path,
    normalize_academic_year,
    normalize_issue_date,
    paragraph_text,
    qn,
    set_paragraph_text,
)


STAFF_TEMPLATE = ASSET_DIR / "staff-appointment-template.docx"
LEADER_BACKGROUND = ASSET_DIR / "leader-appointment-background.png"
DEFAULT_LIST_TEMPLATE = ASSET_DIR / "appointment-approved-list-template.xlsx"
MAX_FIELD_LENGTH = 80
PDF_BATCH_SIZE = 60

HEADER_ALIASES = {
    "internal_key": {"内部人员键", "内部编号"},
    "name": {"姓名", "学生姓名"},
    "class_name": {"班级"},
    "student_id": {"学号"},
    "organization": {"组织", "学生组织"},
    "department": {"部门"},
    "position": {"职务", "任职职务"},
    "kind": {"任职证明类型", "证明类型"},
    "term_start": {"任期开始", "聘期开始"},
    "term_end": {"任期结束", "聘期结束"},
    "approval": {"批准状态", "审核状态"},
}
REQUIRED_HEADERS = {
    "name",
    "class_name",
    "student_id",
    "organization",
    "department",
    "position",
    "kind",
    "term_start",
    "term_end",
    "approval",
}
KIND_ALIASES = {
    "staff": {"staff", "干事", "干事任职证明", "干事证明"},
    "leader": {
        "leader",
        "部长及以上",
        "部长及以上任职证明",
        "部长",
        "主席团",
        "主席团成员",
        "个人聘书",
        "聘请证书",
    },
}


@dataclass(frozen=True)
class AppointmentRecord:
    source_row: int
    name: str
    class_name: str
    student_id: str
    organization: str
    department: str
    position: str
    kind: str
    term_start: str
    term_end: str


@dataclass(frozen=True)
class AppointmentResult:
    output_dir: Path
    staff_proof_count: int
    leader_certificate_count: int
    person_count: int
    pdf_count: int


def _plain(value: str, label: str, row: int, *, required: bool = True) -> str:
    normalized = unicodedata.normalize("NFKC", value).strip()
    if required and not normalized:
        raise InputValidationError(f"{label}为空，位于行 {row}。")
    if len(normalized) > MAX_FIELD_LENGTH or any(ord(char) < 32 for char in normalized):
        raise InputValidationError(f"{label}格式异常，位于行 {row}。")
    if normalized.startswith(("=", "+", "@")):
        raise InputValidationError(f"{label}含公式式内容，位于行 {row}。")
    return normalized


def _name(value: str, row: int) -> str:
    normalized = re.sub(r"\s+", "", unicodedata.normalize("NFKC", value)).strip()
    allowed = all(char.isalpha() or char in {"·", "•", "-", "'", "’"} for char in normalized)
    if len(normalized) < 2 or len(normalized) > 30 or not allowed:
        raise InputValidationError(f"姓名格式异常，位于行 {row}。")
    return normalized


def _kind(value: str, row: int) -> str:
    normalized = _normalize_header(value)
    for canonical, aliases in KIND_ALIASES.items():
        if normalized in {_normalize_header(alias) for alias in aliases}:
            return canonical
    raise InputValidationError(
        f"任职证明类型无法识别，位于行 {row}；请填写“干事”或“部长及以上”。"
    )


def _term(value: str, label: str, row: int, *, required: bool) -> tuple[str, tuple[int, int] | None]:
    normalized = unicodedata.normalize("NFKC", value).strip()
    if not normalized:
        if required:
            raise InputValidationError(f"{label}为空，位于行 {row}。")
        return "", None
    match = re.fullmatch(r"(20\d{2})\s*(?:[-/.年]\s*(0?[1-9]|1[0-2]))\s*月?", normalized)
    if not match:
        raise InputValidationError(f"{label}格式错误，位于行 {row}；请用 YYYY-MM。")
    year, month = int(match.group(1)), int(match.group(2))
    return f"{year}年{month}月", (year, month)


def _csv_rows(path: Path) -> list[tuple[int, list[str]]]:
    text = _decode_text_file(path)
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=",\t;，")
    except csv.Error:
        dialect = csv.excel
    return [
        (index, [_safe_text(cell) for cell in row])
        for index, row in enumerate(csv.reader(io.StringIO(text), dialect), start=1)
    ]


def _xlsx_rows(path: Path, sheet_name: str | None) -> list[tuple[int, list[str]]]:
    try:
        with zipfile.ZipFile(path) as package:
            shared = _xlsx_shared_strings(package)
            sheet_path = _xlsx_sheet_path(package, sheet_name)
            sheet = ET.fromstring(package.read(sheet_path))
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        raise InputValidationError("XLSX 文件损坏或不是有效的 OOXML 工作簿。") from exc
    if any(True for _ in sheet.iter(S + "f")):
        raise InputValidationError("任职批准名单含公式；请粘贴为值后重试。")
    sheet_data = sheet.find(S + "sheetData")
    if sheet_data is None:
        raise InputValidationError("XLSX 工作表为空。")
    rows: list[tuple[int, list[str]]] = []
    for ordinal, row in enumerate(sheet_data.findall(S + "row"), start=1):
        try:
            row_number = int(row.attrib.get("r", ordinal))
        except ValueError:
            row_number = ordinal
        cells: dict[int, str] = {}
        for cell in row.findall(S + "c"):
            match = re.match(r"([A-Za-z]+)", cell.attrib.get("r", ""))
            column = _letters_to_index(match.group(1)) if match else len(cells)
            cells[column] = _xlsx_cell_value(cell, shared)
        width = max(cells) + 1 if cells else 0
        rows.append((row_number, [cells.get(index, "") for index in range(width)]))
    return rows


def _header_map(values: Sequence[str]) -> dict[str, int]:
    aliases = {
        _normalize_header(alias): canonical
        for canonical, names in HEADER_ALIASES.items()
        for alias in names
    }
    mapping: dict[str, int] = {}
    unsupported: list[str] = []
    for index, value in enumerate(values):
        normalized = _normalize_header(value)
        if not normalized:
            continue
        canonical = aliases.get(normalized)
        if canonical is None:
            unsupported.append(value.strip())
            continue
        if canonical in mapping:
            raise InputValidationError("任职批准名单含重复表头。")
        mapping[canonical] = index
    if unsupported:
        raise InputValidationError("任职批准名单含不支持的扩展字段；请使用内置空白模板。")
    missing = REQUIRED_HEADERS - set(mapping)
    if missing:
        labels = {
            key: next(iter(HEADER_ALIASES[key])) for key in sorted(missing)
        }
        raise InputValidationError("任职批准名单缺少必需表头：" + "、".join(labels.values()) + "。")
    return mapping


def read_appointment_records(path: Path, sheet_name: str | None = None) -> list[AppointmentRecord]:
    if not path.is_file():
        raise InputValidationError("任职批准名单不存在或不是普通文件。")
    suffix = path.suffix.casefold()
    if suffix in {".csv", ".tsv"}:
        rows = _csv_rows(path)
    elif suffix == ".xlsx":
        rows = _xlsx_rows(path, sheet_name)
    else:
        raise InputValidationError("任职批准名单仅支持 CSV、TSV 或 XLSX。")
    rows = _trim_outer_blank_rows(rows)
    if len(rows) < 2:
        raise InputValidationError("任职批准名单没有数据。")
    mapping = _header_map(rows[0][1])
    records: list[AppointmentRecord] = []
    unapproved: list[int] = []
    incomplete: list[int] = []
    for row_number, values in rows[1:]:
        if not any(cell.strip() for cell in values):
            incomplete.append(row_number)
            continue
        get = lambda key: values[mapping[key]] if mapping[key] < len(values) else ""
        approval = _normalize_header(get("approval"))
        if approval not in APPROVED_STATUS_VALUES:
            unapproved.append(row_number)
            continue
        kind = _kind(get("kind"), row_number)
        name = _name(get("name"), row_number)
        organization = _plain(get("organization"), "组织", row_number)
        department = _plain(get("department"), "部门", row_number, required=kind == "staff")
        position = _plain(get("position"), "职务", row_number)
        class_name = _plain(get("class_name"), "班级", row_number, required=kind == "staff")
        student_id = _plain(get("student_id"), "学号", row_number, required=kind == "staff")
        if kind == "staff":
            if "干事" not in position:
                raise InputValidationError(
                    f"干事任职证明中出现非干事职务，位于行 {row_number}；部长及以上请改用对应类型。"
                )
            if not re.fullmatch(r"[0-9A-Za-z-]{4,30}", student_id):
                raise InputValidationError(f"学号格式异常，位于行 {row_number}。")
        term_start, start_key = _term(
            get("term_start"), "任期开始", row_number, required=kind == "leader"
        )
        term_end, end_key = _term(
            get("term_end"), "任期结束", row_number, required=kind == "leader"
        )
        if start_key is not None and end_key is not None and start_key >= end_key:
            raise InputValidationError(f"任期结束时间必须晚于开始时间，位于行 {row_number}。")
        records.append(
            AppointmentRecord(
                source_row=row_number,
                name=name,
                class_name=class_name,
                student_id=student_id,
                organization=organization,
                department=department,
                position=position,
                kind=kind,
                term_start=term_start,
                term_end=term_end,
            )
        )
    if incomplete:
        raise InputValidationError(
            "发现空白数据行 " + _format_row_numbers(incomplete) + "；为避免漏开，已停止生成。"
        )
    if unapproved:
        raise InputValidationError(
            "发现未明确批准的条目，位于行 " + _format_row_numbers(unapproved) + "；不会自动筛掉。"
        )
    if not records:
        raise InputValidationError("任职批准名单没有可生成记录。")
    duplicate_rows: list[int] = []
    seen: dict[tuple[str, ...], int] = {}
    for record in records:
        key = (
            record.kind,
            record.name,
            record.student_id,
            record.organization,
            record.department,
            record.position,
        )
        if key in seen:
            duplicate_rows.extend((seen[key], record.source_row))
        else:
            seen[key] = record.source_row
    if duplicate_rows:
        raise InputValidationError(
            "发现重复任职记录，位于行 " + _format_row_numbers(sorted(set(duplicate_rows))) + "。"
        )
    return records


def _set_cell_text(cell: ET.Element, text: str) -> None:
    paragraphs = cell.findall(qn("p"))
    if not paragraphs:
        paragraph = ET.SubElement(cell, qn("p"))
    else:
        paragraph = paragraphs[0]
        for extra in paragraphs[1:]:
            cell.remove(extra)
    set_paragraph_text(paragraph, text)


def _staff_table(root: ET.Element, records: Sequence[AppointmentRecord]) -> None:
    matches = [
        table for table in root.iter(qn("tbl")) if "{{staff_rows}}" in paragraph_text(table)
    ]
    if len(matches) != 1:
        raise TemplateError(f"干事模板名单锚点数量应为 1，实际为 {len(matches)}。")
    table = matches[0]
    rows = table.findall(qn("tr"))
    placeholder = next((row for row in rows if "{{staff_rows}}" in paragraph_text(row)), None)
    if placeholder is None:
        raise TemplateError("干事模板缺少名单占位行。")
    index = list(table).index(placeholder)
    table.remove(placeholder)
    for offset, record in enumerate(records):
        row = copy.deepcopy(placeholder)
        cells = row.findall(qn("tc"))
        if len(cells) != 4:
            raise TemplateError("干事模板名单表必须为 4 列。")
        for cell, value in zip(
            cells, (record.name, record.class_name, record.position, record.student_id)
        ):
            _set_cell_text(cell, value)
        table.insert(index + offset, row)


def build_staff_docx(
    records: Sequence[AppointmentRecord], academic_year: str, issue_date: str, output: Path
) -> None:
    root, start_tag = _load_template_document(STAFF_TEMPLATE)
    first = records[0]
    body = _find_paragraph(root, lambda text: "{{academic_year}}" in text, "干事任职正文")
    date = _find_paragraph(root, lambda text: "{{issue_date}}" in text, "出具日期")
    statement = (
        f"兹证明以下学生于{academic_year}，在数学与统计学院"
        f"{first.organization}{first.department}任职。"
    )
    set_paragraph_text(body, statement)
    set_paragraph_text(date, issue_date)
    _staff_table(root, records)
    _apply_global_document_font(root)
    _write_docx_package(
        STAFF_TEMPLATE,
        output,
        {"word/document.xml": _serialize_document(root, start_tag)},
    )


def _leader_issue_date(value: str) -> str:
    match = re.fullmatch(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", value)
    if match is None:
        raise InputValidationError("出具日期格式异常。")
    return f"{match.group(1)}  年 {int(match.group(2))} 月 {int(match.group(3))} 日"


def _leader_role(record: AppointmentRecord) -> str:
    return f"{record.organization}{record.department}{record.position}"


def _leader_role_font_size(role: str) -> float:
    # The underlined role field runs to the comma at about x=396.  Keep the
    # reference size for ordinary roles and shrink only unusually long labels.
    size = min(17.474, 270.0 / max(len(role), 1))
    if size < 11.5:
        raise InputValidationError(
            f"部长及以上职务文字过长，位于行 {record.source_row}；请简化组织/部门/职务。"
        )
    return size


def _leader_name_font_size(name: str, row: int) -> float:
    size = min(17.474, 54.0 / max(len(name), 1))
    if size < 12.0:
        raise InputValidationError(
            f"部长及以上姓名无法放入原稿姓名栏，位于行 {row}。"
        )
    return size


def build_leader_svg(record: AppointmentRecord, issue_date: str, output: Path) -> None:
    """Create a one-page vector overlay on the exact approved PDF artwork."""
    if not LEADER_BACKGROUND.is_file():
        raise TemplateError("缺少部长及以上证书 PDF 背景资产。")
    background = base64.b64encode(LEADER_BACKGROUND.read_bytes()).decode("ascii")
    role = _leader_role(record)
    role_size = _leader_role_font_size(role)
    name_size = _leader_name_font_size(record.name, record.source_row)
    values = {
        "name": escape(record.name),
        "role": escape(role),
        "start": escape(record.term_start),
        "end": escape(record.term_end),
        "date": escape(_leader_issue_date(issue_date)),
    }
    svg = f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
     width="595.5pt" height="842.25pt" viewBox="0 0 595.5 842.25">
  <image x="0" y="0" width="595.5" height="842.25" preserveAspectRatio="none"
         xlink:href="data:image/png;base64,{background}"/>
  <g font-family="Noto Serif SC" fill="#000000">
    <text x="90.387" y="342.270" font-size="17.474">兹聘用</text>
    <text x="166.266" y="342.270" font-size="{name_size:.3f}">{values['name']}</text>
    <text x="220.446" y="342.295" font-size="17.474">担任北京工商大学数学与统计学院</text>
    <text x="115.556" y="380.914" font-size="{role_size:.3f}">{values['role']}</text>
    <text x="281.627" y="383.017" font-size="17.474">，聘期一年，自{values['start']}至</text>
    <text x="66.383" y="436.563" font-size="17.474">{values['end']}止。</text>
    <text x="448.855" y="675.729" text-anchor="middle" font-size="14.585" letter-spacing="2">{values['date']}</text>
  </g>
</svg>
'''
    output.write_text(svg, encoding="utf-8")


def _fontconfig(temp_dir: Path) -> tuple[Path, dict[str, str]]:
    root = ET.Element("fontconfig")
    # Every visible run is normalized to the bundled font.  Limiting Fontconfig
    # to this directory avoids a slow first-use scan of an Agent's system fonts
    # and makes substitution impossible by construction.
    node = ET.SubElement(root, "dir")
    node.text = str(BUNDLED_FONT_DIR)
    cache = temp_dir / "cache"
    cache.mkdir()
    cache_node = ET.SubElement(root, "cachedir")
    cache_node.text = str(cache)
    config = temp_dir / "fonts.conf"
    config.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
    environment = os.environ.copy()
    environment["FONTCONFIG_FILE"] = str(config)
    return config, environment


def _chunks(values: Sequence[Path], size: int) -> Iterable[Sequence[Path]]:
    for start in range(0, len(values), size):
        yield values[start : start + size]


def convert_sources_to_pdf(
    pairs: Sequence[tuple[Path, Path]], soffice: Path
) -> list[Path]:
    _validate_bundled_font_asset()
    with (
        tempfile.TemporaryDirectory(prefix="appointment-pdf-") as pdf_dir,
        tempfile.TemporaryDirectory(prefix="appointment-lo-profile-") as profile_dir,
        tempfile.TemporaryDirectory(prefix="appointment-fontconfig-") as font_dir,
    ):
        _, environment = _fontconfig(Path(font_dir))
        sources = [source for source, _ in pairs]
        for chunk in _chunks(sources, PDF_BATCH_SIZE):
            command = [
                str(soffice),
                f"-env:UserInstallation={Path(profile_dir).as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                pdf_dir,
                *(str(path) for path in chunk),
            ]
            try:
                completed = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=180,
                    check=False,
                    env=environment,
                )
            except subprocess.TimeoutExpired as exc:
                raise CertificateError("任职证明批量 PDF 转换超时。") from exc
            if completed.returncode != 0:
                raise CertificateError("LibreOffice 未能批量导出任职证明 PDF。")
        final_paths: list[Path] = []
        for source, destination in pairs:
            generated = Path(pdf_dir) / (source.stem + ".pdf")
            if not generated.is_file():
                raise CertificateError("任职证明 PDF 数量不完整。")
            pdf_bytes = generated.read_bytes()
            families = {
                name.split(b"+")[-1]
                for name in re.findall(rb"/FontName\s*/([^\s/<>\[\]()]+)", pdf_bytes)
            }
            if not families or families != {b"NotoSerifSC-Regular"}:
                raise CertificateError(
                    "任职证明 PDF 检测到字体替换；PDF 未交付。"
                )
            os.replace(generated, destination)
            final_paths.append(destination)
        return final_paths


def generate_appointments(
    *,
    list_path: Path,
    academic_year: str,
    issue_date: str,
    output_dir: Path,
    sheet_name: str | None = None,
    make_pdf: bool = True,
    overwrite: bool = False,
) -> AppointmentResult:
    records = read_appointment_records(list_path, sheet_name)
    normalized_year = normalize_academic_year(academic_year)
    normalized_date = normalize_issue_date(issue_date)
    staff_groups: dict[tuple[str, str], list[AppointmentRecord]] = {}
    leaders: list[AppointmentRecord] = []
    for record in records:
        if record.kind == "staff":
            staff_groups.setdefault((record.organization, record.department), []).append(record)
        else:
            leaders.append(record)
    if leaders and not make_pdf:
        raise CertificateError("部长及以上证书按 PDF 原稿直接制作，不支持 --no-pdf。")
    planned: list[tuple[str, object, Path]] = []
    for (organization, department), group in staff_groups.items():
        filename = _safe_filename_component(organization + department) + "干事任职证明.docx"
        planned.append(("staff", group, output_dir / filename))
    for index, record in enumerate(leaders, start=1):
        label = _safe_filename_component(record.organization + record.department + record.position)
        filename = f"{index:03d}_{_safe_filename_component(record.name)}_{label}聘请证书.pdf"
        planned.append(("leader", record, output_dir / filename))
    targets = [path for _, _, path in planned]
    if make_pdf:
        targets.extend(path.with_suffix(".pdf") for kind, _, path in planned if kind == "staff")
    existing = [path for path in targets if path.exists()]
    if existing and not overwrite:
        raise CertificateError("目标目录存在同名任职证明；请更换目录，或明确使用 --overwrite。")
    soffice = _find_soffice() if make_pdf else None
    if make_pdf and soffice is None:
        raise CertificateError("未找到 LibreOffice/soffice；正常制证必须同时生成 PDF。")
    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_paths: list[Path] = []
    with tempfile.TemporaryDirectory(prefix="leader-certificate-svg-") as temp_dir:
        conversion_pairs: list[tuple[Path, Path]] = []
        for index, (kind, payload, path) in enumerate(planned, start=1):
            if kind == "staff":
                build_staff_docx(payload, normalized_year, normalized_date, path)
                if make_pdf:
                    conversion_pairs.append((path, path.with_suffix(".pdf")))
            else:
                svg_path = Path(temp_dir) / f"leader-{index:03d}.svg"
                build_leader_svg(payload, normalized_date, svg_path)
                conversion_pairs.append((svg_path, path))
        if conversion_pairs and soffice:
            pdf_paths = convert_sources_to_pdf(conversion_pairs, soffice)
    return AppointmentResult(
        output_dir=output_dir,
        staff_proof_count=len(staff_groups),
        leader_certificate_count=len(leaders),
        person_count=len(records),
        pdf_count=len(pdf_paths),
    )


def create_list_template(destination: Path) -> None:
    headers = [
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
    instructions = [
        ["任职证明批准名单填写说明"],
        ["一张表可同时生成干事批量任职证明和部长及以上个人聘请证书。"],
        ["干事", "班级、学号、组织、部门、职务必填；职务填写干事。"],
        ["部长及以上", "组织、职务、任期开始、任期结束必填；部门可留空。"],
        ["主席团", "主席、副主席、团委副书记等均使用“部长及以上”类型。"],
        ["日期格式", "任期开始和结束统一填写 YYYY-MM。"],
        ["批准状态", "统一填写“已批准”；其他状态会使整批停止。"],
    ]
    from generate_certificate import create_approved_list_template

    # Reuse the proven OOXML workbook writer, then replace its two sheet bodies
    # and workbook-facing labels without adding an openpyxl dependency.
    with tempfile.TemporaryDirectory(prefix="appointment-list-") as temp_dir:
        base = Path(temp_dir) / "base.xlsx"
        create_approved_list_template(base)
        with zipfile.ZipFile(base) as package:
            infos = package.infolist()
            parts = {info.filename: package.read(info.filename) for info in infos}
        parts["xl/worksheets/sheet1.xml"] = _sheet_xml(
            [headers], [16, 14, 14, 17, 20, 18, 16, 18, 14, 14, 15]
        )
        parts["xl/worksheets/sheet2.xml"] = _sheet_xml(instructions, [20, 86])
        parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
            "证明批准名单".encode(), "任职证明批准名单".encode()
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        handle = tempfile.NamedTemporaryFile(
            prefix=destination.stem + "-", suffix=".xlsx", dir=destination.parent, delete=False
        )
        temp_path = Path(handle.name)
        handle.close()
        try:
            with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as output:
                for info in infos:
                    output.writestr(info, parts[info.filename])
            os.replace(temp_path, destination)
        finally:
            if temp_path.exists():
                temp_path.unlink()


def _sheet_xml(rows: Sequence[Sequence[str]], widths: Sequence[float]) -> bytes:
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
    data = ET.SubElement(worksheet, S + "sheetData")
    for row_number, values in enumerate(rows, start=1):
        row = ET.SubElement(data, S + "row", {"r": str(row_number)})
        for column, value in enumerate(values):
            reference = _column_index_to_letters(column) + str(row_number)
            row.append(_xlsx_inline_cell(reference, value, style=1 if row_number == 1 else 0))
    if rows and rows[0]:
        ET.SubElement(
            worksheet,
            S + "autoFilter",
            {"ref": f"A1:{_column_index_to_letters(len(rows[0]) - 1)}1"},
        )
    return ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="一张批准名单快速生成干事批量任职证明和部长及以上个人聘请证书。"
    )
    sub = parser.add_subparsers(dest="command", required=True)
    generate = sub.add_parser("generate", help="批量生成任职证明")
    generate.add_argument("--list", required=True, type=Path, dest="list_path")
    generate.add_argument("--sheet", help="XLSX 工作表名；默认读取第一张可见表")
    generate.add_argument("--academic-year", required=True, help="例如 2025-2026")
    generate.add_argument("--issue-date", required=True, help="YYYY-MM-DD")
    generate.add_argument("--output-dir", required=True, type=Path)
    pdf = generate.add_mutually_exclusive_group()
    pdf.add_argument("--pdf", dest="make_pdf", action="store_true", help="同时生成 PDF（默认）")
    pdf.add_argument("--no-pdf", dest="make_pdf", action="store_false", help="仅用于诊断")
    generate.set_defaults(make_pdf=True)
    generate.add_argument("--overwrite", action="store_true")
    template = sub.add_parser("create-list-template", help="生成空白任职证明批准名单")
    template.add_argument("--output", type=Path, default=DEFAULT_LIST_TEMPLATE)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _build_parser().parse_args(sys.argv[1:] if argv is None else argv)
    try:
        if args.command == "create-list-template":
            create_list_template(args.output)
            print("已生成空白任职证明批准名单；不含真实人员信息。")
            return 0
        result = generate_appointments(
            list_path=args.list_path,
            academic_year=args.academic_year,
            issue_date=args.issue_date,
            output_dir=args.output_dir,
            sheet_name=args.sheet,
            make_pdf=args.make_pdf,
            overwrite=args.overwrite,
        )
        print(
            f"生成完成：{result.person_count} 人；干事批量证明 {result.staff_proof_count} 份，"
            f"部长及以上个人证书 {result.leader_certificate_count} 份，PDF {result.pdf_count} 份。"
        )
        print(f"输出目录：{result.output_dir.resolve()}")
        if result.pdf_count:
            print("日志未显示姓名；PDF 已检查嵌入 Noto Serif SC 且无替代字体。")
        else:
            print("日志未显示姓名；本次仅生成干事 DOCX，未执行 PDF 检查。")
        return 0
    except CertificateError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"文件操作失败：{exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
