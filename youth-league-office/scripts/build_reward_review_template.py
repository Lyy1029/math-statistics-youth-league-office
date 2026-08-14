#!/usr/bin/env python3
"""Build the privacy-safe internal reward eligibility review workbook.

The workbook is authored with the bundled @oai/artifact-tool runtime. This
script intentionally contains no real person data and does not read any source
folders. It only writes the requested asset plus temporary visual previews.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import re
import subprocess
import tempfile
import zipfile
from xml.etree import ElementTree


NODE = Path(
    "/Users/lyy29001/.cache/codex-runtimes/"
    "codex-primary-runtime/dependencies/node/bin/node"
)
NODE_MODULES = Path(
    "/Users/lyy29001/.cache/codex-runtimes/"
    "codex-primary-runtime/dependencies/node/node_modules"
)
PACKAGE_ROOT = Path(__file__).resolve().parents[1]
OUTPUT = PACKAGE_ROOT / "assets" / "reward-eligibility-review-template.xlsx"
BUILD_DIR = Path(tempfile.gettempdir()) / "reward_artifact_build"
MAIN_PREVIEW = Path(tempfile.gettempdir()) / "reward-review-main.png"
NOTES_PREVIEW = Path(tempfile.gettempdir()) / "reward-review-notes.png"


JAVASCRIPT = r'''
import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = process.env.REWARD_REVIEW_OUTPUT;
const mainPreviewPath = process.env.REWARD_REVIEW_MAIN_PREVIEW;
const notesPreviewPath = process.env.REWARD_REVIEW_NOTES_PREVIEW;

const workbook = Workbook.create();
const review = workbook.worksheets.add("审核台账");
const guide = workbook.worksheets.add("填写说明与选项");

const navy = "#17365D";
const blue = "#D9EAF7";
const lightBlue = "#EAF3F8";
const red = "#C00000";
const lightRed = "#FCE4D6";
const amber = "#FFF2CC";
const amberText = "#7F6000";
const green = "#E2F0D9";
const greenText = "#375623";
const purple = "#E4DFEC";
const gray = "#666666";
const line = "#B8C2CC";
const white = "#FFFFFF";

review.showGridLines = false;
review.mergeCells("A1:O1");
review.mergeCells("A2:O2");
review.mergeCells("A3:O3");
review.getRange("A1").values = [["团委办公室｜内部奖励资格审核表"]];
review.getRange("A2").values = [[
  "强制隔离：本表不得直接作为证书生成器输入。复核完成并经授权批准后，必须单独导出“证明批准名单”；证书生成器只能读取批准名单。"
]];
review.getRange("A3").values = [[
  "本表不内置资格阈值、不自动判定资格。分值仅供内部审核记录，不进入证明正文；“待确认”表示必须补充依据或完成人工复核。"
]];

review.getRange("A1:O1").format = {
  fill: navy,
  font: { bold: true, color: white },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
review.getRange("A2:O2").format = {
  fill: lightRed,
  font: { bold: true, color: red },
  wrapText: true,
  horizontalAlignment: "left",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "medium", color: red },
};
review.getRange("A3:O3").format = {
  fill: blue,
  font: { color: navy },
  wrapText: true,
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
review.getRange("A1:O1").format.rowHeight = 30;
review.getRange("A2:O2").format.rowHeight = 46;
review.getRange("A3:O3").format.rowHeight = 38;

const headers = [[
  "活动",
  "方案版本",
  "适用组别",
  "内部人员键",
  "姓名（受限）",
  "报名/签到证据",
  "出勤/请假/缺勤",
  "是否参演",
  "资格结果",
  "分值",
  "证明类型",
  "复核状态",
  "复核人",
  "复核日期",
  "异常说明",
]];
review.getRange("A4:O4").values = headers;
review.getRange("A4:O4").format = {
  fill: navy,
  font: { bold: true, color: white },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: line },
};
review.getRange("A4:O4").format.rowHeight = 42;

review.getRange("A5:O5").values = [[
  "", "", "", "", "", "", "", "", "待确认", "", "", "待确认", "", null, ""
]];
const table = review.tables.add("A4:O5", true, "RewardEligibilityReviewTable");
table.style = "TableStyleMedium2";
table.showBandedRows = true;
table.showFilterButton = true;

review.freezePanes.freezeRows(4);
review.freezePanes.freezeColumns(5);

const widths = [20, 14, 14, 18, 18, 24, 16, 12, 15, 10, 18, 15, 14, 14, 30];
for (let i = 0; i < widths.length; i += 1) {
  review.getRangeByIndexes(0, i, 30, 1).format.columnWidth = widths[i];
}
review.getRange("A5:O30").format = {
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    insideHorizontal: { style: "thin", color: "#E6E6E6" },
    bottom: { style: "thin", color: line },
  },
};
review.getRange("D5:E104").format.numberFormat = "@";
review.getRange("E5:E104").format.fill = purple;
review.getRange("J5:J104").format.numberFormat = "0.00";
review.getRange("N5:N104").format.numberFormat = "yyyy-mm-dd";
review.getRange("J5:J104").format.horizontalAlignment = "right";
review.getRange("H5:I104").format.horizontalAlignment = "center";
review.getRange("L5:N104").format.horizontalAlignment = "center";

review.getRange("C5:C504").dataValidation = {
  rule: { type: "list", values: ["本科生", "研究生", "其他", "不适用", "待确认"] },
};
review.getRange("G5:G504").dataValidation = {
  rule: { type: "list", values: ["出勤", "请假", "缺勤", "混合情况", "不适用", "待确认"] },
};
review.getRange("H5:H504").dataValidation = {
  rule: { type: "list", values: ["是", "否", "不适用", "待确认"] },
};
review.getRange("I5:I504").dataValidation = {
  rule: { type: "list", values: ["拟符合", "拟不符合", "待补充材料", "不适用", "待确认"] },
};
review.getRange("K5:K504").dataValidation = {
  rule: { type: "list", values: ["活动参与证明", "志愿服务证明", "获奖证明", "其他", "不适用", "待确认"] },
};
review.getRange("L5:L504").dataValidation = {
  rule: { type: "list", values: ["待复核", "复核中", "已复核", "退回补充", "不适用", "待确认"] },
};

review.getRange("A5:O504").conditionalFormats.add("containsText", {
  text: "待确认",
  format: { fill: amber, font: { bold: true, color: amberText } },
});
review.getRange("I5:I504").conditionalFormats.add("containsText", {
  text: "拟符合",
  format: { fill: green, font: { bold: true, color: greenText } },
});
review.getRange("I5:I504").conditionalFormats.add("containsText", {
  text: "拟不符合",
  format: { fill: lightRed, font: { bold: true, color: red } },
});
review.getRange("L5:L504").conditionalFormats.add("containsText", {
  text: "已复核",
  format: { fill: green, font: { bold: true, color: greenText } },
});

guide.showGridLines = false;
guide.mergeCells("A1:F1");
guide.mergeCells("A2:F2");
guide.getRange("A1").values = [["内部奖励资格审核表｜填写说明与强制流程"]];
guide.getRange("A2").values = [[
  "禁止将“审核台账”直接连接证书生成器。只有经人工复核、授权批准并单独导出的“证明批准名单”才可作为证书生成输入。"
]];
guide.getRange("A1:F1").format = {
  fill: navy,
  font: { bold: true, color: white },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
guide.getRange("A2:F2").format = {
  fill: lightRed,
  font: { bold: true, color: red },
  wrapText: true,
  borders: { preset: "outside", style: "medium", color: red },
};
guide.getRange("A1:F1").format.rowHeight = 30;
guide.getRange("A2:F2").format.rowHeight = 48;

guide.getRange("A4:F4").values = [["阶段", "责任角色", "输入", "动作", "输出", "强制控制"]];
guide.getRange("A5:F9").values = [
  ["1. 录入", "经授权经办人", "活动方案与证据", "使用内部人员键建立审核记录；姓名仅在必要时录入", "待审核记录", "不得录入身份证号、手机号等无关敏感字段"],
  ["2. 初审", "审核人", "报名、签到、请假、参演等证据", "按当前获批方案人工判断；填写依据和异常", "拟定资格结果", "本表不设置资格阈值，也不自动判定"],
  ["3. 复核", "独立复核人", "拟定结果及证据", "复查依据、结果、证明类型并填写复核人和日期", "已复核记录", "发现缺证或冲突时退回补充，不得带病通过"],
  ["4. 批准导出", "有权批准人", "已复核记录", "按最小字段单独导出“证明批准名单”并冻结版本", "证明批准名单", "审核台账与批准名单物理/逻辑隔离，记录版本和批准人"],
  ["5. 证明生成", "证明制作人", "证明批准名单", "按批准字段生成证明", "证明文件", "证书生成器只读批准名单；分值绝不进入证明正文"],
];
guide.getRange("A4:F9").format = {
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: line },
};
guide.getRange("A4:F4").format = {
  fill: navy,
  font: { bold: true, color: white },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: line },
};
guide.getRange("A5:A9").format.fill = lightBlue;
guide.getRange("F5:F9").format.fill = amber;
guide.getRange("A4:F4").format.rowHeight = 28;
guide.getRange("A5:F9").format.rowHeight = 48;

guide.getRange("A11:D11").values = [["字段", "是否必填", "用途", "填写与隐私要求"]];
guide.getRange("A12:D26").values = [
  ["活动", "是", "标识审核对象", "使用正式活动名称"],
  ["方案版本", "是", "绑定适用规则", "填写已批准方案的版本号/日期，不填写阈值"],
  ["适用组别", "是", "区分规则适用范围", "从下拉选择；不确定时选“待确认”"],
  ["内部人员键", "是", "跨表关联的最小标识", "使用系统内部键；不得以身份证号、手机号替代"],
  ["姓名（受限）", "按需", "必要的人工作业", "受限访问；导出批准名单时遵循最小化原则"],
  ["报名/签到证据", "是", "审计依据", "填写证据编号、受控路径或核验摘要，不粘贴公开链接"],
  ["出勤/请假/缺勤", "是", "记录事实状态", "按证据选择；混合情况在异常说明中解释"],
  ["是否参演", "按方案", "记录参演事实", "不适用时明确选择“不适用”"],
  ["资格结果", "是", "记录人工初审结果", "只能人工选择；“拟符合”不等于最终批准"],
  ["分值", "按方案", "内部审核辅助", "仅内部记录，不设固定满分/阈值，不进入证明正文"],
  ["证明类型", "是", "约束批准输出", "按批准范围选择，不得自行扩张证明内容"],
  ["复核状态", "是", "控制工作流", "未完成独立复核不得进入批准名单"],
  ["复核人", "复核时", "责任追溯", "填写经授权角色/人员；不得由初审人冒填"],
  ["复核日期", "复核时", "版本追溯", "使用 yyyy-mm-dd"],
  ["异常说明", "有异常时", "解释缺证、冲突或例外", "不写无关敏感信息；必要证据存受控位置"],
];
guide.getRange("A11:D26").format = {
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: line },
};
guide.getRange("A11:D11").format = {
  fill: navy,
  font: { bold: true, color: white },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: line },
};
guide.getRange("A12:A26").format.fill = lightBlue;
guide.getRange("D12:D26").format.fill = "#F7F9FB";

guide.getRange("F11:F16").values = [
  ["关键原则"],
  ["待确认优先"],
  ["人工判断"],
  ["独立复核"],
  ["批准名单隔离"],
  ["分值不进正文"],
];
guide.getRange("F11:F16").format = {
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: line },
};
guide.getRange("F11").format = { fill: navy, font: { bold: true, color: white } };
guide.getRange("F12:F16").format.fill = amber;
guide.getRange("F12:F16").format.font = { bold: true, color: amberText };

const guideWidths = [17, 17, 27, 42, 24, 34];
for (let i = 0; i < guideWidths.length; i += 1) {
  guide.getRangeByIndexes(0, i, 30, 1).format.columnWidth = guideWidths[i];
}
guide.freezePanes.freezeRows(2);

const reviewInspect = await workbook.inspect({
  kind: "table",
  range: "审核台账!A1:O5",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 16,
  maxChars: 6000,
});
const errorInspect = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(reviewInspect.ndjson);
console.log(errorInspect.ndjson);

const mainPreview = await workbook.render({
  sheetName: "审核台账",
  range: "A1:O12",
  scale: 1,
  format: "png",
});
await fs.writeFile(mainPreviewPath, new Uint8Array(await mainPreview.arrayBuffer()));
const notesPreview = await workbook.render({
  sheetName: "填写说明与选项",
  range: "A1:F26",
  scale: 1,
  format: "png",
});
await fs.writeFile(notesPreviewPath, new Uint8Array(await notesPreview.arrayBuffer()));

await fs.mkdir(outputPath.substring(0, outputPath.lastIndexOf("/")), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
'''


def _add_print_settings(path: Path) -> None:
    """Add explicit A4 landscape print settings using Open XML."""
    with zipfile.ZipFile(path, "r") as source:
        members = {name: source.read(name) for name in source.namelist()}

    sheet_name = "xl/worksheets/sheet1.xml"
    sheet = members[sheet_name].decode("utf-8")
    root_match = re.search(r"<(?P<prefix>[A-Za-z_][\w.-]*:)?worksheet\b[^>]*>", sheet)
    if not root_match:
        raise RuntimeError("Worksheet root element not found")
    prefix = root_match.group("prefix") or ""
    q = re.escape(prefix)
    sheet = re.sub(
        rf"<{q}sheetPr(?:\s[^>]*)?/>|<{q}sheetPr(?:\s[^>]*)?>.*?</{q}sheetPr>",
        "",
        sheet,
        flags=re.S,
    )
    sheet = re.sub(
        rf"<{q}printOptions(?:\s[^>]*)?/>|<{q}printOptions(?:\s[^>]*)?>.*?</{q}printOptions>",
        "",
        sheet,
        flags=re.S,
    )
    sheet = re.sub(rf"<{q}pageMargins(?:\s[^>]*)?/>", "", sheet)
    sheet = re.sub(rf"<{q}pageSetup(?:\s[^>]*)?/>", "", sheet)
    root_match = re.search(r"<(?P<prefix>[A-Za-z_][\w.-]*:)?worksheet\b[^>]*>", sheet)
    root_end = root_match.end()
    sheet_pr = f'<{prefix}sheetPr><{prefix}pageSetUpPr fitToPage="1"/></{prefix}sheetPr>'
    sheet = sheet[:root_end] + sheet_pr + sheet[root_end:]
    pane = (
        f'<{prefix}pane xSplit="5" ySplit="4" topLeftCell="F5" '
        'activePane="bottomRight" state="frozen"/>'
    )
    view_close = f"</{prefix}sheetView>"
    if pane not in sheet:
        if view_close in sheet:
            sheet = sheet.replace(view_close, pane + view_close, 1)
        else:
            view_pattern = rf"(<{q}sheetView\b[^>]*)\s*/>"
            sheet, replacements = re.subn(
                view_pattern,
                rf"\1>{pane}{view_close}",
                sheet,
                count=1,
            )
            if replacements != 1:
                raise RuntimeError("Worksheet view element not found")
    print_xml = (
        f'<{prefix}printOptions horizontalCentered="0" verticalCentered="0"/>'
        f'<{prefix}pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>'
        f'<{prefix}pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
    )
    marker = sheet.find(f"<{prefix}tableParts")
    if marker < 0:
        marker = sheet.rfind(f"</{prefix}worksheet>")
    if marker < 0:
        raise RuntimeError("Worksheet closing element not found")
    sheet = sheet[:marker] + print_xml + sheet[marker:]
    members[sheet_name] = sheet.encode("utf-8")

    workbook_name = "xl/workbook.xml"
    workbook = members[workbook_name].decode("utf-8")
    workbook_root = re.search(r"<(?P<prefix>[A-Za-z_][\w.-]*:)?workbook\b[^>]*>", workbook)
    if not workbook_root:
        raise RuntimeError("Workbook root element not found")
    workbook_prefix = workbook_root.group("prefix") or ""
    wq = re.escape(workbook_prefix)
    names = (
        f'<{workbook_prefix}definedNames>'
        f'<{workbook_prefix}definedName name="_xlnm.Print_Area" localSheetId="0">'
        "'审核台账'!$A$1:$O$30"
        f'</{workbook_prefix}definedName>'
        f'<{workbook_prefix}definedName name="_xlnm.Print_Titles" localSheetId="0">'
        "'审核台账'!$1:$4"
        f'</{workbook_prefix}definedName>'
        f'</{workbook_prefix}definedNames>'
    )
    defined_close = f"</{workbook_prefix}definedNames>"
    if re.search(rf"<{wq}definedNames\b", workbook):
        contents = names[names.find(">") + 1 : names.rfind("<")]
        workbook = workbook.replace(defined_close, contents + defined_close, 1)
    else:
        marker = workbook.find(f"<{workbook_prefix}calcPr")
        if marker < 0:
            marker = workbook.rfind(f"</{workbook_prefix}workbook>")
        if marker < 0:
            raise RuntimeError("Workbook closing element not found")
        workbook = workbook[:marker] + names + workbook[marker:]
    members[workbook_name] = workbook.encode("utf-8")

    temp_path = path.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for name, payload in members.items():
            target.writestr(name, payload)
    temp_path.replace(path)


def _self_check(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path, "r") as archive:
        names = archive.namelist()
        workbook = archive.read("xl/workbook.xml").decode("utf-8")
        sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        table = archive.read("xl/tables/table1.xml").decode("utf-8")
        combined = "\n".join(
            archive.read(name).decode("utf-8", errors="ignore")
            for name in names
            if name.endswith(".xml")
        )
        invalid_xml = []
        for name in names:
            if name.endswith((".xml", ".rels")):
                try:
                    ElementTree.fromstring(archive.read(name))
                except ElementTree.ParseError as error:
                    invalid_xml.append({"file": name, "error": str(error)})

    expected_headers = [
        "活动", "方案版本", "适用组别", "内部人员键", "姓名（受限）",
        "报名/签到证据", "出勤/请假/缺勤", "是否参演", "资格结果", "分值",
        "证明类型", "复核状态", "复核人", "复核日期", "异常说明",
    ]
    missing_headers = [header for header in expected_headers if header not in table]
    checks = {
        "output_exists": path.is_file() and path.stat().st_size > 0,
        "sheet_names_present": "审核台账" in workbook and "填写说明与选项" in workbook,
        "xml_well_formed": not invalid_xml,
        "invalid_xml": invalid_xml,
        "headers_complete": not missing_headers,
        "missing_headers": missing_headers,
        "validation_blocks": len(re.findall(r"<(?:[A-Za-z_][\w.-]*:)?dataValidation\b", sheet)) >= 6,
        "freeze_panes": "xSplit=\"5\"" in sheet and "ySplit=\"4\"" in sheet,
        "conditional_formatting": "conditionalFormatting" in sheet and "待确认" in combined,
        "print_setup": "orientation=\"landscape\"" in sheet and "fitToWidth=\"1\"" in sheet,
        "print_area_and_titles": "_xlnm.Print_Area" in workbook and "_xlnm.Print_Titles" in workbook,
        "no_qualification_formula": "<f" not in sheet,
        "no_fixed_twenty_point_rule": "20分" not in combined and "满分20" not in combined,
        "isolation_warning": "不得直接作为证书生成器输入" in combined and "证明批准名单" in combined,
        "score_exclusion_note": "分值" in combined and "不进入证明正文" in combined,
        "no_person_examples": all(token not in combined for token in ("甲乙丙", "丁戊", "示例电话值")),
    }
    checks["passed"] = all(
        value
        for key, value in checks.items()
        if key not in {"missing_headers", "invalid_xml", "passed"}
    )
    return checks


def main() -> None:
    if not NODE.is_file() or not NODE_MODULES.is_dir():
        raise SystemExit("Bundled spreadsheet runtime is unavailable.")
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    link = BUILD_DIR / "node_modules"
    if link.is_symlink():
        if link.resolve() != NODE_MODULES.resolve():
            raise SystemExit(f"Unexpected node_modules link: {link}")
    elif link.exists():
        raise SystemExit(f"Refusing to replace existing path: {link}")
    else:
        link.symlink_to(NODE_MODULES, target_is_directory=True)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env.update(
        {
            "REWARD_REVIEW_OUTPUT": str(OUTPUT),
            "REWARD_REVIEW_MAIN_PREVIEW": str(MAIN_PREVIEW),
            "REWARD_REVIEW_NOTES_PREVIEW": str(NOTES_PREVIEW),
        }
    )
    completed = subprocess.run(
        [str(NODE), "--input-type=module", "--eval", JAVASCRIPT],
        cwd=BUILD_DIR,
        env=env,
        text=True,
        capture_output=True,
        check=False,
    )
    if completed.stdout:
        print(completed.stdout.rstrip())
    if completed.returncode:
        if completed.stderr:
            print(completed.stderr.rstrip())
        raise SystemExit(completed.returncode)
    inspect_sidecar = Path(str(OUTPUT) + ".inspect.ndjson")
    if inspect_sidecar.exists():
        inspect_sidecar.unlink()

    _add_print_settings(OUTPUT)
    result = _self_check(OUTPUT)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result["passed"]:
        raise SystemExit("Workbook self-check failed.")
    print(f"Created {OUTPUT}")
    print(f"Visual previews: {MAIN_PREVIEW}, {NOTES_PREVIEW}")


if __name__ == "__main__":
    main()
