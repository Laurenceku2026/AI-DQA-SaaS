"""Build English DFMEA template shells from Chinese templates (same layout/sheet names)."""
from __future__ import annotations

import os

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES = os.path.join(HERE, "templates")

DQA_T1_REPLACEMENTS = [
    ("设计失效模式和后果分析", "Design Failure Mode and Effects Analysis"),
    ("（DFMEA)", "(DFMEA)"),
    ("系统", "System"),
    ("子系统", "Subsystem"),
    ("部件", "Part"),
    ("设计责任", "Design Responsibility"),
    ("型号/项目", "Model / Project"),
    ("关键日期", "Key Date"),
    ("核心小组", "Core Team"),
    ("要求", "Requirement"),
    ("潜在失效模式", "Potential Failure Mode"),
    ("潜在失效    后果", "Potential Failure Effect"),
    ("潜在失效　　后果", "Potential Failure Effect"),
    ("严重度(S)", "Severity (S)"),
    ("级别", "Class"),
    ("潜在失效原因/机理", "Potential Failure Cause / Mechanism"),
    ("发生度(O)", "Occurrence (O)"),
    ("现行设计控制", "Current Design Controls"),
    ("预防", "Prevention"),
    ("探测", "Detection"),
    ("探测度(D)", "Detection (D)"),
    ("风险优先数(RPN)", "RPN"),
    ("建议措施", "Recommended Actions"),
    ("责任及目标完成日期", "Responsibility & Target Date"),
    ("采取的措施", "Actions Taken"),
    ("S", "S"),
]

DQA_T2_REPLACEMENTS = [
    ("设计失效模式及后果分析（DFMEA）  （标准表格）", "Design FMEA (DFMEA) — Standard Form"),
    ("系统分析    步骤1 - 策划与准备", "System Analysis    Step 1 — Planning & Preparation"),
    ("公司名称：", "Company Name:"),
    ("工厂地址：", "Plant Address:"),
    ("顾客名称：", "Customer Name:"),
    ("年型/项目：", "Model Year / Project:"),
    ("持续改善", "Continuous Improvement"),
    ("系统分析", "System Analysis"),
    ("步骤2 - structural analysis （结构分析）", "Step 2 — Structural Analysis"),
    ("步骤3 - Functional analysis （功能分析）", "Step 3 — Functional Analysis"),
    ("问题#", "Issue #"),
    ("历史/变更授权", "History / Change Authorization"),
    ("(适用时}", "(if applicable)"),
    ("1.上一高层级", "1. Higher Level"),
    ("2.关注要素", "2. Focus Element"),
    ("3.下一低层级或特性类型", "3. Next Lower Level or Characteristic Type"),
    ("1.上一高层级功能及要求", "1. Higher Level Function & Requirements"),
    ("2.关注要素功能及要求", "2. Focus Element Function & Requirements"),
    ("3.下一低层级功能及要求或特性", "3. Next Lower Level Function / Characteristic"),
    ("失效影响(FE)", "Failure Effect (FE)"),
    ("严重度(S)", "Severity (S)"),
    ("关注要素的失效模式(FM)", "Focus Element Failure Mode (FM)"),
    ("失效原因(FC)", "Failure Cause (FC)"),
    ("现行预防控制(PC)", "Current Prevention Control (PC)"),
    ("发生度(O)", "Occurrence (O)"),
    ("现行探测控制(DC)", "Current Detection Control (DC)"),
    ("探测度(D)", "Detection (D)"),
    ("措施优先级(AP)", "Action Priority (AP)"),
    ("特殊特性", "Special Characteristics"),
    ("筛选器代码(可选)", "Filter Code (optional)"),
    ("预防措施", "Prevention Actions"),
    ("探测措施", "Detection Actions"),
    ("责任人", "Responsible"),
    ("目标完成日期", "Target Completion Date"),
    ("状态", "Status"),
    ("采取基于证据的措施", "Evidence-Based Actions Taken"),
    ("完成日期", "Completion Date"),
    ("严重度(S)", "Severity (S)"),
    ("探测度(D)", "Detection (D)"),
    ("备注", "Remarks"),
]


def _translate(value: str, rules: list[tuple[str, str]]) -> str:
    text = str(value)
    for src, dst in rules:
        text = text.replace(src, dst)
    return text


def _translate_workbook(src_path: str, dst_path: str, rules: list[tuple[str, str]]) -> None:
    wb = load_workbook(src_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None or not isinstance(cell.value, str):
                    continue
                translated = _translate(cell.value, rules)
                if translated != cell.value:
                    cell.value = translated
    wb.save(dst_path)
    print("Wrote", dst_path)


def main() -> None:
    pairs = [
        ("模板1-DFMEA旧版.xlsx", "Template-1-DFMEA-Legacy.xlsx", DQA_T1_REPLACEMENTS),
        ("模板2-DFMEA新版.xlsx", "Template-2-DFMEA-New.xlsx", DQA_T2_REPLACEMENTS),
    ]
    for src_name, dst_name, rules in pairs:
        src = os.path.join(TEMPLATES, src_name)
        dst = os.path.join(TEMPLATES, dst_name)
        if not os.path.isfile(src):
            raise FileNotFoundError(src)
        _translate_workbook(src, dst, rules)


if __name__ == "__main__":
    main()
