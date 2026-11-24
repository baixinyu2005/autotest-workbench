"""Excel 导入与标准测试计划生成。

支持两种来源：
1. 老师发的“功能清单”Excel：自动展开合并单元格，得到需求和基础 CRUD 用例骨架；
2. 本系统的“自动化用例”标准 Excel：读取带步骤和预期结果的可执行用例。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .generator import generate_cases
from .models import Project, Requirement, Step, TestCase
from .accounts import extract_accounts_from_folder, extract_accounts_from_workbook
from .spec_detect import FolderScan, SheetCandidate, scan_folder, scan_workbook


class ImportError(Exception):
    pass


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------
def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace("\u3000", " ").strip()


def _grid(ws) -> list[list[Any]]:
    """把工作表复制成矩阵，并把合并单元格的值填充到整个区域。"""
    max_row, max_col = ws.max_row, ws.max_column
    grid = [[None] * (max_col + 1) for _ in range(max_row + 1)]
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            grid[row][col] = ws.cell(row, col).value
    for rng in ws.merged_cells.ranges:
        value = ws.cell(rng.min_row, rng.min_col).value
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                grid[row][col] = value
    return grid


def _find_col(
    grid: list[list[Any]],
    header_row: int,
    *,
    exact: str | None = None,
    contains: tuple[str, ...] = (),
    rows: int = 3,
) -> int | None:
    for row in range(header_row, min(header_row + rows, len(grid))):
        for col in range(1, len(grid[row])):
            text = _text(grid[row][col])
            if not text:
                continue
            if exact is not None and text == exact:
                return col
            if contains and all(word in text for word in contains):
                return col
    return None


def _read_project_name(wb, fallback: str) -> str:
    """只扫描每个 sheet 的前 16 行，避免大工作簿被完整复制。"""
    for ws in wb.worksheets:
        for row in range(1, min(16, ws.max_row) + 1):
            for col in range(1, min(ws.max_column, 12) + 1):
                text = _text(ws.cell(row, col).value)
                if text.startswith("项目："):
                    value = text.split("：", 1)[1].strip()
                    if value:
                        return value
                if text == "项目名称":
                    for offset in range(1, 6):
                        value = _text(ws.cell(row, col + offset).value)
                        if value:
                            return value
    return fallback


# ---------------------------------------------------------------------------
# 老师功能清单
# ---------------------------------------------------------------------------
def _find_checklist_sheet(wb):
    for ws in wb.worksheets:
        if ws.title == "测试内容":
            return ws
    for ws in wb.worksheets:
        grid = _grid(ws)
        for row in range(1, min(20, len(grid))):
            texts = [_text(v) for v in grid[row]]
            if "序号" in texts and any("实际" in t and "功能" in t for t in texts):
                return ws
    return None


def _parse_checklist(wb, ws, fallback_name: str = "需求导入") -> Project:
    grid = _grid(ws)
    header_row = None
    for row in range(1, min(25, len(grid))):
        texts = [_text(v) for v in grid[row]]
        if "序号" in texts and any("实际" in t and "功能" in t for t in texts):
            header_row = row
            break
    if header_row is None:
        raise ImportError("找不到功能清单表头（需要“序号”和“实际的应用功能”列）")

    seq_col = _find_col(grid, header_row, exact="序号", rows=1) or 1
    system_col = _find_col(grid, header_row, exact="系统", rows=4)
    module_col = _find_col(grid, header_row, exact="模块", rows=4)
    function_col = _find_col(grid, header_row, exact="功能", rows=4)
    point_col = _find_col(grid, header_row, exact="功能点", rows=4)
    actual_col = _find_col(grid, header_row, contains=("实际", "功能"), rows=2)

    if not all((system_col, module_col, function_col, point_col, actual_col)):
        raise ImportError(
            "功能清单列定位失败：需要 系统、模块、功能、功能点、实际的应用功能 五列"
        )

    max_col = ws.max_column
    remark_col = actual_col + 1 if actual_col < max_col else None
    subheader = header_row + 1
    has_subheader = subheader < len(grid) and any(
        _text(grid[subheader][col]) in ("系统", "模块", "功能", "功能点")
        for col in range(1, max_col + 1)
    )
    data_start = header_row + 2 if has_subheader else header_row + 1

    requirements: list[Requirement] = []
    last_system = last_module = last_function = ""
    counter = 0

    for row in range(data_start, ws.max_row + 1):
        point = _text(grid[row][point_col])
        actual = _text(grid[row][actual_col])
        if not point and not actual:
            continue

        system = _text(grid[row][system_col]) or last_system
        module = _text(grid[row][module_col]) or last_module
        function = _text(grid[row][function_col]) or last_function
        last_system, last_module, last_function = system, module, function

        counter += 1
        remark = _text(grid[row][remark_col]) if remark_col else ""
        requirements.append(
            Requirement(
                req_id=f"REQ-{counter:04d}",
                seq=_text(grid[row][seq_col]) or str(counter),
                system=system,
                module=module,
                function=function,
                point=point,
                actual_features=actual,
                remark=remark,
                source_sheet=ws.title,
                source_row=row,
            )
        )

    if not requirements:
        raise ImportError("功能清单没有解析到有效数据行")

    project = Project.create(
        name=_read_project_name(wb, fallback_name),
        source_excel="",
        source_sheet=ws.title,
    )
    project.requirements = requirements
    generate_cases(project)
    project.notes = (
        f"已从“{ws.title}”导入 {len(requirements)} 条功能；"
        f"已按基础 CRUD 规则生成 {project.case_count} 个用例骨架。"
        "未配置步骤的用例不会伪造通过。"
    )
    return project


# ---------------------------------------------------------------------------
# 标准自动化用例
# ---------------------------------------------------------------------------
STANDARD_HEADERS = [
    ("需求ID", "req_id"),
    ("系统", "system"),
    ("模块", "module"),
    ("功能", "function"),
    ("功能点", "point"),
    ("用例名称", "case_name"),
    ("步骤序号", "step_no"),
    ("操作", "action"),
    ("目标", "target"),
    ("输入值", "value"),
    ("预期结果", "expected"),
    ("说明", "description"),
    ("超时毫秒", "timeout_ms"),
    ("需规备注", "remark"),
    ("序号", "seq"),
]

ACTION_ALIASES = {
    "goto": "goto",
    "打开网址": "goto",
    "访问": "goto",
    "click": "click",
    "点击": "click",
    "fill": "fill",
    "填写": "fill",
    "输入": "fill",
    "select": "select",
    "选择": "select",
    "press": "press",
    "按键": "press",
    "wait": "wait",
    "等待": "wait",
    "assert_text": "assert_text",
    "断言文本": "assert_text",
    "assert_visible": "assert_visible",
    "断言可见": "assert_visible",
    "assert_url": "assert_url",
    "断言网址": "assert_url",
    "assert_count": "assert_count",
    "断言元素数量": "assert_count",
    "断言数量": "assert_count",
    "assert_sorted": "assert_sorted",
    "断言排序": "assert_sorted",
    "assert_duplicate_max": "assert_duplicate_max",
    "断言重复上限": "assert_duplicate_max",
    "assert_attribute": "assert_attribute",
    "断言属性": "assert_attribute",
    "assert_value": "assert_value",
    "断言输入值": "assert_value",
    "assert_hidden": "assert_hidden",
    "断言隐藏": "assert_hidden",
    "emulate_media": "emulate_media",
    "切换媒体类型": "emulate_media",
    "reload": "reload",
    "刷新页面": "reload",
    "screenshot": "screenshot",
    "截图": "screenshot",
}


def _header_map(ws, required=("需求ID", "操作")) -> tuple[int, dict[str, int]]:
    for row in range(1, min(30, ws.max_row) + 1):
        values = {_text(ws.cell(row, col).value): col for col in range(1, ws.max_column + 1)}
        if all(any(key in text for text in values) for key in required):
            return row, values
    raise ImportError("找不到“自动化用例”表头（需要 需求ID、操作 等列）")


def _col(values: dict[str, int], name: str) -> int | None:
    for text, col in values.items():
        if text == name:
            return col
    return None


def _make_step(ws, row: int, cols: dict[str, int]) -> Step | None:
    action_raw = _text(ws.cell(row, cols["action"]).value) if "action" in cols else ""
    if not action_raw:
        return None
    action = ACTION_ALIASES.get(action_raw.lower()) or ACTION_ALIASES.get(action_raw)
    if not action:
        return None
    target = _text(ws.cell(row, cols["target"]).value) if "target" in cols else ""
    value = _text(ws.cell(row, cols["value"]).value) if "value" in cols else ""
    expected = _text(ws.cell(row, cols["expected"]).value) if "expected" in cols else ""
    description = (
        _text(ws.cell(row, cols["description"]).value) if "description" in cols else ""
    )
    timeout_raw = _text(ws.cell(row, cols["timeout_ms"]).value) if "timeout_ms" in cols else ""
    try:
        timeout_ms = int(float(timeout_raw)) if timeout_raw else 15000
    except ValueError:
        timeout_ms = 15000
    return Step(
        action=action,
        target=target,
        value=value,
        expected=expected,
        description=description,
        timeout_ms=timeout_ms,
    )


def _parse_standard(wb, sheet_name: str = "自动化用例", fallback_name: str = "自动化测试") -> Project:
    ws = wb[sheet_name]
    header_row, values = _header_map(ws)
    cols: dict[str, int] = {}
    for header, key in STANDARD_HEADERS:
        col = _col(values, header)
        if col:
            cols[key] = col

    missing = [key for key in ("req_id", "action") if key not in cols]
    if missing:
        raise ImportError(f"标准用例缺少必要列：{missing}")

    project = Project.create(
        name=_read_project_name(wb, fallback_name),
        source_excel="",
        source_sheet=sheet_name,
    )

    # 前置步骤
    if "前置步骤" in wb.sheetnames:
        setup_ws = wb["前置步骤"]
        try:
            setup_header, setup_values = _header_map(setup_ws)
            setup_cols = {
                key: _col(setup_values, header)
                for header, key in STANDARD_HEADERS
                if _col(setup_values, header)
            }
            for row in range(setup_header + 1, setup_ws.max_row + 1):
                step = _make_step(setup_ws, row, setup_cols)
                if step:
                    project.setup_steps.append(step)
        except ImportError:
            pass

    groups: dict[tuple[str, str], dict[str, Any]] = {}
    for row in range(header_row + 1, ws.max_row + 1):
        req_id = _text(ws.cell(row, cols["req_id"]).value)
        step = _make_step(ws, row, cols)
        if not req_id and step is None:
            continue
        if not req_id:
            continue
        case_name = _text(ws.cell(row, cols["case_name"]).value) if "case_name" in cols else ""
        case_name = case_name or f"{req_id} 自动化用例"
        key = (req_id, case_name)
        if key not in groups:
            groups[key] = {
                "system": _text(ws.cell(row, cols["system"]).value) if "system" in cols else "",
                "module": _text(ws.cell(row, cols["module"]).value) if "module" in cols else "",
                "function": _text(ws.cell(row, cols["function"]).value) if "function" in cols else "",
                "point": _text(ws.cell(row, cols["point"]).value) if "point" in cols else "",
                "remark": _text(ws.cell(row, cols["remark"]).value) if "remark" in cols else "",
                "seq": _text(ws.cell(row, cols["seq"]).value) if "seq" in cols else "",
                "steps": [],
                "expected": [],
                "descriptions": [],
            }
        group = groups[key]
        if step:
            group["steps"].append(step)
            if step.expected:
                group["expected"].append(step.expected)
            if step.description:
                group["descriptions"].append(step.description)
        else:
            # 未配置步骤的用例仍然保留预期和说明，便于后续补充。
            row_expected = (
                _text(ws.cell(row, cols["expected"]).value) if "expected" in cols else ""
            )
            row_description = (
                _text(ws.cell(row, cols["description"]).value)
                if "description" in cols
                else ""
            )
            if row_expected:
                group["expected"].append(row_expected)
            if row_description:
                group["descriptions"].append(row_description)

    if not groups:
        raise ImportError("标准用例中没有解析到有效用例")

    requirements: dict[str, Requirement] = {}
    for index, ((req_id, case_name), group) in enumerate(groups.items(), start=1):
        requirement = requirements.get(req_id)
        if requirement is None:
            requirement = Requirement(
                req_id=req_id,
                seq=group["seq"] or str(index),
                system=group["system"],
                module=group["module"],
                function=group["function"],
                point=group["point"] or case_name,
                actual_features="",
                remark=group["remark"],
                source_sheet=sheet_name,
            )
            requirements[req_id] = requirement
        steps: list[Step] = group["steps"]
        expected = "；".join(dict.fromkeys(group["expected"])) or "按步骤中的断言结果判定。"
        requirement.cases.append(
            TestCase(
                case_id=f"{req_id}-C{len(requirement.cases) + 1:02d}",
                requirement_id=req_id,
                name=case_name,
                kind="自动化",
                steps=steps,
                expected=expected,
                description="\n".join(dict.fromkeys(group["descriptions"])),
                status="pending" if steps else "not_configured",
            )
        )

    project.requirements = list(requirements.values())
    project.notes = (
        f"已导入 {len(project.requirements)} 条需求、"
        f"{project.case_count} 个可执行用例、{len(project.setup_steps)} 个前置步骤。"
    )
    return project


# ---------------------------------------------------------------------------
# 通用验收/功能需求表
# ---------------------------------------------------------------------------
ACCEPTANCE_COLUMNS = {
    "case_id": ("用例编号", "用例ID", "测试用例编号"),
    "req_id": ("关联需求", "需求编号", "需求ID"),
    "scenario": ("场景", "测试场景", "用例名称"),
    "steps": ("操作步骤", "测试步骤", "步骤"),
    "expected": ("预期结果", "期望结果"),
    "priority": ("优先级",),
    "self_test": ("自测结果", "测试结果"),
    "module": ("一级模块", "模块"),
    "function": ("功能名称", "功能"),
}

FUNCTIONAL_COLUMNS = {
    "req_id": ("需求编号", "需求ID"),
    "module": ("一级模块", "模块"),
    "function": ("功能名称", "功能"),
    "description": ("需求描述", "需求说明"),
    "io": ("输入/输出", "输入输出"),
    "priority": ("优先级",),
    "acceptance": ("验收要点", "验收标准"),
}


def _columns_by_keywords(ws, header_row: int, spec: dict[str, tuple[str, ...]]) -> dict[str, int]:
    cols: dict[str, int] = {}
    for row in (header_row, header_row + 1):
        if row > ws.max_row:
            continue
        for col in range(1, min(ws.max_column, 80) + 1):
            text = _text(ws.cell(row, col).value)
            if not text:
                continue
            for key, variants in spec.items():
                if key not in cols and any(v in text for v in variants):
                    cols[key] = col
    return cols


SUMMARY_PREFIXES = ("合计", "总计", "说明", "注：", "注:", "备注")


def _is_summary_row(*values: str) -> bool:
    first = next((value for value in values if value), "")
    return bool(first) and first.startswith(SUMMARY_PREFIXES)


def _parse_acceptance(
    wb, candidate: SheetCandidate, fallback_name: str
) -> Project:
    ws = wb[candidate.sheet]
    cols = _columns_by_keywords(ws, candidate.header_row, ACCEPTANCE_COLUMNS)
    if "case_id" not in cols and "scenario" not in cols:
        raise ImportError("验收表缺少用例编号/场景列")

    project = Project.create(
        name=_read_project_name(wb, fallback_name),
        source_excel="",
        source_sheet=candidate.sheet,
    )
    requirements: dict[str, Requirement] = {}
    seen_case_ids: set[str] = set()
    counter = 0

    for row in range(candidate.header_row + 1, ws.max_row + 1):
        def value(key: str) -> str:
            return _text(ws.cell(row, cols[key]).value) if key in cols else ""

        case_id = value("case_id")
        req_id = value("req_id")
        scenario = value("scenario")
        steps_text = value("steps")
        expected = value("expected")
        if _is_summary_row(case_id, scenario) and not (steps_text or expected):
            continue
        if not any((case_id, req_id, scenario, steps_text, expected)):
            continue

        counter += 1
        req_id = req_id or f"REQ-{counter:04d}"
        case_id = case_id or f"TC-{counter:04d}"
        base_id = case_id
        suffix = 2
        while case_id in seen_case_ids:
            case_id = f"{base_id}-{suffix}"
            suffix += 1
        seen_case_ids.add(case_id)

        requirement = requirements.get(req_id)
        if requirement is None:
            requirement = Requirement(
                req_id=req_id,
                seq=req_id,
                system=project.name,
                module=value("module") or "验收测试",
                function=value("function") or "验收",
                point=scenario or case_id,
                actual_features=steps_text or scenario,
                remark=value("priority"),
                source_sheet=candidate.sheet,
                source_row=row,
            )
            requirements[req_id] = requirement

        description_parts = [steps_text] if steps_text else []
        if value("self_test"):
            description_parts.append(f"开发自测：{value('self_test')}")
        if value("priority"):
            description_parts.append(f"优先级：{value('priority')}")
        requirement.cases.append(
            TestCase(
                case_id=case_id,
                requirement_id=req_id,
                name=scenario or case_id,
                kind="验收",
                steps=[],
                expected=expected or "按操作步骤人工确认",
                description="\n".join(description_parts),
                status="not_configured",
            )
        )

    if not requirements:
        raise ImportError("验收表中没有解析到有效用例")
    project.requirements = list(requirements.values())
    project.notes = (
        f"已自动识别验收/测试用例表“{candidate.sheet}”："
        f"{project.case_count} 条用例。"
        "验收表提供场景、步骤和预期，但没有页面选择器，"
        "需要补充可执行步骤后才能自动断言。"
    )
    return project


def _parse_functional(
    wb, candidate: SheetCandidate, fallback_name: str, generate: bool = True
) -> Project:
    ws = wb[candidate.sheet]
    cols = _columns_by_keywords(ws, candidate.header_row, FUNCTIONAL_COLUMNS)
    if "req_id" not in cols:
        raise ImportError("功能需求表缺少需求编号列")

    project = Project.create(
        name=_read_project_name(wb, fallback_name),
        source_excel="",
        source_sheet=candidate.sheet,
    )
    requirements: dict[str, Requirement] = {}
    counter = 0

    for row in range(candidate.header_row + 1, ws.max_row + 1):
        def value(key: str) -> str:
            return _text(ws.cell(row, cols[key]).value) if key in cols else ""

        req_id = value("req_id")
        description = value("description") or value("acceptance")
        if _is_summary_row(req_id, description):
            continue
        if not req_id and not description:
            continue
        counter += 1
        req_id = req_id or f"REQ-{counter:04d}"
        if req_id in requirements:
            # 同一需求编号出现多次时，合并描述，避免丢失。
            if description and description not in requirements[req_id].actual_features:
                requirements[req_id].actual_features += "\n" + description
            continue
        requirements[req_id] = Requirement(
            req_id=req_id,
            seq=req_id,
            system=project.name,
            module=value("module") or "功能需求",
            function=value("function") or "功能",
            point=value("function") or req_id,
            actual_features=description,
            remark=value("priority"),
            source_sheet=candidate.sheet,
            source_row=row,
        )

    if not requirements:
        raise ImportError("功能需求表中没有解析到有效需求")
    project.requirements = list(requirements.values())
    if generate:
        generate_cases(project)
        project.notes = (
            f"已自动识别功能需求表“{candidate.sheet}”："
            f"{len(project.requirements)} 条需求，"
            f"按基础 CRUD 规则生成 {project.case_count} 个用例骨架。"
            "没有页面步骤的用例不会伪造通过。"
        )
    else:
        project.notes = (
            f"已自动识别功能需求表“{candidate.sheet}”："
            f"{len(project.requirements)} 条需求。"
        )
    return project


def _merge_projects(functional: Project, acceptance: Project) -> Project:
    """合并功能需求与验收用例，验收用例按关联需求挂到功能需求下面。"""
    req_map = {req.req_id: req for req in functional.requirements}
    for acc_req in acceptance.requirements:
        target = req_map.get(acc_req.req_id)
        if target is None:
            functional.requirements.append(acc_req)
            req_map[acc_req.req_id] = acc_req
            continue
        for case in acc_req.cases:
            case.requirement_id = target.req_id
            target.cases.append(case)
        if not target.actual_features:
            target.actual_features = acc_req.actual_features

    # 避免不同表里重复的用例编号。
    seen: set[str] = set()
    for case in functional.iter_cases():
        if case.case_id in seen:
            original = case.case_id
            suffix = 2
            while f"{original}-{suffix}" in seen:
                suffix += 1
            case.case_id = f"{original}-{suffix}"
        seen.add(case.case_id)

    functional.notes = (
        f"已自动合并功能需求表“{functional.source_sheet}”"
        f"和验收表“{acceptance.source_sheet}”："
        f"{len(functional.requirements)} 条需求，{functional.case_count} 条验收用例。"
        "验收用例需要补充页面选择器后才能自动执行。"
    )
    functional.source_sheet = f"{functional.source_sheet} + {acceptance.source_sheet}"
    return functional


# ---------------------------------------------------------------------------
# 对外入口
# ---------------------------------------------------------------------------
def import_workbook(path: str | Path) -> Project:
    source = Path(path)
    if not source.exists():
        raise ImportError(f"文件不存在：{source}")
    if source.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ImportError("目前只支持 .xlsx / .xlsm，请先将 .xls 另存为 .xlsx")

    candidates = scan_workbook(source)
    if not candidates:
        raise ImportError(
            "无法自动识别这个 Excel 的需规工作表。"
            "支持的格式：标准自动化用例、验收/测试用例、功能需求、功能清单。"
        )

    try:
        wb = openpyxl.load_workbook(source, data_only=True, read_only=False)
    except Exception as exc:
        raise ImportError(f"Excel 打不开：{exc}") from exc

    try:
        standard = next((item for item in candidates if item.kind == "standard"), None)
        functional = next((item for item in candidates if item.kind == "functional"), None)
        acceptance = next((item for item in candidates if item.kind == "acceptance"), None)
        checklist = next((item for item in candidates if item.kind == "checklist"), None)

        if standard is not None:
            project = _parse_standard(wb, standard.sheet, source.stem)
            detected = standard
        elif functional is not None and acceptance is not None:
            project = _merge_projects(
                _parse_functional(wb, functional, source.stem, generate=False),
                _parse_acceptance(wb, acceptance, source.stem),
            )
            detected = functional
        elif acceptance is not None:
            project = _parse_acceptance(wb, acceptance, source.stem)
            detected = acceptance
        elif functional is not None:
            project = _parse_functional(wb, functional, source.stem, generate=True)
            detected = functional
        elif checklist is not None:
            project = _parse_checklist(wb, wb[checklist.sheet], source.stem)
            detected = checklist
        else:
            raise ImportError("没有找到可导入的需规工作表")

        project.source_excel = str(source)
        project.notes = f"[自动识别] {detected.describe()}\n" + project.notes
        accounts = extract_accounts_from_workbook(wb)
        if accounts:
            project.accounts = accounts
            project.notes += f"\n已识别 {len(accounts)} 个测试账号及权限。"
        return project
    finally:
        wb.close()


def import_folder(path: str | Path) -> tuple[Project, FolderScan]:
    """扫描文件夹，自动选择最像需规的文件/工作表并导入。"""
    scan = scan_folder(path)
    if scan.best is None:
        raise ImportError(
            "文件夹里没有识别到需规 Excel。"
            "请确认文件为 .xlsx/.xlsm，并且包含验收用例、功能需求或功能清单表头。"
        )
    project = import_workbook(scan.best.path)
    if not project.accounts:
        folder_accounts = extract_accounts_from_folder(path)
        if folder_accounts:
            project.accounts = folder_accounts
            project.notes += f"\n从文件夹其他 Excel 识别到 {len(folder_accounts)} 个测试账号。"
    project.notes = (
        f"[文件夹扫描] 共识别 {len(scan.candidates)} 个候选，"
        f"已选择：{scan.best.describe()}\n" + project.notes
    )
    return project, scan


# ---------------------------------------------------------------------------
# 标准用例模板与导出
# ---------------------------------------------------------------------------
def _style_header(ws, row: int = 1) -> None:
    fill = PatternFill("solid", fgColor="1E3A8A")
    font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="CBD5E1")
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)


def _set_widths(ws, widths: list[int]) -> None:
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width


def create_standard_template(path: str | Path) -> Path:
    target = Path(path)
    wb = openpyxl.Workbook()

    guide = wb.active
    guide.title = "说明"
    guide_lines = [
        ["桌面自动化测试校验系统 - 标准用例模板"],
        [""],
        ["1. 在“自动化用例”sheet 中，一行代表一个步骤。"],
        ["2. 同一个“需求ID + 用例名称”的多行会合并为一个测试用例。"],
        ["3. “操作”支持：goto、click、fill、select、press、wait、assert_text、assert_visible、assert_url、assert_count、assert_sorted、assert_duplicate_max、assert_attribute、assert_value、assert_hidden、emulate_media、reload、screenshot。"],
        ["4. “目标”填写网址或 Playwright 选择器，例如 input[name='username']、text=保存。"],
        ["5. “输入值”用于填写内容、按键、等待毫秒数或断言文本。"],
        ["6. “前置步骤”会在所有用例之前执行一次，适合放登录步骤。"],
        ["7. 未配置步骤的功能会标记为“未配置”，不会伪造通过。"],
    ]
    for line in guide_lines:
        guide.append(line)
    guide["A1"].font = Font(size=16, bold=True, color="1E3A8A")
    _set_widths(guide, [110])
    for row in guide.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="center")

    setup = wb.create_sheet("前置步骤")
    setup.append([header for header, _ in STANDARD_HEADERS])
    _style_header(setup)
    _set_widths(setup, [14, 16, 16, 20, 24, 24, 10, 14, 42, 28, 40, 30, 12, 28, 8])
    setup.freeze_panes = "A2"

    cases = wb.create_sheet("自动化用例")
    cases.append([header for header, _ in STANDARD_HEADERS])
    _style_header(cases)
    _set_widths(cases, [14, 16, 16, 20, 24, 24, 10, 14, 42, 28, 40, 30, 12, 28, 8])
    cases.freeze_panes = "A2"
    cases.auto_filter.ref = f"A1:{get_column_letter(len(STANDARD_HEADERS))}1"

    wb.save(target)
    return target


def export_project_plan(project: Project, path: str | Path) -> Path:
    """把当前项目导出为可继续编辑的标准用例 Excel。

    - 已配置步骤的用例：一个步骤一行；
    - 尚未配置步骤的用例：导出一行元数据但“操作”留空；
      重新导入后仍会保留为“未配置”，不会丢失用例。
    """
    target = Path(path)
    wb = openpyxl.Workbook()

    guide = wb.active
    guide.title = "说明"
    guide.append(["从 AutoTest Workbench 导出的可编辑测试计划"])
    guide.append(["项目：" + project.name])
    guide.append(["导入文件：" + project.source_excel])
    guide.append([""])
    guide.append(["1. 在“自动化用例”中给空操作行补充步骤。"])
    guide.append(["2. 一个“需求ID + 用例名称”可包含多行步骤。"])
    guide.append(["3. 支持 goto / click / fill / select / press / wait / assert_text / assert_visible / assert_url / assert_count / assert_sorted / assert_duplicate_max / assert_attribute / assert_value / assert_hidden / emulate_media / reload / screenshot。"])
    guide.append(["4. 修改完成后，在本程序中重新导入该文件即可执行。"])
    guide["A1"].font = Font(size=16, bold=True, color="1E3A8A")
    _set_widths(guide, [110])

    setup = wb.create_sheet("前置步骤")
    setup.append([header for header, _ in STANDARD_HEADERS])
    for index, step in enumerate(project.setup_steps, start=1):
        setup.append(
            [
                "SETUP", "前置步骤", "", "", "", "前置步骤", index,
                step.action, step.target, step.value, step.expected,
                step.description, step.timeout_ms, "", "",
            ]
        )
    _style_header(setup)
    _set_widths(setup, [14, 16, 16, 20, 24, 24, 10, 14, 42, 28, 40, 30, 12, 28, 8])
    setup.freeze_panes = "A2"

    cases = wb.create_sheet("自动化用例")
    cases.append([header for header, _ in STANDARD_HEADERS])
    for requirement in project.requirements:
        for case in requirement.cases:
            base = [
                requirement.req_id,
                requirement.system,
                requirement.module,
                requirement.function,
                requirement.point,
                case.name,
            ]
            if case.steps:
                for index, step in enumerate(case.steps, start=1):
                    cases.append(
                        base
                        + [
                            index,
                            step.action,
                            step.target,
                            step.value,
                            step.expected or case.expected,
                            step.description,
                            step.timeout_ms,
                            requirement.remark,
                            requirement.seq,
                        ]
                    )
            else:
                # 保留未配置用例：操作留空，重新导入后仍能识别需求，但不会伪造步骤。
                cases.append(
                    base
                    + [
                        "",
                        "",
                        "",
                        "",
                        case.expected,
                        case.description,
                        "",
                        requirement.remark,
                        requirement.seq,
                    ]
                )
    _style_header(cases)
    _set_widths(cases, [14, 16, 16, 20, 24, 24, 10, 14, 42, 28, 40, 30, 12, 28, 8])
    cases.freeze_panes = "A2"
    cases.auto_filter.ref = f"A1:{get_column_letter(len(STANDARD_HEADERS))}{max(1, cases.max_row)}"

    if project.accounts:
        account_sheet = wb.create_sheet("测试账号")
        account_sheet.append(["账号", "密码", "显示名称", "角色", "权限"])
        for account in project.accounts:
            account_sheet.append(
                [
                    account.username,
                    account.password,
                    account.display_name,
                    account.role,
                    "、".join(account.permissions),
                ]
            )
        _style_header(account_sheet)
        _set_widths(account_sheet, [18, 14, 18, 18, 80])
        account_sheet.freeze_panes = "A2"

    wb.save(target)
    return target


def create_demo_plan(path: str | Path, base_url: str = "http://127.0.0.1:5000") -> Path:
    """生成一个针对本机志愿填报系统的可执行演示计划。"""
    target = Path(path)
    wb = openpyxl.Workbook()

    guide = wb.active
    guide.title = "说明"
    guide.append(["演示测试计划：志愿填报管理系统"])
    guide.append(["前置步骤会以管理员身份登录；执行前请确认系统已启动并重置演示数据。"])
    guide["A1"].font = Font(size=16, bold=True, color="1E3A8A")
    _set_widths(guide, [100])

    setup = wb.create_sheet("前置步骤")
    setup.append([header for header, _ in STANDARD_HEADERS])
    setup_rows = [
        ["SETUP", "志愿填报系统", "管理员", "登录", "管理员登录", "管理员登录", 1, "goto", base_url + "/admin/login", "", "", "打开管理员登录页", 15000],
        ["SETUP", "志愿填报系统", "管理员", "登录", "管理员登录", "管理员登录", 2, "fill", "input[name='username']", "admin", "", "填写管理员用户名", 15000],
        ["SETUP", "志愿填报系统", "管理员", "登录", "管理员登录", "管理员登录", 3, "fill", "input[name='password']", "admin123", "", "填写管理员密码", 15000],
        ["SETUP", "志愿填报系统", "管理员", "登录", "管理员登录", "管理员登录", 4, "click", "button[type='submit']", "", "", "点击登录", 15000],
        ["SETUP", "志愿填报系统", "管理员", "登录", "管理员登录", "管理员登录", 5, "assert_url", "", "/admin/applications", "", "登录后进入审核页", 15000],
    ]
    for row in setup_rows:
        setup.append(row)
    _style_header(setup)
    _set_widths(setup, [14, 16, 16, 20, 24, 24, 10, 14, 42, 28, 40, 30, 12, 28, 8])
    setup.freeze_panes = "A2"

    cases = wb.create_sheet("自动化用例")
    cases.append([header for header, _ in STANDARD_HEADERS])
    case_rows = [
        ["R001", "志愿填报系统", "管理员", "志愿审核", "志愿审核页面", "管理员查看审核页", 1, "goto", base_url + "/admin/applications", "", "", "打开审核页", 15000],
        ["R001", "志愿填报系统", "管理员", "志愿审核", "志愿审核页面", "管理员查看审核页", 2, "assert_text", "text=志愿审核", "志愿审核", "", "页面标题正确", 15000],
        ["R001", "志愿填报系统", "管理员", "志愿审核", "志愿审核页面", "管理员查看审核页", 3, "assert_visible", "text=待审核", "", "", "状态标签可见", 15000],
        ["R002", "志愿填报系统", "管理员", "志愿审核", "待审核志愿列表", "查看演示学生志愿", 1, "goto", base_url + "/admin/applications", "", "", "打开审核页", 15000],
        ["R002", "志愿填报系统", "管理员", "志愿审核", "待审核志愿列表", "查看演示学生志愿", 2, "assert_text", "text=张明", "张明", "", "看到待审核学生张明", 15000],
        ["R002", "志愿填报系统", "管理员", "志愿审核", "待审核志愿列表", "查看演示学生志愿", 3, "assert_visible", "text=李小雨", "", "", "看到待审核学生李小雨", 15000],
        ["R003", "志愿填报系统", "学生", "登录", "学生登录页", "学生登录页可打开", 1, "goto", base_url + "/student/login", "", "", "打开学生登录页", 15000],
        ["R003", "志愿填报系统", "学生", "登录", "学生登录页", "学生登录页可打开", 2, "assert_text", "text=学生登录", "学生登录", "", "页面标题正确", 15000],
        ["R003", "志愿填报系统", "学生", "登录", "学生登录页", "学生登录页可打开", 3, "assert_visible", "text=注册学生账号", "", "", "注册入口可见", 15000],
    ]
    for row in case_rows:
        cases.append(row)
    _style_header(cases)
    _set_widths(cases, [14, 16, 16, 20, 24, 24, 10, 14, 42, 28, 40, 30, 12, 28, 8])
    cases.freeze_panes = "A2"
    cases.auto_filter.ref = f"A1:{get_column_letter(len(STANDARD_HEADERS))}1"

    wb.save(target)
    return target
