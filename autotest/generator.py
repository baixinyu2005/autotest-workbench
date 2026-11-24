"""根据老师 Excel 中的功能描述，生成基础 CRUD 用例骨架。

这里只生成“测试意图/用例名称”，不伪造点击步骤。真正的元素定位需要
在拿到被测系统页面后补全；未配置步骤的用例在报告中明确显示为“未配置”。
"""

from __future__ import annotations

import re

from .models import Project, Requirement, TestCase


ACTION_KEYWORDS: list[tuple[str, tuple[str, ...]]] = [
    ("新增", ("新增", "创建", "增加")),
    ("修改", ("修改", "编辑")),
    ("审核", ("审核", "审批")),
    ("删除", ("删除",)),
    ("核销", ("核销",)),
    ("查询", ("查询", "查看", "实时库存")),
    ("生成", ("生成",)),
    ("执行", ("执行",)),
    ("维护", ("维护",)),
    ("配置", ("配置", "设置")),
    ("启用", ("启用",)),
    ("搭建", ("搭建",)),
    ("支撑", ("支撑",)),
]

SPLIT_RE = re.compile(r"[、，,;；/\n]+")
ACTION_PREFIXES = ("由", "按", "进行", "支持", "实时")


def _clean(text: str) -> str:
    return str(text or "").replace("\u3000", " ").strip()


def action_of(part: str) -> str | None:
    """识别一个功能片段对应的操作类型。"""
    text = _clean(part)
    for action, words in ACTION_KEYWORDS:
        if any(word in text for word in words):
            return action
    return None


def split_feature_actions(actual_features: str) -> list[str]:
    """把“客户新增、修改、审核、删除”拆成有序、去重的动作。"""
    actions: list[str] = []
    for part in SPLIT_RE.split(_clean(actual_features)):
        if not part:
            continue
        action = action_of(part)
        if action and action not in actions:
            actions.append(action)
    return actions


def expected_for(action: str, label: str) -> str:
    if action == "查询":
        return f"{label}查询应返回正确结果，页面应显示对应数据或列表。"
    if action == "删除":
        return f"{label}删除应执行成功，记录应从列表或单据中消失。"
    if action in ("审核", "审批"):
        return f"{label}审核应执行成功，单据状态或页面提示应变为已审核。"
    if action == "生成":
        return f"{label}生成操作应成功，并产生对应的下游单据或数据。"
    if action == "执行":
        return f"{label}执行应成功完成，页面应显示执行结果或处理状态。"
    if action == "配置":
        return f"{label}配置应保存成功，并能在页面或后续流程中生效。"
    if action == "查询":
        return f"{label}查询应返回正确结果。"
    return f"{label}{action}应执行成功，页面应显示成功提示，列表/单据状态应发生变化。"


def build_cases(requirement: Requirement) -> list[TestCase]:
    actions = split_feature_actions(requirement.actual_features)
    label = requirement.point or requirement.function or "该功能"

    if not actions:
        actions = ["功能验证"]

    cases: list[TestCase] = []
    for index, action in enumerate(actions, start=1):
        case_id = f"{requirement.req_id}-C{index:02d}"
        kind = action if action != "功能验证" else "功能验证"
        cases.append(
            TestCase(
                case_id=case_id,
                requirement_id=requirement.req_id,
                name=f"{label} - {action}",
                kind=kind,
                expected=expected_for(action, label),
                description=f"依据需规“{requirement.actual_features}”自动生成的基础 {action} 用例骨架。",
                status="not_configured",
            )
        )
    return cases


def generate_cases(project: Project) -> None:
    for requirement in project.requirements:
        requirement.cases = build_cases(requirement)


def summarize(project: Project) -> dict[str, int]:
    total_cases = 0
    configured = 0
    for requirement in project.requirements:
        for case in requirement.cases:
            total_cases += 1
            if case.configured:
                configured += 1
    return {
        "requirements": len(project.requirements),
        "cases": total_cases,
        "configured": configured,
        "unconfigured": total_cases - configured,
    }
