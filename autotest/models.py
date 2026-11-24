"""核心数据模型：需求、用例、步骤、项目。"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Iterator


STATUS_LABELS = {
    "pending": "待运行",
    "pass": "通过",
    "fail": "不通过",
    "not_configured": "未配置",
    "skipped": "跳过",
    "error": "执行错误",
}

STATUS_ORDER = ("pass", "fail", "not_configured", "skipped", "error", "pending")

ACTION_LABELS = {
    "goto": "打开网址",
    "click": "点击",
    "fill": "填写",
    "select": "选择",
    "press": "按键",
    "wait": "等待",
    "assert_text": "断言文本",
    "assert_visible": "断言可见",
    "assert_url": "断言网址",
    "assert_count": "断言元素数量",
    "assert_sorted": "断言数值排序",
    "assert_duplicate_max": "断言重复上限",
    "assert_attribute": "断言属性",
    "assert_value": "断言输入值",
    "assert_hidden": "断言隐藏",
    "emulate_media": "切换媒体类型",
    "reload": "刷新页面",
    "screenshot": "截图",
}


@dataclass
class Step:
    action: str
    target: str = ""
    value: str = ""
    expected: str = ""
    description: str = ""
    timeout_ms: int = 15000

    def to_dict(self) -> dict[str, Any]:
        return {
            "action": self.action,
            "target": self.target,
            "value": self.value,
            "expected": self.expected,
            "description": self.description,
            "timeout_ms": self.timeout_ms,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Step":
        return cls(
            action=str(data.get("action", "")),
            target=str(data.get("target", "")),
            value=str(data.get("value", "")),
            expected=str(data.get("expected", "")),
            description=str(data.get("description", "")),
            timeout_ms=int(data.get("timeout_ms", 15000) or 15000),
        )


@dataclass
class TestCase:
    case_id: str
    requirement_id: str
    name: str
    kind: str = ""
    steps: list[Step] = field(default_factory=list)
    expected: str = ""
    description: str = ""
    status: str = "pending"
    actual: str = ""
    error: str = ""
    screenshot: str = ""
    duration_s: float = 0.0
    executed: bool = False

    @property
    def configured(self) -> bool:
        return bool(self.steps)

    def reset_result(self) -> None:
        self.status = "pending" if self.steps else "not_configured"
        self.actual = ""
        self.error = ""
        self.screenshot = ""
        self.duration_s = 0.0
        self.executed = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "case_id": self.case_id,
            "requirement_id": self.requirement_id,
            "name": self.name,
            "kind": self.kind,
            "steps": [step.to_dict() for step in self.steps],
            "expected": self.expected,
            "description": self.description,
            "status": self.status,
            "actual": self.actual,
            "error": self.error,
            "screenshot": self.screenshot,
            "duration_s": self.duration_s,
            "executed": self.executed,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TestCase":
        case = cls(
            case_id=str(data.get("case_id", "")),
            requirement_id=str(data.get("requirement_id", "")),
            name=str(data.get("name", "")),
            kind=str(data.get("kind", "")),
            steps=[Step.from_dict(item) for item in data.get("steps", [])],
            expected=str(data.get("expected", "")),
            description=str(data.get("description", "")),
            status=str(data.get("status", "pending")),
            actual=str(data.get("actual", "")),
            error=str(data.get("error", "")),
            screenshot=str(data.get("screenshot", "")),
            duration_s=float(data.get("duration_s", 0.0) or 0.0),
            executed=bool(data.get("executed", False)),
        )
        return case


@dataclass
class TestAccount:
    username: str
    password: str = ""
    display_name: str = ""
    role: str = ""
    permissions: list[str] = field(default_factory=list)
    source: str = ""

    def label(self) -> str:
        name = self.display_name or self.username
        role = f"｜{self.role}" if self.role and self.role != name else ""
        return f"{name}（{self.username}）{role}"

    def to_dict(self, include_password: bool = False) -> dict[str, Any]:
        return {
            "username": self.username,
            "password": self.password if include_password else "",
            "display_name": self.display_name,
            "role": self.role,
            "permissions": list(self.permissions),
            "source": self.source,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TestAccount":
        return cls(
            username=str(data.get("username", "")),
            password=str(data.get("password", "")),
            display_name=str(data.get("display_name", "")),
            role=str(data.get("role", "")),
            permissions=[str(item) for item in data.get("permissions", [])],
            source=str(data.get("source", "")),
        )


@dataclass
class Requirement:
    req_id: str
    seq: str
    system: str = ""
    module: str = ""
    function: str = ""
    point: str = ""
    actual_features: str = ""
    remark: str = ""
    source_sheet: str = ""
    source_row: int = 0
    cases: list[TestCase] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "req_id": self.req_id,
            "seq": self.seq,
            "system": self.system,
            "module": self.module,
            "function": self.function,
            "point": self.point,
            "actual_features": self.actual_features,
            "remark": self.remark,
            "source_sheet": self.source_sheet,
            "source_row": self.source_row,
            "cases": [case.to_dict() for case in self.cases],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Requirement":
        return cls(
            req_id=str(data.get("req_id", "")),
            seq=str(data.get("seq", "")),
            system=str(data.get("system", "")),
            module=str(data.get("module", "")),
            function=str(data.get("function", "")),
            point=str(data.get("point", "")),
            actual_features=str(data.get("actual_features", "")),
            remark=str(data.get("remark", "")),
            source_sheet=str(data.get("source_sheet", "")),
            source_row=int(data.get("source_row", 0) or 0),
            cases=[TestCase.from_dict(item) for item in data.get("cases", [])],
        )


def new_project_id() -> str:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"project_{stamp}_{uuid.uuid4().hex[:6]}"


@dataclass
class Project:
    project_id: str
    name: str
    source_excel: str
    source_sheet: str = ""
    imported_at: str = ""
    setup_steps: list[Step] = field(default_factory=list)
    requirements: list[Requirement] = field(default_factory=list)
    accounts: list[TestAccount] = field(default_factory=list)
    notes: str = ""

    @classmethod
    def create(cls, name: str, source_excel: str, source_sheet: str = "") -> "Project":
        return cls(
            project_id=new_project_id(),
            name=name,
            source_excel=source_excel,
            source_sheet=source_sheet,
            imported_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )

    def iter_cases(self) -> Iterator[TestCase]:
        for requirement in self.requirements:
            yield from requirement.cases

    @property
    def case_count(self) -> int:
        return sum(1 for _ in self.iter_cases())

    @property
    def configured_case_count(self) -> int:
        return sum(1 for case in self.iter_cases() if case.configured)

    def status_counts(self) -> dict[str, int]:
        counts = {key: 0 for key in STATUS_ORDER}
        for case in self.iter_cases():
            counts[case.status] = counts.get(case.status, 0) + 1
        return counts

    def reset_results(self) -> None:
        for case in self.iter_cases():
            case.reset_result()

    def to_dict(self) -> dict[str, Any]:
        return {
            "project_id": self.project_id,
            "name": self.name,
            "source_excel": self.source_excel,
            "source_sheet": self.source_sheet,
            "imported_at": self.imported_at,
            "setup_steps": [step.to_dict() for step in self.setup_steps],
            "requirements": [req.to_dict() for req in self.requirements],
            "accounts": [account.to_dict(include_password=False) for account in self.accounts],
            "notes": self.notes,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Project":
        return cls(
            project_id=str(data.get("project_id", "")),
            name=str(data.get("name", "")),
            source_excel=str(data.get("source_excel", "")),
            source_sheet=str(data.get("source_sheet", "")),
            imported_at=str(data.get("imported_at", "")),
            setup_steps=[Step.from_dict(item) for item in data.get("setup_steps", [])],
            requirements=[
                Requirement.from_dict(item) for item in data.get("requirements", [])
            ],
            accounts=[
                TestAccount.from_dict(item) for item in data.get("accounts", [])
            ],
            notes=str(data.get("notes", "")),
        )


def status_label(status: str) -> str:
    return STATUS_LABELS.get(status, status)
