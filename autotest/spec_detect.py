"""需规 Excel 自动发现。

真实需规工作表名和表头并不固定。本模块只做“候选识别”：
- 标准自动化用例表
- 验收标准 / 测试用例表
- 功能需求表
- 功能清单（系统/模块/功能/功能点/实际应用功能）

支持扫描单个工作簿，也支持扫描一个文件夹并选出最像需规的候选。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl


KIND_LABELS = {
    "standard": "标准自动化用例",
    "acceptance": "验收/测试用例",
    "functional": "功能需求",
    "checklist": "功能清单",
}

KIND_PRIORITY = {"standard": 4, "acceptance": 3, "functional": 2, "checklist": 1}


def _norm(value: Any) -> str:
    if value is None:
        return ""
    return "".join(str(value).replace("\u3000", " ").split())


def _row_texts_values(values: tuple[Any, ...], max_cols: int = 80) -> dict[str, int]:
    result: dict[str, int] = {}
    for index, value in enumerate(values[:max_cols], start=1):
        text = _norm(value)
        if text:
            result[text] = index
    return result


def _row_texts(ws, row: int) -> dict[str, int]:
    result: dict[str, int] = {}
    for col in range(1, min(ws.max_column, 80) + 1):
        text = _norm(ws.cell(row, col).value)
        if text:
            result[text] = col
    return result


def _has(texts: dict[str, int], *keywords: str) -> bool:
    return any(all(keyword in text for keyword in keywords) for text in texts)


def _score_standard(texts: dict[str, int]) -> int:
    score = 0
    if _has(texts, "需求ID"):
        score += 50
    if _has(texts, "操作"):
        score += 30
    if _has(texts, "目标"):
        score += 15
    if _has(texts, "用例名称"):
        score += 10
    return score


def _score_acceptance(texts: dict[str, int]) -> int:
    score = 0
    if _has(texts, "用例编号") or _has(texts, "用例ID"):
        score += 40
    if _has(texts, "操作步骤") or _has(texts, "测试步骤") or _has(texts, "步骤"):
        score += 35
    if _has(texts, "预期结果") or _has(texts, "期望结果"):
        score += 25
    if _has(texts, "场景"):
        score += 15
    if _has(texts, "关联需求") or _has(texts, "需求编号"):
        score += 10
    return score


def _score_functional(texts: dict[str, int]) -> int:
    score = 0
    if _has(texts, "需求编号") or _has(texts, "需求ID"):
        score += 45
    if _has(texts, "需求描述") or _has(texts, "需求说明"):
        score += 30
    if _has(texts, "功能名称") or _has(texts, "功能"):
        score += 15
    if _has(texts, "验收要点") or _has(texts, "验收标准"):
        score += 10
    if _has(texts, "一级模块") or _has(texts, "模块"):
        score += 5
    return score


def _score_checklist(texts: dict[str, int]) -> int:
    score = 0
    if _has(texts, "序号"):
        score += 20
    if _has(texts, "功能点"):
        score += 30
    if _has(texts, "实际", "功能"):
        score += 30
    if _has(texts, "系统") or _has(texts, "模块"):
        score += 10
    return score


@dataclass
class SheetCandidate:
    path: Path
    sheet: str
    kind: str
    header_row: int
    score: int
    rows: int
    title: str

    @property
    def kind_label(self) -> str:
        return KIND_LABELS.get(self.kind, self.kind)

    @property
    def priority(self) -> int:
        return KIND_PRIORITY.get(self.kind, 0)

    def describe(self) -> str:
        return (
            f"{self.kind_label}｜{self.path.name} → {self.sheet}｜"
            f"约 {self.rows} 行｜置信度 {self.score}"
        )


@dataclass
class FolderScan:
    candidates: list[SheetCandidate]

    @property
    def best(self) -> SheetCandidate | None:
        return self.candidates[0] if self.candidates else None

    def top(self, count: int = 8) -> list[SheetCandidate]:
        return self.candidates[:count]


def _sheet_title_from_rows(rows: list[tuple[Any, ...]], fallback: str) -> str:
    for row in rows[:4]:
        for value in row[:10]:
            text = str(value or "").strip()
            if text:
                return text[:80]
    return fallback


def _scan_sheet(ws, path: Path) -> SheetCandidate | None:
    max_rows = min(35, int(ws.max_row or 0))
    max_cols = min(80, int(ws.max_column or 1))
    rows: list[tuple[Any, ...]] = []
    if max_rows:
        rows = list(
            ws.iter_rows(
                min_row=1,
                max_row=max_rows,
                max_col=max_cols,
                values_only=True,
            )
        )

    best_kind = ""
    best_row = 0
    best_score = 0

    for index, row_values in enumerate(rows):
        row = index + 1
        single = _row_texts_values(row_values, max_cols)
        combined = dict(single)
        if index + 1 < len(rows):
            combined.update(_row_texts_values(rows[index + 1], max_cols))

        for kind, scorer, texts in (
            ("standard", _score_standard, single),
            ("acceptance", _score_acceptance, single),
            ("functional", _score_functional, single),
            ("checklist", _score_checklist, combined),
        ):
            score = scorer(texts)
            if score > best_score:
                best_kind, best_row, best_score = kind, row, score

    if not best_kind or best_score < {
        "standard": 80,
        "acceptance": 70,
        "functional": 70,
        "checklist": 65,
    }.get(best_kind, 70):
        return None

    title = str(ws.title)
    title_bonus = 0
    if "验收" in title or "测试用例" in title:
        title_bonus = 20
    elif "功能需求" in title or "需求规格" in title:
        title_bonus = 12
    elif "测试内容" in title or "功能清单" in title:
        title_bonus = 10

    return SheetCandidate(
        path=path,
        sheet=ws.title,
        kind=best_kind,
        header_row=best_row,
        score=best_score + title_bonus,
        rows=max(0, int(ws.max_row or 0) - best_row),
        title=_sheet_title_from_rows(rows, ws.title),
    )


def scan_workbook(path: str | Path) -> list[SheetCandidate]:
    source = Path(path)
    try:
        wb = openpyxl.load_workbook(source, data_only=True, read_only=True)
    except Exception:
        return []
    candidates: list[SheetCandidate] = []
    try:
        for ws in wb.worksheets:
            candidate = _scan_sheet(ws, source)
            if candidate:
                candidates.append(candidate)
    finally:
        wb.close()
    candidates.sort(key=lambda item: (item.priority, item.score, item.rows), reverse=True)
    return candidates


def scan_folder(folder: str | Path, max_files: int = 100) -> FolderScan:
    root = Path(folder)
    if not root.exists() or not root.is_dir():
        return FolderScan([])

    def collect(paths) -> list[SheetCandidate]:
        candidates: list[SheetCandidate] = []
        count = 0
        for path in paths:
            if count >= max_files:
                break
            if not path.is_file() or path.suffix.lower() not in (".xlsx", ".xlsm"):
                continue
            if path.name.startswith("~$"):
                continue
            count += 1
            try:
                candidates.extend(scan_workbook(path))
            except Exception:
                continue
        return candidates

    # 先扫描文件夹顶层，通常需规就在这一层；只有顶层没有候选时才递归。
    candidates = collect(root.glob("*"))
    if not candidates:
        candidates = collect(root.rglob("*"))

    candidates.sort(key=lambda item: (item.priority, item.score, item.rows), reverse=True)
    return FolderScan(candidates)
