"""从需规工作簿/文件夹中提取测试账号、角色和权限。

支持常见表头：
    账号 / 用户名 / 登录账号
    密码 / 初始密码 / 密码规则
    显示名称 / 姓名 / 角色
    状态 / 权限

也支持权限矩阵：
    功能 | 管理员1 | 管理员2 | ...
    行内用 ✓ / 是 / 支持 表示有权限。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import openpyxl

from .models import TestAccount


ACCOUNT_KEYS = ("账号", "用户名", "登录账号", "用户账号")
PASSWORD_KEYS = ("密码规则", "初始密码", "密码")
DISPLAY_KEYS = ("显示名称", "姓名", "用户名称")
ROLE_KEYS = ("角色", "权限角色", "用户角色")
STATUS_KEYS = ("状态", "账号状态")
PERMISSION_KEYS = ("权限", "主要权限", "功能权限")
SUMMARY_PREFIXES = ("合计", "总计", "说明", "注：", "注:", "备注")
USERNAME_RE = re.compile(r"^[A-Za-z0-9_.@-]{2,40}$")
DIGITS_RE = re.compile(r"\d{4,}")


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace("\u3000", " ").strip()


def _norm(value: Any) -> str:
    return "".join(_text(value).split())


def _col(texts: dict[str, int], variants: tuple[str, ...], max_len: int = 12) -> int | None:
    for text, col in texts.items():
        if len(text) > max_len:
            continue
        if any(variant in text for variant in variants):
            return col
    return None


def _header_col(texts: dict[str, int], variants: tuple[str, ...]) -> int | None:
    """账号/密码列必须是短表头，避免把“账号和密码格式校验”误当成表头。"""
    for text, col in texts.items():
        for variant in variants:
            if text == variant:
                return col
            if text.startswith(variant) and len(text) <= len(variant) + 6:
                return col
    return None


def _password_from_rule(username: str, rule: str) -> str:
    rule = _text(rule)
    if not rule:
        return ""
    if "后四位" in rule or "后4位" in rule:
        return username[-4:]
    found = DIGITS_RE.search(rule)
    return found.group(0) if found else ""


def _is_summary(value: str) -> bool:
    return bool(value) and value.startswith(SUMMARY_PREFIXES)


def _permission_cell(value: Any) -> bool:
    text = _text(value)
    if not text:
        return False
    return text not in ("—", "-", "﹣", "×", "否", "无", "N", "n")


def extract_accounts_from_workbook(wb) -> list[TestAccount]:
    accounts: dict[str, TestAccount] = {}

    for ws in wb.worksheets:
        max_row = min(int(ws.max_row or 0), 120)
        max_col = min(int(ws.max_column or 0), 40)
        if max_row <= 0:
            continue

        for row in range(1, max_row + 1):
            texts: dict[str, int] = {}
            for col in range(1, max_col + 1):
                text = _norm(ws.cell(row, col).value)
                if text:
                    texts[text] = col

            account_col = _header_col(texts, ACCOUNT_KEYS)
            password_col = _header_col(texts, PASSWORD_KEYS)
            if account_col is None or password_col is None or account_col == password_col:
                continue

            display_col = _col(texts, DISPLAY_KEYS)
            role_col = _col(texts, ROLE_KEYS)
            status_col = _col(texts, STATUS_KEYS)
            permission_col = _col(texts, PERMISSION_KEYS)

            for data_row in range(row + 1, max_row + 1):
                username = _text(ws.cell(data_row, account_col).value)
                if not username or _is_summary(username):
                    continue
                if not USERNAME_RE.match(username):
                    continue
                status = _text(ws.cell(data_row, status_col).value) if status_col else ""
                if status and any(word in status for word in ("停用", "禁用", "锁定")):
                    continue
                display_name = (
                    _text(ws.cell(data_row, display_col).value) if display_col else ""
                )
                role = _text(ws.cell(data_row, role_col).value) if role_col else display_name
                password = _password_from_rule(
                    username, _text(ws.cell(data_row, password_col).value)
                )
                direct_permission = (
                    _text(ws.cell(data_row, permission_col).value)
                    if permission_col
                    else ""
                )
                permissions = (
                    [item.strip() for item in re.split(r"[、,，;；]", direct_permission) if item.strip()]
                    if direct_permission
                    else []
                )
                accounts.setdefault(
                    username,
                    TestAccount(
                        username=username,
                        password=password,
                        display_name=display_name or username,
                        role=role,
                        permissions=permissions,
                        source=ws.title,
                    ),
                )

        # 权限矩阵：功能 | 管理员1 | 管理员2 ...
        for row in range(1, max_row + 1):
            headers: dict[str, int] = {}
            for col in range(1, max_col + 1):
                text = _text(ws.cell(row, col).value)
                if text:
                    headers[text] = col
            if not any("功能" in text for text in headers):
                continue
            account_columns: dict[str, int] = {}
            for header, col in headers.items():
                for account in accounts.values():
                    if header in (account.display_name, account.role, account.username):
                        account_columns[account.username] = col
            if not account_columns:
                continue
            for data_row in range(row + 1, max_row + 1):
                feature = _text(ws.cell(data_row, 1).value)
                if not feature or _is_summary(feature):
                    continue
                for username, col in account_columns.items():
                    if _permission_cell(ws.cell(data_row, col).value):
                        permissions = accounts[username].permissions
                        if feature not in permissions:
                            permissions.append(feature)
            break

    return list(accounts.values())


def extract_accounts_from_folder(folder: str | Path, max_files: int = 30) -> list[TestAccount]:
    root = Path(folder)
    if not root.exists():
        return []
    accounts: dict[str, TestAccount] = {}
    count = 0
    for path in root.glob("*"):
        if count >= max_files:
            break
        if not path.is_file() or path.suffix.lower() not in (".xlsx", ".xlsm"):
            continue
        if path.name.startswith("~$"):
            continue
        count += 1
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        except Exception:
            continue
        try:
            for account in extract_accounts_from_workbook(wb):
                accounts.setdefault(account.username, account)
        finally:
            wb.close()
    return list(accounts.values())
