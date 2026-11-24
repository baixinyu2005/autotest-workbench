"""Tkinter 桌面界面。"""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from .excel_import import (
    create_demo_plan,
    create_standard_template,
    export_project_plan,
    import_folder,
    import_workbook,
)
from .models import ACTION_LABELS, Project, Requirement, Step, status_label
from .runner import RunOptions, RunResult, WebTestRunner, check_web
from .storage import AppStorage, is_on_c_drive


COLORS = {
    "bg": "#f3f6fb",
    "sidebar": "#0f172a",
    "sidebar_active": "#1e293b",
    "sidebar_text": "#cbd5e1",
    "white": "#ffffff",
    "text": "#1f2937",
    "muted": "#667085",
    "border": "#dbe3ee",
    "primary": "#2563eb",
    "primary_dark": "#1d4ed8",
    "success": "#047857",
    "danger": "#b42318",
    "warning": "#b45309",
    "soft": "#f8fafc",
}

FONT = "Microsoft YaHei UI"


class CaseEditorDialog(tk.Toplevel):
    """编辑一条需求下的测试步骤。

    只有配置了步骤的用例才会进入可执行状态；空步骤仍标记为“未配置”。
    """

    def __init__(self, parent: tk.Misc, requirement: Requirement, on_save):
        super().__init__(parent)
        self.title(f"编辑用例 - {requirement.point or requirement.function}")
        self.geometry("1140x740")
        self.minsize(980, 640)
        self.configure(bg=COLORS["bg"])
        self.transient(parent)
        self.grab_set()

        self.requirement = requirement
        self.on_save = on_save
        self.case_steps: dict[str, list[Step]] = {
            case.case_id: [Step.from_dict(step.to_dict()) for step in case.steps]
            for case in requirement.cases
        }
        self.current_case_id: str | None = None
        self.current_steps: list[Step] = []
        self.editing_index: int | None = None
        self.display_to_code = {
            f"{label} ({code})": code for code, label in ACTION_LABELS.items()
        }
        self.code_to_display = {code: display for display, code in self.display_to_code.items()}

        self._build()
        if requirement.cases:
            first = requirement.cases[0].case_id
            self.case_tree.selection_set(first)
            self.case_tree.focus(first)
            self._show_case(first)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    # ------------------------------------------------------------------
    @staticmethod
    def _card(parent: tk.Widget, padx: int = 12, pady: int = 12) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            padx=padx,
            pady=pady,
        )

    def _build(self) -> None:
        header = tk.Frame(self, bg=COLORS["bg"], padx=22, pady=8)
        header.pack(fill="x", pady=(16, 8))
        tk.Label(
            header,
            text=self.requirement.point or self.requirement.function or "未命名需求",
            bg=COLORS["bg"],
            fg=COLORS["text"],
            font=(FONT, 17, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            header,
            text=(
                f"{self.requirement.system} / {self.requirement.module} / "
                f"{self.requirement.function}　需求序号：{self.requirement.seq}"
            ),
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=(FONT, 9),
            anchor="w",
        ).pack(fill="x", pady=(3, 0))
        if self.requirement.actual_features:
            tk.Label(
                header,
                text="需规描述：" + self.requirement.actual_features,
                bg=COLORS["bg"],
                fg=COLORS["muted"],
                font=(FONT, 9),
                anchor="w",
                wraplength=1050,
                justify="left",
            ).pack(fill="x", pady=(3, 0))

        body = tk.Frame(self, bg=COLORS["bg"], padx=22, pady=8)
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        left = self._card(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.grid_rowconfigure(1, weight=1)
        ttk.Label(left, text="用例列表", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        self.case_tree = ttk.Treeview(
            left,
            columns=("name", "steps", "status"),
            show="headings",
            selectmode="browse",
            height=16,
        )
        for col, title, width in (
            ("name", "用例", 170),
            ("steps", "步骤", 45),
            ("status", "状态", 65),
        ):
            self.case_tree.heading(col, text=title)
            self.case_tree.column(col, width=width, minwidth=40, anchor="w")
        self.case_tree.grid(row=1, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.case_tree.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        self.case_tree.configure(yscrollcommand=scroll.set)
        self.case_tree.bind("<<TreeviewSelect>>", self._on_case_select)
        self._fill_cases()

        right = self._card(body)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        ttk.Label(right, text="步骤明细", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        self.step_tree = ttk.Treeview(
            right,
            columns=("no", "action", "target", "value", "expected", "timeout"),
            show="headings",
            selectmode="browse",
            height=9,
        )
        for col, title, width in (
            ("no", "序号", 45),
            ("action", "操作", 90),
            ("target", "目标", 250),
            ("value", "输入值", 150),
            ("expected", "预期", 150),
            ("timeout", "超时", 65),
        ):
            self.step_tree.heading(col, text=title)
            self.step_tree.column(col, width=width, minwidth=40, anchor="w")
        self.step_tree.grid(row=1, column=0, sticky="nsew")
        sy = ttk.Scrollbar(right, orient="vertical", command=self.step_tree.yview)
        sy.grid(row=1, column=1, sticky="ns")
        sx = ttk.Scrollbar(right, orient="horizontal", command=self.step_tree.xview)
        sx.grid(row=2, column=0, sticky="ew")
        self.step_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.step_tree.bind("<<TreeviewSelect>>", self._on_step_select)

        form = tk.Frame(right, bg=COLORS["white"])
        form.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        form.grid_columnconfigure(1, weight=1)
        form.grid_columnconfigure(3, weight=1)

        self.action_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.value_var = tk.StringVar()
        self.expected_var = tk.StringVar()
        self.description_var = tk.StringVar()
        self.timeout_var = tk.StringVar(value="15000")

        ttk.Label(form, text="操作", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(
            form,
            textvariable=self.action_var,
            values=list(self.display_to_code.keys()),
            state="readonly",
            width=20,
        ).grid(row=0, column=1, sticky="w", padx=(8, 16), pady=4)
        ttk.Label(form, text="超时毫秒", style="Card.TLabel").grid(row=0, column=2, sticky="w", pady=4)
        ttk.Spinbox(
            form, from_=1000, to=120000, increment=1000,
            textvariable=self.timeout_var, width=10,
        ).grid(row=0, column=3, sticky="w", padx=(8, 0), pady=4)

        ttk.Label(form, text="目标选择器/网址", style="Card.TLabel").grid(
            row=1, column=0, sticky="w", pady=4
        )
        ttk.Entry(form, textvariable=self.target_var, font=(FONT, 9)).grid(
            row=1, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=4, ipady=3
        )

        ttk.Label(form, text="输入值/断言文本", style="Card.TLabel").grid(
            row=2, column=0, sticky="w", pady=4
        )
        ttk.Entry(form, textvariable=self.value_var, font=(FONT, 9)).grid(
            row=2, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=4, ipady=3
        )

        ttk.Label(form, text="预期结果", style="Card.TLabel").grid(
            row=3, column=0, sticky="w", pady=4
        )
        ttk.Entry(form, textvariable=self.expected_var, font=(FONT, 9)).grid(
            row=3, column=1, sticky="ew", padx=(8, 16), pady=4, ipady=3
        )
        ttk.Label(form, text="说明", style="Card.TLabel").grid(
            row=3, column=2, sticky="w", pady=4
        )
        ttk.Entry(form, textvariable=self.description_var, font=(FONT, 9)).grid(
            row=3, column=3, sticky="ew", padx=(8, 0), pady=4, ipady=3
        )

        buttons = tk.Frame(form, bg=COLORS["white"])
        buttons.grid(row=4, column=0, columnspan=4, sticky="w", pady=(10, 0))
        ttk.Button(
            buttons, text="添加步骤", style="Primary.TButton", command=self._save_step
        ).pack(side="left")
        ttk.Button(
            buttons, text="更新选中", style="Secondary.TButton", command=self._save_step
        ).pack(side="left", padx=8)
        ttk.Button(
            buttons, text="删除选中", style="Secondary.TButton", command=self._delete_step
        ).pack(side="left")
        ttk.Button(
            buttons, text="上移", style="Secondary.TButton", command=lambda: self._move_step(-1)
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            buttons, text="下移", style="Secondary.TButton", command=lambda: self._move_step(1)
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            buttons, text="清空表单", style="Secondary.TButton", command=self._clear_form
        ).pack(side="left", padx=(8, 0))

        hint = tk.Label(
            right,
            text=(
                "提示：goto 的“目标”填网址；click/fill 等填 Playwright 选择器，"
                "例如 button:has-text('新增')、input[name='code']。空步骤用例不会被算作通过。"
            ),
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 8),
            anchor="w",
            justify="left",
            wraplength=820,
        )
        hint.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))

        footer = tk.Frame(self, bg=COLORS["bg"], padx=22, pady=8)
        footer.pack(fill="x", pady=(4, 16))
        ttk.Button(
            footer, text="保存并关闭", style="Primary.TButton", command=self._save_and_close
        ).pack(side="right")
        ttk.Button(
            footer, text="取消", style="Secondary.TButton", command=self.destroy
        ).pack(side="right", padx=(0, 8))
        tk.Label(
            footer,
            text=f"需求 {self.requirement.req_id}　共 {len(self.requirement.cases)} 个用例",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=(FONT, 9),
        ).pack(side="left")

    # ------------------------------------------------------------------
    def _fill_cases(self) -> None:
        for item in self.case_tree.get_children():
            self.case_tree.delete(item)
        for case in self.requirement.cases:
            steps = self.case_steps.get(case.case_id, [])
            self.case_tree.insert(
                "",
                "end",
                iid=case.case_id,
                values=(
                    case.name,
                    len(steps),
                    "可执行" if steps else "未配置",
                ),
            )

    def _fill_steps(self) -> None:
        for item in self.step_tree.get_children():
            self.step_tree.delete(item)
        for index, step in enumerate(self.current_steps):
            self.step_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    index + 1,
                    ACTION_LABELS.get(step.action, step.action),
                    step.target,
                    step.value,
                    step.expected,
                    step.timeout_ms,
                ),
            )

    def _on_case_select(self, _event=None) -> None:
        selected = self.case_tree.selection()
        if not selected:
            return
        self._show_case(selected[0])

    def _show_case(self, case_id: str) -> None:
        self.current_case_id = case_id
        self.current_steps = self.case_steps.setdefault(case_id, [])
        self.editing_index = None
        self._clear_form()
        self._fill_steps()

    def _on_step_select(self, _event=None) -> None:
        selected = self.step_tree.selection()
        if not selected:
            return
        try:
            index = int(selected[0])
        except ValueError:
            return
        if index < 0 or index >= len(self.current_steps):
            return
        step = self.current_steps[index]
        self.editing_index = index
        self.action_var.set(self.code_to_display.get(step.action, step.action))
        self.target_var.set(step.target)
        self.value_var.set(step.value)
        self.expected_var.set(step.expected)
        self.description_var.set(step.description)
        self.timeout_var.set(str(step.timeout_ms))

    def _clear_form(self) -> None:
        self.editing_index = None
        self.action_var.set(self.code_to_display.get("goto", "打开网址 (goto)"))
        self.target_var.set("")
        self.value_var.set("")
        self.expected_var.set("")
        self.description_var.set("")
        self.timeout_var.set("15000")
        self.step_tree.selection_remove(self.step_tree.selection())

    def _read_step(self) -> Step | None:
        if not self.current_case_id:
            messagebox.showinfo("未选择用例", "请先选择左侧的用例。", parent=self)
            return None
        display = self.action_var.get().strip()
        action = self.display_to_code.get(display, display)
        if not action:
            messagebox.showwarning("缺少操作", "请选择操作类型。", parent=self)
            return None
        target = self.target_var.get().strip()
        value = self.value_var.get().strip()
        expected = self.expected_var.get().strip()
        if action in ("goto", "click", "fill", "select", "press", "assert_text", "assert_visible") and not target:
            messagebox.showwarning("缺少目标", f"{ACTION_LABELS.get(action, action)} 需要填写目标。", parent=self)
            return None
        if action == "assert_url" and not (value or expected):
            messagebox.showwarning("缺少断言", "assert_url 需要填写输入值或预期结果。", parent=self)
            return None
        try:
            timeout_ms = max(1000, min(120000, int(float(self.timeout_var.get()))))
        except ValueError:
            messagebox.showwarning("超时无效", "超时毫秒必须是数字。", parent=self)
            return None
        return Step(
            action=action,
            target=target,
            value=value,
            expected=expected,
            description=self.description_var.get().strip(),
            timeout_ms=timeout_ms,
        )

    def _save_step(self) -> None:
        step = self._read_step()
        if step is None:
            return
        if self.editing_index is None:
            self.current_steps.append(step)
        else:
            self.current_steps[self.editing_index] = step
        self.editing_index = None
        self._clear_form()
        self._fill_steps()
        self._fill_cases()
        self.step_tree.selection_remove(self.step_tree.selection())

    def _delete_step(self) -> None:
        selected = self.step_tree.selection()
        if not selected:
            messagebox.showinfo("未选择步骤", "请先选择要删除的步骤。", parent=self)
            return
        index = int(selected[0])
        if 0 <= index < len(self.current_steps):
            del self.current_steps[index]
            self.editing_index = None
            self._clear_form()
            self._fill_steps()
            self._fill_cases()

    def _move_step(self, direction: int) -> None:
        selected = self.step_tree.selection()
        if not selected:
            return
        index = int(selected[0])
        new_index = index + direction
        if new_index < 0 or new_index >= len(self.current_steps):
            return
        self.current_steps[index], self.current_steps[new_index] = (
            self.current_steps[new_index],
            self.current_steps[index],
        )
        self._fill_steps()
        self.step_tree.selection_set(str(new_index))
        self.step_tree.focus(str(new_index))
        self.editing_index = None
        self._clear_form()

    def _save_and_close(self) -> None:
        for case in self.requirement.cases:
            steps = self.case_steps.get(case.case_id, [])
            case.steps = [Step.from_dict(step.to_dict()) for step in steps]
            case.status = "pending" if case.steps else "not_configured"
            case.actual = ""
            case.error = ""
            case.screenshot = ""
            case.duration_s = 0.0
            case.executed = False
        self.on_save()
        self.destroy()


class AutoTestApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("AutoTest Workbench - 桌面自动化测试校验系统")
        self.geometry("1320x860")
        self.minsize(1120, 720)
        self.configure(bg=COLORS["bg"])

        self.storage = AppStorage()
        self.project: Project | None = None
        self.run_thread: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.events: queue.Queue[tuple[Any, ...]] = queue.Queue()
        self.last_report: Path | None = None
        self.last_run_dir: Path | None = None
        self.live_counts = {
            "pass": 0,
            "fail": 0,
            "not_configured": 0,
            "skipped": 0,
        }
        self.report_map: dict[str, Path] = {}
        self.nav_buttons: dict[str, tk.Button] = {}
        self.pages: dict[str, tk.Frame] = {}
        self.page_meta: dict[str, tuple[str, str]] = {}

        self._build_style()
        self._build_layout()
        self._load_config()
        self.show_page("dashboard")
        self.after(100, self._poll_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # 样式与基础布局
    # ------------------------------------------------------------------
    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=(FONT, 10))
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["white"])
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"])
        style.configure("Card.TLabel", background=COLORS["white"], foreground=COLORS["text"])
        style.configure(
            "Muted.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=(FONT, 9),
        )
        style.configure(
            "CardMuted.TLabel",
            background=COLORS["white"],
            foreground=COLORS["muted"],
            font=(FONT, 9),
        )
        style.configure(
            "Title.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=(FONT, 19, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=(FONT, 9),
        )
        style.configure(
            "CardTitle.TLabel",
            background=COLORS["white"],
            foreground=COLORS["text"],
            font=(FONT, 12, "bold"),
        )
        style.configure(
            "Primary.TButton",
            background=COLORS["primary"],
            foreground="white",
            borderwidth=0,
            padding=(16, 9),
            font=(FONT, 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLORS["primary_dark"]), ("disabled", "#93c5fd")],
            foreground=[("disabled", "#eff6ff")],
        )
        style.configure(
            "Secondary.TButton",
            background=COLORS["white"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            borderwidth=1,
            padding=(13, 8),
        )
        style.map("Secondary.TButton", background=[("active", "#f8fafc")])
        style.configure(
            "Danger.TButton",
            background="#b42318",
            foreground="white",
            borderwidth=0,
            padding=(13, 8),
            font=(FONT, 10, "bold"),
        )
        style.map("Danger.TButton", background=[("active", "#912018"), ("disabled", "#fca5a5")],
                  foreground=[("disabled", "#fef2f2")])
        style.configure(
            "Treeview",
            background=COLORS["white"],
            fieldbackground=COLORS["white"],
            foreground=COLORS["text"],
            rowheight=30,
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            background="#f8fafc",
            foreground="#475569",
            font=(FONT, 9, "bold"),
            padding=(8, 7),
        )
        style.map(
            "Treeview",
            background=[("selected", "#dbeafe")],
            foreground=[("selected", "#1e3a8a")],
        )
        style.configure(
            "Horizontal.TProgressbar",
            background=COLORS["primary"],
            troughcolor="#e2e8f0",
            borderwidth=0,
        )

    def _build_layout(self) -> None:
        self.sidebar = tk.Frame(self, bg=COLORS["sidebar"], width=225)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = tk.Frame(self.sidebar, bg=COLORS["sidebar"], padx=22, pady=24)
        brand.pack(fill="x")
        logo = tk.Label(
            brand,
            text="AT",
            bg=COLORS["primary"],
            fg="white",
            font=(FONT, 13, "bold"),
            width=3,
            height=1,
        )
        logo.pack(side="left", padx=(0, 10))
        brand_text = tk.Frame(brand, bg=COLORS["sidebar"])
        brand_text.pack(side="left", fill="x")
        tk.Label(
            brand_text,
            text="AutoTest",
            bg=COLORS["sidebar"],
            fg="white",
            font=(FONT, 13, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            brand_text,
            text="Workbench",
            bg=COLORS["sidebar"],
            fg="#94a3b8",
            font=(FONT, 9),
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            self.sidebar,
            text="测试工作区",
            bg=COLORS["sidebar"],
            fg="#64748b",
            font=(FONT, 8, "bold"),
            anchor="w",
            padx=25,
            pady=8,
        ).pack(fill="x", pady=(8, 5))

        nav_items = [
            ("dashboard", "概览"),
            ("wizard", "测试向导"),
            ("import", "① 导入需规"),
            ("run", "② 网址与账号"),
            ("reports", "测试报告"),
            ("settings", "设置"),
        ]
        for key, text in nav_items:
            button = tk.Button(
                self.sidebar,
                text=text,
                anchor="w",
                padx=26,
                pady=11,
                bd=0,
                relief="flat",
                bg=COLORS["sidebar"],
                fg=COLORS["sidebar_text"],
                activebackground=COLORS["sidebar_active"],
                activeforeground="white",
                font=(FONT, 10),
                cursor="hand2",
                command=lambda page=key: self.show_page(page),
            )
            button.pack(fill="x")
            self.nav_buttons[key] = button

        tk.Frame(self.sidebar, bg=COLORS["sidebar"]).pack(fill="both", expand=True)
        storage_box = tk.Frame(self.sidebar, bg="#111c31", padx=18, pady=14)
        storage_box.pack(fill="x", side="bottom")
        tk.Label(
            storage_box,
            text="数据存储",
            bg="#111c31",
            fg="#94a3b8",
            font=(FONT, 8, "bold"),
            anchor="w",
        ).pack(fill="x")
        self.sidebar_storage = tk.Label(
            storage_box,
            text="",
            bg="#111c31",
            fg="#cbd5e1",
            font=(FONT, 8),
            anchor="w",
            justify="left",
            wraplength=180,
        )
        self.sidebar_storage.pack(fill="x", pady=(4, 0))
        self.sidebar_storage.configure(text=str(self.storage.root))

        self.main = tk.Frame(self, bg=COLORS["bg"])
        self.main.pack(side="left", fill="both", expand=True)
        self.main.grid_rowconfigure(1, weight=1)
        self.main.grid_columnconfigure(0, weight=1)

        topbar = tk.Frame(self.main, bg=COLORS["bg"], padx=30, pady=14)
        topbar.grid(row=0, column=0, sticky="ew", pady=(20, 8))
        titles = tk.Frame(topbar, bg=COLORS["bg"])
        titles.pack(side="left", fill="x", expand=True)
        self.page_title = ttk.Label(titles, text="概览", style="Title.TLabel")
        self.page_title.pack(anchor="w")
        self.page_subtitle = ttk.Label(titles, text="", style="Subtitle.TLabel")
        self.page_subtitle.pack(anchor="w", pady=(2, 0))

        self.top_action = ttk.Button(
            topbar,
            text="打开数据目录",
            style="Secondary.TButton",
            command=lambda: self.storage.open_in_explorer(),
        )
        self.top_action.pack(side="right", padx=(12, 0), pady=4)

        self.content = tk.Frame(self.main, bg=COLORS["bg"])
        self.content.grid(row=1, column=0, sticky="nsew", padx=30, pady=(4, 14))
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="就绪")
        statusbar = tk.Frame(self.main, bg="#eef2f7", height=28)
        statusbar.grid(row=2, column=0, sticky="ew")
        tk.Label(
            statusbar,
            textvariable=self.status_var,
            bg="#eef2f7",
            fg=COLORS["muted"],
            font=(FONT, 8),
            anchor="w",
            padx=30,
        ).pack(fill="x", pady=5)

        self._build_dashboard()
        self._build_import_page()
        self._build_run_page()
        self._build_wizard_page()
        self._build_reports_page()
        self._build_settings_page()

        self.page_meta = {
            "dashboard": ("概览", "导入需规、执行自动化测试并查看结果"),
            "wizard": ("测试向导", "按 1 → 2 → 3 完成一次自动化测试"),
            "import": ("需求导入", "支持送审单位的功能清单 Excel 和标准自动化用例 Excel"),
            "run": ("测试执行", "使用本机 Edge / Chrome 自动操作待测网页"),
            "reports": ("测试报告", "查看历史运行结果和失败证据"),
            "settings": ("设置", "配置数据目录、浏览器和默认参数"),
        }

    # ------------------------------------------------------------------
    # 页面构建
    # ------------------------------------------------------------------
    def _page(self, key: str) -> tk.Frame:
        frame = tk.Frame(self.content, bg=COLORS["bg"])
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.pages[key] = frame
        return frame

    @staticmethod
    def _card(parent: tk.Widget, padx: int = 18, pady: int = 16) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            padx=padx,
            pady=pady,
        )

    def _build_dashboard(self) -> None:
        page = self._page("dashboard")
        page.grid_rowconfigure(1, weight=1)
        page.grid_columnconfigure(0, weight=1)

        hero = tk.Frame(page, bg="#0f172a", padx=28, pady=24)
        hero.grid(row=0, column=0, sticky="new")
        tk.Label(
            hero,
            text="桌面自动化测试校验系统",
            bg="#0f172a",
            fg="white",
            font=(FONT, 22, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            hero,
            text="导入需求 Excel，自动执行网页操作，逐条比对实际结果并生成可追溯的测试报告。",
            bg="#0f172a",
            fg="#cbd5e1",
            font=(FONT, 10),
            anchor="w",
        ).pack(fill="x", pady=(6, 0))
        buttons = tk.Frame(hero, bg="#0f172a")
        buttons.pack(fill="x", pady=(18, 0))
        ttk.Button(
            buttons,
            text="测试向导",
            style="Primary.TButton",
            command=lambda: self.show_page("wizard"),
        ).pack(side="left")
        ttk.Button(
            buttons,
            text="导入需求 Excel",
            style="Secondary.TButton",
            command=lambda: self.show_page("import"),
        ).pack(side="left", padx=10)
        ttk.Button(
            buttons,
            text="开始测试",
            style="Secondary.TButton",
            command=lambda: self.show_page("run"),
        ).pack(side="left")
        ttk.Button(
            buttons,
            text="查看报告",
            style="Secondary.TButton",
            command=lambda: self.show_page("reports"),
        ).pack(side="left", padx=10)

        body = tk.Frame(page, bg=COLORS["bg"])
        body.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        for column in range(4):
            body.grid_columnconfigure(column, weight=1)
        body.grid_rowconfigure(1, weight=1)

        self.dash_cards: dict[str, tk.StringVar] = {}
        cards = [
            ("storage", "数据目录", "非 C 盘存储"),
            ("requirements", "需求条目", "等待导入"),
            ("cases", "测试用例", "等待生成"),
            ("report", "最近测试", "暂无运行"),
        ]
        for index, (key, title, initial) in enumerate(cards):
            card = self._card(body, padx=18, pady=15)
            card.grid(
                row=0,
                column=index,
                sticky="nsew",
                padx=(0 if index == 0 else 7, 0 if index == 3 else 7),
            )
            ttk.Label(card, text=title, style="CardMuted.TLabel").pack(anchor="w")
            var = tk.StringVar(value=initial)
            self.dash_cards[key] = var
            tk.Label(
                card,
                textvariable=var,
                bg=COLORS["white"],
                fg=COLORS["text"],
                font=(FONT, 11, "bold") if key == "storage" else (FONT, 13, "bold"),
                anchor="w",
                wraplength=250 if key == "storage" else 240,
                justify="left",
            ).pack(anchor="w", pady=(5, 0))

        info = self._card(body, padx=24, pady=22)
        info.grid(row=1, column=0, columnspan=4, sticky="nsew", pady=(14, 0))
        ttk.Label(info, text="工作流程", style="CardTitle.TLabel").pack(anchor="w")
        steps = [
            "1. 导入送审单位给的 Excel，系统自动展开合并单元格并识别需求。",
            "2. 功能清单会生成基础 CRUD 用例骨架；带操作步骤的标准用例可直接执行。",
            "3. 在“测试执行”中填写待测网址，使用 Edge/Chrome 自动点击、填写、断言。",
            "4. 失败时保存截图，并生成 HTML + Excel 报告，明确写出实际表现和需求差异。",
        ]
        for text in steps:
            ttk.Label(info, text=text, style="Card.TLabel", wraplength=980).pack(
                anchor="w", pady=(8, 0)
            )

        warning = tk.Frame(info, bg="#fffbeb", padx=14, pady=10)
        warning.pack(fill="x", pady=(18, 0))
        tk.Label(
            warning,
            text="重要：没有配置页面选择器的需求不会被算作“通过”，报告会标记为“未配置”。",
            bg="#fffbeb",
            fg="#92400e",
            font=(FONT, 9),
            anchor="w",
            wraplength=980,
            justify="left",
        ).pack(fill="x")

    def _build_wizard_page(self) -> None:
        page = self._page("wizard")
        page.grid_columnconfigure(1, weight=1)
        page.grid_rowconfigure(0, weight=1)

        stepper = self._card(page, padx=14, pady=18)
        stepper.grid(row=0, column=0, sticky="ns", padx=(0, 14))
        ttk.Label(stepper, text="测试流程", style="CardTitle.TLabel").pack(
            anchor="w", pady=(0, 12)
        )
        self.wizard_step_buttons: dict[int, tk.Button] = {}
        for step, title in ((1, "导入需规"), (2, "软件网址与账号"), (3, "开始测试")):
            button = tk.Button(
                stepper,
                text=f"{step}  {title}",
                anchor="w",
                padx=12,
                pady=10,
                bd=0,
                relief="flat",
                bg=COLORS["white"],
                fg=COLORS["text"],
                activebackground="#eff6ff",
                activeforeground=COLORS["primary"],
                font=(FONT, 10),
                cursor="hand2",
                command=lambda target=step: self._wizard_go(target),
            )
            button.pack(fill="x", pady=3)
            self.wizard_step_buttons[step] = button
        tk.Label(
            stepper,
            text="导入后按下一步，\n最后生成报告。",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 8),
            justify="left",
            anchor="w",
        ).pack(fill="x", pady=(16, 0))

        host = tk.Frame(page, bg=COLORS["bg"])
        host.grid(row=0, column=1, sticky="nsew")
        host.grid_rowconfigure(0, weight=1)
        host.grid_columnconfigure(0, weight=1)
        self.wizard_frames: dict[int, tk.Frame] = {}
        for step in (1, 2, 3):
            frame = tk.Frame(host, bg=COLORS["white"])
            frame.grid(row=0, column=0, sticky="nsew")
            self.wizard_frames[step] = frame

        # Step 1
        step1 = self.wizard_frames[1]
        tk.Label(
            step1,
            text="第 1 步：导入需规",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=(FONT, 19, "bold"),
            anchor="w",
        ).pack(fill="x", padx=24, pady=(24, 4))
        tk.Label(
            step1,
            text="选择 Excel，或扫描整个文件夹。系统会自动识别需规工作表、账号和权限。",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 9),
            anchor="w",
        ).pack(fill="x", padx=24)
        buttons = tk.Frame(step1, bg=COLORS["white"])
        buttons.pack(fill="x", padx=24, pady=(18, 0))
        ttk.Button(
            buttons, text="选择 Excel", style="Primary.TButton",
            command=lambda: (self._choose_excel(), self._refresh_wizard()),
        ).pack(side="left")
        ttk.Button(
            buttons, text="扫描文件夹", style="Secondary.TButton",
            command=lambda: (self._choose_folder(), self._refresh_wizard()),
        ).pack(side="left", padx=8)
        self.wizard_project_var = tk.StringVar(value="尚未导入需规。")
        tk.Label(
            step1,
            textvariable=self.wizard_project_var,
            bg="#f8fafc",
            fg=COLORS["text"],
            font=(FONT, 11, "bold"),
            anchor="w",
            justify="left",
            padx=14,
            pady=14,
        ).pack(fill="x", padx=24, pady=(20, 0))
        self.wizard_accounts_var = tk.StringVar(value="")
        tk.Label(
            step1,
            textvariable=self.wizard_accounts_var,
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 9),
            anchor="w",
            justify="left",
            wraplength=720,
        ).pack(fill="x", padx=24, pady=(10, 0))
        ttk.Button(
            step1, text="下一步：填写软件网址", style="Primary.TButton",
            command=self._wizard_next_from_spec,
        ).pack(anchor="e", padx=24, pady=24)

        # Step 2
        step2 = self.wizard_frames[2]
        tk.Label(
            step2,
            text="第 2 步：软件网址与测试账号",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=(FONT, 19, "bold"),
            anchor="w",
        ).pack(fill="x", padx=24, pady=(24, 4))
        tk.Label(
            step2,
            text="填写被测软件网址，并选择需规里识别出的测试账号。",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 9),
            anchor="w",
        ).pack(fill="x", padx=24)

        form = tk.Frame(step2, bg=COLORS["white"])
        form.pack(fill="x", padx=24, pady=(18, 0))
        form.grid_columnconfigure(1, weight=1)
        ttk.Label(form, text="待测网址", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", pady=6
        )
        ttk.Entry(form, textvariable=self.url_var, font=(FONT, 10)).grid(
            row=0, column=1, sticky="ew", padx=(12, 0), pady=6, ipady=4
        )
        ttk.Label(form, text="测试账号", style="Card.TLabel").grid(
            row=1, column=0, sticky="w", pady=6
        )
        self.wizard_account_combo = ttk.Combobox(
            form,
            textvariable=self.account_var,
            values=[],
            state="readonly",
            width=32,
        )
        self.wizard_account_combo.grid(row=1, column=1, sticky="w", padx=(12, 0), pady=6)
        self.wizard_account_combo.bind("<<ComboboxSelected>>", self._on_account_change)
        ttk.Label(form, text="用户名 / 密码", style="Card.TLabel").grid(
            row=2, column=0, sticky="w", pady=6
        )
        credential_row = tk.Frame(form, bg=COLORS["white"])
        credential_row.grid(row=2, column=1, sticky="ew", padx=(12, 0), pady=6)
        ttk.Entry(credential_row, textvariable=self.username_var, width=18).pack(
            side="left", ipady=3
        )
        ttk.Entry(credential_row, textvariable=self.password_var, width=18, show="*").pack(
            side="left", padx=(8, 0), ipady=3
        )
        ttk.Label(form, text="权限", style="Card.TLabel").grid(
            row=3, column=0, sticky="nw", pady=6
        )
        tk.Label(
            form,
            textvariable=self.permission_var,
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 9),
            anchor="w",
            justify="left",
            wraplength=700,
        ).grid(row=3, column=1, sticky="w", padx=(12, 0), pady=6)
        ttk.Label(form, text="浏览器 / 无头", style="Card.TLabel").grid(
            row=4, column=0, sticky="w", pady=6
        )
        browser_row = tk.Frame(form, bg=COLORS["white"])
        browser_row.grid(row=4, column=1, sticky="w", padx=(12, 0), pady=6)
        ttk.Combobox(
            browser_row,
            textvariable=self.browser_var,
            values=["Edge", "Chrome"],
            state="readonly",
            width=10,
        ).pack(side="left")
        ttk.Checkbutton(
            browser_row, text="无头模式", variable=self.headless_var
        ).pack(side="left", padx=(16, 0))

        nav2 = tk.Frame(step2, bg=COLORS["white"])
        nav2.pack(fill="x", padx=24, pady=24)
        ttk.Button(
            nav2, text="上一步", style="Secondary.TButton",
            command=lambda: self._wizard_go(1),
        ).pack(side="left")
        ttk.Button(
            nav2, text="下一步：开始测试", style="Primary.TButton",
            command=self._wizard_next_from_url,
        ).pack(side="right")

        # Step 3
        step3 = self.wizard_frames[3]
        tk.Label(
            step3,
            text="第 3 步：开始测试",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=(FONT, 19, "bold"),
            anchor="w",
        ).pack(fill="x", padx=24, pady=(24, 4))
        tk.Label(
            step3,
            text="确认信息后开始执行。测试日志和实时进度会显示在“测试执行”页。",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 9),
            anchor="w",
        ).pack(fill="x", padx=24)
        self.wizard_summary_var = tk.StringVar(value="")
        tk.Label(
            step3,
            textvariable=self.wizard_summary_var,
            bg="#f8fafc",
            fg=COLORS["text"],
            font=(FONT, 10),
            anchor="w",
            justify="left",
            padx=14,
            pady=14,
            wraplength=760,
        ).pack(fill="x", padx=24, pady=(20, 0))
        nav3 = tk.Frame(step3, bg=COLORS["white"])
        nav3.pack(fill="x", padx=24, pady=24)
        ttk.Button(
            nav3, text="上一步", style="Secondary.TButton",
            command=lambda: self._wizard_go(2),
        ).pack(side="left")
        ttk.Button(
            nav3, text="开始测试", style="Primary.TButton",
            command=self._wizard_start,
        ).pack(side="right")
        ttk.Button(
            nav3, text="查看测试报告", style="Secondary.TButton",
            command=lambda: self.show_page("reports"),
        ).pack(side="right", padx=(0, 8))

        self.wizard_current_step = 1
        self._wizard_go(1)

    def _wizard_go(self, step: int) -> None:
        step = max(1, min(3, step))
        self.wizard_current_step = step
        for number, frame in self.wizard_frames.items():
            if number == step:
                frame.tkraise()
        for number, button in self.wizard_step_buttons.items():
            if number == step:
                button.configure(bg="#dbeafe", fg="#1e3a8a", font=(FONT, 10, "bold"))
            else:
                button.configure(bg=COLORS["white"], fg=COLORS["text"], font=(FONT, 10))
        self._refresh_wizard()

    def _wizard_next_from_spec(self) -> None:
        if not self.project:
            messagebox.showwarning("未导入需规", "请先导入需规 Excel 或扫描文件夹。")
            return
        self._wizard_go(2)

    def _wizard_next_from_url(self) -> None:
        if not self.url_var.get().strip().startswith(("http://", "https://", "file://")):
            messagebox.showwarning(
                "网址格式不正确",
                "请填写软件网址（http://、https:// 或 file://）。",
            )
            return
        self._wizard_go(3)

    def _wizard_start(self) -> None:
        if not self.project:
            messagebox.showwarning("未导入需规", "请先导入需规 Excel。")
            self._wizard_go(1)
            return
        if not self.url_var.get().strip():
            messagebox.showwarning("网址为空", "请填写被测软件网址。")
            self._wizard_go(2)
            return
        self.show_page("run")
        self._start_test()

    def _refresh_wizard(self) -> None:
        if not hasattr(self, "wizard_project_var"):
            return
        if self.project:
            self.wizard_project_var.set(
                f"{self.project.name}\n"
                f"需求 {len(self.project.requirements)} 条　·　"
                f"用例 {self.project.case_count} 个　·　"
                f"已配置 {self.project.configured_case_count} 个"
            )
            if self.project.accounts:
                labels = [account.label() for account in self.project.accounts]
                self.wizard_accounts_var.set(
                    "识别到测试账号：\n" + "\n".join(f"• {label}" for label in labels)
                )
            else:
                self.wizard_accounts_var.set(
                    "未识别到测试账号；可以在第 2 步手动填写用户名和密码。"
                )
        else:
            self.wizard_project_var.set("尚未导入需规。")
            self.wizard_accounts_var.set("")
        if hasattr(self, "wizard_account_combo"):
            labels = list(self.account_map.keys())
            self.wizard_account_combo.configure(values=labels)
        account = self.account_map.get(self.account_var.get())
        account_text = account.label() if account else "手动填写"
        self.wizard_summary_var.set(
            f"需规：{self.project.name if self.project else '未导入'}\n"
            f"网址：{self.url_var.get().strip() or '未填写'}\n"
            f"账号：{account_text}\n"
            f"用例：{self.project.case_count if self.project else 0} 个，"
            f"已配置：{self.project.configured_case_count if self.project else 0} 个"
        )

    def _build_import_page(self) -> None:
        page = self._page("import")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)

        top = self._card(page, padx=20, pady=18)
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="导入需求 Excel", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            top,
            text="支持送审单位的功能清单（自动展开合并单元格）和带步骤/预期的标准自动化用例。",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(3, 12))

        row = tk.Frame(top, bg=COLORS["white"])
        row.pack(fill="x")
        self.import_path_var = tk.StringVar()
        entry = ttk.Entry(row, textvariable=self.import_path_var, font=(FONT, 10))
        entry.pack(side="left", fill="x", expand=True, ipady=5)
        ttk.Button(
            row, text="选择文件", style="Secondary.TButton", command=self._choose_excel
        ).pack(side="left", padx=(10, 0))
        ttk.Button(
            row, text="扫描文件夹", style="Secondary.TButton", command=self._choose_folder
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            row, text="导入并解析", style="Primary.TButton", command=self._import_excel
        ).pack(side="left", padx=(8, 0))

        export_row = tk.Frame(top, bg=COLORS["white"])
        export_row.pack(fill="x", pady=(10, 0))
        self.export_plan_btn = ttk.Button(
            export_row,
            text="导出可编辑测试计划",
            style="Secondary.TButton",
            command=self._export_plan,
            state="disabled",
        )
        self.export_plan_btn.pack(side="left")
        self.edit_case_btn = ttk.Button(
            export_row,
            text="编辑选中需求用例",
            style="Secondary.TButton",
            command=self._edit_selected_requirement,
            state="disabled",
        )
        self.edit_case_btn.pack(side="left", padx=(8, 0))
        ttk.Label(
            export_row,
            text="双击需求行也可以编辑步骤和选择器。",
            style="CardMuted.TLabel",
        ).pack(side="left", padx=(10, 0))

        self.import_summary_var = tk.StringVar(value="尚未导入需求文件。")
        ttk.Label(
            top,
            textvariable=self.import_summary_var,
            style="Card.TLabel",
            font=(FONT, 10, "bold"),
        ).pack(anchor="w", pady=(13, 0))
        self.import_note_var = tk.StringVar(value="")
        ttk.Label(
            top,
            textvariable=self.import_note_var,
            style="CardMuted.TLabel",
            wraplength=1050,
            justify="left",
        ).pack(anchor="w", pady=(3, 0))

        tree_card = self._card(page, padx=14, pady=14)
        tree_card.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        tree_card.grid_rowconfigure(1, weight=1)
        tree_card.grid_columnconfigure(0, weight=1)
        ttk.Label(tree_card, text="需求与用例预览", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        tree_frame = tk.Frame(tree_card, bg=COLORS["white"])
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        columns = ("seq", "system", "module", "function", "point", "cases", "status", "actual", "remark")
        headings = {
            "seq": "序号",
            "system": "系统",
            "module": "模块",
            "function": "功能",
            "point": "功能点",
            "cases": "用例数",
            "status": "可执行",
            "actual": "实际应用功能 / 需求描述",
            "remark": "需规备注",
        }
        widths = {
            "seq": 95,
            "system": 130,
            "module": 85,
            "function": 90,
            "point": 105,
            "cases": 55,
            "status": 70,
            "actual": 230,
            "remark": 120,
        }
        self.import_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", selectmode="browse", height=8
        )
        for col in columns:
            self.import_tree.heading(col, text=headings[col])
            self.import_tree.column(col, width=widths[col], minwidth=45, anchor="w")
        self.import_tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.import_tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.import_tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.import_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.import_tree.tag_configure("unconfigured", background="#fffbeb")
        self.import_tree.bind("<Double-1>", lambda _event: self._edit_selected_requirement())

        self.import_empty = tk.Label(
            tree_card,
            text="导入 Excel 后，这里会显示每条需求和对应的用例数量。",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 10),
        )
        self.import_empty.grid(row=2, column=0, pady=(10, 4))

    def _build_run_page(self) -> None:
        page = self._page("run")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)

        config_card = self._card(page, padx=20, pady=18)
        config_card.grid(row=0, column=0, sticky="ew")
        ttk.Label(config_card, text="测试配置", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=8, sticky="w"
        )

        self.url_var = tk.StringVar(value="http://192.168.1.10:8080")
        self.browser_var = tk.StringVar(value="Edge")
        self.headless_var = tk.BooleanVar(value=True)
        self.timeout_var = tk.StringVar(value="15")
        self.account_var = tk.StringVar()
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.permission_var = tk.StringVar(value="未识别到测试账号，请手动填写")
        self.account_map: dict[str, Any] = {}

        ttk.Label(config_card, text="待测网址", style="Card.TLabel").grid(
            row=1, column=0, sticky="w", pady=(14, 0)
        )
        url_entry = ttk.Entry(config_card, textvariable=self.url_var, font=(FONT, 10))
        url_entry.grid(row=1, column=1, columnspan=5, sticky="ew", padx=(10, 12), pady=(14, 0), ipady=4)
        config_card.grid_columnconfigure(1, weight=1)

        ttk.Label(config_card, text="浏览器", style="Card.TLabel").grid(
            row=1, column=6, sticky="w", pady=(14, 0)
        )
        ttk.Combobox(
            config_card,
            textvariable=self.browser_var,
            values=["Edge", "Chrome"],
            state="readonly",
            width=8,
        ).grid(row=1, column=7, sticky="w", padx=(8, 0), pady=(14, 0))

        options = tk.Frame(config_card, bg=COLORS["white"])
        options.grid(row=2, column=0, columnspan=8, sticky="ew", pady=(12, 0))
        ttk.Checkbutton(
            options,
            text="无头模式（后台运行，不显示浏览器窗口）",
            variable=self.headless_var,
        ).pack(side="left")
        ttk.Label(options, text="单步超时", style="Card.TLabel").pack(
            side="left", padx=(22, 6)
        )
        ttk.Spinbox(
            options,
            from_=5,
            to=120,
            textvariable=self.timeout_var,
            width=6,
        ).pack(side="left")
        ttk.Label(options, text="秒", style="Card.TLabel").pack(side="left", padx=(4, 0))

        account_row = tk.Frame(config_card, bg=COLORS["white"])
        account_row.grid(row=3, column=0, columnspan=8, sticky="ew", pady=(10, 0))
        ttk.Label(account_row, text="测试账号", style="Card.TLabel").pack(side="left")
        self.account_combo = ttk.Combobox(
            account_row,
            textvariable=self.account_var,
            values=[],
            state="readonly",
            width=28,
        )
        self.account_combo.pack(side="left", padx=(8, 14))
        self.account_combo.bind("<<ComboboxSelected>>", self._on_account_change)
        ttk.Label(account_row, text="用户名", style="Card.TLabel").pack(side="left")
        ttk.Entry(account_row, textvariable=self.username_var, width=14).pack(
            side="left", padx=(6, 12)
        )
        ttk.Label(account_row, text="密码", style="Card.TLabel").pack(side="left")
        ttk.Entry(account_row, textvariable=self.password_var, width=14, show="*").pack(
            side="left", padx=(6, 12)
        )
        ttk.Label(
            account_row,
            textvariable=self.permission_var,
            style="CardMuted.TLabel",
        ).pack(side="left", fill="x", expand=True)

        actions = tk.Frame(config_card, bg=COLORS["white"])
        actions.grid(row=4, column=0, columnspan=8, sticky="ew", pady=(14, 0))
        ttk.Button(
            actions, text="检查网页连通性", style="Secondary.TButton", command=self._check_connectivity
        ).pack(side="left")
        self.start_btn = ttk.Button(
            actions, text="开始测试", style="Primary.TButton", command=self._start_test
        )
        self.start_btn.pack(side="left", padx=10)
        self.stop_btn = ttk.Button(
            actions, text="停止", style="Danger.TButton", command=self._stop_test, state="disabled"
        )
        self.stop_btn.pack(side="left")
        self.open_report_btn = ttk.Button(
            actions,
            text="打开上次报告",
            style="Secondary.TButton",
            command=self._open_last_report,
            state="disabled",
        )
        self.open_report_btn.pack(side="right")

        summary = tk.Frame(page, bg=COLORS["bg"])
        summary.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        for index in range(4):
            summary.grid_columnconfigure(index, weight=1)
        self.summary_vars: dict[str, tk.StringVar] = {}
        summary_items = [
            ("pass", "通过", COLORS["success"]),
            ("fail", "不通过", COLORS["danger"]),
            ("not_configured", "未配置", COLORS["warning"]),
            ("skipped", "跳过 / 错误", COLORS["muted"]),
        ]
        for index, (key, label, color) in enumerate(summary_items):
            card = self._card(summary, padx=16, pady=10)
            card.grid(
                row=0,
                column=index,
                sticky="ew",
                padx=(0 if index == 0 else 6, 0 if index == 3 else 6),
            )
            var = tk.StringVar(value="0")
            self.summary_vars[key] = var
            tk.Label(
                card,
                textvariable=var,
                bg=COLORS["white"],
                fg=color,
                font=(FONT, 20, "bold"),
                anchor="w",
            ).pack(anchor="w")
            ttk.Label(card, text=label, style="CardMuted.TLabel").pack(anchor="w")

        log_card = self._card(page, padx=14, pady=12)
        log_card.grid(row=2, column=0, sticky="nsew", pady=(14, 0))
        log_card.grid_rowconfigure(2, weight=1)
        log_card.grid_columnconfigure(0, weight=1)
        ttk.Label(log_card, text="执行日志", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.progress_var = tk.StringVar(value="0 / 0")
        ttk.Label(log_card, textvariable=self.progress_var, style="CardMuted.TLabel").grid(
            row=0, column=1, sticky="e"
        )
        log_card.grid_columnconfigure(1, weight=0)
        self.progress = ttk.Progressbar(log_card, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 8))
        self.log_text = tk.Text(
            log_card,
            bg="#0f172a",
            fg="#dbeafe",
            insertbackground="white",
            relief="flat",
            font=("Consolas", 9),
            height=13,
            wrap="word",
            padx=12,
            pady=10,
        )
        self.log_text.grid(row=2, column=0, columnspan=2, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_card, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=2, column=2, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set, state="disabled")

    def _build_reports_page(self) -> None:
        page = self._page("reports")
        page.grid_rowconfigure(1, weight=1)
        page.grid_columnconfigure(0, weight=1)
        top = self._card(page, padx=18, pady=14)
        top.grid(row=0, column=0, sticky="ew")
        left = tk.Frame(top, bg=COLORS["white"])
        left.pack(side="left", fill="x", expand=True)
        ttk.Label(left, text="历史测试报告", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            left,
            text="双击一行可直接打开 HTML 报告；Excel 报告和失败截图在同一个运行目录。",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(3, 0))
        buttons = tk.Frame(top, bg=COLORS["white"])
        buttons.pack(side="right")
        ttk.Button(
            buttons, text="刷新", style="Secondary.TButton", command=self._refresh_reports
        ).pack(side="left")
        ttk.Button(
            buttons, text="打开报告", style="Primary.TButton", command=self._open_selected_report
        ).pack(side="left", padx=8)
        ttk.Button(
            buttons, text="打开目录", style="Secondary.TButton", command=self._open_selected_run_dir
        ).pack(side="left")

        tree_card = self._card(page, padx=12, pady=12)
        tree_card.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        tree_card.grid_rowconfigure(0, weight=1)
        tree_card.grid_columnconfigure(0, weight=1)
        columns = ("time", "project", "total", "pass", "fail", "unconfigured", "url")
        headings = {
            "time": "运行时间",
            "project": "项目",
            "total": "总数",
            "pass": "通过",
            "fail": "不通过",
            "unconfigured": "未配置",
            "url": "待测网址",
        }
        widths = {
            "time": 150,
            "project": 260,
            "total": 70,
            "pass": 65,
            "fail": 70,
            "unconfigured": 80,
            "url": 380,
        }
        self.report_tree = ttk.Treeview(
            tree_card, columns=columns, show="headings", selectmode="browse"
        )
        for col in columns:
            self.report_tree.heading(col, text=headings[col])
            self.report_tree.column(col, width=widths[col], minwidth=50, anchor="w")
        self.report_tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(tree_card, orient="vertical", command=self.report_tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.report_tree.configure(yscrollcommand=yscroll.set)
        self.report_tree.bind("<Double-1>", lambda _event: self._open_selected_report())
        self.report_empty = tk.Label(
            tree_card,
            text="还没有测试报告。执行一次测试后，报告会显示在这里。",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=(FONT, 10),
        )
        self.report_empty.grid(row=1, column=0, pady=(10, 4))

    def _build_settings_page(self) -> None:
        page = self._page("settings")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)

        storage_card = self._card(page, padx=22, pady=20)
        storage_card.grid(row=0, column=0, sticky="ew")
        ttk.Label(storage_card, text="数据存储目录", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            storage_card,
            text="测试日志、截图、下载、缓存、临时文件和报告都会写入这里。默认选择非 C 盘。",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(3, 12))
        row = tk.Frame(storage_card, bg=COLORS["white"])
        row.pack(fill="x")
        self.storage_var = tk.StringVar(value=str(self.storage.root))
        ttk.Entry(row, textvariable=self.storage_var, font=(FONT, 9)).pack(
            side="left", fill="x", expand=True, ipady=4
        )
        ttk.Button(
            row, text="选择目录", style="Secondary.TButton", command=self._choose_storage
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            row, text="应用并保存", style="Primary.TButton", command=self._save_settings
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            row,
            text="打开目录",
            style="Secondary.TButton",
            command=lambda: self.storage.open_in_explorer(Path(self.storage_var.get())),
        ).pack(side="left", padx=(8, 0))
        self.storage_warning = ttk.Label(
            storage_card, text="", style="CardMuted.TLabel", foreground=COLORS["warning"]
        )
        self.storage_warning.pack(anchor="w", pady=(8, 0))

        defaults = self._card(page, padx=22, pady=20)
        defaults.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        ttk.Label(defaults, text="默认参数", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            defaults,
            text="这些参数会保存到当前数据目录的 config.json。",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(3, 12))

        form = tk.Frame(defaults, bg=COLORS["white"])
        form.pack(fill="x")
        ttk.Label(form, text="默认待测网址", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", pady=6
        )
        ttk.Entry(form, textvariable=self.url_var, font=(FONT, 10)).grid(
            row=0, column=1, sticky="ew", padx=(14, 0), pady=6, ipady=4
        )
        form.grid_columnconfigure(1, weight=1)

        ttk.Label(form, text="默认浏览器", style="Card.TLabel").grid(
            row=1, column=0, sticky="w", pady=6
        )
        ttk.Combobox(
            form,
            textvariable=self.browser_var,
            values=["Edge", "Chrome"],
            state="readonly",
            width=10,
        ).grid(row=1, column=1, sticky="w", padx=(14, 0), pady=6)

        ttk.Label(form, text="单步超时（秒）", style="Card.TLabel").grid(
            row=2, column=0, sticky="w", pady=6
        )
        ttk.Spinbox(form, from_=5, to=120, textvariable=self.timeout_var, width=8).grid(
            row=2, column=1, sticky="w", padx=(14, 0), pady=6
        )

        ttk.Checkbutton(
            form,
            text="无头模式（后台运行）",
            variable=self.headless_var,
        ).grid(row=3, column=1, sticky="w", padx=(14, 0), pady=6)

        buttons = tk.Frame(defaults, bg=COLORS["white"])
        buttons.pack(fill="x", pady=(16, 0))
        ttk.Button(
            buttons,
            text="创建标准用例模板",
            style="Secondary.TButton",
            command=self._create_template,
        ).pack(side="left")
        ttk.Button(
            buttons,
            text="创建演示测试计划",
            style="Secondary.TButton",
            command=self._create_demo_plan,
        ).pack(side="left", padx=10)
        ttk.Button(
            buttons, text="保存设置", style="Primary.TButton", command=self._save_settings
        ).pack(side="left")

        about = tk.Frame(defaults, bg="#f8fafc", padx=14, pady=12)
        about.pack(fill="x", pady=(18, 0))
        tk.Label(
            about,
            text=(
                "关于：本程序是桌面自动化测试平台，网页操作依赖本机 Microsoft Edge 或 Google Chrome。"
                "测试 ERP 前需要先能访问目标内网地址，并准备可登录的测试账号。"
                "未配置步骤的需求不会被算作通过。"
            ),
            bg="#f8fafc",
            fg=COLORS["muted"],
            font=(FONT, 9),
            justify="left",
            wraplength=1000,
            anchor="w",
        ).pack(fill="x")
        self._update_storage_warning()

    # ------------------------------------------------------------------
    # 页面切换和配置
    # ------------------------------------------------------------------
    def show_page(self, key: str) -> None:
        page = self.pages.get(key)
        if not page:
            return
        page.tkraise()
        for name, button in self.nav_buttons.items():
            if name == key:
                button.configure(
                    bg=COLORS["sidebar_active"],
                    fg="white",
                    font=(FONT, 10, "bold"),
                )
            else:
                button.configure(
                    bg=COLORS["sidebar"],
                    fg=COLORS["sidebar_text"],
                    font=(FONT, 10),
                )
        title, subtitle = self.page_meta.get(key, ("", ""))
        self.page_title.configure(text=title)
        self.page_subtitle.configure(text=subtitle)
        if key == "reports":
            self._refresh_reports()
        if key == "dashboard":
            self._refresh_dashboard()
        if key == "wizard":
            self._refresh_wizard()
        if key == "settings":
            self.storage_var.set(str(self.storage.root))
            self._update_storage_warning()

    def _load_config(self) -> None:
        config = self.storage.config
        self.url_var.set(str(config.get("base_url", "http://192.168.1.10:8080")))
        self.browser_var.set(
            "Chrome" if config.get("browser", "msedge") == "chrome" else "Edge"
        )
        self.headless_var.set(bool(config.get("headless", True)))
        try:
            self.timeout_var.set(str(max(5, int(config.get("timeout_ms", 15000)) // 1000)))
        except (TypeError, ValueError):
            self.timeout_var.set("15")
        source = str(config.get("source_excel", ""))
        if source:
            self.import_path_var.set(source)
        self.sidebar_storage.configure(text=str(self.storage.root))
        self._refresh_dashboard()

    def _apply_config_from_ui(self) -> None:
        try:
            timeout_s = max(5, min(120, int(float(self.timeout_var.get()))))
        except ValueError:
            timeout_s = 15
        self.storage.config.update(
            {
                "base_url": self.url_var.get().strip(),
                "browser": "chrome" if self.browser_var.get() == "Chrome" else "msedge",
                "headless": bool(self.headless_var.get()),
                "timeout_ms": timeout_s * 1000,
            }
        )
        self.storage.save_config()

    def _browser_channel(self) -> str:
        return "chrome" if self.browser_var.get() == "Chrome" else "msedge"

    def _timeout_ms(self) -> int:
        try:
            return max(5, min(120, int(float(self.timeout_var.get())))) * 1000
        except ValueError:
            return 15000

    def _refresh_accounts(self, project: Project) -> None:
        self.account_map = {account.label(): account for account in project.accounts}
        labels = list(self.account_map.keys())
        self.account_combo.configure(values=labels)
        if hasattr(self, "wizard_account_combo"):
            self.wizard_account_combo.configure(values=labels)
        if labels:
            self.account_var.set(labels[0])
            self._on_account_change()
        else:
            self.account_var.set("")
            self.username_var.set("")
            self.password_var.set("")
            self.permission_var.set("未识别到测试账号，请手动填写")

    def _on_account_change(self, _event=None) -> None:
        account = self.account_map.get(self.account_var.get())
        if account is None:
            return
        self.username_var.set(account.username)
        self.password_var.set(account.password)
        permissions = "、".join(account.permissions)
        if len(permissions) > 120:
            permissions = permissions[:120] + "…"
        self.permission_var.set(
            f"权限：{permissions}" if permissions else "权限：需规未提供权限说明"
        )

    # ------------------------------------------------------------------
    # 导入
    # ------------------------------------------------------------------
    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="选择需求 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.import_path_var.set(path)
            self._import_excel()

    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择需规文件夹")
        if not folder:
            return
        self.configure(cursor="watch")
        self.status_var.set("正在扫描文件夹，自动识别需规...")
        self.update_idletasks()
        try:
            project, scan = import_folder(folder)
            self._finish_import(project, project.source_excel)
            top = "\n".join(
                f"{index}. {candidate.describe()}"
                for index, candidate in enumerate(scan.top(6), start=1)
            )
            messagebox.showinfo(
                "文件夹扫描完成",
                f"识别到 {len(scan.candidates)} 个候选，已自动选择最佳项：\n"
                f"{scan.best.describe()}\n\n候选列表：\n{top}\n\n"
                f"项目：{project.name}\n"
                f"需求：{len(project.requirements)} 条，用例：{project.case_count} 个，"
                f"已配置：{project.configured_case_count} 个。",
            )
        except Exception as exc:
            messagebox.showerror("文件夹导入失败", str(exc))
            self.status_var.set("文件夹导入失败")
        finally:
            self.configure(cursor="")
            self.update_idletasks()

    def _import_excel(self) -> None:
        path = self.import_path_var.get().strip()
        if not path:
            messagebox.showwarning("请选择文件", "请先选择需求 Excel 文件。")
            return
        self.configure(cursor="watch")
        self.status_var.set("正在解析 Excel...")
        self.update_idletasks()
        try:
            project = import_workbook(path)
            self._finish_import(project, path)
            if project.configured_case_count == 0:
                messagebox.showinfo(
                    "已导入",
                    "需规已识别，但没有可直接执行的页面步骤，报告会把它们标记为“未配置”。\n\n"
                    "可以双击需求行补充步骤，或导出可编辑测试计划后填写选择器。",
                )
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            self.status_var.set("导入失败")
        finally:
            self.configure(cursor="")
            self.update_idletasks()

    def _finish_import(self, project: Project, source_path: str) -> None:
        self.project = project
        self.storage.save_project(project)
        self.storage.config["source_excel"] = source_path
        self.storage.save_config()
        self.import_path_var.set(source_path)
        self._populate_requirements(project)
        self._refresh_accounts(project)
        self._refresh_dashboard()
        self._refresh_wizard()
        self.start_btn.configure(state="normal" if project.case_count else "disabled")
        self.export_plan_btn.configure(state="normal" if project.case_count else "disabled")
        self.edit_case_btn.configure(state="normal" if project.case_count else "disabled")
        self.status_var.set(
            f"导入完成：{len(project.requirements)} 条需求，{project.case_count} 个用例，"
            f"{project.configured_case_count} 个已配置"
        )

    def _populate_requirements(self, project: Project) -> None:
        for item in self.import_tree.get_children():
            self.import_tree.delete(item)
        max_preview = 1500
        for req in project.requirements[:max_preview]:
            configured = sum(1 for case in req.cases if case.configured)
            tags = ("unconfigured",) if configured == 0 else ()
            self.import_tree.insert(
                "",
                "end",
                iid=req.req_id,
                values=(
                    req.seq,
                    req.system,
                    req.module,
                    req.function,
                    req.point,
                    f"{configured}/{len(req.cases)}",
                    "可执行" if configured else "待配置",
                    req.actual_features[:160],
                    req.remark[:120],
                ),
                tags=tags,
            )
        if len(project.requirements) > max_preview:
            self.import_tree.insert(
                "",
                "end",
                iid="__preview_more__",
                values=(
                    "…",
                    f"共 {len(project.requirements)} 条",
                    "",
                    "",
                    f"界面仅预览前 {max_preview} 条，项目已全部导入",
                    "",
                    "",
                    "",
                    "",
                ),
                tags=("unconfigured",),
            )
        total = project.case_count
        configured = project.configured_case_count
        self.import_summary_var.set(
            f"{project.name}　·　{len(project.requirements)} 条需求　·　"
            f"{total} 个用例　·　{configured} 个已配置步骤"
        )
        note = project.notes
        if len(project.requirements) > max_preview:
            note += f"\n界面仅预览前 {max_preview} 条需求，项目和报告仍包含全部数据。"
        self.import_note_var.set(note)
        self.import_empty.configure(
            text="" if project.requirements else "没有解析到需求。"
        )

    def _export_plan(self) -> None:
        if not self.project:
            messagebox.showwarning("没有项目", "请先导入需求 Excel。")
            return
        safe_name = "".join(
            "_" if char in '\\/:*?"<>|' else char for char in (self.project.name or "项目")
        )
        path = filedialog.asksaveasfilename(
            title="导出可编辑测试计划",
            defaultextension=".xlsx",
            initialfile=f"{safe_name}_可编辑测试计划.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            export_project_plan(self.project, path)
            self.status_var.set(f"测试计划已导出：{path}")
            messagebox.showinfo(
                "导出成功",
                "已导出可编辑测试计划。\n\n"
                "未配置的用例会保留一行空白操作，填入步骤、选择器和预期结果后，"
                "重新导入即可执行。",
            )
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _edit_selected_requirement(self) -> None:
        if not self.project:
            messagebox.showwarning("没有项目", "请先导入需求 Excel。")
            return
        selection = self.import_tree.selection()
        if not selection:
            messagebox.showinfo("未选择需求", "请先在下方列表选择一条需求。")
            return
        req_id = selection[0]
        requirement = next(
            (item for item in self.project.requirements if item.req_id == req_id),
            None,
        )
        if requirement is None:
            messagebox.showerror("需求不存在", "找不到选中的需求，请重新导入。")
            return
        CaseEditorDialog(self, requirement, self._on_case_editor_saved)

    def _on_case_editor_saved(self) -> None:
        if not self.project:
            return
        self.storage.save_project(self.project)
        self._populate_requirements(self.project)
        self.start_btn.configure(state="normal" if self.project.case_count else "disabled")
        self.export_plan_btn.configure(state="normal" if self.project.case_count else "disabled")
        self._refresh_dashboard()
        self.status_var.set("用例步骤已保存到当前项目")

    # ------------------------------------------------------------------
    # 测试执行
    # ------------------------------------------------------------------
    def _append_log(self, line: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _check_connectivity(self) -> None:
        url = self.url_var.get().strip()
        if not url:
            messagebox.showwarning("网址为空", "请填写待测网址。")
            return
        self.status_var.set("正在检查网页连通性...")
        self._append_log(f"检查连通性：{url}")

        def worker() -> None:
            result = check_web(
                url,
                browser=self._browser_channel(),
                headless=bool(self.headless_var.get()),
                timeout_ms=self._timeout_ms(),
            )
            self.events.put(("connectivity", result, url))

        threading.Thread(target=worker, daemon=True).start()

    def _start_test(self) -> None:
        if not self.project:
            messagebox.showwarning("没有项目", "请先在“需求导入”中导入 Excel。")
            self.show_page("import")
            return
        url = self.url_var.get().strip()
        if not url.startswith(("http://", "https://", "file://")):
            messagebox.showwarning(
                "网址格式不正确",
                "待测地址必须以 http://、https:// 或 file:// 开头。",
            )
            return
        unconfigured = self.project.case_count - self.project.configured_case_count
        if unconfigured:
            if not messagebox.askyesno(
                "存在未配置用例",
                f"当前有 {unconfigured} 个用例没有配置页面步骤，报告中会标记为“未配置”，"
                "它们不会被算作通过。\n\n是否继续执行？",
            ):
                return
        self._apply_config_from_ui()
        self.project.reset_results()
        self.live_counts = {"pass": 0, "fail": 0, "not_configured": 0, "skipped": 0}
        for key, var in self.summary_vars.items():
            var.set("0")
        self.progress["value"] = 0
        self.progress_var.set(f"0 / {self.project.case_count}")
        self._clear_log()
        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.open_report_btn.configure(state="disabled")
        self.status_var.set("测试执行中...")

        options = RunOptions(
            base_url=url,
            browser=self._browser_channel(),
            headless=bool(self.headless_var.get()),
            timeout_ms=self._timeout_ms(),
            username=self.username_var.get().strip(),
            password=self.password_var.get(),
            stop_event=self.stop_event,
            on_log=lambda line: self.events.put(("log", line)),
            on_case=lambda case: self.events.put(("case", case)),
            on_progress=lambda done, total: self.events.put(("progress", done, total)),
        )
        self._append_log("启动测试线程...")

        def worker() -> None:
            try:
                runner = WebTestRunner(self.storage, options)
                result = runner.run(self.project)
                self.events.put(("finished", result, self.project))
            except Exception as exc:
                self.events.put(("run_error", str(exc)))

        self.run_thread = threading.Thread(target=worker, daemon=True)
        self.run_thread.start()

    def _stop_test(self) -> None:
        if self.run_thread and self.run_thread.is_alive():
            self.stop_event.set()
            self._append_log("已请求停止，当前步骤完成后停止。")
            self.status_var.set("正在停止...")
            self.stop_btn.configure(state="disabled")

    def _open_last_report(self) -> None:
        if self.last_report and Path(self.last_report).exists():
            os.startfile(str(self.last_report))  # type: ignore[attr-defined]

    # ------------------------------------------------------------------
    # 报告
    # ------------------------------------------------------------------
    def _refresh_reports(self) -> None:
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)
        self.report_map.clear()
        runs = []
        if self.storage.reports_dir.exists():
            for directory in self.storage.reports_dir.iterdir():
                run_json = directory / "run.json"
                if directory.is_dir() and run_json.exists():
                    try:
                        import json

                        data = json.loads(run_json.read_text(encoding="utf-8"))
                        runs.append((directory.stat().st_mtime, directory, data))
                    except Exception:
                        continue
        runs.sort(key=lambda item: item[0], reverse=True)
        for _mtime, directory, data in runs:
            meta = data.get("metadata", {})
            counts = meta.get("counts", {})
            self.report_tree.insert(
                "",
                "end",
                iid=directory.name,
                values=(
                    meta.get("started_at", ""),
                    meta.get("project_name", ""),
                    meta.get("total", ""),
                    counts.get("pass", 0),
                    counts.get("fail", 0),
                    counts.get("not_configured", 0),
                    meta.get("base_url", ""),
                ),
            )
            self.report_map[directory.name] = directory
        self.report_empty.configure(
            text="" if runs else "还没有测试报告。执行一次测试后，报告会显示在这里。"
        )

    def _selected_run_dir(self) -> Path | None:
        selection = self.report_tree.selection()
        if not selection:
            return None
        return self.report_map.get(selection[0])

    def _open_selected_report(self) -> None:
        run_dir = self._selected_run_dir()
        if not run_dir:
            messagebox.showinfo("未选择", "请先选择一条测试报告。")
            return
        report = run_dir / "report.html"
        if report.exists():
            os.startfile(str(report))  # type: ignore[attr-defined]
        else:
            messagebox.showerror("报告不存在", f"找不到报告文件：\n{report}")

    def _open_selected_run_dir(self) -> None:
        run_dir = self._selected_run_dir()
        if run_dir:
            self.storage.open_in_explorer(run_dir)

    # ------------------------------------------------------------------
    # 设置
    # ------------------------------------------------------------------
    def _choose_storage(self) -> None:
        path = filedialog.askdirectory(
            title="选择数据存储目录",
            initialdir=str(self.storage.root),
        )
        if path:
            self.storage_var.set(path)
            self._update_storage_warning()

    def _update_storage_warning(self) -> None:
        path = Path(self.storage_var.get().strip() or ".")
        if is_on_c_drive(path):
            self.storage_warning.configure(
                text="警告：当前目录在 C 盘，测试日志和缓存可能占用系统盘空间。"
            )
        else:
            self.storage_warning.configure(text="")

    def _save_settings(self) -> None:
        new_root = Path(self.storage_var.get().strip())
        if new_root != self.storage.root:
            if is_on_c_drive(new_root):
                if not messagebox.askyesno(
                    "确认使用 C 盘？",
                    "你选择的目录在 C 盘。测试日志、缓存和截图会占用系统盘空间，是否继续？",
                ):
                    return
            if not messagebox.askyesno(
                "切换数据目录",
                "已有测试数据不会自动迁移到新目录。确定切换吗？",
            ):
                return
            try:
                self.storage.set_root(new_root)
            except Exception as exc:
                messagebox.showerror("切换失败", str(exc))
                return
            self.sidebar_storage.configure(text=str(self.storage.root))
            self._refresh_reports()
        self._apply_config_from_ui()
        self.storage_var.set(str(self.storage.root))
        self._update_storage_warning()
        self.status_var.set(f"设置已保存：{self.storage.root}")
        messagebox.showinfo("保存成功", "设置已保存。")

    def _create_template(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存标准用例模板",
            defaultextension=".xlsx",
            initialfile="自动化测试用例模板.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            create_standard_template(path)
            messagebox.showinfo("模板已生成", f"已生成标准用例模板：\n{path}")
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc))

    def _create_demo_plan(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存演示测试计划",
            defaultextension=".xlsx",
            initialfile="志愿系统演示测试计划.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            create_demo_plan(path, self.url_var.get().strip() or "http://127.0.0.1:5000")
            messagebox.showinfo(
                "演示计划已生成",
                f"已生成演示计划：\n{path}\n\n可在“需求导入”中导入并执行。",
            )
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc))

    # ------------------------------------------------------------------
    # 事件循环
    # ------------------------------------------------------------------
    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "log":
                    self._append_log(str(event[1]))
                elif kind == "case":
                    case = event[1]
                    if case.status in self.live_counts:
                        self.live_counts[case.status] += 1
                        self.summary_vars[case.status].set(str(self.live_counts[case.status]))
                elif kind == "progress":
                    done, total = int(event[1]), int(event[2])
                    self.progress["value"] = (done / total * 100) if total else 0
                    self.progress_var.set(f"{done} / {total}")
                elif kind == "finished":
                    self._on_run_finished(event[1], event[2])
                elif kind == "run_error":
                    self._on_run_error(str(event[1]))
                elif kind == "connectivity":
                    self._on_connectivity(event[1], str(event[2]))
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _on_run_finished(self, result: RunResult, project: Project) -> None:
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.last_report = result.report_html
        self.last_run_dir = result.run_dir
        if result.report_html:
            self.open_report_btn.configure(state="normal")
        counts = result.counts or {}
        for key in ("pass", "fail", "not_configured"):
            self.summary_vars[key].set(str(counts.get(key, 0)))
        skipped = int(counts.get("skipped", 0)) + int(counts.get("error", 0))
        self.summary_vars["skipped"].set(str(skipped))
        self.progress["value"] = 100
        self.progress_var.set(f"{project.case_count} / {project.case_count}")
        if result.error:
            self.status_var.set(f"测试完成，但有执行错误：{result.error}")
            messagebox.showwarning("测试完成", f"测试完成，但存在执行错误：\n{result.error}")
        else:
            self.status_var.set(
                f"测试完成：通过 {counts.get('pass', 0)}，"
                f"不通过 {counts.get('fail', 0)}，未配置 {counts.get('not_configured', 0)}"
            )
        self._refresh_reports()
        self._refresh_dashboard()

    def _on_run_error(self, error: str) -> None:
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.status_var.set("测试执行异常")
        messagebox.showerror("测试执行异常", error)

    def _on_connectivity(self, result: dict[str, Any], url: str) -> None:
        if result.get("ok"):
            text = (
                f"网页可访问。\n\n地址：{result.get('url')}\n"
                f"标题：{result.get('title')}\n状态码：{result.get('status')}"
            )
            self.status_var.set("网页连通性正常")
            messagebox.showinfo("连通性检查", text)
            self._append_log(f"连通性正常：{result.get('title')}（{result.get('url')}）")
        else:
            error = str(result.get("error", "未知错误"))
            self.status_var.set("网页无法访问")
            messagebox.showerror(
                "无法访问",
                f"打不开：{url}\n\n{error}\n\n"
                "请检查网址、内网/VPN、服务器是否启动，以及是否需要登录。",
            )
            self._append_log(f"连通性失败：{error}")

    def _refresh_dashboard(self) -> None:
        self.dash_cards["storage"].set(self._short_path(self.storage.root))
        if self.project:
            self.dash_cards["requirements"].set(str(len(self.project.requirements)))
            self.dash_cards["cases"].set(
                f"{self.project.configured_case_count}/{self.project.case_count} 可执行"
            )
        if self.last_report and Path(self.last_report).exists():
            self.dash_cards["report"].set("已生成")
        elif self.storage.reports_dir.exists() and any(self.storage.reports_dir.iterdir()):
            self.dash_cards["report"].set("有历史报告")

    @staticmethod
    def _short_path(path: Path, limit: int = 34) -> str:
        text = str(path)
        return text if len(text) <= limit else "..." + text[-(limit - 3):]

    def _on_close(self) -> None:
        if self.run_thread and self.run_thread.is_alive():
            if not messagebox.askyesno(
                "测试仍在运行", "测试仍在运行，确定停止并退出吗？"
            ):
                return
            self.stop_event.set()
        self.destroy()
