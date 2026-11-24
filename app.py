"""桌面自动化测试校验系统入口。"""

from __future__ import annotations

import ctypes
import json
import os
import sys
import traceback
from pathlib import Path

APP_VERSION = "0.1.0"


def enable_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # type: ignore[attr-defined]
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # type: ignore[attr-defined]
        except Exception:
            pass


def run_self_test() -> int:
    """打包后自检：存储、Excel、报告、Playwright/Edge。"""
    from autotest.excel_import import create_demo_plan, create_standard_template, import_workbook
    from autotest.storage import AppStorage

    result: dict[str, object] = {"version": APP_VERSION, "ok": False, "checks": []}
    try:
        storage = AppStorage()
        checks = result["checks"]
        assert isinstance(checks, list)

        checks.append({"name": "storage_root", "value": str(storage.root), "ok": True})
        assert not str(storage.root).upper().startswith("C:"), "默认数据目录在 C 盘"

        template = storage.temp_dir / "self_test_template.xlsx"
        create_standard_template(template)
        checks.append({"name": "excel_template", "value": str(template), "ok": template.exists()})

        demo = storage.temp_dir / "self_test_demo.xlsx"
        create_demo_plan(demo, "http://127.0.0.1:5000")
        project = import_workbook(demo)
        checks.append(
            {
                "name": "excel_import",
                "value": f"{len(project.requirements)} requirements, {project.case_count} cases, "
                f"{len(project.setup_steps)} setup steps",
                "ok": project.case_count == 3 and len(project.setup_steps) == 5,
            }
        )

        html = storage.temp_dir / "self_test.html"
        html.write_text("<html><head><title>selftest</title></head><body><h1>ok</h1></body></html>", encoding="utf-8")
        from autotest.models import Project, Requirement, Step, TestCase
        from autotest.runner import RunOptions, WebTestRunner

        project = Project.create("self_test", str(html))
        requirement = Requirement(req_id="ST-001", seq="1", system="self-test", module="engine", function="web", point="local page")
        requirement.cases.append(
            TestCase(
                case_id="ST-001-C01",
                requirement_id="ST-001",
                name="local page assertion",
                expected="h1 text is ok",
                steps=[
                    Step(action="goto", target=html.as_uri()),
                    Step(action="assert_text", target="h1", value="ok"),
                ],
            )
        )
        project.requirements.append(requirement)
        run_result = WebTestRunner(
            storage,
            RunOptions(base_url=html.as_uri(), browser="msedge", headless=True),
        ).run(project)
        checks.append(
            {
                "name": "web_runner",
                "value": f"counts={run_result.counts}, report={bool(run_result.report_html)}",
                "ok": run_result.counts.get("pass") == 1 and bool(run_result.report_html),
            }
        )
        import shutil

        shutil.rmtree(run_result.run_dir, ignore_errors=True)

        result["ok"] = all(bool(item.get("ok")) for item in checks if isinstance(item, dict))
        result["data_root"] = str(storage.root)
    except Exception as exc:
        result["error"] = str(exc)
        result["traceback"] = traceback.format_exc()
    finally:
        try:
            from autotest.storage import AppStorage

            storage = AppStorage()
            log_path = storage.logs_dir / "self_test.json"
            log_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            if sys.stdout is not None:
                print(json.dumps(result, ensure_ascii=False, indent=2))
        except Exception:
            pass
    return 0 if result.get("ok") else 1


def main() -> int:
    enable_dpi_awareness()
    if "--self-test" in sys.argv:
        return run_self_test()

    try:
        from autotest.ui import AutoTestApp

        app = AutoTestApp()
        app.mainloop()
        return 0
    except Exception as exc:
        try:
            from autotest.storage import AppStorage

            storage = AppStorage()
            (storage.logs_dir / "crash.log").write_text(
                f"{exc}\n\n{traceback.format_exc()}", encoding="utf-8"
            )
        except Exception:
            pass
        try:
            from tkinter import messagebox

            messagebox.showerror(
                "启动失败",
                f"程序启动失败：\n{exc}\n\n详细日志已写入数据目录 logs/crash.log。",
            )
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    import multiprocessing

    multiprocessing.freeze_support()
    raise SystemExit(main())
