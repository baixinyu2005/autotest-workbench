"""Playwright 网页测试执行引擎。"""

from __future__ import annotations

import json
import re
import threading
import time
import traceback
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable
from urllib.parse import urljoin, urlparse

from .models import Project, Step, TestCase
from .storage import AppStorage


class RunnerError(Exception):
    pass


@dataclass
class RunOptions:
    base_url: str
    browser: str = "msedge"
    headless: bool = True
    timeout_ms: int = 15000
    username: str = ""
    password: str = ""
    stop_event: threading.Event | None = None
    on_log: Callable[[str], None] | None = None
    on_case: Callable[[TestCase], None] | None = None
    on_progress: Callable[[int, int], None] | None = None
    open_report: bool = False


@dataclass
class RunResult:
    run_dir: Path
    report_html: Path | None = None
    report_xlsx: Path | None = None
    counts: dict[str, int] = field(default_factory=dict)
    error: str = ""
    started_at: str = ""
    finished_at: str = ""


def _short(text: str, limit: int = 300) -> str:
    text = " ".join(str(text or "").split())
    return text if len(text) <= limit else text[:limit] + "..."


def _display_url(url: str) -> str:
    """让 file:// 本地应用的证据更短，同时保留页面路由。"""
    if "index.html" in url:
        return "index.html" + url.split("index.html", 1)[1]
    return url


def _join_url(base_url: str, target: str) -> str:
    target = (target or "").strip()
    if not target:
        return base_url
    if target.startswith(("http://", "https://", "file://", "data:")):
        return target
    return urljoin(base_url.rstrip("/") + "/", target.lstrip("/"))


class WebTestRunner:
    def __init__(self, storage: AppStorage, options: RunOptions):
        self.storage = storage
        self.options = options
        self.log_lines: list[str] = []
        self.run_dir: Path | None = None
        self.logger = None

    # ------------------------------------------------------------------
    def log(self, message: str) -> None:
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
        self.log_lines.append(line)
        if self.logger:
            self.logger.info(message)
        if self.options.on_log:
            self.options.on_log(line)

    def _setup_logging(self, run_dir: Path) -> None:
        import logging

        self.logger = logging.getLogger(f"autotest.run.{run_dir.name}")
        self.logger.setLevel(logging.INFO)
        self.logger.handlers.clear()
        handler = logging.FileHandler(run_dir / "logs" / "run.log", encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        self.logger.addHandler(handler)

    # ------------------------------------------------------------------
    def run(self, project: Project) -> RunResult:
        from playwright.sync_api import sync_playwright

        if not self.options.base_url.strip():
            raise RunnerError("请先填写待测系统网址。")

        run_dir = self.storage.new_run_dir()
        self.run_dir = run_dir
        self._setup_logging(run_dir)
        project.reset_results()

        started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        total = project.case_count
        done = 0

        self.log("开始测试")
        self.log(f"项目：{project.name}")
        self.log(f"网址：{self.options.base_url}")
        self.log(f"浏览器：{self.options.browser}，无头模式：{self.options.headless}")
        self.log(f"测试数据目录：{run_dir}")

        # 没有任何可执行步骤时，不启动浏览器，直接生成“未配置”报告。
        if not any(case.steps for case in project.iter_cases()):
            for case in project.iter_cases():
                case.status = "not_configured"
                case.actual = "未配置可执行步骤，无法验证实际表现。"
                case.executed = False
            finished = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            result = self._finish(project, run_dir, started, finished)
            self.log("没有可执行步骤，已生成未配置报告。")
            return result

        setup_failed = ""
        browser = None
        context = None
        page = None

        try:
            with sync_playwright() as playwright:
                try:
                    browser = playwright.chromium.launch(
                        channel=self.options.browser or "msedge",
                        headless=self.options.headless,
                    )
                except Exception as exc:
                    raise RunnerError(
                        f"浏览器启动失败：{exc}\n"
                        "请确认电脑已安装 Microsoft Edge 或 Google Chrome。"
                    ) from exc

                context_args = {
                    "accept_downloads": True,
                    "locale": "zh-CN",
                    "viewport": {"width": 1440, "height": 960},
                }
                try:
                    context = browser.new_context(
                        downloads_path=str(run_dir / "downloads"),
                        **context_args,
                    )
                except TypeError:
                    context = browser.new_context(**context_args)
                context.set_default_timeout(self.options.timeout_ms)

                # 前置步骤只执行一次，例如登录。
                if project.setup_steps:
                    self.log(f"执行前置步骤（{len(project.setup_steps)} 步）")
                    page = context.new_page()
                    page.set_default_timeout(self.options.timeout_ms)
                    try:
                        for index, step in enumerate(project.setup_steps, start=1):
                            self._execute_step(
                                page, step, index, run_dir / "screenshots", "SETUP"
                            )
                        self.log("前置步骤执行成功")
                    except Exception as exc:
                        setup_failed = _short(str(exc), 500)
                        self.log(f"前置步骤失败：{setup_failed}")
                        try:
                            page.screenshot(
                                path=str(run_dir / "screenshots" / "setup_failed.png"),
                                full_page=True,
                            )
                        except Exception:
                            pass

                for requirement in project.requirements:
                    for case in requirement.cases:
                        if self.options.stop_event and self.options.stop_event.is_set():
                            case.status = "skipped"
                            case.actual = "用户停止执行"
                            case.executed = False
                            done += 1
                            if self.options.on_case:
                                self.options.on_case(case)
                            if self.options.on_progress:
                                self.options.on_progress(done, total)
                            continue

                        if case.status == "not_configured" or not case.steps:
                            case.status = "not_configured"
                            case.actual = "未配置可执行步骤，无法验证实际表现。"
                            case.executed = False
                            self.log(f"[未配置] {case.name}")
                            done += 1
                            if self.options.on_case:
                                self.options.on_case(case)
                            if self.options.on_progress:
                                self.options.on_progress(done, total)
                            continue

                        if setup_failed:
                            case.status = "skipped"
                            case.actual = f"前置步骤失败，未执行用例：{setup_failed}"
                            case.error = setup_failed
                            case.executed = False
                            self.log(f"[跳过] {case.name}")
                            done += 1
                            if self.options.on_case:
                                self.options.on_case(case)
                            if self.options.on_progress:
                                self.options.on_progress(done, total)
                            continue

                        self._run_case(context, case, run_dir)
                        done += 1
                        if self.options.on_case:
                            self.options.on_case(case)
                        if self.options.on_progress:
                            self.options.on_progress(done, total)

        except RunnerError as exc:
            self.log(f"执行中止：{exc}")
            for case in project.iter_cases():
                if case.status == "pending":
                    case.status = "skipped"
                    case.actual = f"执行中止：{exc}"
            finished = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            result = self._finish(project, run_dir, started, finished)
            result.error = str(exc)
            return result
        except Exception as exc:
            self.log(f"执行异常：{exc}\n{traceback.format_exc()}")
            for case in project.iter_cases():
                if case.status == "pending":
                    case.status = "skipped"
                    case.actual = f"执行异常：{exc}"
            finished = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            result = self._finish(project, run_dir, started, finished)
            result.error = str(exc)
            return result
        finally:
            try:
                if browser:
                    browser.close()
            except Exception:
                pass

        finished = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        result = self._finish(project, run_dir, started, finished)
        self.log("测试执行完成")
        return result

    # ------------------------------------------------------------------
    def _run_case(self, context, case: TestCase, run_dir: Path) -> None:
        page = None
        case.started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # type: ignore[attr-defined]
        case.status = "pending"
        start = time.perf_counter()
        self.log(f"[执行] {case.name}")
        try:
            page = context.new_page()
            page.set_default_timeout(self.options.timeout_ms)
            if not case.steps or case.steps[0].action != "goto":
                page.goto(self.options.base_url, wait_until="domcontentloaded")

            for index, step in enumerate(case.steps, start=1):
                if self.options.stop_event and self.options.stop_event.is_set():
                    raise RunnerError("用户停止执行")
                self._execute_step(page, step, index, run_dir / "screenshots", case.case_id)

            case.status = "pass"
            case.actual = (
                f"全部 {len(case.steps)} 个步骤执行成功；"
                f"最终页面：{page.title()}（{_display_url(page.url)}）"
            )
            self.log(f"[通过] {case.name}")
        except Exception as exc:
            case.status = "fail"
            case.error = _short(str(exc), 1000)
            try:
                if page:
                    url = _display_url(page.url)
                    title = page.title()
                else:
                    url = ""
                    title = ""
            except Exception:
                url, title = "", ""
            case.actual = (
                f"执行未通过。当前页面：{title}（{url}）。"
                f"错误：{case.error}"
            )
            self.log(f"[不通过] {case.name}：{case.error}")
            try:
                if page:
                    screenshot = run_dir / "screenshots" / f"{case.case_id}.png"
                    page.screenshot(path=str(screenshot), full_page=True)
                    case.screenshot = str(screenshot.relative_to(run_dir)).replace("\\", "/")
            except Exception:
                pass
        finally:
            case.duration_s = round(time.perf_counter() - start, 2)
            case.executed = True
            try:
                if page:
                    page.close()
            except Exception:
                pass

    # ------------------------------------------------------------------
    def _execute_step(
        self,
        page,
        step: Step,
        index: int,
        screenshot_dir: Path,
        case_id: str,
    ) -> str:
        from playwright.sync_api import expect

        action = (step.action or "").strip().lower()
        target = (step.target or "").strip()
        value = (step.value or "").strip()
        value = value.replace("{{username}}", self.options.username).replace(
            "{{password}}", self.options.password
        )
        expected = (step.expected or "").strip()
        timeout = max(1000, int(step.timeout_ms or self.options.timeout_ms))

        if action == "goto":
            url = _join_url(self.options.base_url, target)
            page.goto(url, wait_until="domcontentloaded", timeout=timeout)
            return f"打开 {url}"

        if action == "screenshot":
            path = screenshot_dir / f"{case_id}_step{index}.png"
            page.screenshot(path=str(path), full_page=True)
            return f"截图 {path.name}"

        if action == "reload":
            page.reload(wait_until="domcontentloaded", timeout=timeout)
            return "刷新页面"

        if action == "emulate_media":
            media = (value or "print").strip().lower()
            if media not in ("print", "screen"):
                raise RunnerError("emulate_media 只支持 print 或 screen")
            page.emulate_media(media=media)
            return f"切换媒体类型：{media}"

        if action == "wait":
            if target:
                page.locator(target).first.wait_for(state="visible", timeout=timeout)
                return f"等待元素 {target}"
            milliseconds = int(value or target or 1000)
            page.wait_for_timeout(max(0, milliseconds))
            return f"等待 {milliseconds} 毫秒"

        if action == "assert_url":
            expected_url = value or expected or target
            if expected_url not in page.url:
                raise AssertionError(
                    f"期望网址包含“{expected_url}”，实际为“{page.url}”"
                )
            return f"网址断言通过：{expected_url}"

        if not target:
            raise RunnerError(f"第 {index} 步缺少目标选择器")

        if action == "click":
            locator = page.locator(target).first
            locator.wait_for(state="visible", timeout=timeout)
            locator.click(timeout=timeout)
            return f"点击 {target}"

        if action == "fill":
            locator = page.locator(target).first
            locator.wait_for(state="visible", timeout=timeout)
            locator.fill(value, timeout=timeout)
            return f"填写 {target}"

        if action == "select":
            page.locator(target).first.select_option(value, timeout=timeout)
            return f"选择 {target} = {value}"

        if action == "press":
            page.locator(target).first.press(value or "Enter", timeout=timeout)
            return f"在 {target} 按下 {value or 'Enter'}"

        if action == "assert_text":
            locator = page.locator(target).first
            locator.wait_for(state="visible", timeout=timeout)
            expected_text = value or expected
            if not expected_text:
                raise RunnerError(f"第 {index} 步断言文本为空")
            actual = locator.inner_text(timeout=timeout)
            if expected_text not in actual:
                raise AssertionError(
                    f"期望文本包含“{expected_text}”，实际为“{_short(actual, 240)}”"
                )
            return f"文本断言通过：{expected_text}"

        if action == "assert_visible":
            page.locator(target).first.wait_for(state="visible", timeout=timeout)
            return f"元素可见：{target}"

        if action == "assert_count":
            expected_raw = value or expected
            if not expected_raw:
                raise RunnerError(f"第 {index} 步断言数量为空")
            expected_count = int(float(expected_raw))
            deadline = time.time() + timeout / 1000
            actual_count = page.locator(target).count()
            while time.time() < deadline:
                actual_count = page.locator(target).count()
                if actual_count == expected_count:
                    return f"元素数量断言通过：{expected_count}"
                page.wait_for_timeout(120)
            raise AssertionError(
                f"期望元素数量为 {expected_count}，实际为 {actual_count}（{target}）"
            )

        if action == "assert_sorted":
            direction = (value or "asc").strip().lower()
            deadline = time.time() + timeout / 1000
            numbers: list[float] = []
            texts: list[str] = []
            while time.time() < deadline:
                texts = page.locator(target).all_inner_texts()
                numbers = []
                for text in texts:
                    found = re.search(r"-?\d+(?:\.\d+)?", text)
                    if found:
                        numbers.append(float(found.group(0)))
                if len(numbers) >= 2:
                    break
                page.wait_for_timeout(120)
            if len(numbers) < 2:
                raise AssertionError(
                    f"排序断言至少需要 2 个数值，实际取到 {len(numbers)} 个（{target}）"
                )
            ordered = (
                all(numbers[i] <= numbers[i + 1] for i in range(len(numbers) - 1))
                if direction != "desc"
                else all(numbers[i] >= numbers[i + 1] for i in range(len(numbers) - 1))
            )
            if not ordered:
                raise AssertionError(
                    f"期望数值按{'降序' if direction == 'desc' else '升序'}排列，"
                    f"实际前几项为 {numbers[:8]}"
                )
            return f"排序断言通过：{direction}，共 {len(numbers)} 项"

        if action == "assert_duplicate_max":
            expected_max = int(float(value or expected or "2"))
            texts = page.locator(target).all_inner_texts()
            counts = Counter(text.strip() for text in texts if text.strip())
            if not counts:
                raise AssertionError(f"重复上限断言没有取到文本（{target}）")
            max_count = max(counts.values())
            if max_count > expected_max:
                offenders = [name for name, count in counts.items() if count > expected_max][:5]
                raise AssertionError(
                    f"期望同一名称最多出现 {expected_max} 次，实际最多 {max_count} 次；"
                    f"超出项：{offenders}"
                )
            return f"重复上限断言通过：最多 {max_count} 次 / 上限 {expected_max}"

        if action == "assert_attribute":
            if "=" not in value:
                raise RunnerError("assert_attribute 的输入值格式应为 attribute=expected")
            attr, _, expected_value = value.partition("=")
            actual_value = page.locator(target).first.get_attribute(attr.strip())
            if actual_value != expected_value:
                raise AssertionError(
                    f"属性 {attr} 期望为“{expected_value}”，实际为“{actual_value}”"
                )
            return f"属性断言通过：{attr}={expected_value}"

        if action == "assert_value":
            expected_value = value or expected
            actual_value = page.locator(target).first.input_value(timeout=timeout)
            if actual_value != expected_value:
                raise AssertionError(
                    f"输入值期望为“{expected_value}”，实际为“{actual_value}”"
                )
            return f"输入值断言通过：{expected_value}"

        if action == "assert_hidden":
            locator = page.locator(target).first
            if locator.is_visible():
                raise AssertionError(f"期望元素不可见，实际仍可见：{target}")
            return f"元素隐藏断言通过：{target}"

        raise RunnerError(f"不支持的操作类型：{action}")

    # ------------------------------------------------------------------
    def _finish(
        self, project: Project, run_dir: Path, started: str, finished: str
    ) -> RunResult:
        from .report import generate_reports

        counts = project.status_counts()
        metadata = {
            "project_name": project.name,
            "source_excel": project.source_excel,
            "source_sheet": project.source_sheet,
            "base_url": self.options.base_url,
            "browser": self.options.browser,
            "headless": self.options.headless,
            "started_at": started,
            "finished_at": finished,
            "counts": counts,
            "total": project.case_count,
        }
        html_path, xlsx_path = generate_reports(project, metadata, run_dir)
        (run_dir / "run.json").write_text(
            json.dumps(
                {
                    "metadata": metadata,
                    "project": project.to_dict(),
                    "report_html": html_path.name,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        self.storage.config["last_report"] = str(html_path)
        self.storage.save_config()
        return RunResult(
            run_dir=run_dir,
            report_html=html_path,
            report_xlsx=xlsx_path,
            counts=counts,
            started_at=started,
            finished_at=finished,
        )


def check_web(
    url: str,
    browser: str = "msedge",
    headless: bool = True,
    timeout_ms: int = 15000,
) -> dict[str, object]:
    """检查网址是否能打开，返回状态、标题和错误。"""
    from playwright.sync_api import sync_playwright

    if not url.strip():
        return {"ok": False, "status": 0, "title": "", "url": "", "error": "网址为空"}
    result: dict[str, object] = {
        "ok": False,
        "status": 0,
        "title": "",
        "url": "",
        "error": "",
    }
    try:
        with sync_playwright() as playwright:
            browser_obj = playwright.chromium.launch(channel=browser, headless=headless)
            context = browser_obj.new_context(locale="zh-CN")
            page = context.new_page()
            page.set_default_timeout(timeout_ms)
            response = page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            result.update(
                {
                    "ok": True,
                    "status": response.status if response else 0,
                    "title": page.title(),
                    "url": page.url,
                }
            )
            browser_obj.close()
    except Exception as exc:
        result["error"] = str(exc)
    return result
