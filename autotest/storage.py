"""本地存储、配置与运行目录管理。

所有缓存、日志、截图、下载、报告和临时文件都放在用户选择的数据根目录。
数据根目录默认选择非 C 盘，路径记录在 Windows 注册表，避免在系统盘写配置文件。
"""

from __future__ import annotations

import json
import os
import shutil
import string
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import winreg
except ImportError:  # pragma: no cover - 非 Windows 环境
    winreg = None


APP_NAME = "AutoTestWorkbench"
REG_PATH = r"Software\AutoTestWorkbench"
DEFAULT_FOLDER_NAME = "AutoTestWorkbenchData"


def _non_c_drives() -> list[tuple[str, int]]:
    """返回可用的非 C 盘及剩余空间（GB），按剩余空间降序。"""
    drives: list[tuple[str, int]] = []
    for letter in string.ascii_uppercase:
        if letter == "C":
            continue
        root = Path(f"{letter}:\\")
        try:
            if not root.exists():
                continue
            free = shutil.disk_usage(str(root)).free // (1024**3)
        except OSError:
            continue
        drives.append((letter, int(free)))
    drives.sort(key=lambda item: item[1], reverse=True)
    return drives


def default_data_root() -> Path:
    """优先选择剩余空间最大的非 C 盘；实在没有才退回用户目录。"""
    for letter, free_gb in _non_c_drives():
        if free_gb >= 5:
            return Path(f"{letter}:\\{DEFAULT_FOLDER_NAME}")
    return Path.home() / DEFAULT_FOLDER_NAME


def _read_registry_root() -> Path | None:
    if winreg is None:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH) as key:
            value, _ = winreg.QueryValueEx(key, "StorageRoot")
            return Path(str(value))
    except (FileNotFoundError, OSError):
        return None


def _write_registry_root(root: Path) -> None:
    if winreg is None:
        return
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH)
    try:
        winreg.SetValueEx(key, "StorageRoot", 0, winreg.REG_SZ, str(root))
    finally:
        winreg.CloseKey(key)


def is_on_c_drive(path: Path) -> bool:
    try:
        return path.resolve().drive.upper() == "C:"
    except OSError:
        return False


class AppStorage:
    """应用数据目录与配置。"""

    DEFAULT_CONFIG: dict[str, Any] = {
        "base_url": "",
        "browser": "msedge",
        "headless": True,
        "timeout_ms": 15000,
        "source_excel": "",
        "last_report": "",
    }

    def __init__(self, root: Path | str | None = None):
        if root is None:
            candidate = _read_registry_root()
        else:
            candidate = Path(root)

        if candidate is None or not self._can_use(candidate):
            candidate = default_data_root()

        self.root = Path(candidate)
        self.ensure_dirs()
        _write_registry_root(self.root)
        self.config = self._load_config()
        self.apply_temp_env()

    @staticmethod
    def _can_use(path: Path) -> bool:
        try:
            path.mkdir(parents=True, exist_ok=True)
            probe = path / ".write_test"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink()
            return True
        except OSError:
            return False

    # ------------------------------------------------------------------
    # 目录
    # ------------------------------------------------------------------
    @property
    def config_path(self) -> Path:
        return self.root / "config.json"

    @property
    def projects_dir(self) -> Path:
        return self.root / "projects"

    @property
    def reports_dir(self) -> Path:
        return self.root / "reports"

    @property
    def logs_dir(self) -> Path:
        return self.root / "logs"

    @property
    def cache_dir(self) -> Path:
        return self.root / "cache"

    @property
    def downloads_dir(self) -> Path:
        return self.root / "downloads"

    @property
    def temp_dir(self) -> Path:
        return self.root / "tmp"

    @property
    def browsers_dir(self) -> Path:
        return self.root / "browsers"

    @property
    def profiles_dir(self) -> Path:
        return self.root / "browser_profiles"

    def ensure_dirs(self) -> None:
        for directory in (
            self.root,
            self.projects_dir,
            self.reports_dir,
            self.logs_dir,
            self.cache_dir,
            self.downloads_dir,
            self.temp_dir,
            self.browsers_dir,
            self.profiles_dir,
        ):
            directory.mkdir(parents=True, exist_ok=True)

    def apply_temp_env(self) -> None:
        """把 Python、Playwright 和下载临时目录指向数据根目录。"""
        temp = str(self.temp_dir)
        os.environ["TMP"] = temp
        os.environ["TEMP"] = temp
        os.environ["TMPDIR"] = temp
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(self.browsers_dir)

    # ------------------------------------------------------------------
    # 配置
    # ------------------------------------------------------------------
    def _load_config(self) -> dict[str, Any]:
        config = dict(self.DEFAULT_CONFIG)
        if self.config_path.exists():
            try:
                loaded = json.loads(self.config_path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    config.update(loaded)
            except (OSError, ValueError):
                pass
        return config

    def save_config(self) -> None:
        self.config_path.write_text(
            json.dumps(self.config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def set_root(self, root: Path | str) -> None:
        new_root = Path(root)
        if not self._can_use(new_root):
            raise OSError(f"目录不可写：{new_root}")
        old_config = dict(self.config)
        self.root = new_root
        self.ensure_dirs()
        _write_registry_root(self.root)
        self.config = dict(self.DEFAULT_CONFIG)
        for key in ("base_url", "browser", "headless", "timeout_ms"):
            if key in old_config:
                self.config[key] = old_config[key]
        self.save_config()
        self.apply_temp_env()

    # ------------------------------------------------------------------
    # 运行目录
    # ------------------------------------------------------------------
    def new_run_dir(self) -> Path:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        run_dir = self.reports_dir / f"run_{stamp}"
        suffix = 1
        while run_dir.exists():
            run_dir = self.reports_dir / f"run_{stamp}_{suffix}"
            suffix += 1
        for sub in ("screenshots", "logs", "downloads"):
            (run_dir / sub).mkdir(parents=True, exist_ok=True)
        return run_dir

    # ------------------------------------------------------------------
    # 项目保存
    # ------------------------------------------------------------------
    def save_project(self, project: Any) -> Path:
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        path = self.projects_dir / f"{project.project_id}.json"
        path.write_text(
            json.dumps(project.to_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return path

    def list_project_files(self) -> list[Path]:
        if not self.projects_dir.exists():
            return []
        return sorted(
            self.projects_dir.glob("*.json"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )

    def open_in_explorer(self, path: Path | None = None) -> None:
        target = Path(path or self.root)
        if sys.platform == "win32":
            os.startfile(str(target))  # type: ignore[attr-defined]
