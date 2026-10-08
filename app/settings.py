"""用户设置持久化（外部工具路径、界面偏好）。

存储位置：
  Windows  %APPDATA%\\OfficeDocAssistant\\settings.json
  macOS    ~/Library/Application Support/OfficeDocAssistant/settings.json
  Linux    ~/.config/OfficeDocAssistant/settings.json
"""
from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


def _config_home() -> Path:
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    d = Path(base) / "OfficeDocAssistant"
    d.mkdir(parents=True, exist_ok=True)
    return d


def config_dir() -> Path:
    """对外暴露配置目录（体检脚本、设置窗口「打开配置目录」都用它）。"""
    return _config_home()


DEFAULTS: dict[str, Any] = {
    # 外部工具路径（空表示自动探测）
    "poppler_dir": "",
    "soffice_path": "",
    "ghostscript_path": "",
    # 界面偏好
    "last_input_dir": "",
    "last_output_dir": "",
    "window_geometry": "",
    "remember_paths": True,
    "verbose_log": False,
    # 行为
    "open_output_when_done": False,
    "keep_temp": False,
}


@dataclass
class Settings:
    data: dict = field(default_factory=lambda: dict(DEFAULTS))

    # -------------------------------------------------- 读写
    @classmethod
    def load(cls) -> "Settings":
        path = _config_home() / "settings.json"
        data = dict(DEFAULTS)
        try:
            if path.exists():
                raw = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(raw, dict):
                    for k, v in raw.items():
                        if k in DEFAULTS and isinstance(v, type(DEFAULTS[k])):
                            data[k] = v
                        elif k in DEFAULTS:
                            data[k] = v
                        else:
                            data[k] = v
        except Exception:
            pass
        return cls(data=data)

    def save(self) -> None:
        try:
            path = _config_home() / "settings.json"
            path.write_text(
                json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception:
            pass

    # -------------------------------------------------- 访问
    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, DEFAULTS.get(key, default))

    def set(self, key: str, value: Any) -> None:
        self.data[key] = value

    def as_dict(self) -> dict:
        return asdict(self)["data"]

    def reset_tools(self) -> None:
        for k in ("poppler_dir", "soffice_path", "ghostscript_path"):
            self.data[k] = ""


SETTINGS = Settings.load()
