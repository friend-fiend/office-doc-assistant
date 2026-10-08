r"""路径工具。

重点解决三件事：
1. Windows 长路径（>260 字符）—— 统一加 \\?\ 前缀；
2. 中文 / 含空格 / 含 emoji 的路径 —— 全程使用 str 与 pathlib，禁止手工拼接；
3. 输出目录策略 —— 永不覆盖原文件，冲突时自动追加序号。
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path

from .config import IS_MACOS, IS_WINDOWS, OUTPUT_SUFFIX

_ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_WIN = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def long_path(path: str | os.PathLike) -> str:
    r"""Windows 下把路径转成扩展长度前缀形式，绕过 MAX_PATH 限制。

    非 Windows 平台原样返回。已经是 \\?\ 形式的不重复添加。
    """
    p = os.fspath(path)
    if not IS_WINDOWS:
        return p
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):  # UNC: \\server\share -> \\?\UNC\server\share
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def open_path(path: str | os.PathLike) -> None:
    """用系统默认方式打开文件或文件夹。"""
    p = os.fspath(path)
    if not os.path.exists(p):
        raise FileNotFoundError(p)
    if IS_WINDOWS:
        os.startfile(p)  # type: ignore[attr-defined]
    elif IS_MACOS:
        subprocess.Popen(["open", p])
    else:
        subprocess.Popen(["xdg-open", p])


def reveal_path(path: str | os.PathLike) -> None:
    """在文件管理器中定位到该项（Windows/macOS 支持）。"""
    p = os.path.abspath(os.fspath(path))
    if IS_WINDOWS:
        subprocess.Popen(["explorer", "/select,", os.path.normpath(p)])
    elif IS_MACOS:
        subprocess.Popen(["open", "-R", p])
    else:
        open_path(os.path.dirname(p))


def sanitize_name(name: str, replacement: str = "_") -> str:
    """清洗文件名中的非法字符，并处理 Windows 保留名。"""
    if not name:
        return "unnamed"
    # 归一化，避免 macOS 上 NFD/NFC 混用导致同名文件不同名
    name = unicodedata.normalize("NFC", name)
    name = _ILLEGAL.sub(replacement, name).strip().strip(".")
    if not name:
        name = "unnamed"
    stem = name.split(".")[0].upper()
    if stem in _RESERVED_WIN:
        name = "_" + name
    return name[:200]


def unique_dir(parent: str | os.PathLike, stem: str) -> str:
    """在 parent 下创建一个不冲突的新目录，返回其路径。

    首次使用 ``<stem>_输出``，之后依次 ``<stem>_输出_2``、``_3`` …
    """
    parent = os.fspath(parent)
    base = sanitize_name(stem) + OUTPUT_SUFFIX
    candidate = os.path.join(parent, base)
    idx = 2
    while os.path.exists(candidate):
        candidate = os.path.join(parent, f"{base}_{idx}")
        idx += 1
    os.makedirs(long_path(candidate), exist_ok=False)
    return candidate


def ensure_dir(path: str | os.PathLike) -> str:
    p = os.fspath(path)
    os.makedirs(long_path(p), exist_ok=True)
    return p


def unique_file(path: str | os.PathLike) -> str:
    """返回一个当前不存在的文件路径（同名则追加 _2、_3 …）。"""
    p = os.path.abspath(os.fspath(path))
    if not os.path.exists(p):
        return p
    root, ext = os.path.splitext(p)
    idx = 2
    while os.path.exists(f"{root}_{idx}{ext}"):
        idx += 1
    return f"{root}_{idx}{ext}"


def temp_dir(prefix: str = "oda_") -> str:
    d = tempfile.mkdtemp(prefix=prefix)
    return d


def cleanup_dir(path: str | os.PathLike) -> None:
    """尽力删除临时目录，失败不抛异常。"""
    try:
        shutil.rmtree(long_path(path), ignore_errors=True)
    except Exception:
        pass


def free_space_mb(path: str | os.PathLike) -> float:
    """返回目标路径所在磁盘的可用空间（MB）。"""
    try:
        probe = os.path.abspath(os.fspath(path))
        while probe and not os.path.exists(probe):
            parent = os.path.dirname(probe)
            if parent == probe:
                break
            probe = parent
        return shutil.disk_usage(probe).free / (1024 * 1024)
    except Exception:
        return float("inf")


def same_file(a: str | os.PathLike, b: str | os.PathLike) -> bool:
    try:
        return os.path.samefile(long_path(a), long_path(b))
    except Exception:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def human_size(num_bytes: float) -> str:
    """人类可读的大小，例如 1.23 MB。"""
    if num_bytes is None:
        return "-"
    n = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(n) < 1024.0 or unit == "TB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.2f} {unit}"
        n /= 1024.0
    return f"{n:.2f} TB"


def path_len(path: str) -> int:
    return len(os.path.abspath(os.fspath(path)))


def is_windows_long_path_enabled() -> bool:
    """查询注册表，判断系统是否已启用长路径支持。"""
    if not IS_WINDOWS:
        return True
    try:
        import winreg  # type: ignore

        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\FileSystem",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "LongPathsEnabled")
            return bool(value)
    except Exception:
        return False


def scripts_dir() -> Path:
    """当前 Python 解释器的 bin/Scripts 目录（用于定位同目录下的外部工具）。"""
    return Path(sys.executable).resolve().parent
