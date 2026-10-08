"""外部依赖探测：Poppler / LibreOffice / Ghostscript。

探测优先级（高 -> 低）：
  1. 用户在"设置"里手动指定的路径
  2. 程序自带目录（vendor/ 或 PyInstaller 解包目录）—— 便于做"绿色便携版"
  3. 与解释器同级的 Scripts 目录
  4. 系统 PATH
  5. 各平台常见安装位置

探测结果缓存在 ``ToolStatus`` 里，界面据此禁用按钮并给出下载指引。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

from .config import IS_MACOS, IS_WINDOWS, app_root, resource_dir
from .settings import SETTINGS

_NO_WINDOW = 0
if IS_WINDOWS:
    _NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)


def _run(cmd: list[str], timeout: int = 12) -> tuple[int, str]:
    """在无窗口模式下执行命令，返回 (returncode, 输出)。"""
    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=timeout,
            creationflags=_NO_WINDOW,
            encoding="utf-8",
            errors="replace",
        )
        return proc.returncode, (proc.stdout or "").strip()
    except FileNotFoundError:
        return 127, ""
    except subprocess.TimeoutExpired:
        return 124, "timeout"
    except Exception as exc:  # pragma: no cover
        return 1, str(exc)


@dataclass
class ToolSpec:
    key: str
    title: str
    # 需要检测的可执行文件名（命中任意一个即可）
    exes: tuple[str, ...]
    # 校验版本用的参数
    version_args: tuple[str, ...]
    # 常见安装目录（相对/绝对路径均可）
    common: tuple[str, ...] = ()
    # 该工具缺失时受影响的功能
    affects: tuple[str, ...] = ()
    help_url: str = ""
    help_text: str = ""
    # True 表示这是一个"目录"型依赖（Poppler 需要整个 bin 目录）
    is_dir: bool = False


@dataclass
class ToolStatus:
    spec: ToolSpec
    available: bool = False
    path: str = ""          # 可执行文件绝对路径
    directory: str = ""     # 所在目录（Poppler 用）
    version: str = ""
    source: str = "none"    # settings | bundled | scripts | path | common
    error: str = ""

    @property
    def display(self) -> str:
        if not self.available:
            return f"✗ {self.spec.title}：未检测到"
        tag = {
            "settings": "手动指定",
            "bundled": "程序自带",
            "scripts": "解释器同级",
            "path": "系统 PATH",
            "common": "默认安装位置",
        }.get(self.source, self.source)
        ver = f" v{self.version}" if self.version else ""
        return f"✓ {self.spec.title}{ver}（{tag}）"


TOOL_SPECS: tuple[ToolSpec, ...] = (
    ToolSpec(
        key="poppler",
        title="Poppler（PDF 转图片）",
        exes=("pdftoppm.exe", "pdftoppm") if IS_WINDOWS else ("pdftoppm",),
        version_args=("-v",),
        common=(
            r"C:\Program Files\poppler\Library\bin",
            r"C:\Program Files\poppler\bin",
            r"C:\poppler\Library\bin",
            r"C:\poppler\bin",
            "/opt/homebrew/bin",
            "/usr/local/bin",
            "/opt/local/bin",
        ),
        affects=("PDF 转图片",),
        help_url="https://github.com/oschwartz10612/poppler-windows/releases",
        help_text=(
            "Windows：下载 Release-Poppler-xx.xx.x-0.zip，解压到 C:\\Program Files\\poppler，"
            "把里面的 Library\\bin 目录填到设置里；\n"
            "macOS：brew install poppler（Apple Silicon 装到 /opt/homebrew/bin）"
        ),
        is_dir=True,
    ),
    ToolSpec(
        key="soffice",
        title="LibreOffice（Office 转 PDF）",
        exes=("soffice.exe", "soffice.com") if IS_WINDOWS else ("soffice",),
        version_args=("--version",),
        common=(
            r"C:\Program Files\LibreOffice\program",
            r"C:\Program Files (x86)\LibreOffice\program",
            "/Applications/LibreOffice.app/Contents/MacOS",
            "/usr/bin",
            "/usr/local/bin",
            "/opt/homebrew/bin",
            "/snap/bin",
        ),
        affects=("Word/Excel/PPT 转 PDF",),
        help_url="https://www.libreoffice.org/download/download-libreoffice/",
        help_text="Windows/macOS 均可从官网下载安装包，一路默认安装即可；安装后重启本程序。",
        is_dir=False,
    ),
    ToolSpec(
        key="ghostscript",
        title="Ghostscript（PDF 深度压缩）",
        exes=("gswin64c.exe", "gswin32c.exe", "gswin64c", "gs") if IS_WINDOWS else ("gs",),
        version_args=("--version",),
        common=(
            r"C:\Program Files\gs\gs10.05.1\bin",
            r"C:\Program Files\gs\gs10.04.0\bin",
            r"C:\Program Files\gs\gs10.03.1\bin",
            r"C:\Program Files\gs\gs10.02.1\bin",
            r"C:\Program Files\gs\gs10.01.2\bin",
            r"C:\Program Files\gs\gs9.56.1\bin",
            "/opt/homebrew/bin",
            "/usr/local/bin",
            "/usr/bin",
        ),
        affects=("PDF 中度/强力压缩",),
        help_url="https://ghostscript.com/releases/gsdnld.html",
        help_text=(
            "Windows：下载 gs10.0x.x-win64.exe 安装（务必勾选「添加到 PATH」或使用默认目录）；\n"
            "macOS：brew install ghostscript。\n"
            "缺失时 PDF 压缩会自动降级为 pypdf 结构优化，效果有限。"
        ),
        is_dir=False,
    ),
)


_CACHE: dict[str, ToolStatus] = {}


def _candidates(spec: ToolSpec) -> list[tuple[str, str]]:
    """生成 (来源标记, 目录) 候选列表。"""
    out: list[tuple[str, str]] = []

    override = (SETTINGS.get(f"{'poppler_dir' if spec.key == 'poppler' else spec.key + '_path'}") or "").strip()
    if override:
        p = os.path.expanduser(override.strip('"').strip("'"))
        out.append(("settings", p if spec.is_dir else os.path.dirname(p) or p))

    roots = [resource_dir(), app_root()]
    for r in roots:
        for sub in ("vendor", os.path.join("vendor", spec.key)):
            out.append(("bundled", str(Path(r) / sub)))

    try:
        out.append(("scripts", str(Path(sys.executable).resolve().parent)))
    except Exception:
        pass

    which_dir = ""
    for exe in spec.exes:
        found = shutil.which(exe)
        if found:
            which_dir = os.path.dirname(found)
            break
    if which_dir:
        out.append(("path", which_dir))

    for c in spec.common:
        out.append(("common", c))

    # 去重并保序
    seen: set[str] = set()
    uniq: list[tuple[str, str]] = []
    for src, d in out:
        if not d:
            continue
        norm = os.path.normcase(os.path.abspath(d))
        if norm in seen:
            continue
        seen.add(norm)
        uniq.append((src, d))
    return uniq


def _parse_version(spec: ToolSpec, text: str) -> str:
    import re

    if not text:
        return ""
    # pdftoppm -v 输出在 stderr，形如 "pdftoppm version 24.08.0"
    m = re.search(r"(\d+\.\d+(?:\.\d+)?)", text)
    return m.group(1) if m else text.splitlines()[0][:40]


def detect(key: str, force: bool = False) -> ToolStatus:
    """检测单个工具。结果会缓存，``force=True`` 时重新探测。"""
    if not force and key in _CACHE:
        return _CACHE[key]

    spec = next((s for s in TOOL_SPECS if s.key == key), None)
    if spec is None:
        raise KeyError(key)

    status = ToolStatus(spec=spec)
    tried: list[str] = []

    for source, directory in _candidates(spec):
        if not directory or not os.path.isdir(directory):
            continue
        for exe in spec.exes:
            candidate = os.path.join(directory, exe)
            tried.append(candidate)
            if not os.path.isfile(candidate):
                continue
            code, out = _run([candidate, *spec.version_args])
            if code in (0, 1) or "version" in out.lower():
                status.available = True
                status.path = candidate
                status.directory = directory
                status.version = _parse_version(spec, out)
                status.source = source
                _CACHE[key] = status
                return status

    # PATH 直接命中再兜一次（有些环境 which 命中但目录不在 common 里）
    for exe in spec.exes:
        found = shutil.which(exe)
        if found:
            code, out = _run([found, *spec.version_args])
            if code in (0, 1) or out:
                status.available = True
                status.path = found
                status.directory = os.path.dirname(found)
                status.version = _parse_version(spec, out)
                status.source = "path"
                _CACHE[key] = status
                return status

    status.error = "未在以下位置找到可执行文件：" + "；".join(tried[:6]) if tried else "未找到"
    _CACHE[key] = status
    return status


def detect_all(force: bool = False) -> dict[str, ToolStatus]:
    if force:
        _CACHE.clear()
    return {spec.key: detect(spec.key, force=force) for spec in TOOL_SPECS}


def get(key: str) -> ToolStatus:
    return _CACHE.get(key) or detect(key)


def is_available(key: str) -> bool:
    return get(key).available


def ghostscript_exe() -> str:
    return get("ghostscript").path


def soffice_exe() -> str:
    return get("soffice").path


def poppler_dir() -> str:
    st = get("poppler")
    return st.directory if st.available else ""


def missing_report() -> list[ToolStatus]:
    return [st for st in detect_all().values() if not st.available]
