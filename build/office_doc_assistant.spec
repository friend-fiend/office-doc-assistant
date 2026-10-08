"""PyInstaller 打包配置（Windows / macOS 共用）。

用法
----
    Windows:  pyinstaller build/office_doc_assistant.spec --noconfirm
    macOS:    pyinstaller build/office_doc_assistant.spec --noconfirm

产物
----
    Windows:  dist/OfficeDocAssistant/OfficeDocAssistant.exe   （onedir）
    macOS:    dist/OfficeDocAssistant.app                       （app bundle）

为什么用 onedir 而不是 onefile
------------------------------
1. onefile 每次启动都要把整个包解压到临时目录，冷启动 3~8 秒；onedir 几乎瞬时。
2. onefile 在异常退出时会残留临时目录（用户磁盘越用越少），onedir 不会。
3. 本程序要调用 Poppler/Ghostscript 等外部程序，onedir 便于把 vendor/ 放在旁边。
4. onefile 更容易被杀毒软件误报。

为什么要在 spec 里显式收集二进制
--------------------------------
pypdf / reportlab / Pillow / py7zr / tkinterdnd2 都带非 .py 资源
（字体、DLL、tkdnd 扩展），PyInstaller 的静态分析经常抓不全，
这里用 collect_all / collect_data_files 显式收，避免"打包后功能缺失"。
"""
from __future__ import annotations

import os
import sys

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

# ------------------------------------------------------------------ 路径
# 注意：PyInstaller 6.20+ 执行 spec 时不再注入 __file__（实测 NameError），
# 但始终保留官方变量 SPECPATH = spec 所在目录；两者都兜底以防再变。
try:
    SPEC_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    try:
        SPEC_DIR = SPECPATH  # noqa: F821  (PyInstaller 注入)
    except NameError:
        SPEC_DIR = os.getcwd()
PROJECT_DIR = os.path.dirname(SPEC_DIR)
ENTRY = os.path.join(PROJECT_DIR, "main.py")

IS_WIN = sys.platform.startswith("win")
IS_MAC = sys.platform == "darwin"
APP_NAME = "OfficeDocAssistant"
DISPLAY_NAME = "办公文档一键处理助手"
VERSION = "1.0.2"


# ------------------------------------------------------------ 版本资源（自动生成）
# 这里曾经踩过一次真实的坑：version_info.txt 是手写文件，spec 里又写了一份 VERSION，
# 于是同一个 exe 有了**两个版本来源**。1.0.1 → 1.0.2 那次只顺手改了 StringStruct
# 里的字符串，`filevers` / `prodvers` 元组留在 (1, 0, 1, 0) 没动，结果是：
#   右键属性页显示 1.0.2.0（读字符串），
#   而 win32api.GetFileVersionInfo 读出 1.0.1.0（读 ffi 元组）。
# 企业资产盘点、升级判断、签名校验都看后者 → 版本号直接判错，而且不报错、不告警。
# 现在由本文件现场渲染，spec 顶部的 VERSION 是唯一权威，结构上不可能再漂移。
def _render_version_info(version: str) -> str:
    """按 VERSION 渲染 VSVersionInfo 文本（ffi 元组与字符串表同源）。"""
    nums = [int(p) for p in version.split(".") if p.strip()]
    nums = (nums + [0, 0, 0, 0])[:4]
    tup = "(" + ", ".join(str(n) for n in nums) + ")"
    fields = [
        ("CompanyName", APP_NAME),
        ("FileDescription", f"{DISPLAY_NAME} - PDF/Office/图片/压缩批量处理工具"),
        ("FileVersion", f"{version}.0"),
        ("InternalName", APP_NAME),
        ("LegalCopyright", "本程序不捆绑任何需付费许可的第三方软件，详见自带 LICENSE-NOTICES.txt"),
        ("OriginalFilename", f"{APP_NAME}.exe"),
        ("ProductName", DISPLAY_NAME),
        ("ProductVersion", f"{version}.0"),
    ]
    rows = "\n".join(
        '          StringStruct({}, {}),'.format(_q(k), _q(v)) for k, v in fields)
    return "\n".join([
        "# UTF-8",
        "# 本文件由 build/office_doc_assistant.spec 在打包时自动生成，请勿手工编辑。",
        "# 版本号唯一来源：spec 顶部的 VERSION。改了 VERSION，重新打包即可同步。",
        "VSVersionInfo(",
        "  ffi=FixedFileInfo(",
        f"    filevers={tup},",
        f"    prodvers={tup},",
        "    mask=0x3F,",
        "    flags=0x0,",
        "    OS=0x40004,          # VOS_NT_WINDOWS32",
        "    fileType=0x1,        # VFT_APP",
        "    subtype=0x0,",
        "    date=(0, 0),",
        "  ),",
        "  kids=[",
        "    StringFileInfo(",
        "      [",
        # 080404b0 = 简体中文 + Unicode；换成别的 StringTable 中文会变问号
        '        StringTable("080404b0", [',
        rows,
        "        ])",
        "      ]",
        "    ),",
        '    VarFileInfo([VarStruct("Translation", [2052, 1200])]),',
        "  ]",
        ")",
        "",
    ])


def _q(text: str) -> str:
    """把字符串转成双引号字面量（版本资源要的是合法的 Python 字符串字面量）。"""
    escaped = text.replace("\\", "\\\\").replace('"', '\\"')
    return '"' + escaped + '"'


def _write_version_info() -> str:
    """渲染并覆盖 build/version_info.txt，返回路径（构建产物，无需入库）。"""
    path = os.path.join(SPEC_DIR, "version_info.txt")
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(_render_version_info(VERSION))
    print(f"[spec] 版本资源已按 VERSION={VERSION} 重新生成：{os.path.basename(path)}")
    return path


# ------------------------------------------------------------------ 数据文件
datas = []
binaries = []
hiddenimports: list[str] = []

# 运行期用不到、但 collect_all 会顺手收进来的"开发资源"，打包时剔除。
#
# 目前只有 pymupdf/mupdf-devel：92 个 C/C++ 头文件 + 2 个 .lib 导入库，约 6.7 MB。
# 它只服务于 pymupdf._mupdf_devel() 这个开发期 API（给 Cython 扩展编译时指路），
# 本程序从不调用，运行期也从不读取 → 收进安装包纯属浪费用户磁盘和下载流量。
# 保留 pymupdf/mupdfcpp64.dll 和 _mupdf.pyd，那才是真正跑 PDF 渲染的东西。
_DROP_DATA_DIRS = ("mupdf-devel",)


def _is_dev_resource(src, dst) -> bool:
    """判断一条 datas 是否属于要剔除的开发资源目录。"""
    hay = (str(src) + "/" + str(dst)).replace("\\", "/")
    return any(f"/{d}/" in hay or hay.endswith("/" + d) for d in _DROP_DATA_DIRS)


# 需要显式收集的第三方包
# pymupdf：PDF 导出 / 批量打印的渲染后端，带 _mupdf.pyd C 扩展，必须 collect_all
for pkg in ("pypdf", "reportlab", "pdf2image", "py7zr", "tkinterdnd2",
            "openpyxl", "docx", "pptx", "et_xmlfile", "pymupdf"):
    try:
        d, b, h = collect_all(pkg)
        if _DROP_DATA_DIRS:
            kept = [it for it in d if not _is_dev_resource(it[0], it[1])]
            if len(kept) != len(d):
                print(f"[spec] {pkg}: 剔除 {len(d) - len(kept)} 项开发资源 "
                      f"({'/'.join(_DROP_DATA_DIRS)})")
            d = kept
            b = [it for it in b if not _is_dev_resource(it[0], it[1])]
        datas += d
        binaries += b
        hiddenimports += h
    except Exception as exc:  # 某个包没装也不应让打包失败
        print(f"[spec] collect_all({pkg}) 跳过：{exc}")

# 兜底：确保渲染后端的关键二进制确实收到了（被上面的过滤误伤时要立刻发现）
if not IS_MAC:
    _pm_ok = [s for s, _d in binaries if os.path.basename(s).lower() == "mupdfcpp64.dll"]
    if not _pm_ok:
        print("[spec] 警告：未收集到 mupdfcpp64.dll，PDF 导出/打印可能不可用")

# tkinterdnd2 的 tkdnd 二进制（tkinterdnd2 自带的 hook 有时抓不到，这里再兜一次）
try:
    datas += collect_data_files("tkinterdnd2")
except Exception:
    pass

# Pillow 的插件与 reportlab 的字体资源
for pkg in ("PIL", "reportlab.pdfbase", "pypdf._codecs"):
    try:
        hiddenimports += collect_submodules(pkg)
    except Exception:
        pass

# pypdf 的加密后端
hiddenimports += [
    "pypdf._crypt_providers._cryptography",
    "pypdf._crypt_providers._fallback",
    "pypdf._crypt_providers._pycryptodome",
    "cryptography",
    "cryptography.hazmat.backends.openssl",
]

if IS_WIN:
    # Windows 下 tkinterdnd2 依赖 tkdnd2.9.5 目录
    hiddenimports += ["tkinterdnd2.TkinterDnD"]

    # ----------------------------------------------------------------
    # pywin32（PDF 批量打印用）
    # ----------------------------------------------------------------
    # 为什么必须显式处理：win32print / win32ui / win32con 是 .pyd 扩展模块，
    # 且它们运行时依赖 pywin32_system32/ 下的 pythoncom3xx.dll、pywintypes3xx.dll。
    # 这两个 DLL 不在包内、也不在系统 PATH 上，PyInstaller 静态分析抓不到，
    # 漏了就会在用户机上抛 "ImportError: DLL load failed"。
    hiddenimports += [
        "win32print", "win32ui", "win32con", "win32gui", "win32api",
        "win32com", "win32com.client", "pythoncom", "pywintypes",
    ]

    # pywin32_system32 是无 __init__.py 的 namespace package，.__file__ 为 None，
    # 只能从 .__path__ 取目录。DLL 放到目标目录 pywin32_system32/ 下与源结构一致。
    try:
        import pywin32_system32  # noqa: F401

        _pw_dirs = list(getattr(pywin32_system32, "__path__", []) or [])
        _pw_count = 0
        for _d in _pw_dirs:
            if not os.path.isdir(_d):
                continue
            for _f in os.listdir(_d):
                if _f.lower().endswith(".dll"):
                    binaries.append((os.path.join(_d, _f), "pywin32_system32"))
                    _pw_count += 1
        print(f"[spec] pywin32 运行时 DLL 已收集：{_pw_count} 个")
    except Exception as _exc:
        print(f"[spec] 警告：pywin32 DLL 收集失败（PDF 打印功能可能不可用）：{_exc}")

# 自带外部工具（可选）：把 Poppler / Ghostscript 绿色版放到 vendor/ 即可一起分发
vendor = os.path.join(PROJECT_DIR, "vendor")
if os.path.isdir(vendor):
    datas.append((vendor, "vendor"))
    print(f"[spec] 检测到 vendor 目录，将一并打包：{vendor}")

# 图标（存在才用，避免因缺图标而打包失败）
icon_path = None
for cand in (os.path.join(PROJECT_DIR, "assets", "app.icns" if IS_MAC else "app.ico"),
             os.path.join(PROJECT_DIR, "assets", "app.ico")):
    if os.path.isfile(cand):
        icon_path = cand
        break

# ------------------------------------------------------------------ 排除无关大包
excludes = [
    "matplotlib", "numpy", "scipy", "pandas", "IPython", "jupyter",
    "pytest", "setuptools", "pip", "wheel",
    "PyQt5", "PyQt6", "PySide2", "PySide6", "wx",
    "tkinter.test", "test", "unittest",
]

# ------------------------------------------------------------------ Analysis
a = Analysis(  # noqa: F821  (PyInstaller 在运行时注入)
    [ENTRY],
    pathex=[PROJECT_DIR],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[SPEC_DIR],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)  # noqa: F821

# ------------------------------------------------------------------ EXE / APP
if IS_MAC:
    exe = EXE(  # noqa: F821
        pyz, a.scripts, [],
        exclude_binaries=True,
        name=APP_NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,               # macOS 上 UPX 会破坏签名
        console=False,           # 触发 .app 打包
        disable_windowed_traceback=False,
        argv_emulation=False,
        # 架构：None=本机架构；可用环境变量 ODA_TARGET_ARCH=universal2/arm64/x86_64 覆盖
        target_arch=os.environ.get("ODA_TARGET_ARCH") or None,
        codesign_identity=os.environ.get("ODA_CODESIGN_ID") or None,  # 正式分发时填 "Developer ID Application: xxx"
        entitlements_file=None,  # 需要沙盒/网络权限时指向 entitlements.plist
        icon=icon_path,
    )

    coll = COLLECT(  # noqa: F821
        exe, a.binaries, a.datas,
        strip=False, upx=False, name=APP_NAME,
    )

    app = BUNDLE(  # noqa: F821
        coll,
        name=f"{APP_NAME}.app",
        icon=icon_path,
        bundle_identifier="com.officedocassistant.desktop",
        version=VERSION,
        info_plist={
            "CFBundleName": APP_NAME,
            "CFBundleDisplayName": DISPLAY_NAME,
            "CFBundleShortVersionString": VERSION,
            "CFBundleVersion": VERSION,
            "NSHumanReadableCopyright": f"{DISPLAY_NAME} v{VERSION}",
            # 关键：不声明这个，Retina 屏上界面是模糊的
            "NSHighResolutionCapable": True,
            # 需求要求支持 macOS 12 及以上
            "LSMinimumSystemVersion": "12.0",
            # 允许读写用户选择的任意文件（打开/保存面板）
            "NSDocumentsFolderUsageDescription": "用于读取和保存你需要处理的文档",
            "NSDesktopFolderUsageDescription": "用于读取和保存桌面上的文档",
            "NSDownloadsFolderUsageDescription": "用于读取和保存下载文件夹中的文档",
            # 隐藏 Dock 里的默认菜单项
            "LSApplicationCategoryType": "public.app-category.productivity",
        },
    )
else:
    exe = EXE(  # noqa: F821
        pyz, a.scripts, [],
        exclude_binaries=True,
        name=APP_NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,               # UPX 会显著提高杀软误报率，这里主动关闭
        console=False,
        disable_windowed_traceback=False,
        icon=icon_path,
        # 版本信息：exe 属性页的"产品名/版本/描述"，也降低杀软误报概率。
        # 由 _write_version_info() 按 VERSION 现场生成，避免"字符串改了、ffi 元组没改"。
        version=_write_version_info() if IS_WIN else None,
        # 长路径 + 高 DPI + UTF-8 代码页都在这个 manifest 里声明
        manifest=os.path.join(SPEC_DIR, "app.manifest") if IS_WIN else None,
    )

    coll = COLLECT(  # noqa: F821
        exe, a.binaries, a.datas,
        strip=False, upx=False, name=APP_NAME,
    )

print("[spec] 打包配置就绪。产物将出现在 dist/ 目录。")
print(f"[spec] 提示：把 Poppler / Ghostscript 解压到 {os.path.join(PROJECT_DIR, 'vendor')} "
      f"后重新打包，即可把外部依赖一起分发。")
