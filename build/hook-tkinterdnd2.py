"""PyInstaller hook：收集 tkinterdnd2 的 tkdnd 二进制。

tkinterdnd2 把 tkdnd 的 Tcl 扩展（.dll / .dylib / .so）放在包目录下的
tkdnd 子目录里，PyInstaller 的默认分析只收 .py，运行时会报：

    TclError: can't find package tkdnd

把本文件放到项目根目录或 spec 的 hookspath 里，PyInstaller 会自动加载。
"""
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

datas = collect_data_files("tkinterdnd2")
binaries = collect_dynamic_libs("tkinterdnd2")
