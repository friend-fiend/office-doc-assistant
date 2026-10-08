"""修复被旧版本改坏的系统打印机默认设置。

背景
----
1.0.1 及更早的版本在每次打印时都会调用 ``win32print.SetPrinter`` 把 DEVMODE
写回打印机 —— 那是**持久**修改系统默认设置，会连带影响所有其它软件的打印
（比如默认纸张被改成了 A3 横向）。1.0.2 已经彻底不再这么做，但已经被改过的
机器需要显式还原一次。

用法
----
    python tools/fix_printer_defaults.py                 # 只看现状，不改任何东西
    python tools/fix_printer_defaults.py --restore 震旦AD289S
    python tools/fix_printer_defaults.py --restore-all

默认恢复成：A4 / 纵向 / 1 份 / 彩色 / 单面（办公文档的常见默认）。

本脚本是**唯一**会写系统打印机状态的入口，日常打印路径绝不调用它。
"""
from __future__ import annotations

import os
import sys

THIS = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(THIS)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

ORIENT = {1: "纵向", 2: "横向"}


def main(argv: list[str]) -> int:
    from app.core import print_ops as po

    if not po.PYWIN32_OK:
        print("仅 Windows 且 pywin32 可用时才能修复打印机默认设置。")
        return 2

    print("=" * 68)
    print("打印机默认设置诊断 / 修复")
    print("=" * 68)

    real = po.list_printers(real_only=True)
    virtual = po.list_virtual_pdf_printers()
    print(f"真实打印机 {len(real)} 台：{real}")
    print(f"虚拟打印设备 {len(virtual)} 台（程序不会向它们打印）：{virtual}")
    print()

    base = po._enum_printers()
    for it in base:
        name = it["name"]
        if it["virtual"]:
            continue
        info = po._printer_default_summary(name)
        if not info:
            print(f"  · {name}：无法读取默认设置")
            continue
        paper_rev = {v: k for k, v in po.PAPER_IDS.items()}
        paper = paper_rev.get(info.get("paper", 0), f"ID{info.get('paper')}")
        print(f"  · {name}")
        print(f"      默认纸张 {paper.upper()} / {ORIENT.get(info.get('orientation'), '?')}"
              f" / {info.get('copies', 1)} 份")

    restore = None
    for i, a in enumerate(argv):
        if a == "--restore" and i + 1 < len(argv):
            restore = [argv[i + 1]]
        elif a == "--restore-all":
            restore = list(real)

    if not restore:
        print("\n只做诊断，未修改任何设置。")
        print("如默认纸张/方向/份数不对（例如被老版本改成了 A3 横向），可执行：")
        print("    python tools/fix_printer_defaults.py --restore \"<打印机名>\"")
        return 0

    print("\n即将恢复为：A4 / 纵向 / 1 份 / 彩色 / 单面")
    for name in restore:
        if name not in real:
            print(f"  跳过 {name!r}：不在真实打印机列表中")
            continue
        try:
            result = po.restore_printer_defaults(name)
            print(f"  ✓ {result}")
        except Exception as exc:
            print(f"  ✗ {name}：{type(exc).__name__}: {exc}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
