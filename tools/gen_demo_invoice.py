"""用真实样票生成一张演示汇总表（开发自测用，只读样票）。

用法::

    python tools/gen_demo_invoice.py 输出.xlsx
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.cancel import CancelToken
from app.core.invoice_ops import INVOICE_SUMMARY_OP, run_invoice_summary
from app.core.registry import OpContext
from app.engine import TaskReport

SRC = r"C:\Users\Administrator\Desktop\测试\发票测试\票据范例"


class StubBus:
    def log(self, *a, **k):
        pass

    def progress(self, *a, **k):
        pass

    def ask_password(self, *a, **k):
        return None


def main() -> int:
    out = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "demo.xlsx")
    files = sorted(
        os.path.join(dp, f)
        for dp, _, fs in os.walk(SRC)
        for f in fs if f.lower().endswith(".pdf")
    )
    ctx = OpContext(StubBus(), CancelToken(), INVOICE_SUMMARY_OP,
                    os.path.dirname(out), TaskReport(op_title=INVOICE_SUMMARY_OP.title))
    ctx.set_total(len(files))
    run_invoice_summary(ctx, files, {
        "project_name": "", "project_level": "1",
        "dedupe": True, "check_amount": True, "split_by_project": False,
    })
    print(f"输入 {len(files)} 张 -> {ctx.report.outputs[0]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
