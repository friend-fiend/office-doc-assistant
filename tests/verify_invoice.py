"""发票整理功能验证脚本（只读解析 + 在临时目录产出汇总 xlsx）。

用法::

    python tests/verify_invoice.py            # 解析全部样票并打印命中率
    python tests/verify_invoice.py --save     # 额外生成汇总 xlsx 到 %TEMP%

不读写用户桌面（样票路径只读），产物只进临时目录。
"""
from __future__ import annotations

import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.cancel import CancelToken  # noqa: E402
from app.core.invoice_ops import (  # noqa: E402
    INVOICE_SUMMARY_OP,
    Attachment,
    assign_projects,
    parse_document,
)
from app.core.registry import OpContext  # noqa: E402
from app.engine import TaskReport  # noqa: E402

SRC = next((a for a in sys.argv[1:] if not a.startswith("-")),
           r"C:\Users\Administrator\Desktop\测试\发票测试\票据范例")


class StubBus:
    def log(self, *a, **k):
        pass

    def progress(self, *a, **k):
        pass

    def ask_password(self, *a, **k):
        return None


def make_ctx(out_dir: str) -> OpContext:
    os.makedirs(out_dir, exist_ok=True)
    return OpContext(StubBus(), CancelToken(), INVOICE_SUMMARY_OP, out_dir,
                     TaskReport(op_title=INVOICE_SUMMARY_OP.title))


def main() -> int:
    files = sorted(glob.glob(os.path.join(SRC, "**", "*.pdf"), recursive=True))
    projects = assign_projects(files)
    print(f"样票 {len(files)} 张，归属示例：{sorted(set(projects.values()))}")

    ok = bad = atts = 0
    hit = {"号码": 0, "价税合计": 0, "类别": 0, "行程": 0, "销售方": 0, "乘车日期": 0}
    issues: list[str] = []
    for f in files:
        r = parse_document(f, projects[f])
        from app.core.invoice_ops import Unreadable
        if isinstance(r, Attachment):
            atts += 1
            continue
        if isinstance(r, Unreadable):
            bad += 1
            issues.append(f"[不可读] {os.path.basename(f)}：{r.reason}")
            continue
        ok += 1
        if r.invoice_no:
            hit["号码"] += 1
        if r.total is not None:
            hit["价税合计"] += 1
        if r.kind:
            hit["类别"] += 1
        if r.trips or r.kind in ("住宿", "餐饮"):
            hit["行程"] += 1
        if r.seller:
            hit["销售方"] += 1
        if r.use_date:
            hit["乘车日期"] += 1
        if r.notes and any("请人工" in n or "未读" in n or "校验" in n for n in r.notes):
            issues.append(f"[注意] {os.path.basename(f)}：{'；'.join(r.notes)}")

    print(f"\n解析成功 {ok} 张发票 / 附件 {atts} 份 / 失败 {bad}")
    for k, v in hit.items():
        print(f"  {k}: {v}/{ok}")
    if issues:
        print("\n需关注的票：")
        for s in issues:
            print("  ", s)

    if "--save" in sys.argv:
        import tempfile
        ctx = make_ctx(os.path.join(tempfile.gettempdir(), "oda_invoice_verify"))
        ctx.set_total(len(files))
        from app.core.invoice_ops import run_invoice_summary
        run_invoice_summary(ctx, files, {
            "project_name": "", "project_level": "1", "dedupe": True,
            "check_amount": True, "split_by_project": False,
        })
        print("\n汇总 xlsx：", ctx.report.outputs[0])
    return 0


if __name__ == "__main__":
    sys.exit(main())
