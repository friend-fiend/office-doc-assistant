"""核心功能自检：不依赖 Tkinter，直接驱动任务引擎跑通全部操作。

用法：
    python tests/selftest.py            # 在当前目录下建 _selftest 沙盒运行
    python tests/selftest.py --keep     # 保留沙盒便于人工查看产物

输出一份通过/失败清单，作为"方案可行"的证据。
"""
from __future__ import annotations

import argparse
import io
import os
import shutil
import sys
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.core import OPS  # noqa: E402
from app.engine import TaskRequest, TaskRunner  # noqa: E402
from app.errors import Cancelled  # noqa: E402
from app.cancel import CancelToken  # noqa: E402


class StubBus:
    """替代 GUI 的消息总线：把日志收进内存。"""

    def __init__(self, verbose: bool = False):
        self.logs: list[tuple[str, str]] = []
        self.progress_events = 0
        self.report = None
        self.verbose = verbose
        self.passwords: dict[str, str] = {}

    def log(self, text: str, level: str = "info") -> None:
        self.logs.append((level, text))
        if self.verbose:
            print(f"      [{level}] {text}")

    def progress(self, done: int, total: int, label: str = "") -> None:
        self.progress_events += 1

    def state(self, text: str) -> None:
        pass

    def done(self, report) -> None:
        self.report = report

    def ask_password(self, path: str, confirm: bool = False) -> str | None:
        return self.passwords.get(os.path.basename(path))


def make_samples(root: str) -> dict[str, list[str]]:
    """生成测试用样例文件。"""
    from PIL import Image
    from pypdf import PdfReader
    from reportlab.pdfgen import canvas

    src = os.path.join(root, "samples")
    os.makedirs(src, exist_ok=True)
    samples: dict[str, list[str]] = {}

    # --- 中文 PDF（3 页，带中文与图形，用于叠加类操作）
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))

    pdfs = []
    for n in (3, 2):
        path = os.path.join(src, f"样例文档{n}页.pdf")
        c = canvas.Canvas(path)
        for i in range(n):
            c.setFont("STSong-Light", 16)
            c.drawString(72, 780, f"这是第 {i + 1} 页的中文标题")
            c.setFont("Helvetica", 10)
            c.drawString(72, 760, f"Page {i + 1} - " + "content " * 40)
            c.rect(72, 600, 200, 80, fill=0)
            c.showPage()
        c.save()
        pdfs.append(path)
    samples["pdf"] = pdfs

    # --- 图片（用于 图片转 PDF）
    imgs = []
    for i, color in enumerate([(220, 60, 60), (60, 160, 90), (60, 90, 220), (230, 180, 40)]):
        p = os.path.join(src, f"图片{i + 1}.png")
        Image.new("RGB", (600, 400), color).save(p)
        imgs.append(p)
    samples["image"] = imgs

    # --- Word / Excel
    from docx import Document
    from openpyxl import Workbook

    d1 = os.path.join(src, "报告_旧版.docx")
    doc = Document()
    doc.add_heading("季度报告", 0)
    for line in ["第一段：销售情况稳定。", "第二段：成本略有上升。", "第三段：下季度目标 100 万。"]:
        doc.add_paragraph(line)
    doc.save(d1)

    d2 = os.path.join(src, "报告_新版.docx")
    doc = Document()
    doc.add_heading("季度报告", 0)
    for line in ["第一段：销售情况稳定增长。", "第二段：成本略有上升。", "第四段：新增海外渠道。"]:
        doc.add_paragraph(line)
    doc.save(d2)
    samples["document"] = [d1, d2]

    xlsx = []
    for i in (1, 2):
        p = os.path.join(src, f"台账{i}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = f"明细{i}"
        ws.append(["日期", "项目", "金额"])
        for r in range(1, 11):
            ws.append([f"2026-0{i}-{r:02d}", f"项目{r}", r * 100])
        wb.create_sheet(f"汇总{i}").append(["合计", 5500])
        wb.save(p)
        xlsx.append(p)
    samples["excel"] = xlsx

    # --- 压缩包
    import zipfile

    zp = os.path.join(src, "素材.zip")
    with zipfile.ZipFile(zp, "w") as zf:
        for p in imgs[:2]:
            zf.write(p, os.path.basename(p))
    samples["archive"] = [zp]

    # --- 加密 PDF（用于解密测试）
    from pypdf import PdfWriter

    enc = os.path.join(src, "加密文档.pdf")
    r = PdfReader(pdfs[0])
    w = PdfWriter()
    w.append(r)
    w.encrypt(user_password="secret", owner_password="secret", algorithm="RC4-128")
    with open(enc, "wb") as fh:
        w.write(fh)
    samples["encrypted"] = [enc]

    return samples


def run_case(name: str, op_key: str, files: list[str], params: dict, out_root: str,
             bus: StubBus, expect_ok: bool = True) -> bool:
    print(f"\n▶ {name}  ({op_key})")
    print(f"  输入 {len(files)} 个文件")
    bus.logs.clear()
    bus.report = None
    bus.progress_events = 0

    req = TaskRequest(op_key=op_key, files=files, params=params, output_parent=out_root)
    runner = TaskRunner(req, bus, CancelToken())
    t0 = time.time()
    runner.run()
    dt = time.time() - t0

    rep = bus.report
    if rep is None:
        print("  ✗ 未收到结果报告")
        return False

    status = "✓" if (rep.ok > 0 and rep.failed == 0) or (not expect_ok and rep.failed > 0) else "?"
    ok = (rep.ok > 0) if expect_ok else (rep.failed > 0 or rep.ok == 0)

    print(f"  {status} {rep.summary_text()}  用时 {dt:.2f}s")
    for k, v in rep.metrics.items():
        print(f"      · {k}: {v}")
    if rep.outputs:
        shown = [o for o in rep.outputs[:3]]
        print(f"      · 输出示例: " + "; ".join(os.path.basename(o) for o in shown))
    for f in rep.failures[:4]:
        print(f"      ! [{f.category}] {os.path.basename(f.path)}: {f.reason}")

    # 校验输出目录里确实有文件
    produced = _count_outputs(rep.output_dir)
    print(f"      · 输出目录文件数: {produced}")
    if expect_ok and produced == 0:
        print("      ✗ 输出目录为空")
        return False
    return ok


def _count_outputs(path: str) -> int:
    if not path or not os.path.isdir(path):
        return 0
    n = 0
    for root, _d, files in os.walk(path):
        n += len(files)
    return n


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep", action="store_true", help="保留沙盒目录")
    ap.add_argument("-v", "--verbose", action="store_true", help="打印全部日志")
    args = ap.parse_args()

    # 每次运行都开一个独立的时间戳沙盒，并把路径写入 LATEST.txt。
    # 这样多次运行不会互相污染（历史产物曾导致产物校验误判），
    # verify_artifacts.py 也能准确找到"这一轮"的输出。
    base = os.path.join(os.getcwd(), "_selftest")
    stamp = time.strftime("run-%Y%m%d-%H%M%S")
    root = os.path.join(base, stamp)
    os.makedirs(root, exist_ok=True)
    try:
        with open(os.path.join(base, "LATEST.txt"), "w", encoding="utf-8") as fh:
            fh.write(root)
    except Exception:
        pass

    print("=" * 72)
    print("办公文档一键处理助手 —— 核心功能自检")
    print("=" * 72)
    print(f"沙盒目录: {root}")
    print(f"已注册操作: {len(OPS)} 个  ->  {', '.join(OPS.keys())}")

    samples = make_samples(root)
    bus = StubBus(verbose=args.verbose)
    out = os.path.join(root, "out")
    os.makedirs(out, exist_ok=True)

    results: list[tuple[str, bool]] = []

    def case(name, op_key, files, params, expect_ok=True, passwords=None):
        if passwords:
            bus.passwords = passwords
        try:
            r = run_case(name, op_key, files, params, out, bus, expect_ok)
        except Exception as exc:
            print(f"  ✗ 抛出异常 {type(exc).__name__}: {exc}")
            traceback.print_exc()
            r = False
        results.append((name, r))

    # ---------- PDF 页面操作
    case("PDF 合并", "pdf_merge", samples["pdf"],
         {"out_name": "合并结果", "add_bookmarks": True})
    case("PDF 拆分（每页一个）", "pdf_split", samples["pdf"][:1],
         {"split_mode": "single"})
    case("PDF 拆分（每 N 页）", "pdf_split", samples["pdf"][:1],
         {"split_mode": "every", "every_n": 2})
    case("PDF 拆分（按范围）", "pdf_split", samples["pdf"][:1],
         {"split_mode": "ranges", "ranges": "1-2,3"})
    case("提取指定页 1-2,3", "pdf_extract", samples["pdf"][:1], {"pages": "1-2,3"})
    case("删除指定页", "pdf_delete_pages", samples["pdf"][:1], {"pages": "2"})
    case("PDF 旋转 90°", "pdf_rotate", samples["pdf"][:1], {"angle": "90"})
    case("PDF 旋转 180°（指定页）", "pdf_rotate", samples["pdf"][:1],
         {"angle": "180", "pages": "1"})

    # ---------- 叠加类
    case("PDF 加页码（中文格式+跳过封面）", "pdf_page_number", samples["pdf"][:1],
         {"pnum_format": "第 {n} 页 / 共 {total} 页", "position": "bottom-center",
          "start_at": 1, "skip_first": 1, "font_size": 10, "margin_mm": 12,
          "font_style": "cjk"})
    case("PDF 水印（平铺+透明）", "pdf_watermark", samples["pdf"][:1],
         {"wm_text": "内部资料 请勿外传", "wm_size": 36, "wm_opacity": 0.18,
          "wm_angle": "45", "wm_color": "red", "wm_tile": True, "wm_pages": "",
          "wm_behind": False})
    case("PDF 页眉页脚", "pdf_header_footer", samples["pdf"][:1],
         {"header_text": "{filename}", "footer_text": "第 {n} 页 / 共 {total} 页",
          "hf_align": "center", "hf_size": 9, "hf_margin": 10})

    # ---------- 压缩 / 加密
    case("PDF 压缩（轻度）", "pdf_compress", samples["pdf"][:1],
         {"level": "light", "strip_metadata": True, "keep_if_bigger": True})
    case("PDF 压缩（中度，无 GS 时降级）", "pdf_compress", samples["pdf"][:1],
         {"level": "medium", "strip_metadata": True, "keep_if_bigger": True})
    case("PDF 加密", "pdf_encrypt", samples["pdf"][:1],
         {"user_pwd": "test123", "user_pwd2": "test123", "owner_pwd": "",
          "allow_print": True, "allow_copy": False, "algorithm": "auto"})
    case("PDF 解密", "pdf_decrypt", samples["encrypted"],
         {"password": "secret"}, passwords={"加密文档.pdf": "secret"})

    # ---------- 格式转换
    case("图片转 PDF（A4 居中）", "image_to_pdf", samples["image"],
         {"out_name": "图片合成", "page_size": "a4", "margin_mm": 10,
          "quality": 90, "sort_natural": True})
    case("图片转 PDF（跟随原图）", "image_to_pdf", samples["image"],
         {"out_name": "原尺寸合成", "page_size": "original", "margin_mm": 0,
          "quality": 90, "sort_natural": True})
    case("PDF 转图片（需 Poppler，缺依赖时应报错）", "pdf_to_image", samples["pdf"][:1],
         {"img_format": "png", "dpi": "150", "pages": "", "each_subfolder": True},
         expect_ok=False)
    case("Office 转 PDF（需 LibreOffice，缺依赖时应报错）", "office_to_pdf",
         samples["document"], {"calc_pages": 0, "keep_subdir": False}, expect_ok=False)

    # ---------- 批量文件
    case("Excel 合并（每文件一表）", "excel_merge", samples["excel"],
         {"merge_mode": "per_sheet", "values_only": False, "skip_hidden_sheets": True})
    case("Excel 合并（追加到同一表）", "excel_merge", samples["excel"],
         {"merge_mode": "append", "sheet_name": "汇总", "values_only": False,
          "skip_hidden_sheets": True})
    case("Excel 拆分（保留公式与样式）", "excel_split", samples["excel"][:1],
         {"each_subfolder": True, "preserve_format": True, "skip_hidden_sheets": True})
    case("Excel 拆分（精简模式）", "excel_split", samples["excel"][:1],
         {"each_subfolder": True, "preserve_format": False, "skip_hidden_sheets": True})
    case("批量重命名（副本模式）", "batch_rename", samples["image"],
         {"name_mode": "original_seq", "prefix": "成品_", "suffix": "", "seq_start": 1,
          "seq_digits": 2, "seq_pos": "before_ext", "seq_sep": "_", "find_text": "",
          "replace_text": "", "use_regex": False, "ext_case": "keep",
          "apply_mode": "copy"})
    case("批量重命名（查找替换）", "batch_rename", samples["image"][:2],
         {"name_mode": "find_replace", "prefix": "", "suffix": "_v2", "seq_start": 1,
          "seq_digits": 2, "seq_pos": "before_ext", "seq_sep": "_", "find_text": "图片",
          "replace_text": "素材", "use_regex": False, "ext_case": "keep",
          "apply_mode": "copy"})

    # ---------- 压缩包
    case("压缩为 ZIP", "archive_compress", samples["image"],
         {"format": "zip", "out_name": "图片打包", "level": "normal", "password": ""})
    case("压缩为 7z", "archive_compress", samples["image"][:2],
         {"format": "7z", "out_name": "7z打包", "level": "max", "password": ""})
    case("解压 ZIP", "archive_extract", samples["archive"], {"password": ""})

    # ---------- 高级
    case("文档对比（Word 两版）", "doc_compare", samples["document"],
         {"ignore_whitespace": True, "ignore_case": False, "context_lines": 2,
          "make_html": True})
    case("文档对比（PDF vs PDF）", "doc_compare", samples["pdf"],
         {"ignore_whitespace": True, "ignore_case": False, "context_lines": 2,
          "make_html": True})

    # ---------- 取消机制
    print("\n▶ 取消机制（在合并过程中取消）")
    bus.logs.clear()
    token = CancelToken()
    token.cancel()  # 立即取消
    req = TaskRequest(op_key="pdf_merge", files=samples["pdf"],
                      params={"out_name": "x", "add_bookmarks": False}, output_parent=out)
    TaskRunner(req, bus, token).run()
    cancelled_ok = bus.report is not None and bus.report.cancelled
    print(f"  {'✓' if cancelled_ok else '✗'} 任务被正确标记为已取消")
    results.append(("取消机制", cancelled_ok))

    # ---------- 损坏文件容错
    print("\n▶ 损坏文件容错（应「失败但任务继续」，不崩溃）")
    broken = os.path.join(root, "samples", "损坏文件.pdf")
    with open(broken, "wb") as fh:
        fh.write(b"%PDF-1.4\n this is not a real pdf \x00\x01\x02")
    bus.logs.clear()
    req = TaskRequest(op_key="pdf_rotate", files=[broken, samples["pdf"][1]],
                      params={"angle": "90", "pages": ""}, output_parent=out)
    TaskRunner(req, bus, CancelToken()).run()
    rep = bus.report
    survived = rep is not None and rep.failed >= 1 and rep.ok >= 1
    print(f"  {'✓' if survived else '✗'} 失败 {rep.failed if rep else '?'} 个，"
          f"成功 {rep.ok if rep else '?'} 个（好文件未被坏文件打断）")
    results.append(("损坏文件容错", survived))

    # ---------- 汇总
    print("\n" + "=" * 72)
    passed = sum(1 for _n, r in results if r)
    total = len(results)
    for n, r in results:
        print(f"  {'通过' if r else '未通过'}  {n}")
    print("-" * 72)
    print(f"总计：{passed}/{total} 通过")
    print(f"沙盒：{root}")
    print("=" * 72)

    if not args.keep:
        print("（如需保留产物用于人工检查，请加 --keep 重新运行）")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
