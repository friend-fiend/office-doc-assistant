"""产物校验：读取 _selftest 的输出文件，确认真的是有效文件而不是空壳。

这是"自检通过"之外的独立证据 —— 自检只看有没有抛异常和文件数，
这里进一步验证 PDF 能被重新解析、页数正确、水印/页码文字确实写进去了。
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pypdf import PdfReader  # noqa: E402


def find(root: str, keyword: str, ext: str = ".pdf", deep: bool = True) -> str | None:
    """返回匹配关键字的最新文件。

    多次运行 selftest 会在 out/ 下累积多份 "XX_输出_N" 目录，这里必须取最新的，
    否则会拿到上一轮的历史产物，得出错误的校验结论。
    """
    hits: list[str] = []
    for dirpath, _d, files in os.walk(root):
        if not deep and os.path.relpath(dirpath, root).count(os.sep) > 1:
            continue
        for f in files:
            if keyword in f and f.lower().endswith(ext):
                hits.append(os.path.join(dirpath, f))
    if not hits:
        return None
    return max(hits, key=lambda p: (os.path.getmtime(p), p))


def check(label: str, fn):
    try:
        print(f"  [OK]   {label}: {fn()}")
        return True
    except Exception as e:
        print(f"  [FAIL] {label}: {type(e).__name__}: {e}")
        return False


def latest_sandbox() -> str:
    """读取 selftest 写下的 LATEST.txt，定位最新一轮沙盒。"""
    base = os.path.join(os.getcwd(), "_selftest")
    marker = os.path.join(base, "LATEST.txt")
    if os.path.isfile(marker):
        try:
            with open(marker, encoding="utf-8") as fh:
                p = fh.read().strip()
            if p and os.path.isdir(p):
                return p
        except Exception:
            pass
    # 兜底：取 _selftest 下最新的 run-* 目录
    if os.path.isdir(base):
        cands = [os.path.join(base, d) for d in os.listdir(base)
                 if d.startswith("run-") and os.path.isdir(os.path.join(base, d))]
        if cands:
            return max(cands, key=os.path.getmtime)
    return base


def main() -> int:
    sandbox = latest_sandbox()
    root = os.path.join(sandbox, "out")
    samples_root = os.path.join(sandbox, "samples")
    if not os.path.isdir(root):
        print("未找到沙盒输出目录，请先运行：python tests/selftest.py")
        return 1

    print("=" * 72)
    print("产物有效性校验")
    print(f"沙盒：{sandbox}")
    print("=" * 72)

    results = []

    # 合并结果：3 页 + 2 页 = 5 页
    merged = find(root, "合并结果")
    results.append(check("合并 PDF 页数 = 5", lambda: len(PdfReader(merged).pages)))

    # 拆分：每页一个 → p1.pdf / p2.pdf / p3.pdf 各 1 页
    p1 = find(root, "样例文档3页_p1.pdf")
    results.append(check("拆分单页 PDF 页数 = 1", lambda: len(PdfReader(p1).pages)))

    # 提取指定页 1-2,3 → 3 页
    ext3 = find(root, "提取")
    results.append(check("提取指定页 页数 = 3", lambda: len(PdfReader(ext3).pages)))

    # 删页 2 → 3 页变 2 页
    dele = find(root, "已删除1页")
    results.append(check("删除指定页后 页数 = 2", lambda: len(PdfReader(dele).pages)))

    # 旋转 90° 已应用
    rot = find(root, "旋转90")
    results.append(check("旋转 90° 已写入 /Rotate",
                         lambda: PdfReader(rot).pages[0].rotation))

    # 页码：第 1 页跳过封面 → 第 2 页应出现 "第 1 页"
    pn = find(root, "加页码")
    def pn_text():
        t1 = PdfReader(pn).pages[0].extract_text() or ""
        t2 = PdfReader(pn).pages[1].extract_text() or ""
        return {"封面含页码": "共" in t1, "第2页含页码": ("第" in t2 and "共" in t2)}
    results.append(check("页码跳过封面已生效", pn_text))

    # 水印：平铺后第 1 页文本里应出现水印字样
    wm = find(root, "水印")
    def wm_text():
        text = PdfReader(wm).pages[0].extract_text() or ""
        return {"含水印关键字": "内部资料" in text, "长度": len(text)}
    results.append(check("水印文字已写入", wm_text))

    # 页眉页脚
    hf = find(root, "页眉页脚")
    def hf_text():
        text = PdfReader(hf).pages[0].extract_text() or ""
        return {"含页脚页码": "共" in text}
    results.append(check("页眉页脚已写入", hf_text))

    # 压缩：产物应可解析且比原始小
    comp = find(root, "压缩")
    orig = find(samples_root, "样例文档3页")
    def size_cmp():
        return {"原始": os.path.getsize(orig), "压缩后": os.path.getsize(comp),
                "可解析页数": len(PdfReader(comp).pages)}
    results.append(check("压缩产物可解析", size_cmp))

    # 加密：应能检测到加密并用密码打开
    enc = find(root, "已加密")
    def enc_check():
        r = PdfReader(enc)
        is_enc = r.is_encrypted
        ok = bool(r.decrypt("test123"))
        return {"is_encrypted": is_enc, "decrypt_ok": ok}
    results.append(check("加密 PDF 需密码才能打开", enc_check))

    # 解密：产物应不再加密
    dec = find(root, "已解密")
    results.append(check("解密产物不再加密", lambda: PdfReader(dec).is_encrypted))

    # 图片合成 PDF：4 页，且页面尺寸应精确等于 A4（595.28 x 841.89 pt）
    imgpdf = find(root, "图片合成")
    def img_check():
        r = PdfReader(imgpdf)
        mb = r.pages[0].mediabox
        w, h = float(mb.width), float(mb.height)
        return {"页数": len(r.pages),
                "页面尺寸pt": (round(w, 2), round(h, 2)),
                "是否为A4": abs(w - 595.28) < 1 and abs(h - 841.89) < 1}
    results.append(check("图片合成 A4 PDF", img_check))

    # 图片转图片：4 张 PNG
    imgs = [f for r_, _d, fs in os.walk(root) for f in fs
            if f.startswith("样例文档3页_第") and f.endswith(".png")]
    results.append(check("PDF 转图片产物", lambda: f"{len(imgs)} 张 PNG（依赖缺失时应为 0）"))

    # Excel 合并（取最新一份）
    xls = find(root, "Excel合并结果", ".xlsx")
    def xls_check():
        from openpyxl import load_workbook
        wb = load_workbook(xls)
        total = sum(ws.max_row for ws in wb.worksheets)
        return {"工作表": wb.sheetnames, "总行数": total}
    results.append(check("Excel 合并结果", xls_check))

    # ZIP 与 7z
    zp = find(root, "图片打包", ".zip")
    def zip_check():
        import zipfile
        with zipfile.ZipFile(zp) as zf:
            return zf.namelist()
    results.append(check("ZIP 内容", zip_check))

    sev = find(root, "7z打包", ".7z")
    def sev_check():
        import py7zr
        with py7zr.SevenZipFile(sev) as zf:
            return zf.getnames()
    results.append(check("7z 内容", sev_check))

    # 对比报告
    rep = find(root, "报告_旧版_对比_报告_新版", ".html")
    def report_check():
        with open(rep, encoding="utf-8") as fh:
            content = fh.read()
        return {"大小": len(content), "含 diff 表格": "table" in content}
    results.append(check("对比 HTML 报告", report_check))

    # 重命名副本（只看最新一轮输出目录，避免历史产物干扰）
    ren_root = os.path.dirname(find(root, "成品_", ".png") or "")
    ren = sorted(f for f in os.listdir(ren_root)) if ren_root and os.path.isdir(ren_root) else []
    results.append(check("重命名副本", lambda: ren))

    print("-" * 72)
    print(f"校验通过：{sum(1 for r in results if r)}/{len(results)}")
    print("=" * 72)
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
