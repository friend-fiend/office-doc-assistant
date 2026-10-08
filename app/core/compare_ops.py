"""文档对比：比较两个版本的 PDF / Word / 纯文本差异。

定位是"两版差异提示"，不是精确排版比对。做法是抽取文本层后做行级 diff：
  · 相似度 = 1 - (差异行数 / 总行数)
  · 输出一份 .txt（unified diff，方便存档/邮件）和一份 .html（红绿高亮，方便阅读）

已知边界（会在结果里明确提示，不装作能处理）：
  扫描件/图片型 PDF 没有文本层，抽取结果为空 —— 报告会提示改用 OCR 工具，
  而不是给出一个"100% 不同"的误导性结论。
"""
from __future__ import annotations

import difflib
import html
import os

from ..errors import ParamError
from ..paths import long_path, sanitize_name, unique_file
from .pdf_pages import _open_with_password
from .registry import OpContext, OpDef, Param, register

COMPARE_OP = OpDef(
    key="doc_compare",
    title="文档对比",
    category="advanced",
    accepts=("document",),
    mode="pair",
    description=(
        "选择两个文件（旧版 / 新版）比较文本差异。\n"
        "支持 PDF、Word(.docx)、纯文本；输出差异报告，含相似度与逐行增删。"
    ),
    params=(
        Param("ignore_whitespace", "忽略空白与空行差异", "bool", default=True),
        Param("ignore_case", "忽略大小写", "bool", default=False),
        Param("context_lines", "上下文行数", "int", default=2, min_value=0, max_value=10),
        Param("make_html", "额外生成 HTML 高亮报告", "bool", default=True),
    ),
)


@register(COMPARE_OP)
def run_compare(ctx: OpContext, files: list[str], params: dict) -> None:
    if len(files) != 2:
        raise ParamError(f"文档对比需要且仅需要 2 个文件，当前选中 {len(files)} 个")

    old_path, new_path = files[0], files[1]
    ctx.set_total(4)

    ctx.log(f"旧版本：{os.path.basename(old_path)}")
    ctx.log(f"新版本：{os.path.basename(new_path)}")

    ctx.progress(1, 4, "读取旧版本")
    old_lines = _extract_lines(ctx, old_path)
    ctx.check()
    ctx.progress(2, 4, "读取新版本")
    new_lines = _extract_lines(ctx, new_path)
    ctx.check()

    if not old_lines and not new_lines:
        raise ParamError(
            "两个文件都没有可抽取的文本层（可能是扫描件或图片型 PDF）。"
            "本工具不含 OCR 功能，建议先用 OCR 工具转成可搜索 PDF 再对比。"
        )
    if not old_lines:
        ctx.warn(old_path, "未抽取到文本（可能是扫描件），差异结果将全部视为新增")
    if not new_lines:
        ctx.warn(new_path, "未抽取到文本（可能是扫描件），差异结果将全部视为删除")

    ignore_ws = bool(params.get("ignore_whitespace", True))
    ignore_case = bool(params.get("ignore_case", False))

    a = [_norm(x, ignore_ws, ignore_case) for x in old_lines]
    b = [_norm(x, ignore_ws, ignore_case) for x in new_lines]

    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    ratio = sm.ratio()
    added = removed = changed = 0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "insert":
            added += j2 - j1
        elif tag == "delete":
            removed += i2 - i1
        elif tag == "replace":
            changed += max(i2 - i1, j2 - j1)

    ctx.progress(3, 4, "生成差异报告")

    ctx_data = int(params.get("context_lines") or 2)
    diff_lines = list(difflib.unified_diff(
        old_lines, new_lines,
        fromfile=os.path.basename(old_path),
        tofile=os.path.basename(new_path),
        fromfiledate="", tofiledate="",
        n=ctx_data, lineterm="",
    ))

    base = f"{sanitize_name(os.path.splitext(os.path.basename(old_path))[0])}_对比_" \
           f"{sanitize_name(os.path.splitext(os.path.basename(new_path))[0])}"

    txt_path = unique_file(os.path.join(ctx.out_dir, base + ".txt"))
    with open(long_path(txt_path), "w", encoding="utf-8") as fh:
        fh.write(f"文档对比报告\n{'=' * 60}\n")
        fh.write(f"旧版本：{old_path}\n新版本：{new_path}\n")
        fh.write(f"旧版本行数：{len(old_lines)}    新版本行数：{len(new_lines)}\n")
        fh.write(f"相似度：{ratio * 100:.2f}%\n")
        fh.write(f"新增 {added} 行，删除 {removed} 行，修改 {changed} 行\n")
        if ignore_ws or ignore_case:
            fh.write(f"比对选项：忽略空白={ignore_ws}  忽略大小写={ignore_case}\n")
        fh.write(f"{'=' * 60}\n\n")
        if diff_lines:
            fh.write("\n".join(diff_lines))
        else:
            fh.write("两个文件的文本内容完全一致。\n")
    ctx.report.outputs.append(txt_path)

    if params.get("make_html", True):
        html_path = unique_file(os.path.join(ctx.out_dir, base + ".html"))
        _write_html_report(html_path, old_path, new_path, old_lines, new_lines,
                           ratio, added, removed, changed, ignore_ws, ignore_case)
        ctx.report.outputs.append(html_path)
        ctx.metric("HTML 报告", html_path)

    ctx.metric("相似度", f"{ratio * 100:.2f}%")
    ctx.metric("新增行", added)
    ctx.metric("删除行", removed)
    ctx.metric("修改行", changed)
    ctx.metric("差异报告", txt_path)

    ctx.progress(4, 4, "完成")
    ctx.ok(txt_path, src=old_path)
    if ratio >= 0.999:
        ctx.log("两个文档的文本内容基本一致。")
    else:
        ctx.log(f"共发现 {added + removed + changed} 处差异，详见报告。")


def _norm(line: str, ignore_ws: bool, ignore_case: bool) -> str:
    s = line
    if ignore_ws:
        s = " ".join(s.split())
    if ignore_case:
        s = s.lower()
    return s


def _extract_lines(ctx: OpContext, path: str) -> list[str]:
    """抽取文本，按行返回。抽取不到内容返回空列表（而不是抛异常）。"""
    lower = path.lower()
    try:
        if lower.endswith(".pdf"):
            return _extract_pdf(ctx, path)
        if lower.endswith(".docx"):
            return _extract_docx(path)
        if lower.endswith(".doc"):
            ctx.log("旧版 .doc 无法直接读取文本，建议先另存为 .docx；已尝试按二进制文本抽取。", "warn")
        with open(long_path(path), "r", encoding="utf-8", errors="replace") as fh:
            return [ln.rstrip("\n") for ln in fh]
    except Exception as exc:
        ctx.log(f"读取 {os.path.basename(path)} 失败：{exc}", "warn")
        return []


def _extract_pdf(ctx: OpContext, path: str) -> list[str]:
    reader = _open_with_password(ctx, path)
    lines: list[str] = []
    for i, page in enumerate(reader.pages):
        ctx.check()
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        if text:
            lines.extend(text.splitlines())
    return [ln for ln in lines if ln.strip()]


def _extract_docx(path: str) -> list[str]:
    from docx import Document

    doc = Document(long_path(path))
    lines = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                lines.append(" | ".join(cells))
    return [ln for ln in lines if ln.strip()]


def _write_html_report(path, old_path, new_path, old_lines, new_lines,
                       ratio, added, removed, changed, ignore_ws, ignore_case) -> None:
    rows = []
    diff = difflib.HtmlDiff(wrapcolumn=78)
    table = diff.make_table(old_lines, new_lines,
                            fromdesc=os.path.basename(old_path),
                            todesc=os.path.basename(new_path),
                            context=True, numlines=3)
    rows.append(table)

    doc = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>文档对比报告</title>
<style>
  body {{ font-family: -apple-system, "Microsoft YaHei", "PingFang SC", sans-serif;
         margin: 24px; background: #f7f8fa; color: #1f2329; }}
  .card {{ background:#fff; border:1px solid #e5e6eb; border-radius:10px;
           padding:18px 22px; margin-bottom:16px; }}
  h1 {{ font-size: 20px; margin: 0 0 12px; }}
  .kv {{ display:flex; flex-wrap:wrap; gap:24px; font-size:14px; color:#4e5969; }}
  .kv b {{ color:#1f2329; }}
  .good {{ color:#0f9d58; }} .bad {{ color:#d93026; }}
  table.diff {{ font-family: Consolas, Menlo, monospace; font-size:12.5px;
                border-collapse:collapse; width:100%; background:#fff; }}
  table.diff td {{ padding:2px 6px; border:1px solid #eef0f3; vertical-align:top;
                   white-space:pre-wrap; word-break:break-all; }}
  .diff_header {{ background:#f2f3f5; color:#86909c; }}
  td.diff_add {{ background:#e6f7e9; }}
  td.diff_sub {{ background:#fdecea; }}
  td.diff_chg {{ background:#fff7e6; }}
</style></head><body>
<div class="card">
  <h1>文档对比报告</h1>
  <div class="kv">
    <div>旧版本：<b>{html.escape(os.path.basename(old_path))}</b></div>
    <div>新版本：<b>{html.escape(os.path.basename(new_path))}</b></div>
  </div>
  <div class="kv" style="margin-top:10px">
    <div>相似度：<b>{ratio * 100:.2f}%</b></div>
    <div class="good">新增 <b>{added}</b> 行</div>
    <div class="bad">删除 <b>{removed}</b> 行</div>
    <div>修改 <b>{changed}</b> 行</div>
  </div>
  <div class="kv" style="margin-top:10px; font-size:12.5px">
    <div>忽略空白：{'是' if ignore_ws else '否'}</div>
    <div>忽略大小写：{'是' if ignore_case else '否'}</div>
  </div>
</div>
<div class="card">{''.join(rows)}</div>
</body></html>"""
    with open(long_path(path), "w", encoding="utf-8") as fh:
        fh.write(doc)
