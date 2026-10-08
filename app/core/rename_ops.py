"""批量重命名（含预览）。

预览是本模块的核心：``build_preview()`` 纯计算新文件名，不碰磁盘，
界面弹出预览表让用户确认后再执行，从根本上避免"手一抖全改错了"。

执行有两种模式，对应需求里"输出不覆盖原文件"的原则：
  · copy   （默认）把文件以新名字"复制"到输出文件夹，原文件不动；
  · inplace        原地重命名，执行前强制二次确认，并自动生成回滚清单。
"""
from __future__ import annotations

import json
import os
import re
import shutil
from dataclasses import dataclass

from ..errors import ParamError
from ..paths import long_path, sanitize_name, unique_dir
from .registry import OpContext, OpDef, Param, register


@dataclass
class RenameRow:
    old_path: str
    new_name: str
    status: str = "ok"        # ok | conflict | unchanged | invalid | empty
    note: str = ""

    @property
    def old_name(self) -> str:
        return os.path.basename(self.old_path)


RENAME_OP = OpDef(
    key="batch_rename",
    title="批量重命名",
    category="batch",
    accepts=("any",),
    mode="per_file",
    confirm=True,
    confirm_text="请先点击「预览重命名」确认结果。使用原地模式时原文件名将被改变。",
    description=(
        "为文件批量添加前缀/后缀/序号，或做查找替换。\n"
        "点击「预览重命名」可在执行前查看每个文件的新名字。"
    ),
    params=(
        Param("name_mode", "命名方式", "choice", default="original_seq",
              choices=(("original_seq", "原名 + 序号"), ("sequence", "纯序号"),
                       ("find_replace", "查找替换"))),
        Param("prefix", "前缀", "entry", default="", width=16),
        Param("suffix", "后缀", "entry", default="", width=16),
        Param("seq_start", "起始序号", "int", default=1, min_value=0, max_value=999999),
        Param("seq_digits", "序号位数", "int", default=2, min_value=1, max_value=6,
              hint="2 表示 01、02 …"),
        Param("seq_pos", "序号位置", "choice", default="before_ext",
              choices=(("before_ext", "放在原名之后，扩展名之前"),
                       ("after_prefix", "紧跟前缀")),
              visible_if=("name_mode", ("original_seq", "sequence"))),
        Param("seq_sep", "序号分隔符", "choice", default="_",
              choices=(("_", "下划线 _"), ("-", "短横线 -"), ("", "无分隔"),
                       (".", "点 ."), (" ", "空格")),
              hint="例如 图片1 + _ + 01 = 图片1_01",
              visible_if=("name_mode", ("original_seq",))),
        Param("find_text", "查找", "entry", default="", width=16,
              visible_if=("name_mode", ("find_replace",))),
        Param("replace_text", "替换为", "entry", default="", width=16,
              visible_if=("name_mode", ("find_replace",))),
        Param("use_regex", "查找使用正则表达式", "bool", default=False,
              visible_if=("name_mode", ("find_replace",))),
        Param("ext_case", "扩展名大小写", "choice", default="keep",
              choices=(("keep", "保持不变"), ("lower", "小写 .pdf"),
                       ("upper", "大写 .PDF"))),
        Param("apply_mode", "执行方式", "choice", default="copy",
              choices=(("copy", "生成重命名副本到输出文件夹（不改原文件）"),
                       ("inplace", "原地重命名（会改变原文件名）"))),
    ),
)


# ------------------------------------------------------------------ 预览
def build_preview(files: list[str], params: dict) -> list[RenameRow]:
    """计算重命名结果，不产生任何磁盘写操作。"""
    mode = str(params.get("name_mode") or "original_seq")
    prefix = str(params.get("prefix") or "")
    suffix = str(params.get("suffix") or "")
    start = int(params.get("seq_start") or 1)
    digits = max(1, min(6, int(params.get("seq_digits") or 2)))
    seq_pos = str(params.get("seq_pos") or "before_ext")
    seq_sep = str(params.get("seq_sep") if params.get("seq_sep") is not None else "_")
    find_text = str(params.get("find_text") or "")
    replace_text = str(params.get("replace_text") or "")
    use_regex = bool(params.get("use_regex"))
    ext_case = str(params.get("ext_case") or "keep")

    rows: list[RenameRow] = []
    for i, path in enumerate(files):
        stem, ext = os.path.splitext(os.path.basename(path))
        seq = str(start + i).zfill(digits)

        try:
            if mode == "find_replace":
                if not find_text:
                    new_stem = stem
                elif use_regex:
                    new_stem = re.sub(find_text, replace_text, stem)
                else:
                    new_stem = stem.replace(find_text, replace_text)
                new_stem = f"{prefix}{new_stem}{suffix}"
            elif mode == "sequence":
                new_stem = f"{prefix}{seq}{suffix}"
            else:  # original_seq
                if seq_pos == "after_prefix":
                    new_stem = f"{prefix}{seq}{seq_sep}{stem}{suffix}"
                else:
                    new_stem = f"{prefix}{stem}{seq_sep}{seq}{suffix}"
        except re.error as exc:
            rows.append(RenameRow(path, "", "invalid", f"正则表达式错误：{exc}"))
            continue

        if ext_case == "lower":
            ext = ext.lower()
        elif ext_case == "upper":
            ext = ext.upper()

        new_name = sanitize_name(new_stem) + ext
        row = RenameRow(path, new_name)
        if not new_stem.strip():
            row.status, row.note = "empty", "新文件名为空"
        elif new_name == os.path.basename(path):
            row.status, row.note = "unchanged", "名称未变化"
        rows.append(row)

    # 冲突检测：批内重名 + 与磁盘既存文件重名（原地模式才检查磁盘）
    seen: dict[str, int] = {}
    for row in rows:
        key = os.path.normcase(row.new_name)
        seen[key] = seen.get(key, 0) + 1
    for row in rows:
        key = os.path.normcase(row.new_name)
        if seen.get(key, 0) > 1:
            row.status, row.note = "conflict", "批次内出现重名"
        elif row.status == "ok" and os.path.normcase(row.new_name) != os.path.normcase(row.old_name):
            sibling = os.path.join(os.path.dirname(row.old_path), row.new_name)
            if os.path.exists(sibling):
                row.status, row.note = "conflict", "目标目录已存在同名文件"
    return rows


# ------------------------------------------------------------------ 执行
@register(RENAME_OP)
def run_batch_rename(ctx: OpContext, files: list[str], params: dict) -> None:
    rows = build_preview(files, params)
    if not rows:
        raise ParamError("没有可重命名的文件")

    apply_mode = str(params.get("apply_mode") or "copy")
    ctx.set_total(len(rows))
    rollback: list[dict] = []

    if apply_mode == "inplace":
        ctx.log(f"原地重命名模式：将直接修改 {len(rows)} 个文件的名称。", "warn")

    for row in rows:
        ctx.check()
        if row.status in ("invalid", "empty"):
            ctx.fail(row.old_path, row.note or "文件名无效")
            ctx.step(row.old_name)
            continue
        if row.status == "conflict":
            ctx.skip(row.old_path, row.note or "目标名称冲突")
            ctx.step(row.old_name)
            continue
        if row.status == "unchanged":
            ctx.skip(row.old_path, "名称未变化")
            ctx.step(row.old_name)
            continue

        try:
            if apply_mode == "inplace":
                target = os.path.join(os.path.dirname(row.old_path), row.new_name)
                os.rename(long_path(row.old_path), long_path(target))
                rollback.append({"from": target, "to": row.old_path})
            else:
                target = os.path.join(ctx.out_dir, row.new_name)
                if os.path.exists(target):
                    from ..paths import unique_file

                    target = unique_file(target)
                shutil.copy2(long_path(row.old_path), long_path(target))
            ctx.ok(target, src=row.old_path)
        except Exception as exc:
            ctx.fail(row.old_path, str(exc))
        ctx.step(row.old_name)

    if rollback:
        log_path = os.path.join(ctx.out_dir, "重命名回滚清单.json")
        try:
            with open(long_path(log_path), "w", encoding="utf-8") as fh:
                json.dump(rollback, fh, ensure_ascii=False, indent=2)
            ctx.metric("回滚清单", log_path)
            ctx.log(f"已生成回滚清单，如需还原请参照：{log_path}")
        except Exception:
            pass
