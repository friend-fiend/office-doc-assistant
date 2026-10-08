"""生成应用图标（Windows .ico / macOS .icns / 通用 .png）。

为什么要自己生成
----------------
打包时如果 ``icon=None``，PyInstaller 会把 exe 的图标留成默认的 Python 图标，
产品属性页也是空白的 —— 用户看到的第一个"这是个正经软件吗"的判断就来自这里。

依赖只有 Pillow（已经是本项目依赖），跨平台都能跑，不需要 ImageMagick。

用法
----
    python build/make_icons.py            # 输出到 assets/
    python build/make_icons.py --out DIR  # 指定输出目录

关于 .icns
----------
Pillow 的 ICNS 写入支持依赖平台，因此这里自己按 ICNS 容器格式写：
    ICNS = b'icns' + 总长度(4, 大端) + [类型(4) + 长度(4, 大端) + PNG 数据] * N
只要嵌入的是 PNG（现代 macOS 完全支持），无需 Apple 的 iconutil。
"""
from __future__ import annotations

import argparse
import os
import struct
import sys

try:
    from PIL import Image, ImageDraw
except ImportError:  # pragma: no cover
    print("需要 Pillow：pip install Pillow", file=sys.stderr)
    raise SystemExit(2)

# ------------------------------------------------------------------ 设计参数
SS = 4  # 超采样倍数：先按 4 倍画再缩，边缘才不会有锯齿
BASE = 1024

# 背景渐变（上浅下深的蓝，办公软件的稳妥配色）
BG_TOP = (77, 141, 246)
BG_BOTTOM = (26, 63, 184)

DOC_FILL = (255, 255, 255)
DOC_EDGE = (222, 232, 246)
FOLD_FILL = (206, 220, 240)
LINE_COLOR = (168, 190, 222)

BADGE_FILL = (255, 154, 46)   # 橙色徽标 = "一键/快速"
BOLT_FILL = (255, 255, 255)

ICO_SIZES = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]

# ICNS 里 PNG 块的类型标识 -> 像素尺寸
ICNS_TYPES = [
    (b"icp4", 16), (b"icp5", 32), (b"icp6", 64),
    (b"ic07", 128), (b"ic08", 256), (b"ic09", 512), (b"ic10", 1024),
]


def _rounded_mask(size: int, radius_ratio: float = 0.215) -> Image.Image:
    """圆角矩形遮罩（L 模式，白色为不透明）。"""
    m = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(m)
    r = int(size * radius_ratio)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=255)
    return m


def _gradient(size: int) -> Image.Image:
    """竖直渐变背景。逐行画线，1000 行级别开销可忽略。"""
    img = Image.new("RGB", (size, size))
    d = ImageDraw.Draw(img)
    for y in range(size):
        t = y / max(size - 1, 1)
        color = tuple(
            int(BG_TOP[i] + (BG_BOTTOM[i] - BG_TOP[i]) * t) for i in range(3)
        )
        d.line([(0, y), (size, y)], fill=color)
    return img


def _document_layer(size: int) -> Image.Image:
    """白纸 + 折角 + 文本线条。返回带透明通道的图层。

    所有坐标都以 1024 为设计基准，按 ``k = size / 1024`` 等比缩放，
    这样超采样（size = 4096）与直接渲染同一套参数。
    """
    k = size / 1024.0

    def P(*vals: float) -> list[float]:
        return [v * k for v in vals]

    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    # 纸张范围（相对 1024 基准）
    x0, y0, x1, y1 = P(268, 214, 700, 786)
    fold = 132 * k                               # 折角边长
    pad = 26 * k                                 # 描边宽度

    # 主体：右下角留出徽标位置，所以纸张稍偏左上
    body = [
        (x0 + pad, y0),
        (x1 - fold, y0),
        (x1, y0 + fold),
        (x1, y1 - pad),
        (x1 - pad, y1),
        (x0 + pad, y1),
        (x0, y1 - pad),
        (x0, y0 + pad),
    ]
    d.polygon(body, fill=DOC_FILL)
    d.line(body + [body[0]], fill=DOC_EDGE, width=max(int(6 * k), 2), joint="curve")

    # 折角三角
    d.polygon([(x1 - fold, y0), (x1, y0 + fold), (x1 - fold, y0 + fold)],
              fill=FOLD_FILL)

    # 文本线条：长度故意参差，像真实文稿
    lines = [
        (y0 + 226 * k, 0.86), (y0 + 302 * k, 0.72), (y0 + 378 * k, 0.86),
        (y0 + 454 * k, 0.60),
    ]
    inner_l, inner_r = x0 + 62 * k, x1 - 62 * k
    span = inner_r - inner_l
    for ly, ratio in lines:
        d.rounded_rectangle(
            [inner_l, ly, inner_l + int(span * ratio), ly + 34 * k],
            radius=17 * k, fill=LINE_COLOR,
        )
    return layer


def _badge_layer(size: int) -> Image.Image:
    """右下角圆形徽标 + 闪电。"""
    k = size / 1024.0
    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    cx, cy, r = 742 * k, 742 * k, 196 * k
    # 先画一圈与背景同色的"隔离环"，让徽标和白纸之间不糊在一起
    ring = 30 * k
    d.ellipse([cx - r - ring, cy - r - ring, cx + r + ring, cy + r + ring],
              fill=(255, 255, 255, 255))
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=BADGE_FILL)

    # 闪电剪影（经典 6 点折线，归一化坐标 -> 实际坐标）
    norm = [(0.35, -1.0), (-0.45, 0.12), (0.0, 0.12),
            (-0.30, 1.0), (0.45, -0.12), (0.02, -0.12)]
    sx, sy = r * 0.62, r * 0.66
    pts = [(cx + nx * sx, cy + ny * sy) for nx, ny in norm]
    d.polygon(pts, fill=BOLT_FILL)
    return layer


def render() -> Image.Image:
    """渲染 1024x1024 的母图（RGBA）。"""
    s = BASE * SS
    bg = _gradient(s)
    bg = bg.convert("RGBA")

    doc = _document_layer(s)
    badge = _badge_layer(s)
    bg.alpha_composite(doc)
    bg.alpha_composite(badge)

    # 圆角裁切
    mask = _rounded_mask(s)
    out = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    out.paste(bg, (0, 0), mask)

    return out.resize((BASE, BASE), Image.LANCZOS)


def save_ico(master: Image.Image, path: str) -> None:
    master.save(path, format="ICO", sizes=ICO_SIZES)


def save_icns(master: Image.Image, path: str) -> None:
    """手写 ICNS 容器：PNG 块拼装，不依赖 iconutil。"""
    chunks = []
    for tag, px in ICNS_TYPES:
        buf = os.path.join(os.path.dirname(path), f".tmp_{px}.png")
        master.resize((px, px), Image.LANCZOS).save(buf, format="PNG")
        with open(buf, "rb") as fh:
            data = fh.read()
        os.remove(buf)
        chunks.append(tag + struct.pack(">I", len(data) + 8) + data)

    payload = b"".join(chunks)
    with open(path, "wb") as fh:
        fh.write(b"icns" + struct.pack(">I", len(payload) + 8) + payload)


def main() -> int:
    ap = argparse.ArgumentParser(description="生成应用图标")
    ap.add_argument("--out", default=None, help="输出目录（默认 <项目>/assets）")
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    out_dir = args.out or os.path.join(os.path.dirname(here), "assets")
    os.makedirs(out_dir, exist_ok=True)

    master = render()

    png = os.path.join(out_dir, "app.png")
    ico = os.path.join(out_dir, "app.ico")
    icns = os.path.join(out_dir, "app.icns")

    master.save(png, format="PNG")
    save_ico(master, ico)
    save_icns(master, icns)

    for p in (png, ico, icns):
        print(f"  {os.path.basename(p):<10} {os.path.getsize(p):>8,} bytes  {p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
