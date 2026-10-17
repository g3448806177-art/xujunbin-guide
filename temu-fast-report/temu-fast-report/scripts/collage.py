#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""TEMU 报单 - 候选图拼总览（规则4：同一组目标图+2-3张候选图拼成一张，一次 Read 完成对比）。

用法:
  python3 collage.py --target images-small/row-3.jpg \
      --cand cand/a.jpg cand/b.jpg cand/c.jpg \
      --labels "志成/20138" "拼步/20138" "善缘/20138" \
      --chosen 2 \
      --out cand/review-row3.jpg

说明:
  - 第一格固定为目标 SKU 图（红框，标"目标"），其后依次为候选图（标候选序号+label）
  - --chosen 指定被选中候选的 1-based 序号，该格加绿框；不指定则候选格统一灰框
  - 每格等宽正方形，底部留文字条；输出单张 jpg，供一次 Read 并排核对
"""
import argparse
import os

from PIL import Image, ImageDraw, ImageFont

CELL = 340          # 每格图片区边长
LABEL_H = 46        # 底部文字条高度
GAP = 12            # 格间距
PAD = 12            # 外边距
_WINDOWS_FONTS = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
FONT_CANDIDATES = [
    # Windows：微软雅黑、黑体、宋体
    os.path.join(_WINDOWS_FONTS, "msyh.ttc"),
    os.path.join(_WINDOWS_FONTS, "msyh.ttf"),
    os.path.join(_WINDOWS_FONTS, "simhei.ttf"),
    os.path.join(_WINDOWS_FONTS, "simsun.ttc"),
    # macOS
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    # Linux
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
]


def load_font(size):
    for fp in FONT_CANDIDATES:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()


def square_fit(img, side):
    """等比裁剪居中到 side×side。"""
    img = img.convert("RGB")
    w, h = img.size
    s = min(w, h)
    img = img.crop(((w - s) // 2, (h - s) // 2, (w - s) // 2 + s, (h - s) // 2 + s))
    return img.resize((side, side), Image.LANCZOS)


def make_cell(img_path, title, subtitle, frame_color, side=CELL):
    """生成一格：上图(带彩色边框)+下文字条。"""
    cell_h = side + LABEL_H
    cell = Image.new("RGB", (side, cell_h), "white")
    draw = ImageDraw.Draw(cell)
    if img_path and os.path.exists(img_path):
        pic = square_fit(Image.open(img_path), side - 8)
        cell.paste(pic, (4, 4))
    else:
        draw.rectangle([4, 4, side - 4, side - 4], outline="#CCCCCC", fill="#F2F2F2")
        draw.text((side // 2 - 30, side // 2), "缺图", fill="#999999", font=load_font(22))
    # 边框（目标红/选中绿/候选灰）
    draw.rectangle([2, 2, side - 3, side - 1], outline=frame_color, width=5)
    f_title = load_font(20)
    f_sub = load_font(17)
    draw.text((8, side + 4), title, fill="#222222", font=f_title)
    if subtitle:
        draw.text((8, side + 25), subtitle[:24], fill="#555555", font=f_sub)
    return cell


def main():
    p = argparse.ArgumentParser(description="候选图拼总览")
    p.add_argument("--target", required=True, help="目标SKU图")
    p.add_argument("--cand", nargs="+", default=[], help="候选图路径(2-3张)")
    p.add_argument("--labels", nargs="*", default=[], help="候选标签(店铺/货号)，顺序与--cand一致")
    p.add_argument("--chosen", type=int, default=0, help="选中候选的1-based序号，0=未选(归网购)")
    p.add_argument("--target-label", default="", help="目标图副标题(颜色/码数)")
    p.add_argument("--out", required=True, help="输出jpg路径")
    args = p.parse_args()

    cells = [make_cell(args.target, "目标 SKU图", args.target_label, (220, 60, 60))]
    for i, c in enumerate(args.cand, start=1):
        label = args.labels[i - 1] if i - 1 < len(args.labels) else ""
        is_chosen = (i == args.chosen)
        title = f"候选{i}" + ("  ★选中" if is_chosen else "")
        color = (40, 170, 80) if is_chosen else (150, 150, 150)
        cells.append(make_cell(c, title, label, color))

    n = len(cells)
    W = PAD * 2 + CELL * n + GAP * (n - 1)
    H = PAD * 2 + CELL + LABEL_H
    canvas = Image.new("RGB", (W, H), "white")
    x = PAD
    for cell in cells:
        canvas.paste(cell, (x, PAD))
        x += CELL + GAP
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    canvas.save(args.out, quality=90)
    print(f"已生成对比图: {args.out}（{n}格）")


if __name__ == "__main__":
    main()
