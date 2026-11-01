#!/usr/bin/env python3
"""crop_figures.py — 從題目截圖精確裁出插圖（worksheet-docx skill）

用法（單張）：
    python crop_figures.py 來源.png --box 左,上,右,下 --out 圖名 [--dir crops]

用法（批次，推薦）：
    boxes.json:
    {
      "fig_spring": {"src": "img-000.png", "box": [625, 0, 905, 250]},
      "fig_graph":  {"src": "img-002.png", "box": [590, 0, 934, 210]}
    }
    python crop_figures.py --spec boxes.json [--dir crops]

src 路徑相對於 boxes.json 所在目錄。
⚠️ 裁完每張都要目視複驗（ReadMediaFile）：左緣是否切到相鄰文字、
   右緣/下緣是否切斷曲線或標籤——通常需 2–3 輪調整 box。
"""
import json
import os
import sys

try:
    from PIL import Image
except ImportError:
    sys.exit("需要 Pillow：pip install pillow")


def crop(src, box, out):
    im = Image.open(src)
    w, h = im.size
    l, t, r, b = box
    if not (0 <= l < r <= w and 0 <= t < b <= h):
        sys.exit(f"box {box} 超出圖片範圍 {w}x{h}：{src}")
    Image.open(src).crop((l, t, r, b)).save(out)
    print(f"{out}  ({r-l}x{b-t})")


def main():
    args = sys.argv[1:]
    outdir = "crops"
    if "--dir" in args:
        i = args.index("--dir")
        outdir = args[i + 1]
        del args[i:i + 2]
    os.makedirs(outdir, exist_ok=True)

    if "--spec" in args:
        spec_path = args[args.index("--spec") + 1]
        spec_dir = os.path.dirname(os.path.abspath(spec_path))
        spec = json.load(open(spec_path, encoding="utf-8"))
        for name, item in spec.items():
            src = item["src"]
            if not os.path.isabs(src):
                src = os.path.join(spec_dir, src)
            crop(src, item["box"], os.path.join(outdir, f"{name}.png"))
    elif "--box" in args:
        # python crop_figures.py 來源.png --box 左,上,右,下 --out 圖名
        src = args[0]
        box = [int(x) for x in args[args.index("--box") + 1].split(",")]
        name = args[args.index("--out") + 1]
        if not name.endswith(".png"):
            name += ".png"
        crop(src, box, os.path.join(outdir, name))
    else:
        sys.exit(__doc__)

    print("完成。請逐張目視複驗裁切結果！")


if __name__ == "__main__":
    main()
