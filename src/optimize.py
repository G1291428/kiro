# -*- coding: utf-8 -*-
"""把官方原图压成适合 PPT 嵌入的尺寸：最长边 1400px，JPEG q82，透明底拍平成白色"""
import os, glob
from PIL import Image

SRC, DST = "assets", "assets_min"
os.makedirs(DST, exist_ok=True)
MAXD = 1400

used = {
    "amalfi_n1", "amalfi_n2", "amalfi_n2t", "milkpal_prod", "linkbar_st_crop",
    "marco_bar", "ubermilk_front", "linkbar_milkpal_hero", "linkbar_milkpal_scene2",
    "st_prod", "marco_milkpal", "marco_inuse", "st_bar",
}

total = 0
for p in sorted(glob.glob(os.path.join(SRC, "*.png"))):
    name = os.path.splitext(os.path.basename(p))[0]
    if name not in used:
        continue
    im = Image.open(p)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    w, h = im.size
    if max(w, h) > MAXD:
        sc = MAXD / max(w, h)
        im = im.resize((int(w * sc), int(h * sc)), Image.LANCZOS)
    out = os.path.join(DST, name + ".jpg")
    im.save(out, "JPEG", quality=82, optimize=True, progressive=True)
    total += os.path.getsize(out)
    print("%-26s %4dx%-4d -> %4dx%-4d  %6.1f KB" %
          (name, w, h, im.size[0], im.size[1], os.path.getsize(out) / 1024))
print("合计 %.2f MB" % (total / 1048576))
