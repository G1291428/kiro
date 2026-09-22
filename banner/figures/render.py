# -*- coding: utf-8 -*-
"""把易拉宝配图的 HTML 稿渲染成印刷级 PNG。

CSS 画布 1400×980，device_scale_factor=3 → 4200×2940 px，
按 300dpi 输出约 35.6×24.9 cm，满足易拉宝幅面。
"""
from pathlib import Path

from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
OUT = HERE.parent
CHROME = "/opt/playwright/chromium-1232/chrome-linux64/chrome"
W, H, SCALE = 1400, 980, 3

JOBS = [
    ("fig05_principle.html", "05_超声空化低温杀菌原理.png"),
    ("fig06_compare.html", "06_传统热杀菌与超声低温杀菌对比.png"),
]


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            executable_path=CHROME,
            args=["--no-sandbox", "--disable-dev-shm-usage",
                  "--font-render-hinting=none", "--force-color-profile=srgb"],
        )
        page = browser.new_page(viewport={"width": W, "height": H},
                                device_scale_factor=SCALE)
        for src, dst in JOBS:
            page.goto((HERE / src).as_uri())
            page.wait_for_timeout(350)
            out = OUT / dst
            page.screenshot(path=str(out), scale="device")
            kb = out.stat().st_size // 1024
            print("%-42s %dx%d px   300dpi=%.1f×%.1fcm   %dKB"
                  % (dst, W * SCALE, H * SCALE, W * SCALE / 300 * 2.54,
                     H * SCALE / 300 * 2.54, kb))
        browser.close()


if __name__ == "__main__":
    main()
