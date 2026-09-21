# -*- coding: utf-8 -*-
"""咖啡液萃取与超声杀菌项目汇报 —— 排版美化。

规则：
  1. 页序与页数不变（24 页）；
  2. 原始图片全部保留并复用（按 sha1 复用 image part，不新增、不替换任何位图）；
  3. 图表 / 表格均为原稿截图，一律等比缩放不裁切，数值零改动；
  4. 仅调整排版、留白、结构、字体与视觉统一性；
  5. 第 6 / 8 / 14 / 15 页为待补充页，只统一页面框架，不填内容。
"""
from __future__ import annotations

import shutil
import sys
import zipfile
from copy import deepcopy
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from theme import (  # noqa: E402
    BLUE, BLUE_LINE, BLUE_TINT, CRIMSON, CW, FONT, FONT_CAL, FONT_EN, GOLD,
    GOLD_LT, GREEN, INK, INK_MID, INK_SOFT, M, NAVY, NAVY_DEEP, NAVY_MID,
    PAPER, RED_HI, RULE_Y, SH, SW, WHITE, I, P, T, brand_rule, card,
    card_head, chip, cover_crop, fill_text, fit_box, foot, grad, head, kpi,
    line, placeholder_hint, put_pic, rect, set_rule_image, soft_shadow,
    takeaway, textbox, write,
)

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "咖啡液萃取与超声杀菌工艺及装备设计研究设备.pptx"
OUTDIR = ROOT / "dist"
OUT = OUTDIR / "咖啡液萃取与超声杀菌工艺及装备设计研究设备_美化版.pptx"
WORK = Path("/projects/sandbox/work/build")

MED: dict[str, tuple[str, int, int]] = {}
USED: set[str] = set()


# ------------------------------------------------------------------ 准备
def extract_media() -> None:
    mdir = WORK / "media"
    if mdir.exists():
        shutil.rmtree(mdir)
    mdir.mkdir(parents=True)
    with zipfile.ZipFile(SRC) as z:
        for n in z.namelist():
            if n.startswith("ppt/media/"):
                base = n.split("/")[-1]
                data = z.read(n)
                p = mdir / base
                p.write_bytes(data)
                try:
                    with Image.open(p) as im:
                        w, h = im.size
                except Exception:
                    continue
                MED[base] = (str(p), w, h)


def reset(slide):
    """清空页面图元，保留 spTree 必需的头两个节点。"""
    spTree = slide.shapes._spTree
    for el in list(spTree):
        tag = el.tag.split("}")[-1]
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        spTree.remove(el)


def renumber(prs):
    """重排 shape id，避免同页 id 冲突。"""
    for slide in prs.slides:
        n = 2
        for el in slide.shapes._spTree.iter(qn("p:cNvPr")):
            el.set("id", str(n))
            n += 1


def clone(slide, sp_el, x, y, w=None, h=None):
    """把原稿的自定义几何形状（分节页斜切块）原样搬到目标页。"""
    el = deepcopy(sp_el)
    slide.shapes._spTree.append(el)
    spPr = el.find(qn("p:spPr"))
    xfrm = spPr.find(qn("a:xfrm"))
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    off.set("x", str(I(x)))
    off.set("y", str(I(y)))
    if w is not None:
        ext.set("cx", str(I(w)))
    if h is not None:
        ext.set("cy", str(I(h)))
    return el


# ------------------------------------------------------------------ 图片放置
def pic_fit(slide, key, x, y, w, h, *, crop=(0.0, 0.0, 0.0, 0.0), **kw):
    """等比放入（不裁切）—— 图表、示意图、报告扫描件走这条。"""
    path, pw, ph = MED[key]
    USED.add(key)
    vw = pw * (1 - crop[0] - crop[1])
    vh = ph * (1 - crop[2] - crop[3])
    bx, by, bw, bh = fit_box(vw, vh, x, y, w, h)
    return put_pic(slide, path, bx, by, bw, bh,
                   crop=crop if any(crop) else None, **kw)


def pic_cover(slide, key, x, y, w, h, *, bias_x=0.5, bias_y=0.5,
              base=(0.0, 0.0, 0.0, 0.0), **kw):
    """裁切铺满 —— 只用于照片类素材，不用于任何含数据的图。"""
    path, pw, ph = MED[key]
    USED.add(key)
    c = cover_crop(pw, ph, w, h, bias_x=bias_x, bias_y=bias_y, base=base)
    return put_pic(slide, path, x, y, w, h, crop=c, **kw)


# ------------------------------------------------------------------ 01 封面
def s01(slide):
    reset(slide)
    rect(slide, 0, SH - 0.078, SW, 0.078, fill=CRIMSON)

    pic_fit(slide, "image2.jpg", M, 0.30, 3.40, 0.84)
    textbox(slide, SW - M - 4.40, 0.60, 4.40, 0.30,
            [P([T("项目汇报", sz=12, b=True, c=NAVY, spc=2.4)], al="r")])
    brand_rule(slide, 1.24)

    # 主视觉：校园夜景横幅 + 深蓝渐变压暗，保证标题可读
    hero_y, hero_h = 1.52, 3.52
    pic_cover(slide, "image1.png", 0, hero_y, SW, hero_h, bias_y=0.72)
    grad(slide, 0, hero_y, SW, hero_h, NAVY_DEEP, NAVY_MID, angle=0, a1=88, a2=76)
    rect(slide, 0, hero_y + hero_h - 0.05, SW, 0.05, fill=CRIMSON)

    rect(slide, (SW - 1.30) / 2, hero_y + 0.42, 1.30, 0.055, fill=WHITE)
    textbox(slide, 0.80, hero_y + 0.72, SW - 1.60, 1.60,
            [P([T("咖啡液萃取与超声杀菌")], al="c", ls=1.12),
             P([T("工艺及装备设计研究设备")], al="c", ls=1.12)],
            sz=39, b=True, c=WHITE, anchor="t")
    textbox(slide, 0.80, hero_y + 2.62, SW - 1.60, 0.42,
            [P([T("Tianjin University of Science ＆ Technology",
                  sz=17, b=True, c="C9DAEA", f=FONT_EN, spc=0.6)], al="c")])

    # 责任人信息：四栏等宽 + 细分隔线
    items = [("项目负责人", "苏  杭"), ("理工科负责人", "宋继田 教授"),
             ("人文社科负责人", "李金茹 副教授"), ("汇报日期", "2026.09.14")]
    tw, x0, iy = 11.60, (SW - 11.60) / 2, 5.42
    iw = tw / 4
    for k, (lab, val) in enumerate(items):
        x = x0 + k * iw
        textbox(slide, x, iy, iw, 0.26,
                [P([T(lab, sz=11, c=INK_SOFT, spc=1.0)], al="c")])
        textbox(slide, x, iy + 0.28, iw, 0.38,
                [P([T(val, sz=15.5, b=True, c=NAVY)], al="c")])
        if k:
            rect(slide, x, iy + 0.06, 0.01, 0.62, fill=BLUE_LINE)

    # 校训：沿用原稿书法体，两侧细线收口
    my = 6.66
    rect(slide, M, my + 0.20, 3.85, 0.012, fill=BLUE_LINE)
    rect(slide, SW - M - 3.85, my + 0.20, 3.85, 0.012, fill=BLUE_LINE)
    textbox(slide, (SW - 4.40) / 2, my - 0.06, 4.40, 0.52,
            [P([T("尚德尚学尚行 爱国爱校爱人", sz=22, c=NAVY, f=FONT_CAL)], al="c")],
            anchor="m")


# ------------------------------------------------------------------ 分节页
DIVIDERS = {
    2: ("01", "项目背景", "PROJECT BACKGROUND",
        "国内咖啡消费五年持续增长，咖啡液成为增速领先的细分品类。"),
    7: ("02", "行业现状", "INDUSTRY STATUS",
        "加工以热杀菌为主，风味损失与高能耗并存，灭菌是关键工序。"),
    12: ("03", "核心优势", "CORE ADVANTAGES",
         "低温超声杀菌叠加多级工艺整合，兼顾风味保留与节能降耗。"),
    19: ("04", "团队成员", "OUR TEAM",
         "机械与能源动力方向研究生团队，双负责人跨学科指导。"),
    22: ("05", "未来规划", "ROADMAP",
         "2026—2028 三步走：中试运行、一体化装备、技术更新。"),
}


def divider(slide, shapes, num, cn, en, lead):
    reset(slide)
    clone(slide, shapes["para"], 0, 0, 7.03, 7.50)
    clone(slide, shapes["strip"], 0, -0.01, 3.04, 7.51)
    rect(slide, SW - 0.13, 0, 0.13, SH, fill=CRIMSON)

    textbox(slide, 0.30, 1.58, 6.50, 0.30,
            [P([T("SECTION", sz=13, b=True, c=WHITE, spc=4.5)], al="c")])
    rect(slide, (0.30 + 6.50) / 2 - 0.55, 2.00, 1.10, 0.035, fill=WHITE)
    textbox(slide, 0.30, 2.20, 6.50, 3.30,
            [P([T(num, sz=196, b=True, c=WHITE)], al="c", ls=1.0)], anchor="m")

    rect(slide, 7.85, 2.34, 1.05, 0.06, fill=CRIMSON)
    textbox(slide, 7.85, 2.54, 5.10, 1.00,
            [P([T(cn, sz=52, b=True, c=NAVY, spc=3.0)], ls=1.0)], anchor="t")
    textbox(slide, 7.85, 3.62, 5.10, 0.34,
            [P([T(en, sz=13.5, b=True, c=INK_SOFT, spc=2.6)])])
    rect(slide, 7.85, 4.06, 0.60, 0.022, fill=BLUE_LINE)
    textbox(slide, 7.85, 4.26, 4.70, 1.30,
            [P([T(lead, sz=14.5, c=INK_MID)], ls=1.5)], anchor="t")


# ------------------------------------------------------------------ 待补充页
PLACEHOLDERS = {
    6: ("01 / 项目背景 · 市场前景", [T("市场前景")], "市场前景"),
    8: ("02 / 行业现状 · 应用场景", [T("应用场景 · 高峰期大流量人群")], "应用场景"),
    14: ("03 / 核心优势 · 竞品分析", [T("竞品分析")], "竞品分析"),
    15: ("03 / 核心优势 · 竞品分析", [T("竞品分析")], "竞品分析"),
}


def placeholder_page(slide, idx, eyebrow, title, label):
    reset(slide)
    from theme import page_frame

    page_frame(slide)
    head(slide, eyebrow, title)
    placeholder_hint(slide, 1.36, 5.06,
                     "【%s】内容待补充 —— 本页仅统一页面框架，未改动内容" % label)
    foot(slide, idx)


# ------------------------------------------------------------------ 03 市场分析①
def market_page(slide, idx, *, eyebrow, title, img, kpis, conclusion, note,
                title_sz=27):
    """三张市场分析页共用版式：左图卡 + 右数据栏 + 底部结论条。"""
    from theme import page_frame

    page_frame(slide)
    head(slide, eyebrow, title, sz=title_sz)

    body_t, body_b = 1.30, 5.96
    lw, rw = 9.10, 2.79
    rx = M + lw + 0.20

    card(slide, M, body_t, lw, body_b - body_t)
    pic_fit(slide, img, M + 0.20, body_t + 0.20, lw - 0.40, body_b - body_t - 0.40)

    n = len(kpis)
    kh = (body_b - body_t - 0.17 * (n - 1)) / n
    for k, (val, unit, lab, acc) in enumerate(kpis):
        kpi(slide, rx, body_t + k * (kh + 0.17), rw, kh, val, lab,
            accent=acc, unit=unit)

    takeaway(slide, 6.12, conclusion, h=0.74)
    foot(slide, idx, note=note)


def s03(slide, idx):
    reset(slide)
    market_page(
        slide, idx,
        eyebrow="01 / 项目背景 · 市场分析",
        title=[T("国内咖啡生豆消耗量五年持续攀升，"),
               T("2025 年达 41.5 万吨", c=RED_HI)],
        title_sz=25,
        img="image3.PNG",
        kpis=[("41.5", "万吨", "2025 年国内咖啡生豆消耗总量", NAVY),
              ("22.5 → 41.5", "", "2021 → 2025 消耗总量（万吨）", NAVY_MID),
              ("+84.44", "%", "总生豆五年总增长率", RED_HI)],
        conclusion=[T("2025 年国内咖啡豆生豆消耗量达到 "),
                    T("41.5 万吨", c=GOLD_LT)],
        note="统计区间：2021—2025 年；数值与原图表完全一致",
    )


def s04(slide, idx):
    reset(slide)
    market_page(
        slide, idx,
        eyebrow="01 / 项目背景 · 市场分析",
        title=[T("咖啡液五年增长 "), T("621.74%", c=RED_HI),
               T("，增速仅次于冻干咖啡")],
        img="image4.PNG",
        kpis=[("621.74", "%", "咖啡液 2021—2025 总增长率", RED_HI),
              ("731.11", "%", "冻干咖啡（全品类增速第一）", NAVY),
              ("-18.84", "%", "三合一速溶（持续萎缩）", INK_SOFT)],
        conclusion=[T("2021—2025 年咖啡液增长率达到 "), T("621.74%", c=GOLD_LT)],
        note="统计区间：2021—2025 年；数值与原表格完全一致",
    )


def s05(slide, idx):
    reset(slide)
    market_page(
        slide, idx,
        eyebrow="01 / 项目背景 · 市场分析",
        title=[T("咖啡液年度消费产值增至 "), T("94.8 亿元", c=RED_HI)],
        img="image5.PNG",
        kpis=[("94.8", "亿元", "2025 年咖啡液年度消费产值", RED_HI),
              ("15.2", "亿元", "2021 年基数", NAVY_MID),
              ("连续五年", "", "2021—2025 年逐年递增", NAVY)],
        conclusion=[T("截至 2025 年，咖啡液年度消费产值达到 "),
                    T("94.8 亿元", c=GOLD_LT)],
        note="统计区间：2021—2025 年；数值与原图表完全一致",
    )


# ------------------------------------------------------------------ 09 产业现状
def s09(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "02 / 行业现状 · 产业现状",
         [T("咖啡液加工路径中，"), T("灭菌", c=RED_HI),
          T("是决定货架期与品质的关键工序")])
    textbox(slide, M, RULE_Y + 0.12, CW, 0.28,
            [P([T("咖啡行业市场分析 —— 市场萃取与灭菌技术分析", sz=13, c=INK_MID)])])

    # 工艺链：四道工序，关键工序实心强调
    steps = [("咖啡粉", None), ("高温萃取", NAVY), ("灭菌", CRIMSON), ("灌装", None)]
    cwid, gap = 2.05, 0.85
    total = 4 * cwid + 3 * gap
    x0, cy = (SW - total) / 2, 1.62
    for k, (txt, solid) in enumerate(steps):
        x = x0 + k * (cwid + gap)
        if solid:
            chip(slide, x, cy, cwid, 0.50, [T(txt, sz=15)], fill=solid)
        else:
            chip(slide, x, cy, cwid, 0.50, [T(txt, sz=15)],
                 fill=WHITE, color=NAVY, line_color=BLUE_LINE)
        if k < 3:
            line(slide, x + cwid + 0.16, cy + 0.25, x + cwid + gap - 0.14,
                 cy + 0.25, color=NAVY_MID, lw=1.5, arrow=True)

    # 两极现状对照
    ph_t, ph_h, pw = 2.34, 3.58, 5.90
    for k, (key, cap, bias) in enumerate([
        ("image6.jpeg", "分散式小批量加工（釜式间歇）", 0.55),
        ("image7.jpeg", "规模化连续加工生产线", 0.50),
    ]):
        x = M + k * (pw + 0.29)
        card(slide, x, ph_t, pw, ph_h, radius=0.03)
        pic_cover(slide, key, x + 0.14, ph_t + 0.14, pw - 0.28, ph_h - 0.86,
                  bias_y=bias)
        bar = rect(slide, x + 0.14, ph_t + ph_h - 0.64, pw - 0.28, 0.50, fill=NAVY)
        fill_text(bar, [P([T(cap, sz=14)], al="c")], b=True, c=WHITE, anchor="m")

    takeaway(slide, 6.10, [
        T("咖啡液加工路径："), T("磨粉 → 高温萃取 → 灭菌 → 灌装", c=GOLD_LT),
        T("；其中 "), T("灭菌", c=GOLD_LT), T(" 是决定产品货架期和品质的关键工序"),
    ], h=0.72, sz=16)
    foot(slide, idx)


# ------------------------------------------------------------------ 10 存在问题
def s10(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "02 / 行业现状 · 存在问题",
         [T("传统热杀菌的代价："), T("风味损失、营养破坏与高能耗", c=RED_HI)])

    # 工艺要求
    band = rect(slide, M, 1.30, CW, 0.70, fill=BLUE_TINT,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.10)
    rect(slide, M, 1.30, 0.06, 0.70, fill=NAVY)
    fill_text(band, [P([
        T("工艺要求　", sz=14.5, b=True, c=NAVY),
        T("为了延长产品质保期，获得更高的商业价值，需要对产品进行符合国家要求的"
          "无菌化消毒处理，同时尽量保存咖啡的风味物质。", sz=13.5, c=INK),
    ], ls=1.3)], anchor="m", pad=(0.26, 0.24, 0.04, 0.04))

    cards = [
        ("巴氏杀菌", "", "温度", "70-90℃",
         [T("特点：", c=INK_MID), T("风味损失少", c=NAVY), T("，"),
          T("质保期短", c=RED_HI)],
         "此方法只能将细菌的数量下降至不会引发疾病的水平，不能消灭可以忍受 70°C "
         "以上的细菌孢子，在牛奶中仍然有微生物存在，一旦温度上升至 10°C 或以上，"
         "细菌孢子又会变得活跃，细菌加快繁殖使牛奶腐坏，因此需要将牛奶维持在 "
         "4-5°C 的低温保存及运输"),
        ("UHT", "超高温瞬时杀菌", "温度", "135-150℃",
         [T("特点：", c=INK_MID), T("风味损失多", c=RED_HI), T("，"),
          T("质保期长", c=NAVY)],
         "UHT 的加热时间更短，能够更好地保持风味。即使这样，UHT 处理的温度也足够"
         "造成梅纳反应，会改变乳品的风味和色泽。经过 UHT 处理的牛奶储存在无菌容器中，"
         "在常温情况下有 6—9 个月的保质期"),
        ("HPP", "超高压杀菌", "压力", "100-1000MPA",
         [T("特点：", c=INK_MID), T("风味损失少、质保期长", c=NAVY), T("，"),
          T("成本高", c=RED_HI)],
         "食品超高压杀菌技术是以水或其他液体介质为传递压力的媒介物质，加在液体中的"
         "压力（100～1000MPa），通过介质，以压力作为能量因子，将放在专门密封超高压"
         "容器内的食品，在常温或者低温 (低于 100℃) 下对食品加压."),
    ]
    ct, ch, cwid, gap = 2.18, 3.66, 3.843, 0.28
    for k, (name, sub, klab, kval, feat, desc) in enumerate(cards):
        x = M + k * (cwid + gap)
        card(slide, x, ct, cwid, ch, radius=0.035)
        hd = [T(name, sz=17)]
        if sub:
            hd.append(T("  " + sub, sz=13, c=GOLD_LT))
        card_head(slide, x, ct, cwid, 0.60, hd, al="c")

        textbox(slide, x + 0.22, ct + 0.74, cwid - 0.44, 0.34,
                [P([T(klab + "　", sz=12, c=INK_SOFT),
                    T(kval, sz=19, b=True, c=NAVY)], ls=1.0)])
        rect(slide, x + 0.22, ct + 1.16, cwid - 0.44, 0.012, fill=BLUE_LINE)
        textbox(slide, x + 0.22, ct + 1.26, cwid - 0.44, 0.58,
                [P(feat, ls=1.3)], sz=13.5, b=True)
        textbox(slide, x + 0.22, ct + 1.92, cwid - 0.44, ch - 2.10,
                [P([T(desc, sz=10.8, c=INK_MID)], al="j", ls=1.42)])

    takeaway(slide, 6.02, [
        T("传统咖啡液加工主要为热杀菌技术，包括巴氏杀菌（TP）、高温短时杀菌（HTST）、"
          "超高温瞬时杀菌（UHT）等，导致"),
        T("风味损失、营养破坏", c=GOLD_LT),
        T("等问题；同时咖啡液加工为高耗能产业，"),
        T("耗水、耗电、耗气等能耗为发达国家 1.9 倍", c=GOLD_LT),
        T("。"),
    ], h=0.86, sz=15)
    foot(slide, idx)


# ------------------------------------------------------------------ 11 解决方法
def s11(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "02 / 行业现状 · 解决方法",
         [T("超声非热辅助杀菌："), T("高效杀菌", c=RED_HI), T(" ＋ "),
          T("风味保留", c=RED_HI)])

    bt, bb = 1.30, 5.90
    lw = 8.10
    card(slide, M, bt, lw, bb - bt, radius=0.03)
    pic_fit(slide, "image8.png", M + 0.18, bt + 0.18, lw - 0.36, 3.42)
    textbox(slide, M + 0.18, bt + 3.66, lw - 0.36, 0.28,
            [P([T("超声探头式与超声水浴式杀菌作用原理", sz=12, c=INK_SOFT)], al="c")])
    tags = ["非热辅助杀菌", "空化效应", "低温处理"]
    tw = (lw - 0.36 - 2 * 0.16) / 3
    for k, tg in enumerate(tags):
        chip(slide, M + 0.18 + k * (tw + 0.16), bt + 4.02, tw, 0.44,
             [T(tg, sz=13)], fill=BLUE_TINT, color=NAVY, line_color=BLUE_LINE)

    rx, rw = M + lw + 0.26, CW - lw - 0.26
    # 图片按语义归位：微生物照片/插画归"安全"，分子结构归"营养"
    blocks = [
        (GREEN, "安全", "灭活有害微生物", ["image9.png", "image12.jpeg"]),
        (NAVY_MID, "营养", "保留原有风味", ["image10.jpeg", "image11.jpeg"]),
    ]
    bh = (bb - bt - 0.16) / 2
    for k, (acc, tag, txt, keys) in enumerate(blocks):
        y = bt + k * (bh + 0.16)
        card(slide, rx, y, rw, bh, radius=0.05)
        rect(slide, rx, y, rw, 0.52, fill=acc, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE,
             adj=0.12)
        hd = slide.shapes[-1]
        fill_text(hd, [P([T(tag, sz=15, b=True, c=WHITE),
                          T("　" + txt, sz=13.5, b=True, c=WHITE)])],
                  anchor="m", pad=(0.20, 0.14, 0, 0))
        iw = (rw - 0.36 - 0.14) / 2
        ih = bh - 0.52 - 0.34
        for j, key in enumerate(keys):
            pic_cover(slide, key, rx + 0.18 + j * (iw + 0.14), y + 0.68, iw, ih,
                      geom="roundRect", adj=0.04)

    takeaway(slide, 6.06, [
        T("超声波", c=GOLD_LT),
        T("作为非热辅助杀菌技术，顺应国家 "),
        T("“食品洁净”和“最少加工”", c=GOLD_LT),
        T(" 的大健康趋势，既能高效杀菌，且能有效保留原有风味。"),
    ], h=0.74, sz=16)
    foot(slide, idx)


# ------------------------------------------------------------------ 13 产品介绍
def s13(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "03 / 核心优势 · 产品介绍",
         [T("产品介绍 · "), T("SK-Ⅱ 萃取与超声杀菌一体化装备", c=NAVY)])

    bt, bh = 1.22, 5.46
    lw = 5.10
    card(slide, M, bt, lw, bh, radius=0.03)
    chip(slide, M + 0.18, bt + 0.16, 1.90, 0.44, [T("产品 SK-Ⅱ", sz=14)], fill=NAVY)
    # 原图外框为深色圆角描边，沿用原稿裁切把边框裁掉
    pic_fit(slide, "image13.png", M + 0.22, bt + 0.74, lw - 0.44, bh - 1.50,
            crop=(0.168, 0.144, 0.110, 0.098))
    bar = rect(slide, M + 0.30, bt + bh - 0.66, lw - 0.60, 0.50, fill=NAVY,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.06)
    fill_text(bar, [P([T("产品模型图", sz=14.5)], al="c")], b=True, c=WHITE, anchor="m")

    rx = M + lw + 0.37
    pic_fit(slide, "image15.png", rx, bt, CW - lw - 0.37, bh,
            geom="roundRect", adj=0.012, line_color=BLUE_LINE)
    foot(slide, idx)


# ------------------------------------------------------------------ 16 核心技术 1
def s16(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "03 / 核心优势 · 核心技术 1",
         [T("核心技术 1 · "), T("杀菌温度低，风味保留度好", c=RED_HI)])

    bt, bb = 1.30, 6.52
    lw = 4.60
    rx = M + lw + 0.20
    rw = CW - lw - 0.20

    # 左上：杀菌温度对照
    h1 = 2.52
    card(slide, M, bt, lw, h1, radius=0.035)
    pic_fit(slide, "image20.png", M + 0.18, bt + 0.16, lw - 0.36, h1 - 0.66,
            crop=(0.088, 0.081, 0.0, 0.037))
    textbox(slide, M + 0.18, bt + h1 - 0.46, lw - 0.36, 0.30,
            [P([T("杀菌温度对照：", sz=12, c=INK_SOFT),
                T("85℃ → 56℃", sz=13, b=True, c=RED_HI)], al="c")])

    # 左下：第三方检测报告
    y2 = bt + h1 + 0.18
    h2 = bb - y2
    card(slide, M, y2, lw, h2, radius=0.035)
    card_head(slide, M, y2, lw, 0.52, [T("第三方检验检测报告", sz=14.5)], al="c")
    iw = (lw - 0.36 - 2 * 0.12) / 3
    for k, key in enumerate(["image18.jpeg", "image17.jpeg", "image16.jpeg"]):
        pic_fit(slide, key, M + 0.18 + k * (iw + 0.12), y2 + 0.64, iw,
                h2 - 1.04, line_color=BLUE_LINE)
    textbox(slide, M + 0.18, y2 + h2 - 0.36, lw - 0.36, 0.28,
            [P([T("天津市食品安全检测技术研究院", sz=11, c=INK_SOFT)], al="c")])

    # 右上：技术说明
    th = 1.58
    textbox(slide, rx, bt + 0.02, rw, 0.40,
            [P([T("突破传统热杀菌限制，实现低温高效灭菌", sz=18, b=True, c=NAVY)])])
    textbox(slide, rx, bt + 0.50, rw, th - 0.50,
            [P([T("超声杀菌技术通过物理能场作用，破坏微生物细胞结构，降低杀菌温度，"
                  "实现高效杀菌。在 "),
                T("56℃", b=True, c=RED_HI),
                T(" 低温条件下达到良好杀菌效果，同时显著降低风味物质损失，"
                  "提升产品品质，实现营养、香气及口感的有效保留。")],
               al="j", ls=1.52)], sz=14.5, c=INK)

    # 右下：GC-MS 风味物质对照
    gy = bt + th + 0.20
    gh = bb - gy
    card(slide, rx, gy, rw, gh, radius=0.035)
    card_head(slide, rx, gy, rw, 0.52,
              [T("风味物质总离子流色谱对照（超声杀菌 / 传统热杀菌）", sz=14)], al="c")
    pic_fit(slide, "image19.png", rx + 0.18, gy + 0.64, rw - 0.36, gh - 0.82)
    foot(slide, idx)


# ------------------------------------------------------------------ 17 核心技术 2
def s17(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "03 / 核心优势 · 核心技术 2",
         [T("核心技术 2 · "), T("多级工艺整合实现节能降耗", c=RED_HI)])

    band = rect(slide, M, 1.28, CW, 1.12, fill=BLUE_TINT,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.06)
    rect(slide, M, 1.28, 0.06, 1.12, fill=NAVY)
    fill_text(band, [
        P([T("工艺协同　", sz=14.5, b=True, c=NAVY),
           T("高效萃取 ＋ 低温超声杀菌", sz=14.5, b=True, c=RED_HI),
           T("，实现低温高效处理，大幅降低传统热杀菌与萃取的高能耗",
             sz=14, c=INK)], ls=1.3, sa=4),
        P([T("智能匹配　", sz=14.5, b=True, c=NAVY),
           T("传感闭环 ＋ PLC 精准调控", sz=14.5, b=True, c=RED_HI),
           T("，实现按需供能与动态匹配，综合降低运行能耗 ", sz=14, c=INK),
           T("20%~30%", sz=14.5, b=True, c=RED_HI)], ls=1.3),
    ], anchor="m", pad=(0.28, 0.24, 0.06, 0.06))

    ct, cb = 2.58, 6.34
    side_w, hub_w = 3.90, 3.30
    hub_x = M + side_w + 0.50
    rx = hub_x + hub_w + 0.50
    rh = (cb - ct - 0.16) / 2

    # 中枢：技术集成
    card(slide, hub_x, ct, hub_w, cb - ct, fill=PAPER, radius=0.04)
    pic_fit(slide, "image13.png", hub_x + 0.22, ct + 0.24, hub_w - 0.44,
            cb - ct - 1.00, crop=(0.143, 0.127, 0.165, 0.143))
    hb = rect(slide, hub_x + 0.45, cb - 0.68, hub_w - 0.90, 0.52, fill=NAVY,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.08)
    fill_text(hb, [P([T("技术集成", sz=15.5, spc=1.2)], al="c")], b=True, c=WHITE,
              anchor="m")

    mods = [
        (M, ct, "工业级萃取泵", "高效萃取", ["image27.png"]),
        (M, ct + rh + 0.16, "超声杀菌系统", "低温超声杀菌",
         ["image21.png", "image22.jpeg"]),
        (rx, ct, "温度、流量传感系统", "传感闭环", ["image23.png", "image24.png"]),
        (rx, ct + rh + 0.16, "PLC 管控系统", "精准调控",
         ["image25.png", "image26.png"]),
    ]
    for x, y, name, note, keys in mods:
        card(slide, x, y, side_w, rh, radius=0.045)
        ibw = 1.62
        if len(keys) == 1:
            pic_fit(slide, keys[0], x + 0.14, y + 0.16, ibw, rh - 0.32)
        else:
            sw = (ibw - 0.08) / 2
            for j, key in enumerate(keys):
                pic_fit(slide, key, x + 0.14 + j * (sw + 0.08), y + 0.16, sw,
                        rh - 0.32)
        tx = x + ibw + 0.26
        textbox(slide, tx, y + rh / 2 - 0.42, side_w - ibw - 0.40, 0.44,
                [P([T(name, sz=15, b=True, c=NAVY)], ls=1.15)], anchor="b")
        rect(slide, tx, y + rh / 2 + 0.06, 0.42, 0.022, fill=CRIMSON)
        textbox(slide, tx, y + rh / 2 + 0.18, side_w - ibw - 0.40, 0.34,
                [P([T(note, sz=12.5, c=INK_MID)])])
        # 指向中枢的连接线
        if x < hub_x:
            line(slide, x + side_w + 0.08, y + rh / 2, hub_x - 0.08, y + rh / 2,
                 color=NAVY_MID, lw=1.5, arrow=True)
        else:
            line(slide, x - 0.08, y + rh / 2, hub_x + hub_w + 0.08, y + rh / 2,
                 color=NAVY_MID, lw=1.5, arrow=True)
    foot(slide, idx)


# ------------------------------------------------------------------ 18 核心技术 3
def s18(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "03 / 核心优势 · 核心技术 3",
         [T("核心技术 3 · "), T("萃取与超声杀菌一体化装备", c=RED_HI)])

    band = rect(slide, M, 1.28, CW, 0.96, fill=BLUE_TINT,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.07)
    rect(slide, M, 1.28, 0.06, 0.96, fill=NAVY)
    fill_text(band, [P([
        T("本装备将"), T("萃取混合", b=True, c=NAVY), T("与"),
        T("超声空化灭菌", b=True, c=NAVY),
        T("两大核心技术集于一体，采用全封闭管道化连续作业模式，实现物料萃取与杀菌"
          "工艺的连续化闭环运行，在有效保留物料原有品质的同时大幅降低系统综合能耗。"),
    ], ls=1.42)], sz=14.5, c=INK, anchor="m", pad=(0.28, 0.26, 0.06, 0.06))

    # 流程图为宽幅图，卡片拉通整幅宽度，图形在卡内等比居中（不裁切）
    dy, dh = 2.32, 4.24
    card(slide, M, dy, CW, dh, radius=0.03)
    card_head(slide, M, dy, CW, 0.50,
              [T("萃取 — 超声杀菌一体化工艺流程图", sz=14)], al="c")
    pic_fit(slide, "image28.png", M + 0.18, dy + 0.62, CW - 0.36, dh - 0.80,
            crop=(0.0, 0.0, 0.032, 0.0))
    foot(slide, idx)


# ------------------------------------------------------------------ 20 团队成员
TEAM = [
    ("image29.png", "苏  杭", "项目负责人", "机械工程专业研究生",
     "获第十七届“挑战杯”天津科技大学大学生课外学术科技作品竞赛市赛一等奖；"
     "获天津市知识产权创新创业发明与设计大赛“超声联合杀菌装置”校赛一等奖"),
    ("image31.jpeg", "袁祎晨", None, "能源动力专业研究生",
     "获“初级工程师”证书；曾参加多项学科知识比赛；曾获国家电网北方客服中心优秀见习生。"),
    ("image30.jpeg", "雷国豪", None, "机械工程专业研究生",
     "曾获第一届天津市节能减排社会实践与科技竞赛市赛三等奖；熟练运用各类办公软件。"),
    ("image34.jpeg", "李丽梅", None, "机械专业博士生",
     "发表论文 6 篇，授权发明专利 9 篇"),
    ("image33.jpeg", "张善槟", None, "能源动力专业研究生",
     "获天津科技大学第一届校友创新创业大赛学生组二等奖"),
    ("image32.jpeg", "马建斌", None, "能源动力专业研究生",
     "获天津科技大学第一届校友创新创业大赛学生组二等奖"),
]


def s20(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "04 / 团队成员", [T("研究生团队 · "), T("机械 ＋ 能源动力交叉配置", c=NAVY)])

    bt, bb, rgap = 1.30, 6.58, 0.20
    cwid, gap = 3.843, 0.28
    rh = (bb - bt - rgap) / 2
    for k, (key, name, role, major, ach) in enumerate(TEAM):
        col, row = k % 3, k // 3
        x = M + col * (cwid + gap)
        y = bt + row * (rh + rgap)
        card(slide, x, y, cwid, rh, radius=0.04)
        if role:
            rect(slide, x, y + 0.10, 0.06, rh - 0.20, fill=CRIMSON)
        d = 1.16
        pic_cover(slide, key, x + 0.20, y + 0.18, d, d, bias_y=0.34, geom="ellipse")
        tx = x + 0.20 + d + 0.22
        twid = cwid - (tx - x) - 0.18
        textbox(slide, tx, y + 0.20, twid, 0.38,
                [P([T(name, sz=17.5, b=True, c=NAVY, spc=0.6)], ls=1.0)])
        if role:
            chip(slide, tx, y + 0.62, 1.30, 0.32, [T(role, sz=11.5)], fill=GOLD)
            textbox(slide, tx, y + 0.98, twid, 0.30,
                    [P([T(major, sz=12.5, c=INK_MID)])])
        else:
            textbox(slide, tx, y + 0.64, twid, 0.32,
                    [P([T(major, sz=13, c=INK_MID)])])
        rect(slide, x + 0.20, y + 1.44, cwid - 0.40, 0.012, fill=BLUE_LINE)
        textbox(slide, x + 0.20, y + 1.54, cwid - 0.40, rh - 1.66,
                [P([T(ach, sz=10.4, c=INK_MID)], al="j", ls=1.38)])
    foot(slide, idx)


# ------------------------------------------------------------------ 21 指导教师
MENTORS = [
    ("image35.jpeg", "宋继田", "理工科负责人",
     "天津科技大学教授；博士生导师；全国蒸发结晶学组专家",
     "天津市轻工工程学会副理事长，天津市节能协会专家组成员，发表学术论文 90 余篇，"
     "主持完成省部级和横向科研项目 20 余项，申请并授权专利近 50 项。荣获 2019 年"
     "天津市专利优秀奖；获第三届中国创新挑战赛（宁夏）优胜奖（一等奖）。"),
    ("image36.png", "李金茹", "人文社科负责人",
     "天津科技大学副教授；硕士生导师",
     "天津市无形资产研究会理事，承担、参与多项各级科研课题，发表论文 20 余篇，"
     "出版专著 1 部，主编出版等多部教材。撰写的教学案例入选中国专业学位案例库，"
     "并荣获中国专业学位案例中心优秀案例。擅长案例教学，承担多项学校研究生案例"
     "教学教改项目，注重校企合作。"),
]


def s21(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "04 / 团队成员", [T("指导教师 · "), T("理工 ＋ 人文社科双负责人", c=NAVY)])

    bt = 1.34
    ch = 2.52
    for k, (key, name, role, title, bio) in enumerate(MENTORS):
        y = bt + k * (ch + 0.24)
        card(slide, M, y, CW, ch, radius=0.035)
        rect(slide, M, y + 0.12, 0.06, ch - 0.24, fill=NAVY)
        pw_, ph_ = 1.58, ch - 0.44
        pic_cover(slide, key, M + 0.26, y + 0.22, pw_, ph_, bias_y=0.38,
                  geom="roundRect", adj=0.05, line_color=BLUE_LINE)
        tx = M + 0.26 + pw_ + 0.34
        twid = CW - (tx - M) - 0.30
        textbox(slide, tx, y + 0.24, twid, 0.44,
                [P([T(name, sz=23, b=True, c=NAVY, spc=1.2)], ls=1.0)])
        chip(slide, tx + 1.62, y + 0.32, 2.05, 0.34, [T(role, sz=12)], fill=GOLD)
        textbox(slide, tx, y + 0.78, twid, 0.34,
                [P([T(title, sz=15, b=True, c=INK)])])
        rect(slide, tx, y + 1.18, twid, 0.012, fill=BLUE_LINE)
        textbox(slide, tx, y + 1.30, twid, ch - 1.52,
                [P([T(bio, sz=13, c=INK_MID)], al="j", ls=1.52)])
    foot(slide, idx)


# ------------------------------------------------------------------ 23 未来规划
ROADMAP = [
    ("2026", "天津", "成立公司",
     ["打开市场积累技术经验", "将“萃取＋超声杀菌”落到实处进行中试运行"]),
    ("2027", "华北", "一体化设备开发",
     ["开发出萃取、超声低温杀菌一体化装备", "树立企业品牌　做行业先行者"]),
    ("2028", "全国", "更新产品技术",
     ["完成咖啡萃取杀菌的“交钥匙”工程"]),
]
STEP_TINT = [BLUE, NAVY_MID, NAVY]


def s23(slide, idx):
    from theme import page_frame

    reset(slide)
    page_frame(slide)
    head(slide, "05 / 未来规划", [T("三年三步走 · "), T("从天津到华北再到全国", c=NAVY)])

    cwid, gap = 3.843, 0.28
    tops = [5.10, 4.26, 3.42]      # 台阶上沿，等距上升
    riser, plat_h, ground = 0.26, 0.28, 5.86

    # 左上角补一张"扩张路径"小卡，平衡阶梯造成的空白
    band = rect(slide, M, 1.34, cwid, 0.70, fill=BLUE_TINT,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.10)
    rect(slide, M, 1.34, 0.055, 0.70, fill=NAVY)
    fill_text(band, [P([T("扩张路径　", sz=12.5, b=True, c=NAVY),
                        T("天津 → 华北 → 全国", sz=13.5, b=True, c=INK)])],
              anchor="m", pad=(0.24, 0.16, 0, 0))

    for k in range(3):
        x = M + k * (cwid + gap)
        y = tops[k]
        base = tops[k - 1] if k else ground
        rect(slide, x, y, riser, base - y + plat_h, fill=STEP_TINT[k])   # 竖向踏面
        rect(slide, x, y, cwid + gap, plat_h, fill=STEP_TINT[k])         # 水平台面

        year, region, title, lines = ROADMAP[k]
        mh = 1.62
        my = y - 0.22 - mh
        mcard = card(slide, x, my, cwid, mh, radius=0.04)
        rect(slide, x, my, cwid, 0.055, fill=STEP_TINT[k])
        paras = [P([T(title, sz=17, b=True, c=NAVY)], ls=1.1, sa=6)]
        for ln in lines:
            paras.append(P([T("· " + ln, sz=12.5, c=INK_MID)], ls=1.38, sa=2))
        fill_text(mcard, paras, anchor="m", pad=(0.22, 0.20, 0.12, 0.10))

        ax = x + riser + 0.72
        chip(slide, ax - 0.50, y + plat_h + 0.14, 1.00, 0.38,
             [T(region, sz=13)], fill=WHITE, color=NAVY, line_color=NAVY)
        line(slide, ax, y + plat_h + 0.56, ax, 6.14,
             color=NAVY_MID, lw=1.25, dash="dash", arrow=True)
        textbox(slide, ax - 1.20, 6.22, 2.40, 0.46,
                [P([T(year, sz=25, b=True, c=NAVY), T(" 年", sz=15, b=True, c=NAVY)],
                   al="c", ls=1.0)], anchor="m")
    foot(slide, idx)


# ------------------------------------------------------------------ 24 致谢
def s24(slide, shapes):
    reset(slide)
    clone(slide, shapes["para"], 6.30, -0.014, 7.03, 7.50)
    clone(slide, shapes["strip"], 6.30, -0.014, 3.04, 7.51)
    rect(slide, 0, 0, 0.13, SH, fill=CRIMSON)

    rect(slide, M + 0.18, 2.70, 1.30, 0.06, fill=CRIMSON)
    textbox(slide, M + 0.18, 2.92, 5.20, 0.82,
            [P([T("感谢各位专家聆听", sz=40, b=True, c=NAVY, spc=1.6)], ls=1.0)],
            anchor="t")
    textbox(slide, M + 0.18, 3.82, 5.20, 0.70,
            [P([T("THANKS", sz=40, b=True, c=BLUE, spc=8.0)], ls=1.0)], anchor="t")
    rect(slide, M + 0.18, 4.66, 4.40, 0.012, fill=BLUE_LINE)
    textbox(slide, M + 0.18, 4.82, 5.20, 0.80,
            [P([T("咖啡液萃取与超声杀菌工艺及装备设计研究设备", sz=14, b=True, c=INK)],
               sa=4),
             P([T("天津科技大学　·　2026.09.14", sz=12.5, c=INK_SOFT)])], anchor="t")
    pic_fit(slide, "image2.jpg", M + 0.18, 6.05, 2.70, 0.68)


# ------------------------------------------------------------------ main
def main():
    WORK.mkdir(parents=True, exist_ok=True)
    extract_media()
    set_rule_image(MED["image14.png"][0])
    USED.add("image14.png")
    prs = Presentation(str(SRC))
    slides = list(prs.slides)
    assert len(slides) == 24, len(slides)

    # 先取出分节页的斜切几何，供 5 个分节页与致谢页复用
    shapes = {}
    for sp in slides[1].shapes:
        if sp.name == "平行四边形 7":
            shapes["para"] = deepcopy(sp._element)
        elif sp.name == "矩形 8":
            shapes["strip"] = deepcopy(sp._element)
    assert "para" in shapes and "strip" in shapes

    s01(slides[0])
    for n, (num, cn, en, lead) in DIVIDERS.items():
        divider(slides[n - 1], shapes, num, cn, en, lead)
    for n, (eb, title, label) in PLACEHOLDERS.items():
        placeholder_page(slides[n - 1], n, eb, title, label)
    s03(slides[2], 3)
    s04(slides[3], 4)
    s05(slides[4], 5)
    s09(slides[8], 9)
    s10(slides[9], 10)
    s11(slides[10], 11)
    s13(slides[12], 13)
    s16(slides[15], 16)
    s17(slides[16], 17)
    s18(slides[17], 18)
    s20(slides[19], 20)
    s21(slides[20], 21)
    s23(slides[22], 23)
    s24(slides[23], shapes)

    renumber(prs)
    OUTDIR.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))

    # 校验：原始位图是否全部仍在使用
    missing = sorted(set(MED) - USED)
    print("图片素材 %d 个，已复用 %d 个" % (len(MED), len(USED)))
    if missing:
        print("!! 未被复用：", missing)
    else:
        print("原始图片全部复用（hdphoto1.wdp 为 image22.jpeg 的 HD 副本图层，"
              "随图片一并保留在包内）")
    print("输出：", OUT)


if __name__ == "__main__":
    main()
