# -*- coding: utf-8 -*-
"""商用自动奶泡系统 竞品分析 PPT — 科研蓝白风"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from theme import *  # noqa

A = "assets_min/%s.jpg"
prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)

L = PP_ALIGN.LEFT
R = PP_ALIGN.RIGHT
C = PP_ALIGN.CENTER


# =====================================================================
# 01 封面
# =====================================================================
s = new_slide(prs)
rect(s, 7.05, 0, W - 7.05, H, fill=TINT)
rect(s, 0, 0, W, 0.13, fill=NAVY)

label(s, 0.85, 1.02, 5.8, "COMMERCIAL MILK SYSTEMS / CHINA FOCUS", 10, True, CYAN, spc=1.6)
label(s, 0.85, 1.42, 6.0, "商用自动奶泡系统", 38, True, NAVY, h=0.75)
label(s, 0.85, 2.18, 6.0, "竞品分析报告", 38, True, NAVY, h=0.75)
rect(s, 0.87, 3.10, 1.30, 0.055, fill=CYAN)
label(s, 0.85, 3.34, 5.9,
      "7 个型号　·　5 个品牌　·　9 条可核实价格记录　·　中国市场优先",
      12.5, True, BLUE, h=0.3)

kpis = [("2600–3400", "W", "样本额定功率区间"),
        ("120–250", "杯/h", "标称出品能力"),
        ("¥36.8k–60k", "", "跨市场价格带")]
for i, (num, unit, cap) in enumerate(kpis):
    x = 0.85 + i * 2.02
    vline(s, x - 0.13, 4.08, 0.86, LINE, 1.0) if i else None
    rich(s, x, 4.10, 1.95, 0.42,
         [(num, 19, True, NAVY), ("  " + unit, 10.5, True, CYAN)])
    label(s, x, 4.60, 1.95, cap, 9.5, False, GRAY, h=0.26)

label(s, 0.85, 5.42, 5.9,
      "阿玛菲 AMF-N1 / N2 / N2T　·　Linkbar MilkPal / Single Touch Milk　·　"
      "Marco MilkPal　·　Übermilk ONE",
      10, False, GRAY, h=0.5, line=1.5)
hline(s, 0.85, 6.08, 5.9, LINE, 0.75)
label(s, 0.85, 6.22, 5.9,
      "产品开发与市场定位　|　统一检索日 2026-09-18\n"
      "参数与图片取自各品牌官方产品页、规格表与说明书；价格按原币种与结算口径登记",
      8.8, False, LGRAY, h=0.6, line=1.5)

label(s, 7.62, 0.62, 5.2, "样本机型 / SPECIMEN PLATE", 9.5, True, BLUE, spc=1.2)
hline(s, 7.62, 0.92, 5.12, LINE, 0.75)
cards = [("amalfi_n1", "阿玛菲 AMF-N1", "台上整机 · 2600W"),
         ("amalfi_n2t", "阿玛菲 AMF-N2T", "台下系统 · 双拨杆龙头"),
         ("marco_milkpal", "Marco MilkPal", "爱尔兰 · 3100W"),
         ("ubermilk_front", "Übermilk ONE", "德国 · 250 杯/h")]
for i, (img, name, sub) in enumerate(cards):
    cx = 7.62 + (i % 2) * 2.66
    cy = 1.18 + (i // 2) * 2.92
    rect(s, cx, cy, 2.46, 2.70, fill=WHITE, line_color=LINE)
    place_img(s, A % img, cx + 0.10, cy + 0.10, 2.26, 1.90, pad=0.02)
    rect(s, cx + 0.14, cy + 2.05, 0.38, 0.035, fill=CYAN)
    label(s, cx + 0.14, cy + 2.16, 2.20, name, 10, True, NAVY, h=0.24)
    label(s, cx + 0.14, cy + 2.40, 2.20, sub, 8.5, False, GRAY, h=0.22)

cover_footer(s, 1)


# =====================================================================
# 02 竞品版图
# =====================================================================
s = new_slide(prs)
header(s, "01", "竞品版图：7 个型号，两种形态，三条价格口径",
       "国内 5 个型号定义主战场，海外 2 台提供运营与文档基准；MilkPal 按型号分列，Marco 与 Linkbar 共享同源平台")

grid = [
    ("amalfi_n1", "阿玛菲 AMF-N1", "台上整机", "2600W · 10.5kg · 390×180×340", "contain", 0.5),
    ("amalfi_n2", "阿玛菲 AMF-N2", "台下系统 + 触屏龙头", "2600W · 130×240×260（台上件）", "contain", 0.5),
    ("amalfi_n2t", "阿玛菲 AMF-N2T", "台下系统 + 双拨杆龙头", "2600W · 100×220×280（台上件）", "contain", 0.5),
    ("milkpal_prod", "Linkbar MilkPal", "台上自动供奶/发泡", "3000W · 30kg · 120 杯/h", "contain", 0.50),
    ("linkbar_st_crop", "Linkbar Single Touch Milk", "分体台下主机 + 台上龙头", "3400W · 19.4kg · 120 杯/h", "cover", 0.20),
    ("marco_bar", "Marco MilkPal 1000090", "台上自动供奶/发泡", "3100W · 30kg · 120 份/h", "cover", 0.50),
    ("ubermilk_front", "Übermilk ONE", "台上自动微泡系统", "3100W · 25kg · 峰值 250 杯/h", "contain", 0.50),
]
cw_, gap = 2.845, 0.238
for i, (img, name, form, spec, mode, focus) in enumerate(grid):
    cx = ML + (i % 4) * (cw_ + gap)
    cy = 1.62 + (i // 4) * 2.62
    rect(s, cx, cy, cw_, 2.42, fill=TINT2, line_color=LINE)
    if mode == "cover":
        place_img(s, A % img, cx + 0.001, cy + 0.001, cw_ - 0.002, 1.42, mode="cover", focus=focus)
    else:
        rect(s, cx, cy, cw_, 1.42, fill=WHITE)
        place_img(s, A % img, cx, cy, cw_, 1.42, pad=0.07)
    hline(s, cx, cy + 1.42, cw_, LINE, 0.75)
    label(s, cx + 0.13, cy + 1.54, cw_ - 0.26, name, 10.5, True, NAVY, h=0.24)
    label(s, cx + 0.13, cy + 1.80, cw_ - 0.26, form, 8.6, False, CYAN, h=0.22)
    label(s, cx + 0.13, cy + 2.03, cw_ - 0.26, spec, 8.8, True, GRAY, h=0.24)

nx = ML + 3 * (cw_ + gap)
rect(s, nx, 4.24, cw_, 2.42, fill=NAVY)
label(s, nx + 0.20, 4.44, cw_ - 0.40, "版图怎么读", 12, True, WHITE, h=0.3)
rect(s, nx + 0.20, 4.76, 0.36, 0.035, fill=CYAN)
tf = textbox(s, nx + 0.20, 4.90, cw_ - 0.40, 1.66)
for k, t in enumerate(["阿玛菲一次铺开台上与台下三种外观，走国产整机路线。",
                       "Linkbar 同时握住台上 MilkPal 与台下 ST-Milk，并与 Marco 共享 Double Favor 平台。",
                       "Übermilk 单独成一条海外基准，靠数据记录说话。"]):
    para(tf, "— " + t, 8.8, False, TINT, space_after=6, line=1.32, first=(k == 0))

source_note(s, "图片：各品牌官方产品页原图（S01/S03/S04/S05/S07/S11/S12）　·　"
               "参数：官方规格图与说明书（S02/S06/S08/S09/S10/S12）")
cover_footer(s, 2)


# =====================================================================
# 03 核心判断
# =====================================================================
s = new_slide(prs)
header(s, "02", "核心判断：打泡是入场券，运营性才是战场",
       "四台样本已把冷热奶泡与参数预设做成标配，胜负手转到高峰稳定性、清洁人工、奶路维护与国内售后")

concl = [
    ("01", "功能同质化已经发生",
     "N1、MilkPal、ST-Milk、Marco 同时提供热奶 / 热泡 / 冷奶 / 冷泡四种输出；Linkbar 与 Marco 再叠加 "
     "4 类型 + 16 组自定义参数。把研发资源压在“会不会打泡”上，收益递减。"),
    ("02", "价格分成两层，国产握着成本空间",
     "国内可核实锚点落在 ¥36,800；境外整机折算到人民币在 ¥47,900–60,300。同配置国产方案有 25%–40% "
     "的定价余量，用来换装机量与服务密度。"),
    ("03", "清洁与停机决定复购",
     "ST-Milk 每日碱洗、每周酸性深洗；Marco 日终周期约 30 分钟。把日常人工动作压到 5 分钟以内并留下"
     "完成记录，就是一条能被门店验证的卖点。"),
]
for i, (n, t, b) in enumerate(concl):
    y = 1.68 + i * 1.71
    rect(s, ML, y, 8.05, 1.52, fill=TINT2)
    rect(s, ML, y, 0.055, 1.52, fill=CYAN)
    label(s, ML + 0.28, y + 0.20, 0.6, n, 16, True, CYAN)
    label(s, ML + 0.92, y + 0.19, 6.9, t, 13, True, NAVY, h=0.3)
    label(s, ML + 0.92, y + 0.56, 6.95, b, 9.8, False, GRAY, h=0.85, line=1.46)

px = ML + 8.35
place_img(s, A % "linkbar_milkpal_hero", px, 1.68, 3.74, 2.55, mode="cover", focus=0.30)
label(s, px, 4.31, 3.74, "门店实景：MilkPal 落在吧台工作面上（Linkbar 官方图）",
      8.2, False, LGRAY, h=0.24)

rect(s, px, 4.66, 3.74, 2.00, fill=NAVY)
label(s, px + 0.22, 4.84, 3.3, "需求侧背景", 10.5, True, CYAN, spc=0.8)
rich(s, px + 0.22, 5.14, 3.3, 0.45,
     [("36,310", 22, True, WHITE), (" 家", 11, True, TINT)])
label(s, px + 0.22, 5.66, 3.34,
      "瑞幸 2026 年二季度末门店数，环比再增 8.1%。奶咖出杯量持续放大，"
      "把万元级奶泡系统推进单店预算。",
      8.5, False, TINT, h=0.80, line=1.42)

source_note(s, "结论依据：官方规格与说明书比对（S02/S06/S08/S09/S10）；价格记录见第 05 页；"
               "门店数据：Luckin Coffee 2026 年第二季度财报")
cover_footer(s, 3)


# =====================================================================
# 04 参数总览
# =====================================================================
s = new_slide(prs)
header(s, "03", "参数总览：统一口径后的七行对照",
       "功率取各品牌自家口径，尺寸标注轴向，产能统一到 200mL/杯再比较")

hdr = ["型号 / 形态", "额定功率", "重量", "外形尺寸 mm", "标称出品能力"]
cols = [3.30, 1.20, 1.12, 4.40, 2.07]
body = [
    ["阿玛菲 AMF-N1　台上整机", "2600W", "10.5kg", "390×180×340（官网 W×D×H）", "实测口径补齐"],
    ["阿玛菲 AMF-N2　台下+触屏龙头", "2600W", "以铭牌为准", "130×240×260（台上部件）", "实测口径补齐"],
    ["阿玛菲 AMF-N2T　台下+双拨杆龙头", "2600W", "以铭牌为准", "100×220×280（台上部件）", "实测口径补齐"],
    ["Linkbar MilkPal　台上整机", "3000W", "30kg", "534×195×388（官方规格图）", "120 杯/h @200mL"],
    ["Linkbar Single Touch Milk　分体", "3400W", "19.4kg", "龙头 116×137×250　主机 320×500×147", "120 杯/h @200mL"],
    ["Marco MilkPal 1000090　台上整机", "3100W", "30kg", "195×535×388（W×D×H）出水口至托盘 180", "120 份/h @200mL"],
    ["Übermilk ONE　台上整机", "3100W", "25kg", "180×485×540（W×D×H）", "峰值 250 杯/h"],
]
table(s, ML, 1.58, CW, cols, None, hdr, body,
      head_size=10, body_size=8.8, head_h=0.40, row_h=0.42,
      align=[L, C, C, L, L])

thumbs = [("milkpal_prod", 0.50, "Linkbar MilkPal", "3000W　30kg　534×195×388", "120 杯/h @200mL　4 类型 + 16 自定义"),
          ("linkbar_st_crop", 0.50, "Linkbar ST-Milk", "3400W　龙头 1.4kg + 主机 18kg", "120 杯/h @200mL　台面仅留龙头")]
for i, (img, fo, nm, l1, l2) in enumerate(thumbs):
    cx = ML + i * 3.14
    rect(s, cx, 5.52, 2.90, 1.10, fill=WHITE, line_color=LINE)
    place_img(s, A % img, cx + 0.005, 5.525, 1.28, 1.09, mode="cover", focus=fo)
    label(s, cx + 1.40, 5.62, 1.44, nm, 8.6, True, NAVY, h=0.22)
    label(s, cx + 1.40, 5.86, 1.46, l1 + "\n" + l2, 7.5, False, GRAY, h=0.66, line=1.44)

nb = ML + 6.28
rect(s, nb, 5.56, 5.81, 1.06, fill=NAVY)
label(s, nb + 0.20, 5.70, 5.4, "读表规则", 10, True, CYAN, spc=0.8)
label(s, nb + 0.20, 5.96, 5.45,
      "功率取各品牌自家口径：Linkbar 用官方规格图 3000W，Marco 用 220V 规格表 3100W，"
      "ST-Milk 用交付版本 V0.7.3 的 3400W。台上尺寸对应台面部件，整套占地按"
      "“台上件 + 台下主机 + 冷藏柜 + 检修空间”四项求和。",
      8.4, False, TINT, h=0.60, line=1.42)
cover_footer(s, 4)


# =====================================================================
# 05 三个数字分开看
# =====================================================================
s = new_slide(prs)
header(s, "04", "功率、产能、占地：三组数字各自成立",
       "把同一台机器的三个指标分开画，才能看出加热负载、净产能与吧台代价之间的关系")

MODELS = ["阿玛菲 N1", "阿玛菲 N2", "阿玛菲 N2T", "Linkbar MilkPal",
          "Linkbar ST-Milk", "Marco MilkPal", "Übermilk ONE"]
panels = [
    ("额定功率", "W", [2600, 2600, 2600, 3000, 3400, 3100, 3100],
     ["2600", "2600", "2600", "3000", "3400", "3100", "3100"], 3600,
     "3400W 的 ST-Milk 与 2600W 的 N1 之间是峰值加热与制冷配置差，实际电费按连续出杯工况计量。"),
    ("标称出品", "杯/h @200mL", [0, 0, 0, 120, 120, 120, 250],
     ["实测补齐", "实测补齐", "实测补齐", "120", "120", "120", "250*"], 280,
     "三台国产/爱尔兰机型统一在 200mL/杯的 120 杯；Übermilk 的 250 杯按其自有口径记录（*单杯量另计）。"),
    ("台面占地", "cm²（W×D）", [702, 312, 220, 1041, 159, 1043, 873],
     ["702", "312", "220", "1041", "159", "1043", "873"], 1150,
     "ST-Milk 台面只留 159cm² 的龙头，主机入柜；台上整机把 700–1040cm² 直接压在吧台上。"),
]
pw = 3.86
for pi, (title, unit, vals, labs, vmax, note) in enumerate(panels):
    px = ML + pi * (pw + 0.245)
    rect(s, px, 1.60, pw, 4.36, fill=TINT2, line_color=LINE)
    rect(s, px, 1.60, pw, 0.42, fill=NAVY)
    rich(s, px + 0.16, 1.70, pw - 0.32, 0.26,
         [(title, 11, True, WHITE), ("　" + unit, 8.5, False, TINT)])
    bar_x = px + 1.28
    bar_max = pw - 1.28 - 0.72
    for i, v in enumerate(vals):
        y = 2.20 + i * 0.452
        label(s, px + 0.14, y - 0.015, 1.10, MODELS[i], 8.0, False, GRAY,
              h=0.22, align=R)
        rect(s, bar_x, y + 0.075, bar_max, 0.065, fill=RGBColor(0xE2, 0xEC, 0xF6))
        if v > 0:
            wl = bar_max * v / vmax
            col = CYAN if i < 3 else BLUE
            rect(s, bar_x, y - 0.01, max(wl, 0.05), 0.22, fill=col)
            label(s, bar_x + max(wl, 0.05) + 0.07, y - 0.005, 0.72, labs[i],
                  8.4, True, NAVY, h=0.22)
        else:
            rect(s, bar_x, y - 0.01, 0.28, 0.22, fill=RGBColor(0xD6, 0xE2, 0xEF))
            label(s, bar_x + 0.35, y - 0.005, 1.10, labs[i], 8.0, False, LGRAY, h=0.22)
    hline(s, px + 0.16, 5.18, pw - 0.32, LINE, 0.75)
    label(s, px + 0.16, 5.32, pw - 0.32, note, 8.3, False, GRAY, h=0.60, line=1.42)

rect(s, ML, 6.12, CW, 0.46, fill=NAVY)
label(s, ML + 0.18, 6.22, CW - 0.36,
      "落到测试表：连续 30 杯的 P95 耗时与合格率替代瞬时流量；整套占地按“台上件 + 台下主机 + 冷藏柜 + 检修空间”四项求和；"
      "耗电按 60 分钟混合订单实测。",
      9.2, True, WHITE, h=0.28)
source_note(s, "数值来源：各品牌官方规格图 / 规格表 / 说明书（S02/S06/S08/S09/S12）；占地按官网标注 W×D 计算")
cover_footer(s, 5)


# =====================================================================
# 06 价格全景
# =====================================================================
s = new_slide(prs)
header(s, "05", "价格全景：9 条可核实记录与一个定价窗口",
       "境外标价保留原币与税运口径，人民币等值仅用于画价带；国内定价以同配置询价落地")

hdr = ["对象 / 市场", "可核实金额", "口径与日期"]
cols = [2.60, 1.95, 2.85]
body = [
    ["Marco MilkPal　美国", "$7,500 / $6,750", "标价 / 三家经销商同价，2026-09"],
    ["Marco MilkPal　爱尔兰", "€6,068", "未含 VAT（Blue Butterfly）"],
    ["Marco MilkPal　新加坡", "S$9,950", "税运另计（Stellar M）"],
    ["Übermilk ONE　法国", "€7,260 起", "未税；安装 €350、160 天耗材 €390"],
    ["Linkbar ST-Milk　新加坡", "S$9,500", "奶模块无配件（Stellar M）"],
    ["ST-Milk + 1 冷藏柜", "S$10,600", "同页配置，非中国交付"],
    ["Linkbar Double Favor　中国", "¥36,800", "2024-08 中国零售价（前代）"],
    ["Linkbar ST-Milk　报道价", "¥43,000", "2024-08，地区表述为 overseas"],
    ["阿玛菲 N1 / N2 / N2T", "项目询价", "江苏主体直供，同配置报价落地"],
]
table(s, ML, 1.60, 7.40, cols, None, hdr, body,
      head_size=9.5, body_size=8.8, head_h=0.38, row_h=0.375,
      align=[L, L, L])

bx = ML + 7.72
label(s, bx, 1.62, 4.38, "人民币等值价带（$1≈7.1　€1≈8.3　S$1≈5.5）", 9.5, True, BLUE, h=0.26)
hline(s, bx, 1.92, 4.38, LINE, 0.75)
bands = [("Linkbar 中国历史锚点", 36800, "¥36,800", CYAN),
         ("Marco 美国成交", 47900, "≈¥47,900", BLUE),
         ("Marco 爱尔兰未税", 50400, "≈¥50,400", BLUE),
         ("Linkbar ST 新加坡", 52300, "≈¥52,300", BLUE),
         ("Marco 新加坡", 54700, "≈¥54,700", BLUE),
         ("ST-Milk + 冷藏柜", 58300, "≈¥58,300", BLUE),
         ("Übermilk 法国未税", 60300, "≈¥60,300", BLUE)]
b0, bmax, vmax = bx + 1.62, 1.72, 66000
for i, (name, v, lab, col) in enumerate(bands):
    y = 2.10 + i * 0.395
    label(s, bx, y + 0.015, 1.50, name, 8.4, False, GRAY, h=0.22, align=R)
    rect(s, b0, y + 0.075, bmax, 0.055, fill=RGBColor(0xE2, 0xEC, 0xF6))
    rect(s, b0, y, bmax * v / vmax, 0.205, fill=col)
    label(s, b0 + bmax * v / vmax + 0.08, y + 0.005, 1.0, lab, 8.6, True, NAVY, h=0.22)

rect(s, bx, 4.96, 4.38, 1.66, fill=NAVY)
label(s, bx + 0.22, 5.12, 3.9, "目标价带", 10.5, True, CYAN, spc=0.8)
rich(s, bx + 0.22, 5.38, 3.9, 0.44, [("¥30,000 – ¥37,000", 21, True, WHITE)])
label(s, bx + 0.22, 5.90, 3.96,
      "下沿取 24 个月回收倒推的可承受投入 ¥30,760，上沿贴国内历史锚点 ¥36,800。"
      "这条带子把境外等值价让出 20% 以上，换装机量与服务密度。",
      8.5, False, TINT, h=0.62, line=1.40)

source_note(s, "价格来源：Pro Coffee Gear / Coffee Machine Depot / Voltage Restaurant Supply（$，2026-09）、"
               "Blue Butterfly（€）、Stellar M（S$）、Lomi（€）、Daily Coffee News 2024-08（¥）")
cover_footer(s, 6)


# =====================================================================
# 07 阿玛菲三兄弟
# =====================================================================
s = new_slide(prs)
header(s, "06", "阿玛菲 AMF-N1 / N2 / N2T：国产台上台下同时铺开",
       "同一套厚膜加热 + 齿轮泵平台做出三种吧台形态，4.3 英寸屏统一控制泡厚、温度与出奶速度")

amalfi = [
    ("amalfi_n1", "AMF-N1", "台上整机",
     [("额定", "2600W　220V / 50–60Hz"),
      ("尺寸", "390×180×340mm　10.5kg"),
      ("菜单", "热奶 / 热泡 / 冷奶 / 冷泡，可新增"),
      ("技术", "厚膜加热 + 齿轮泵，官方称无水奶泡"),
      ("接口", "进水 / 奶管 / 排水，随机 2 根奶管"),
      ("清洁", "放置奶管与药液容器，配缺奶检测")],
     "说明书与清洁流程公开可取，作为国产台上基准优先借样。"),
    ("amalfi_n2", "AMF-N2", "台下系统 + 触屏龙头",
     [("额定", "2600W"),
      ("尺寸", "130×240×260mm（台上部件）"),
      ("菜单", "热泡 / 冷泡"),
      ("技术", "厚膜加热 + 齿轮泵 + 全铝机身"),
      ("控制", "4.3″ 屏，泡厚 / 温度 / 出奶速度可调"),
      ("采购", "按龙头+主机+冰箱+管路+安装整套报价")],
     "释放台面的方案，和 ST-Milk 正面比吧台集成与检修空间。"),
    ("amalfi_n2t", "AMF-N2T", "台下系统 + 双拨杆龙头",
     [("额定", "2600W"),
      ("尺寸", "100×220×280mm（台上部件）"),
      ("菜单", "热泡 / 冷泡"),
      ("技术", "厚膜加热 + 齿轮泵 + 全铝机身"),
      ("交互", "双拨杆台面操作，厚度/温度/速度可调"),
      ("核对", "页面标题 N2T、参数字段 N2，以铭牌定版")],
     "双拨杆把高峰操作压成一个动作，交互差异的候选方向。"),
]
cw2 = 3.86
for i, (img, name, form, rowsx, note) in enumerate(amalfi):
    cx = ML + i * (cw2 + 0.245)
    rect(s, cx, 1.60, cw2, 4.98, fill=TINT2, line_color=LINE)
    rect(s, cx + 0.001, 1.601, cw2 - 0.002, 1.58, fill=WHITE)
    place_img(s, A % img, cx, 1.60, cw2, 1.58, pad=0.10)
    hline(s, cx, 3.18, cw2, LINE, 0.75)
    rich(s, cx + 0.18, 3.30, cw2 - 0.36, 0.28,
         [("阿玛菲 ", 10, True, GRAY), (name, 13.5, True, NAVY)])
    label(s, cx + 0.18, 3.62, cw2 - 0.36, form, 9, True, CYAN, h=0.24)
    for j, (k, v) in enumerate(rowsx):
        y = 3.94 + j * 0.375
        label(s, cx + 0.18, y, 0.52, k, 8.2, True, BLUE, h=0.22)
        label(s, cx + 0.74, y, cw2 - 0.94, v, 8.4, False, DARK, h=0.34, line=1.3)
    hline(s, cx + 0.18, 6.18, cw2 - 0.36, LINE, 0.75)
    label(s, cx + 0.18, 6.27, cw2 - 0.36, note, 8.4, True, NAVY, h=0.28, line=1.35)

source_note(s, "参数与图片：阿玛菲官方产品页 S01 / S03 / S04，N1 英文说明书 S02（规格、菜单、缺奶检测、清洁）；"
               "N1 手册限定乳制品，植物奶适配取厂家书面确认")
cover_footer(s, 7)


# =====================================================================
# 08 Linkbar 双形态
# =====================================================================
s = new_slide(prs)
header(s, "07", "Linkbar：台上 MilkPal 与台下 ST-Milk 各守一条形态",
       "两台都做冷热奶泡、支持植物奶、标 120 杯/h @200mL；差别落在吧台改造量与检修可达性")

lk = [
    ("linkbar_milkpal_scene2", "MilkPal", "台上自动供奶 / 发泡",
     [("额定", "3000W　220–230V / 50Hz　30kg"),
      ("尺寸", "534×195×388mm（官方规格图）"),
      ("菜单", "热奶 / 热泡 / 冷奶 / 冷泡，官网明确支持植物奶"),
      ("预设", "4 种类型 + 16 组自定义参数"),
      ("产能", "120 杯/h @200mL（官方标称）"),
      ("价格", "前代 Double Favor 中国零售 ¥36,800（2024-08）")],
     "免开孔落台，装机最快；多菜单适配奶咖+茶饮混合菜单。"),
    ("st_prod", "Single Touch Milk", "分体台下主机 + 台上龙头",
     [("额定", "3400W　220V / 16A（V0.7.3）"),
      ("尺寸", "龙头 116×137×250　主机 320×500×147mm"),
      ("重量", "龙头 1.4kg + 主机 18kg = 19.4kg（不含冷藏柜）"),
      ("预设", "4 种类型 + 16 组自定义参数，支持植物奶"),
      ("清洗", "每日碱性清洗 + 每周酸性深洗，24h 触发强制清洁"),
      ("价格", "新加坡奶模块 S$9,500；+1 冷藏柜 S$10,600")],
     "台面只留龙头，适合新店与吧台改造；交付要带开孔图与管线图。"),
]
for i, (img, name, form, rowsx, note) in enumerate(lk):
    cx = ML + i * 6.10
    rect(s, cx, 1.60, 5.99, 4.30, fill=TINT2, line_color=LINE)
    place_img(s, A % img, cx + 0.001, 1.601, 2.55, 4.298, mode="cover",
              focus=0.45 if i == 0 else 0.52)
    tx = cx + 2.72
    rich(s, tx, 1.82, 3.12, 0.30, [("LINKBAR ", 9, True, CYAN, 1.0)])
    label(s, tx, 2.06, 3.16, name, 15, True, NAVY, h=0.34)
    label(s, tx, 2.42, 3.16, form, 9, True, GRAY, h=0.24)
    hline(s, tx, 2.70, 3.10, LINE, 0.75)
    for j, (k, v) in enumerate(rowsx):
        y = 2.82 + j * 0.415
        label(s, tx, y, 0.50, k, 8.2, True, BLUE, h=0.22)
        label(s, tx + 0.56, y, 2.56, v, 8.4, False, DARK, h=0.40, line=1.3)
    rect(s, tx, 5.36, 3.10, 0.42, fill=NAVY)
    label(s, tx + 0.14, 5.44, 2.86, note, 8.3, True, WHITE, h=0.28, line=1.3)

rect(s, ML, 6.06, CW, 0.52, fill=TINT)
rect(s, ML, 6.06, 0.055, 0.52, fill=CYAN)
label(s, ML + 0.22, 6.16, CW - 0.44,
      "同源关系：2024 年 8 月行业媒体披露 Linkbar 与 Marco 就 Double Favor 版本达成全球市场合作。"
      "两者共享核心平台，因此把比较落点放在版本号、区域配置、渠道覆盖与服务包这四项上。",
      9.0, True, NAVY, h=0.34, line=1.35)

source_note(s, "参数与图片：Linkbar 官方产品页 S05 / S07、官方规格图 S06、ST-Milk V0.7.3 说明书 S08；"
               "价格：Stellar M 新加坡页 S15、Daily Coffee News 2024-08 S13", y=6.70)
cover_footer(s, 8)


# =====================================================================
# 09 海外基准
# =====================================================================
s = new_slide(prs)
header(s, "08", "海外基准：Marco 交文档，Übermilk 交运行数据",
       "两台机器提供两种可借鉴能力——完整到 15A 与管径的交付口径，以及把出量、运行与清洁周期留成记录")

ov = [
    ("marco_bar", "Marco MilkPal 1000090", "爱尔兰 · 同源合作参照",
     [("额定", "3100W　220V / 50Hz / 15A　30kg"),
      ("尺寸", "195×535×388mm，出水口至托盘 180mm"),
      ("管路", "1⁄4″ 快接进水，1½″ 排水；冰箱与滤芯单列计价"),
      ("菜单", "热奶 / 热泡 / 冷奶 / 冷泡，支持乳奶与替代奶"),
      ("配方", "官方宣传 25 配方，V7 手册记 16 类，按版本核"),
      ("清洁", "日终机器周期约 30 分钟，闲置 3/6 分钟自动冲洗")],
     "$7,500 标价　|　$6,750 成交价\n€6,068 未含 VAT　|　S$9,950 税运另计",
     "把规格表做到可直接施工的程度，是国内交付包可以照搬的动作。"),
    ("ubermilk_front", "Übermilk ONE", "德国 · 海外运营基准",
     [("额定", "3100W　230V　25kg"),
      ("尺寸", "180×485×540mm（台高 540）"),
      ("奶源", "外接冷藏奶源，官网建议距离不超过 2m"),
      ("菜单", "热微泡 / 拉花用途明确，份量·温度·稠度可调"),
      ("产能", "峰值 250 杯/h（单杯量按其自有口径）"),
      ("数据", "Übermilk Analytics 记录出量、运行与清洁周期")],
     "€7,260 起（未税）\n安装 €350　|　160 天耗材包 €390",
     "运行与清洁数据留痕，为连锁提供可审计的出品与保养证据。"),
]
for i, (img, name, form, rowsx, price, note) in enumerate(ov):
    cx = ML + i * 6.10
    rect(s, cx, 1.60, 5.99, 4.94, fill=TINT2, line_color=LINE)
    if i == 0:
        place_img(s, A % img, cx + 0.001, 1.601, 2.50, 2.10, mode="cover", focus=0.45)
    else:
        rect(s, cx + 0.001, 1.601, 2.50, 2.10, fill=WHITE)
        place_img(s, A % img, cx, 1.60, 2.50, 2.10, pad=0.10)
    tx = cx + 2.68
    label(s, tx, 1.78, 3.20, name, 13.5, True, NAVY, h=0.30)
    label(s, tx, 2.12, 3.20, form, 8.8, True, CYAN, h=0.24)
    rect(s, tx, 2.42, 3.14, 0.02, fill=LINE)
    label(s, tx, 2.56, 3.20, "可核实标价", 8.2, True, BLUE, h=0.22)
    label(s, tx, 2.80, 3.22, price, 9.6, True, NAVY, h=0.62, line=1.38)
    for j, (k, v) in enumerate(rowsx):
        y = 3.92 + j * 0.375
        label(s, cx + 0.20, y, 0.50, k, 8.2, True, BLUE, h=0.22)
        label(s, cx + 0.76, y, 5.00, v, 8.6, False, DARK, h=0.24)
    rect(s, cx + 0.001, 6.10, 5.988, 0.44, fill=NAVY)
    label(s, cx + 0.20, 6.19, 5.60, note, 8.8, True, WHITE, h=0.28)

source_note(s, "参数与图片：Marco 1000090 规格表 S09、说明书 V7 S10、美国产品页 S11；Übermilk 官方站 S12；"
               "价格：美国三家经销商（2026-09）、Blue Butterfly S17、Stellar M S16、Lomi S18", y=6.66)
cover_footer(s, 9)


# =====================================================================
# 10 运营节奏
# =====================================================================
s = new_slide(prs)
header(s, "09", "运营节奏：清洗周期与停机成本决定门店口碑",
       "两份官方说明书写清了闲置冲洗与日终周期，把机器占用时间和员工人工时间分开计，就能算出真实停机成本")

label(s, ML, 1.62, 7.6, "官方清洗节奏对照", 11, True, BLUE, h=0.26)
hline(s, ML, 1.92, 7.6, LINE, 0.75)
tl = [("Linkbar ST-Milk　V0.7.3", CYAN,
       [("闲置 3 min", "快洗 约 10 s"), ("闲置 6 min", "30 s 倒数 + 标准清洗 约 25 s"),
        ("每日", "碱性清洗"), ("每周", "酸性深洗"), ("满 24 h", "触发强制清洁")]),
      ("Marco MilkPal　V7 手册", BLUE,
       [("闲置 3 min", "冲洗 约 15 s"), ("闲置 6 min", "冲洗 约 25 s"),
        ("日终", "机器完整周期 约 30 min"), ("计时口径", "机器占用与人工分别计"), ("", "")])]
for i, (name, col, items) in enumerate(tl):
    y = 2.06 + i * 1.72
    rect(s, ML, y, 7.60, 1.54, fill=TINT2, line_color=LINE)
    rect(s, ML, y, 0.05, 1.54, fill=col)
    label(s, ML + 0.22, y + 0.16, 3.6, name, 10.5, True, NAVY, h=0.26)
    for j, (k, v) in enumerate(items):
        if not k:
            continue
        bx2 = ML + 0.22 + j * 1.49
        rect(s, bx2, y + 0.52, 1.34, 0.30, fill=col)
        label(s, bx2 + 0.08, y + 0.585, 1.20, k, 8.2, True, WHITE, h=0.22)
        label(s, bx2, y + 0.90, 1.38, v, 8.2, False, GRAY, h=0.52, line=1.34)

place_img(s, A % "marco_inuse", ML + 7.94, 1.62, 4.16, 2.34, mode="cover", focus=0.5)
label(s, ML + 7.94, 4.02, 4.16, "日常操作：奶缸直接取奶，员工动作从蒸汽棒切成一次按压（Marco 官方图）",
      8.2, False, LGRAY, h=0.40, line=1.4)

rect(s, ML + 7.94, 4.52, 4.16, 2.02, fill=NAVY)
label(s, ML + 8.16, 4.70, 3.7, "开发目标", 10.5, True, CYAN, spc=0.8)
rich(s, ML + 8.16, 4.96, 3.7, 0.42, [("≤ 5 min / 日", 20, True, WHITE)])
label(s, ML + 8.16, 5.44, 3.74, "员工人工清洁时间", 8.6, False, TINT, h=0.22)
hline(s, ML + 8.16, 5.74, 3.72, RGBColor(0x2A, 0x4C, 0x6B), 0.75)
label(s, ML + 8.16, 5.86, 3.76,
      "自动清洁引导员工按顺序完成并留下记录；可拆部件伸手可及；异常直接报到具体位置。",
      8.4, False, TINT, h=0.58, line=1.42)

items = [("开店准备", "预热与首杯等待、昨夜残留、奶源温度、开机冲洗耗水"),
         ("高峰连续", "连续 30 杯、冷热交替、换奶后首杯质量与最慢一杯"),
         ("闲置恢复", "清洗中来单能否接、首杯温度与稠度、残奶回流规则"),
         ("故障维修", "到场时效、备机、常换件价格、平均修复时间进合同")]
label(s, ML, 5.52, 7.6, "统一测试要抓的四个时刻", 10.5, True, BLUE, h=0.24)
hline(s, ML, 5.80, 7.60, LINE, 0.75)
for i, (k, v) in enumerate(items):
    x = ML + i * 1.92
    label(s, x, 5.92, 1.80, k, 9.2, True, NAVY, h=0.24)
    label(s, x, 6.18, 1.82, v, 8.0, False, GRAY, h=0.56, line=1.38)

source_note(s, "清洗数据：Linkbar ST-Milk V0.7.3 说明书第 18–20 页（S08）；Marco MilkPal V7 说明书第 23 页（S10）", y=6.76)
cover_footer(s, 10)


# =====================================================================
# 11 投入回收与路线
# =====================================================================
s = new_slide(prs)
header(s, "10", "投入回收与开发路线：用门店数据接管假设",
       "情景模型按整套投入 ¥35,000、26 天/月、人工 ¥30/小时、牛奶 ¥12/L 计算，输入全部可替换")

hdr = ["情景", "奶类饮品/日", "净收益/月", "静态回收期"]
cols = [1.55, 1.70, 1.60, 1.60]
body = [["低使用", "100 杯", "−¥66", "—"],
        ["基准", "250 杯", "¥1,282", "约 27.3 个月"],
        ["高使用", "400 杯", "¥3,686", "约 9.5 个月"]]
label(s, ML, 1.62, 6.45, "投入回收情景模型", 11, True, BLUE, h=0.26)
table(s, ML, 1.94, 6.45, cols, None, hdr, body,
      head_size=9.5, body_size=9.5, head_h=0.38, row_h=0.42, align=[L, C, C, C])
rect(s, ML, 3.62, 6.45, 1.02, fill=TINT)
label(s, ML + 0.16, 3.72, 6.15,
      "月净收益 =日杯数 × 营业天数 × ( 每杯净节省秒数 ÷ 3600 × 时薪 × 兑现比例 + 每杯净节奶 L × 奶价 ) "
      "− 额外清洁人工 − 清洁耗材与维护。人工兑现比例设 0 时，基准场景月净收益 ¥740、回收期约 47.3 个月。",
      8.5, False, NAVY, h=0.84, line=1.44)

label(s, ML, 4.86, 6.45, "先验证的主场景", 11, True, BLUE, h=0.26)
hline(s, ML, 5.16, 6.45, LINE, 0.75)
scene = [("客群", "独立咖啡店与区域连锁中，每天 200–400 杯奶类饮品的门店"),
         ("痛点", "员工轮换频繁、高峰出品波动、清洁步骤靠人盯"),
         ("价值主张", "让不同熟练度的员工稳定出奶，把清洁与停机成本摊到明面上")]
for i, (k, v) in enumerate(scene):
    y = 5.30 + i * 0.46
    label(s, ML, y, 0.92, k, 8.8, True, CYAN, h=0.22)
    label(s, ML + 1.00, y, 3.42, v, 8.8, False, DARK, h=0.42, line=1.34)

place_img(s, A % "st_bar", ML + 4.62, 5.26, 1.83, 1.40, mode="cover", focus=0.45)
label(s, ML + 4.62, 6.70, 1.90, "一次按压完成出奶（Linkbar 官方图）", 7.2, False, LGRAY, h=0.2)

rx = ML + 6.78
label(s, rx, 1.62, 5.32, "开发优先级", 11, True, BLUE, h=0.26)
hline(s, rx, 1.92, 5.32, LINE, 0.75)
road = [("P0", NAVY, "冷热奶泡重复性 · 定量校准 · 缺奶与堵塞提示 · 清洁引导 · "
                     "可维护奶路 · 冷藏整合 · 国内备件与售后流程"),
        ("P1", BLUE, "奶种配方管理 · 权限与使用记录 · 台下安装套件；双奶路依访谈结论决定"),
        ("P2", CYAN, "联网远程下发配方与故障诊断，配合连锁付费意愿推进")]
for i, (p, col, v) in enumerate(road):
    y = 2.06 + i * 0.86
    rect(s, rx, y, 0.62, 0.62, fill=col)
    label(s, rx, y + 0.17, 0.62, p, 13, True, WHITE, h=0.3, align=C)
    label(s, rx + 0.78, y + 0.02, 4.52, v, 8.8, False, DARK, h=0.60, line=1.40)

label(s, rx, 4.72, 5.32, "决策顺序", 11, True, BLUE, h=0.26)
hline(s, rx, 5.02, 5.32, LINE, 0.75)
steps = [("1", "拿同配置含税整套报价", "主机 / 龙头 · 冰箱 · 净水减压 · 安装耗材 · 运输 · 培训 · 税率 · 保修"),
         ("2", "借样做统一工况测试", "定量与温控误差 · 连续 30 杯 P95 · 泡质盲评 · 净损耗 · 人工清洁计时"),
         ("3", "访谈 10–15 家门店", "覆盖不同规模与奶咖占比，验证支付意愿与回收期门槛"),
         ("4", "定台上/台下主产品与售价", "用实测与报价替换模型输入，锁定 ¥30,000–37,000 价带")]
for i, (n, t, v) in enumerate(steps):
    y = 5.14 + i * 0.45
    rect(s, rx, y + 0.015, 0.20, 0.20, fill=CYAN)
    label(s, rx, y + 0.035, 0.20, n, 8, True, WHITE, h=0.18, align=C)
    label(s, rx + 0.30, y, 1.68, t, 8.6, True, NAVY, h=0.22)
    label(s, rx + 2.02, y + 0.01, 3.30, v, 8.0, False, GRAY, h=0.36, line=1.3)

cover_footer(s, 11)

os.makedirs("out", exist_ok=True)
out = "out/商用自动奶泡系统_竞品分析.pptx"
prs.save(out)
print("saved:", out, os.path.getsize(out), "bytes,", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
