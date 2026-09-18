# 商用自动奶泡系统 竞品分析

11 页竞品分析演示文稿，科研蓝白风格，参数与价格为主轴。统一检索日 2026-09-18。

## 下载

| 文件 | 体积 | 说明 |
|---|---|---|
| [`dist/milk-system-competitive-analysis.pptx`](dist/milk-system-competitive-analysis.pptx) | 1.73 MB | 可编辑演示文稿，16:9，表格与图形均为原生对象 |
| [`dist/milk-system-competitive-analysis.pdf`](dist/milk-system-competitive-analysis.pdf) | 1.99 MB | 同内容 PDF，直接翻阅或打印 |

点进文件页后用右上角的下载按钮保存。

## 页面结构

| 页 | 内容 |
|---|---|
| 01 | 封面：4 张官方产品图，功率 2600–3400W、产能 120–250 杯/h、价格带 ¥36.8k–60k |
| 02 | 竞品版图：7 个型号图片墙，每张卡直标功率 / 重量 / 尺寸 / 产能 |
| 03 | 核心判断：功能同质化、价格两层结构、清洁与停机决定复购 |
| 04 | 参数总览：7 行统一口径对照表 + 读表规则 |
| 05 | 功率 / 产能 / 台面占地三组条形图 |
| 06 | 价格全景：9 条可核实记录 + 人民币等值价带 + 目标价带 ¥30,000–37,000 |
| 07 | 阿玛菲 AMF-N1 / N2 / N2T |
| 08 | Linkbar MilkPal 与 Single Touch Milk，含同源合作说明 |
| 09 | 海外基准：Marco MilkPal 1000090 / Übermilk ONE |
| 10 | 运营节奏：两份官方说明书的清洗周期对照 |
| 11 | 投入回收三档情景 + P0/P1/P2 开发优先级 + 四步决策顺序 |

## 关键数据

**参数**（各品牌官方规格图 / 规格表 / 说明书）

| 型号 | 功率 | 重量 | 尺寸 mm | 标称产能 |
|---|---|---|---|---|
| 阿玛菲 AMF-N1 | 2600W | 10.5kg | 390×180×340 | 实测口径补齐 |
| 阿玛菲 AMF-N2 | 2600W | 以铭牌为准 | 130×240×260（台上件） | 实测口径补齐 |
| 阿玛菲 AMF-N2T | 2600W | 以铭牌为准 | 100×220×280（台上件） | 实测口径补齐 |
| Linkbar MilkPal | 3000W | 30kg | 534×195×388 | 120 杯/h @200mL |
| Linkbar ST-Milk | 3400W | 19.4kg | 龙头 116×137×250　主机 320×500×147 | 120 杯/h @200mL |
| Marco MilkPal 1000090 | 3100W | 30kg | 195×535×388 | 120 份/h @200mL |
| Übermilk ONE | 3100W | 25kg | 180×485×540 | 峰值 250 杯/h |

**价格**（原币种与结算口径）

| 对象 | 金额 | 口径 |
|---|---|---|
| Marco MilkPal 美国 | $7,500 / $6,750 | 标价 / 三家经销商同价，2026-09 |
| Marco MilkPal 爱尔兰 | €6,068 | 未含 VAT |
| Marco MilkPal 新加坡 | S$9,950 | 税运另计 |
| Übermilk ONE 法国 | €7,260 起 | 未税；安装 €350、160 天耗材 €390 |
| Linkbar ST-Milk 新加坡 | S$9,500 / S$10,600 | 奶模块 / +1 冷藏柜 |
| Linkbar Double Favor 中国 | ¥36,800 | 2024-08 中国零售价（前代） |
| 阿玛菲 N1/N2/N2T | 项目询价 | 江苏主体直供 |

## 重新构建

```bash
pip install python-pptx pillow
cd src && python3 build.py          # 输出到 out/
```

- `theme.py` —— 配色、字体、表格、条形图、图片裁切等基础组件
- `build.py` —— 11 页内容与版式
- `optimize.py` —— 把官方原图压到最长边 1400px / JPEG q82（`assets_min/`）
- `make_download_page.py` —— 把产物包成单个 HTML 文本文件，用于二进制下载受限的环境

`src/assets_min/` 里是 13 张已压缩的品牌官方原图，抓取自阿玛菲官网、Linkbar Shopify 商品接口、Marco 美国站与 Übermilk 德国站。

## 资料来源

阿玛菲官方产品页与 N1 英文说明书；Linkbar MilkPal 官方页与规格图、ST-Milk V0.7.3 说明书；Marco MilkPal 1000090 规格表与 V7 说明书、美国产品页；Übermilk 德国官网；经销商报价页 Blue Butterfly（爱尔兰）、Stellar M（新加坡）、Lomi（法国）、Pro Coffee Gear / Coffee Machine Depot / Voltage Restaurant Supply（美国）；行业媒体 Daily Coffee News 2024-08；门店规模数据取自 Luckin Coffee 2026 年第二季度财报。

外部资料均经转述压缩，原始链接见演示文稿各页脚注。
