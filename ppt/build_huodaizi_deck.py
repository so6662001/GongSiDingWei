# -*- coding: utf-8 -*-
"""
生成《货袋子平台销售培训》PPT
内容来源：../货袋子平台销售话术手册.md
运行：python3 build_huodaizi_deck.py
输出：../货袋子平台销售培训.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------- 设计系统 ----------------
NAVY = RGBColor(0x0F, 0x2B, 0x46)      # 主色·深钢蓝
NAVY_D = RGBColor(0x08, 0x1A, 0x2B)    # 更深
NAVY_L = RGBColor(0x1B, 0x44, 0x6B)    # 稍浅
ORANGE = RGBColor(0xF2, 0x6A, 0x21)    # 强调·橙
GOLD = RGBColor(0xFF, 0xB0, 0x20)      # 次强调
LIGHT = RGBColor(0xF4, 0xF6, 0xF9)     # 浅底
LINE = RGBColor(0xD8, 0xDF, 0xE7)      # 分隔线
GRAY = RGBColor(0x5A, 0x6B, 0x7B)      # 次要文字
DARKTXT = RGBColor(0x1D, 0x2A, 0x36)   # 正文
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREENOK = RGBColor(0x1E, 0x8E, 0x5A)
REDNO = RGBColor(0xC0, 0x39, 0x2B)

FONT = "Microsoft YaHei"

W = Inches(13.333)
H = Inches(7.5)

prs = Presentation()
prs.slide_width = W
prs.slide_height = H

_page = {"n": 0}


# ---------------- 基础工具 ----------------
def new_slide(bg=WHITE):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bgshape = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    bgshape.fill.solid()
    bgshape.fill.fore_color.rgb = bg
    bgshape.line.fill.background()
    bgshape.shadow.inherit = False
    return s


def rect(s, l, t, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE):
    sh = s.shapes.add_shape(shape, l, t, w, h)
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    return sh


def text(s, content, l, t, w, h, size=18, color=DARKTXT, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.25, font=FONT,
         italic=False):
    """content: str 或 [(片段, 是否加粗, 颜色), ...]"""
    tb = s.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    if isinstance(content, str):
        content = [(content, bold, color)]
    for seg, b, c in content:
        r = p.add_run()
        r.text = seg
        r.font.size = Pt(size)
        r.font.bold = b
        r.font.italic = italic
        r.font.color.rgb = c
        r.font.name = font
    return tb


def paras(s, items, l, t, w, h, size=17, color=DARKTXT, spacing=1.5,
          space_after=10, align=PP_ALIGN.LEFT):
    """items: [str 或 [(片段,粗,色),...]]"""
    tb = s.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_after = Pt(space_after)
        segs = [(it, False, color)] if isinstance(it, str) else it
        for seg, b, c in segs:
            r = p.add_run()
            r.text = seg
            r.font.size = Pt(size)
            r.font.bold = b
            r.font.color.rgb = c
            r.font.name = FONT
    return tb


def footer(s, label="货袋子平台销售培训"):
    _page["n"] += 1
    rect(s, Inches(0), H - Inches(0.42), W, Pt(1.2), fill=LINE)
    text(s, label, Inches(0.62), H - Inches(0.36), Inches(6), Inches(0.3),
         size=10, color=GRAY)
    text(s, str(_page["n"]), W - Inches(1.0), H - Inches(0.36), Inches(0.4),
         Inches(0.3), size=10, color=GRAY, align=PP_ALIGN.RIGHT)


def title_bar(s, kicker, title, sub=None):
    """内容页标题区"""
    rect(s, Inches(0.62), Inches(0.52), Inches(0.09), Inches(0.42), fill=ORANGE)
    if kicker:
        text(s, kicker, Inches(0.86), Inches(0.5), Inches(8), Inches(0.28),
             size=12, color=ORANGE, bold=True)
    text(s, title, Inches(0.86), Inches(0.78), Inches(11.6), Inches(0.62),
         size=30, color=NAVY, bold=True)
    y = Inches(1.5)
    if sub:
        text(s, sub, Inches(0.86), Inches(1.44), Inches(11.6), Inches(0.4),
             size=14, color=GRAY)
        y = Inches(1.95)
    return y


def section_page(num, title, sub, items=None):
    s = new_slide(NAVY)
    rect(s, Inches(0), Inches(0), Inches(0.22), H, fill=ORANGE)
    text(s, num, Inches(1.0), Inches(1.6), Inches(4.2), Inches(2.4),
         size=110, color=RGBColor(0x24, 0x4C, 0x71), bold=True)
    text(s, title, Inches(1.05), Inches(3.55), Inches(10), Inches(0.9),
         size=44, color=WHITE, bold=True)
    rect(s, Inches(1.1), Inches(4.62), Inches(1.5), Pt(3), fill=ORANGE)
    text(s, sub, Inches(1.05), Inches(4.95), Inches(9.6), Inches(0.9),
         size=17, color=RGBColor(0xA9, 0xC2, 0xD8))
    if items:
        text(s, "  ·  ".join(items), Inches(1.05), Inches(5.75), Inches(11),
             Inches(0.6), size=13, color=RGBColor(0x7E, 0x9C, 0xB8))
    footer(s)
    return s


def quote_page(quote, tag=None, note=None, bg=NAVY):
    s = new_slide(bg)
    rect(s, Inches(0), Inches(0), Inches(0.22), H, fill=ORANGE)
    text(s, "“", Inches(1.0), Inches(0.9), Inches(2), Inches(1.6),
         size=96, color=RGBColor(0x2C, 0x57, 0x7C), bold=True)
    if tag:
        text(s, tag, Inches(1.15), Inches(2.05), Inches(10), Inches(0.34),
             size=13, color=GOLD, bold=True)
    text(s, quote, Inches(1.15), Inches(2.5), Inches(10.9), Inches(2.6),
         size=32, color=WHITE, bold=True, spacing=1.42)
    if note:
        rect(s, Inches(1.2), Inches(5.45), Inches(1.2), Pt(3), fill=ORANGE)
        text(s, note, Inches(1.15), Inches(5.75), Inches(10.6), Inches(0.9),
             size=15, color=RGBColor(0xA9, 0xC2, 0xD8), spacing=1.35)
    footer(s)
    return s


def table(s, data, l, t, widths, row_h=0.46, header_h=0.5, size=13,
          header_fill=NAVY, header_color=WHITE, zebra=True, first_col_bold=True,
          align_center_cols=None):
    """data: [[h1,h2..],[r1..],...]  widths: inches列表"""
    rows, cols = len(data), len(data[0])
    total_w = Inches(sum(widths))
    shape = s.shapes.add_table(rows, cols, l, t, total_w,
                               Inches(header_h + row_h * (rows - 1)))
    tbl = shape.table
    tbl.first_row = True
    tbl.horz_banding = False
    for j, wd in enumerate(widths):
        tbl.columns[j].width = Inches(wd)
    tbl.rows[0].height = Inches(header_h)
    for i in range(1, rows):
        tbl.rows[i].height = Inches(row_h)
    align_center_cols = align_center_cols or []
    for i in range(rows):
        for j in range(cols):
            cell = tbl.cell(i, j)
            cell.text = ""
            cell.margin_left = Inches(0.11)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.045)
            cell.margin_bottom = Inches(0.045)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = header_fill
            elif zebra and i % 2 == 0:
                cell.fill.fore_color.rgb = LIGHT
            else:
                cell.fill.fore_color.rgb = WHITE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if (j in align_center_cols or i == 0) else PP_ALIGN.LEFT
            p.line_spacing = 1.15
            val = data[i][j]
            segs = [(val, False, DARKTXT)] if isinstance(val, str) else val
            for seg, b, c in segs:
                r = p.add_run()
                r.text = seg
                r.font.size = Pt(size if i else size + 0.5)
                r.font.name = FONT
                if i == 0:
                    r.font.bold = True
                    r.font.color.rgb = header_color
                else:
                    r.font.bold = b or (first_col_bold and j == 0)
                    r.font.color.rgb = c
    return tbl


def card(s, l, t, w, h, head, body, head_fill=NAVY, accent=ORANGE,
         head_size=16, body_size=13):
    rect(s, l, t, w, h, fill=WHITE, line=LINE, lw=1.1)
    rect(s, l, t, w, Inches(0.5), fill=head_fill)
    rect(s, l, t, Inches(0.055), Inches(0.5), fill=accent)
    text(s, head, l + Inches(0.2), t + Inches(0.1), w - Inches(0.3), Inches(0.3),
         size=head_size, color=WHITE, bold=True)
    paras(s, body, l + Inches(0.2), t + Inches(0.68), w - Inches(0.4),
          h - Inches(0.8), size=body_size, spacing=1.35, space_after=6)


def kpi_row(s, items, t, l=Inches(0.9), gap=Inches(0.28), h=Inches(1.55),
            w=None, num_size=32):
    """items: [(数字, 说明)]"""
    n = len(items)
    if w is None:
        total = W - Inches(1.8)
        w = Emu(int((total - gap * (n - 1)) / n))
    x = l
    for num, desc in items:
        rect(s, x, t, w, h, fill=LIGHT, line=LINE, lw=1.0)
        rect(s, x, t, w, Pt(4), fill=ORANGE)
        text(s, num, x, t + Inches(0.26), w, Inches(0.66), size=num_size, color=NAVY,
             bold=True, align=PP_ALIGN.CENTER)
        text(s, desc, x, t + Inches(0.98), w, Inches(0.4), size=12.5,
             color=GRAY, align=PP_ALIGN.CENTER)
        x = Emu(int(x + w + gap))


# ================= 封面 =================
s = new_slide(NAVY)
rect(s, Inches(0), Inches(0), Inches(0.26), H, fill=ORANGE)
rect(s, Inches(7.3), Inches(0), Inches(6.033), H, fill=NAVY_D)
text(s, "内部培训 · 一线执行", Inches(1.1), Inches(1.35), Inches(6), Inches(0.35),
     size=14, color=GOLD, bold=True)
text(s, "货袋子平台", Inches(1.05), Inches(1.85), Inches(6.2), Inches(1.0),
     size=54, color=WHITE, bold=True)
text(s, "销售培训手册", Inches(1.05), Inches(2.82), Inches(6.2), Inches(1.0),
     size=54, color=ORANGE, bold=True)
rect(s, Inches(1.12), Inches(3.98), Inches(1.8), Pt(4), fill=ORANGE)
paras(s, [
    [("三天上手：会接待、会开通、会成交 ", False, RGBColor(0xC6, 0xD8, 0xE8)),
     ("980", True, WHITE),
     ("，并讲清 1980 / 9800 / 19800", False, RGBColor(0xC6, 0xD8, 0xE8))],
    [("www.huodaizi.com", False, RGBColor(0x7E, 0x9C, 0xB8))],
], Inches(1.05), Inches(4.35), Inches(5.9), Inches(1.2), size=15, spacing=1.4)

text(s, "不做交易 · 不抽成 · 不碰你的客户", Inches(7.95), Inches(2.55),
     Inches(4.9), Inches(0.5), size=20, color=GOLD, bold=True,
     align=PP_ALIGN.CENTER)
for i, (k, v) in enumerate([("2,100+", "入驻商家"), ("20,000", "平台用户"),
                            ("1,000+", "每日搜索"), ("150万吨+", "日活跃库存")]):
    cx = Inches(7.95) + Inches(2.45) * (i % 2)
    cy = Inches(3.35) + Inches(1.0) * (i // 2)
    text(s, k, cx, cy, Inches(2.3), Inches(0.45), size=26, color=WHITE,
         bold=True, align=PP_ALIGN.CENTER)
    text(s, v, cx, cy + Inches(0.48), Inches(2.3), Inches(0.3), size=12,
         color=RGBColor(0x8E, 0xAC, 0xC6), align=PP_ALIGN.CENTER)
footer(s)

# ================= 培训目标 =================
s = new_slide()
y = title_bar(s, "TRAINING GOAL", "这份培训要达到什么",
              "三天之后，新销售必须能独立跑单——不是“懂产品”，是“会动作”")
kpi_row(s, [("背 5 句", "五档一句话零错误"), ("背 4 个数", "平台数据零错误"),
            ("问 3 个问题", "三问法自然问出并总结"),
            ("做 3 个动作", "开通·铺资源·出海报")], y, num_size=23)
rect(s, Inches(0.9), Inches(4.05), Inches(11.55), Inches(2.3), fill=LIGHT,
     line=LINE, lw=1.1)
rect(s, Inches(0.9), Inches(4.05), Inches(0.07), Inches(2.3), fill=ORANGE)
text(s, "出师标准（四条全达标才可独立跑单）", Inches(1.15), Inches(4.28),
     Inches(10), Inches(0.35), size=17, color=NAVY, bold=True)
paras(s, [
    [("① ", True, ORANGE), ("五档一句话零错误；", False, DARKTXT),
     ("② ", True, ORANGE), ("平台四个数据零错误；", False, DARKTXT)],
    [("③ ", True, ORANGE), ("三问法能自然问出并总结客户的话；", False, DARKTXT)],
    [("④ ", True, ORANGE),
     ("能独立完成“当场开通 + 铺满资源 + 做出一张行情海报”三个动作。", False, DARKTXT)],
], Inches(1.15), Inches(4.78), Inches(11), Inches(1.4), size=15.5, spacing=1.4,
    space_after=8)
footer(s)

# ================= 目录 =================
s = new_slide()
y = title_bar(s, "CONTENTS", "培训目录")
toc = [("01", "地基", "先懂三件事 + 平台底气"),
       ("02", "五档产品", "价值定位与话术包"),
       ("03", "价格结构", "线索单价里的秘密"),
       ("04", "销售路径", "从陌拜到 9,800"),
       ("05", "痛点沟通", "三问法与三个算账"),
       ("06", "抗拒解除", "17 个异议标准答案"),
       ("07", "成交动作", "三前提与三句话"),
       ("08", "培训与红线", "考核标准与 9 条纪律")]
for i, (num, name, desc) in enumerate(toc):
    col, row = i % 2, i // 2
    l = Inches(0.9) + Inches(5.95) * col
    t = Inches(2.15) + Inches(1.12) * row
    rect(s, l, t, Inches(5.6), Inches(0.95), fill=LIGHT, line=LINE, lw=1.0)
    rect(s, l, t, Inches(0.06), Inches(0.95), fill=ORANGE)
    text(s, num, l + Inches(0.28), t + Inches(0.2), Inches(0.9), Inches(0.55),
         size=26, color=ORANGE, bold=True)
    text(s, name, l + Inches(1.25), t + Inches(0.16), Inches(4), Inches(0.35),
         size=18, color=NAVY, bold=True)
    text(s, desc, l + Inches(1.25), t + Inches(0.55), Inches(4.2), Inches(0.3),
         size=12.5, color=GRAY)
footer(s)

# ================= 01 地基 =================
section_page("01", "地基", "讲错了，后面全错",
             ["三句地基", "平台底气", "海报机制", "团队与创始人"])

quote_page("我们不抽成、不碰你的客户、不跟你抢生意。\n"
           "买家看到你，直接联系你、你自己谈、自己收钱。\n"
           "我们只负责让买家看到你、相信你。",
           tag="地基一 · 我们卖的不是交易",
           note="这是全场第一个必答题的标准答案，也是我们对交易平台的结构性差异化。")

s = new_slide()
y = title_bar(s, "地基二", "五档不是“功能多少”，是四段不同的价值")
table(s, [
    ["档位", "它真正解决的问题", "客户心里的那句话"],
    ["标准版 ￥0", [("有没有", True, NAVY)], "网上搜不到我"],
    ["专业版 ￥980", [("有没有用", True, ORANGE)], "发了半天不知道有没有效果"],
    ["认证版 ￥1,980", [("敢不敢", True, ORANGE)], "陌生买家不敢找我"],
    ["旗舰版 ￥9,800", [("多不多", True, NAVY)], "排在第三页等于没挂"],
    ["尊享版 ￥19,800", [("是不是第一眼", True, NAVY)], "线下有名号，线上没体现"],
], Inches(0.9), y, [3.0, 3.6, 4.95], row_h=0.62, header_h=0.52, size=14.5)
rect(s, Inches(0.9), Inches(6.05), Inches(11.55), Inches(0.72), fill=NAVY)
text(s, "记忆口诀： 0 免费占位 → 980 看效果 → 1980 建信任 → 9800 抢流量 → 19800 做品牌",
     Inches(1.15), Inches(6.24), Inches(11), Inches(0.35), size=16.5,
     color=WHITE, bold=True)
footer(s)

quote_page("免费版能看到有多少条线索，但看不到具体是谁。",
           tag="地基三 · 最重要的功能是“看得见、摸不着”",
           note="所以第一次拜访不要卖 980。只做三件事：让他免费开通、把库存铺上去、教他发一张海报。\n"
                "销售不靠说服，靠等线索自己出现。")

s = new_slide()
y = title_bar(s, "平台底气 ①", "先给硬数据：“你们平台到底有没有人”",
              "这是全场第二个必答题。以公司每月统一发布口径为准。")
kpi_row(s, [("2,100+", "入驻商家"), ("20,000", "平台用户"),
            ("1,000+", "每日搜索次数"), ("150万吨+", "日活跃库存")], y,
        h=Inches(1.7))
rect(s, Inches(0.9), Inches(4.25), Inches(11.55), Inches(1.0), fill=NAVY)
rect(s, Inches(0.9), Inches(4.25), Inches(0.07), Inches(1.0), fill=ORANGE)
text(s, "讲数据必须带趋势：“而且这个量每个月都在往上走。”",
     Inches(1.2), Inches(4.42), Inches(11), Inches(0.35), size=18,
     color=WHITE, bold=True)
text(s, "静态数字让人评估，增长趋势让人怕错过。",
     Inches(1.2), Inches(4.83), Inches(11), Inches(0.3), size=13.5,
     color=RGBColor(0xA9, 0xC2, 0xD8))
paras(s, [[("红线：", True, REDNO),
           ("只用公司每月统一发布的数据；不许把“用户数”说成“买家数”、“搜索次数”说成“询盘数”——一次被识破，全部信任状失效。",
            False, DARKTXT)]],
      Inches(0.95), Inches(5.6), Inches(11.4), Inches(0.6), size=14)
footer(s)

s = new_slide()
y = title_bar(s, "平台底气 ②", "六大流量来源：回答“我上来之后谁会看到我”")
srcs = [("3,000 家老客户\n在平台上找货", "最硬的一条，必讲：不是买来的流量，是我们服务多年的真实钢贸企业", True),
        ("30 多位全国销售员", "天天在各地市场上一家一家推", False),
        ("每天两场行情直播", "早盘一场、下午一场，买家养成了看行情的习惯", False),
        ("每天近 30 条短视频", "平台自己在往外推，您不是孤军奋战", False),
        ("1 万多件 T 恤", "发给驾驶员与库房工人 —— 流动品宣", False),
        ("商家自己用海报裂变", "前五条是平台给您导流，这一条是您自己拓客", False)]
for i, (h, d, hot) in enumerate(srcs):
    col, row = i % 3, i // 3
    l = Inches(0.9) + Inches(3.93) * col
    t = y + Inches(2.05) * row
    fill = ORANGE if hot else NAVY
    rect(s, l, t, Inches(3.66), Inches(1.82), fill=WHITE, line=LINE, lw=1.1)
    rect(s, l, t, Inches(3.66), Pt(4.5), fill=fill)
    text(s, h, l + Inches(0.2), t + Inches(0.24), Inches(3.3), Inches(0.62),
         size=15.5, color=NAVY if not hot else ORANGE, bold=True, spacing=1.2)
    text(s, d, l + Inches(0.2), t + Inches(0.98), Inches(3.3), Inches(0.72),
         size=11.5, color=GRAY, spacing=1.3)
rect(s, Inches(0.9), Inches(6.15), Inches(11.55), Inches(0.62), fill=NAVY)
text(s, "一句话收口：“平台推流 + 您自己曝光，两边一起来。”",
     Inches(1.15), Inches(6.29), Inches(11), Inches(0.35), size=16.5,
     color=WHITE, bold=True)
footer(s)

quote_page("我们印了一万多件 T 恤，发给了跑车的驾驶员和库房的工人。\n"
           "这些人每天在多少个库房、多少家公司之间来回跑——\n"
           "他们就是这行里流动性最强的一群人。",
           tag="开场破冰故事",
           note="它有画面、有规模感，但真正的作用是证明我们在真花钱做事。\n"
                "在钢贸这种熟人行业，“看得见的投入”比任何 PPT 都可信。")

s = new_slide()
y = title_bar(s, "话术方法论", "把“机会”讲成“损失”", "这是免费开通说服力的分水岭")
rect(s, Inches(0.9), y, Inches(5.6), Inches(2.0), fill=WHITE, line=LINE, lw=1.2)
rect(s, Inches(0.9), y, Inches(5.6), Pt(5), fill=GRAY)
text(s, "✕  机会框架", Inches(1.15), y + Inches(0.24), Inches(5), Inches(0.35),
     size=17, color=GRAY, bold=True)
text(s, "“您上来可以获得曝光。”", Inches(1.15), y + Inches(0.72), Inches(5.1),
     Inches(0.5), size=17, color=DARKTXT)
text(s, "客户的反应：可有可无，以后再说。", Inches(1.15), y + Inches(1.35),
     Inches(5.1), Inches(0.4), size=13, color=GRAY)

rect(s, Inches(6.85), y, Inches(5.6), Inches(2.0), fill=WHITE, line=ORANGE, lw=1.6)
rect(s, Inches(6.85), y, Inches(5.6), Pt(5), fill=ORANGE)
text(s, "✓  损失框架", Inches(7.1), y + Inches(0.24), Inches(5), Inches(0.35),
     size=17, color=ORANGE, bold=True)
text(s, "“每天 1,000 多次搜索，一年 36 万多次。只要有一次搜的是您常做的规格而您没挂上去，这个客户就去别家了。”",
     Inches(7.1), y + Inches(0.68), Inches(5.1), Inches(1.15), size=15,
     color=DARKTXT, spacing=1.3)

rect(s, Inches(0.9), Inches(4.55), Inches(11.55), Inches(0.95), fill=LIGHT,
     line=LINE, lw=1.0)
rect(s, Inches(0.9), Inches(4.55), Inches(0.06), Inches(0.95), fill=ORANGE)
text(s, "“平台上已经挂了 150 万吨货，买家现在搜到的都是别人家的——您的库存还在自己仓库里躺着。”",
     Inches(1.2), Inches(4.78), Inches(11), Inches(0.5), size=16,
     color=NAVY, bold=True)
rect(s, Inches(0.9), Inches(5.75), Inches(11.55), Inches(0.82), fill=NAVY)
text(s, "免费开通的说服逻辑升级：从“您先试试看” → “您不挂上去，这 1,000 次搜索里搜到您的机会是零”。",
     Inches(1.15), Inches(5.95), Inches(11), Inches(0.42), size=15.5,
     color=WHITE, bold=True)
footer(s)

s = new_slide()
y = title_bar(s, "海报机制 · 980 的价值地基", "为什么发我们的海报有人看，发他自己的没人看",
              "商家自己发朋友圈最大的问题，不是没有工具，是没有内容。")
table(s, [
    ["", "他自己发", "用我们的海报"],
    ["内容", "“我有货、我有货”", [("一半是今日行情，一半是他的现货资源", True, ORANGE)]],
    ["买家为什么打开", "没理由，看两次就屏蔽", [("买家每天都要看价格 —— 行情是刚需", True, ORANGE)]],
    ["结果", "发了等于没发", [("为了看行情点开，顺手看到了他的货", True, ORANGE)]],
], Inches(0.9), y, [2.4, 4.1, 5.05], row_h=0.68, header_h=0.5, size=14.5)
rect(s, Inches(0.9), Inches(5.05), Inches(11.55), Inches(0.85), fill=NAVY)
rect(s, Inches(0.9), Inches(5.05), Inches(0.07), Inches(0.85), fill=ORANGE)
text(s, "核心机制：行情数据由平台免费提供 —— 商家用“买家想看的内容”，带自己想卖的货。",
     Inches(1.2), Inches(5.27), Inches(11), Inches(0.42), size=17,
     color=WHITE, bold=True)
paras(s, [[("它同时解掉那个高频抗拒（“我自己也会做图”）：", False, DARKTXT),
           ("“您会做图，但您没有行情数据。”", True, ORANGE)]],
      Inches(0.95), Inches(6.15), Inches(11.4), Inches(0.5), size=16)
footer(s)

quote_page("您让销售天天发“我有货”，客户看两次就屏蔽了。\n"
           "但行情不一样——买家每天都得看价格。\n"
           "我们把行情数据免费给您用，海报一半行情、一半您的现货：\n"
           "客户为了看行情点开，顺手就看到了您的货。",
           tag="必背话术 · 逐字复述",
           note="Day 0 当场做海报时，必须边做边讲这套逻辑。")

s = new_slide()
y = title_bar(s, "信任状 ③", "团队与创始人：为什么这个平台靠得住",
              "用法纪律：这些是“降低风险感”的信任状，不是“制造购买欲”的卖点——只在客户提出疑虑时用。")
card(s, Inches(0.9), y, Inches(5.6), Inches(1.95), "近 40 位研发工程师 · 公司近 100 人",
     [[("用在：", True, ORANGE), ("“你们公司多大”“会不会做两年就没了”", False, DARKTXT)],
      [("“光研发工程师就有近 40 位，公司近 100 人。平台每个月都在更新——您现在上，后面加的新功能都是白得的。”",
        False, DARKTXT)]])
card(s, Inches(6.85), y, Inches(5.6), Inches(1.95), "客服团队 + 客户成功团队",
     [[("用在：", True, ORANGE), ("“买了会不会没人管”", False, DARKTXT)],
      [("三层服务：全员每周运营指导课 → 认证版起有在线人工客服 → 旗舰/尊享配专属顾问与 1 对 1 客户成功经理。",
        False, DARKTXT)]])
rect(s, Inches(0.9), Inches(4.55), Inches(11.55), Inches(2.05), fill=NAVY)
rect(s, Inches(0.9), Inches(4.55), Inches(0.07), Inches(2.05), fill=ORANGE)
text(s, "创始人是钢贸老板出身：做过销售、管过财务、当过总经理",
     Inches(1.2), Inches(4.75), Inches(11), Inches(0.38), size=19,
     color=GOLD, bold=True)
paras(s, [
    [("用法一 · 打“不懂行”：", True, WHITE),
     ("“平台上每个功能，都是他当年被折磨过之后想出来的，不是互联网公司拍脑袋做的。”",
      False, RGBColor(0xC6, 0xD8, 0xE8))],
    [("用法二 · 解释“为什么不做交易”（更重要）：", True, WHITE),
     ("“他自己就是钢贸老板出身，最清楚商家最怕平台抢客户，所以我们从第一天就定死了：不做交易、不抽成、不碰您的客户。”",
      False, RGBColor(0xC6, 0xD8, 0xE8))],
], Inches(1.2), Inches(5.22), Inches(11.05), Inches(1.25), size=13.5,
    spacing=1.35, space_after=6)
footer(s)

quote_page("说清动机之后，“不做交易”就从一个承诺，变成了立场。\n而立场比承诺可信。",
           tag="为什么创始人这张牌威力这么大", bg=NAVY_D)

# ================= 02 五档 =================
section_page("02", "五档产品", "价值定位 · 一句话 · 三句痛点 · 下单引导",
             ["0 元", "980", "1,980", "9,800", "19,800"])

s = new_slide()
y = title_bar(s, "全景表", "五档价值阶梯（背下来这张表）")
table(s, [
    ["档位", "一句话介绍（必须背）", "线索", "核心武器"],
    ["标准版\n￥0 永久", "先免费开个店，把库存挂上去，让买家能搜到你。",
     [("0", True, GRAY), ("\n只见条数", False, GRAY)],
     "不限上架 · 平台曝光 · 数据看板 · 海报 7 天体验"],
    ["专业版\n￥980/年", [("让你的销售每天发的朋友圈，第一次能看见效果。", True, ORANGE)],
     [("3", True, ORANGE)], "行情+现货海报 · 经营数据看板 · 工商认证 · 信用分"],
    ["认证版\n￥1,980/年", [("让陌生买家敢联系你。", True, ORANGE)],
     [("10", True, ORANGE)], "五维认证 · 精装店铺 · 认证加权排名 · 推荐 1 次"],
    ["旗舰版\n￥9,800/年", "让平台把流量推给你，而不是等买家碰巧找到你。",
     [("30", True, NAVY)], "推荐 7 次 · Banner 3 次 · 置顶 5 次 · 专属顾问 · 上镜位"],
    ["尊享版\n￥19,800/年", "在这个平台上，成为买家第一眼看到的那家。",
     [("50", True, NAVY)], "品牌店 · 首页 Banner · 首页置顶 · 1 对 1 客户成功"],
], Inches(0.72), y, [1.95, 3.85, 0.95, 5.15], row_h=0.72, header_h=0.48,
    size=12.5, align_center_cols=[2])
footer(s)

tiers = [
    dict(name="标准版", price="￥0 / 永久", pos="入场券：让企业先在平台上“存在”，并亲眼看到平台能带来什么。",
         one="先免费开个店，把库存挂上去，让买家能搜到你。一分钱不花，永久有效。",
         pains=[
             "我们平台每天 1,000 多次搜索。只要有一次搜的是您常做的规格而您没挂上去，这客户就去别家了。",
             "您的库存现在只有老客户知道——货压在库里不是没人要，是没人看见。平台上已经挂了 150 万吨。",
             "免费全挂上去，还蹭平台推流：3,000 家老客户在找货、30 多个销售员在推、每天两场直播、30 条短视频。"],
         cta="这个不要钱，也不用签合同，我现在用您的营业执照帮您开好，两分钟。您把常备规格发我，我顺手挂上去。",
         hi=False),
    dict(name="专业版", price="￥980 / 年", pos="让“看不见效果的营销动作”第一次变得可衡量。核心是营销工具 + 数据看板，不是工商认证。",
         one="行情数据免费给您，海报一半行情、一半您的现货——客户为了看行情点开，顺手看到您的货；而且谁看了、从哪来、哪个员工最有效，全看得见。",
         pains=[
             "您让销售天天发朋友圈，发的是不是都是“我有货”？客户看两次就屏蔽了——这个动作做了等于没做。",
             "而且您不知道有没有用，他自己也不知道。好销售最怕干没意义的事——人不是被累走的，是被“不知道为什么要做”磨走的。",
             "上了两件事变了：内容变了（行情带货），效果看得见了（谁扫的、从哪来、哪个员工最有效）。"],
         cta="980 一年，一天不到 3 块钱，等于给全公司销售配一套能看见效果的拓客工具，还带 3 条平台线索。我现在给您开，今天就能发第一张海报。",
         hi=True),
    dict(name="认证版", price="￥1,980 / 年", pos="把“陌生买家不敢联系你”这个问题解决掉。全线性价比最高的一档。",
         one="让陌生买家敢联系你——五维认证加精装店铺，扫码进来一眼就看明白你是谁、有多少实力。",
         pains=[
             "买家看到两家价格差不多，一家有认证、店铺清清楚楚，另一家什么都没有——他会打给谁？",
             "钢贸买家最怕遇到不靠谱的、怕一货多卖、怕付了钱提不到货。他不是不想找您，是不敢找您。",
             "五维认证 + 精装店铺让买家敢联系您；认证还加权排名，同样的货您排得更靠前。"],
         cta="从 980 到 1980 只多 1,000 块，线索从 3 条变 10 条——多的这 7 条相当于一条 143 块，还白得五维认证和精装店铺。",
         hi=True),
    dict(name="旗舰版", price="￥9,800 / 年", pos="从“等买家找到你”变成“平台主动把买家推给你”。卖的是曝光位，不是线索单价。",
         one="让平台把流量推给你——推荐位 7 次、找钢材 Banner 3 次、资源置顶 5 次，再配一个专属运营顾问。",
         pains=[
             "平台上货越来越多，买家看得过来的就前面那几家。您的资源挂上去了，但排在第三页——等于没挂。",
             "发海报是“自己找客户”；推荐位和置顶位是“把找货的买家直接送到您面前”。两件事的效率不一样。",
             "30 条线索之外，专属顾问会告诉您什么时间挂什么规格、什么价位发海报最有效——这套方法比位置更值钱。"],
         cta="9,800 一年，一个月 800 多块，换平台主动给您导流 + 一个懂行的顾问带着做。前面效果已经出来了，放大就靠这一档。",
         hi=False),
    dict(name="尊享版", price="￥19,800 / 年", pos="在平台上确立第一梯队的品牌地位，并配 1 对 1 客户成功。",
         one="在这个平台上，成为买家第一眼看到的那家——品牌专属店铺、首页 Banner、首页置顶，加 1 对 1 客户成功经理。",
         pains=[
             "您在本地是有名号的，但线上买家看您和看一家小贸易商没区别——线下的品牌优势没被体现出来。",
             "买家看价格，也看“跟谁买最稳”。第一眼看到谁，谁就有先发优势，这个位置不是有钱就能随时买到。",
             "品牌店 + 首页位置 + 15 次推荐 + 50 条线索 + 1 对 1 经理全年跟着——把线上做成第二个销售渠道。"],
         cta="这一档一年只放有限名额（按平台位置容量核定）。您是这个区域的头部，位置被同行占了，再想上就得等下一年。",
         hi=False),
]
for tr in tiers:
    s = new_slide()
    accent = ORANGE if tr["hi"] else NAVY
    y = title_bar(s, "产品话术包", f"{tr['name']}　{tr['price']}", tr["pos"])
    rect(s, Inches(0.9), y, Inches(11.55), Inches(1.12), fill=NAVY)
    rect(s, Inches(0.9), y, Inches(0.07), Inches(1.12), fill=ORANGE)
    text(s, "一句话介绍", Inches(1.2), y + Inches(0.12), Inches(4), Inches(0.28),
         size=12, color=GOLD, bold=True)
    text(s, tr["one"], Inches(1.2), y + Inches(0.44), Inches(11.0), Inches(0.6),
         size=15.5, color=WHITE, bold=True, spacing=1.28)
    ty = y + Inches(1.42)
    text(s, "三句话讲痛点与价值", Inches(0.95), ty, Inches(6), Inches(0.3),
         size=13.5, color=ORANGE, bold=True)
    for i, p in enumerate(tr["pains"]):
        t = ty + Inches(0.42) + Inches(0.78) * i
        rect(s, Inches(0.9), t, Inches(11.55), Inches(0.68), fill=LIGHT,
             line=LINE, lw=0.9)
        rect(s, Inches(0.9), t, Inches(0.055), Inches(0.68), fill=accent)
        text(s, str(i + 1), Inches(1.1), t + Inches(0.16), Inches(0.35),
             Inches(0.35), size=16, color=accent, bold=True)
        text(s, p, Inches(1.52), t + Inches(0.1), Inches(10.75), Inches(0.52),
             size=13, color=DARKTXT, spacing=1.28)
    ct = ty + Inches(0.42) + Inches(0.78) * 3 + Inches(0.14)
    rect(s, Inches(0.9), ct, Inches(11.55), Inches(0.82), fill=ORANGE)
    text(s, "引导下单： " + tr["cta"], Inches(1.2), ct + Inches(0.16),
         Inches(11.0), Inches(0.55), size=13.5, color=WHITE, bold=True,
         spacing=1.25)
    footer(s)

# ================= 03 价格结构 =================
section_page("03", "价格结构", "线索单价里藏着一个甜蜜点",
             ["327 元", "198 元", "327 元", "396 元"])

s = new_slide()
y = title_bar(s, "定价武器", "单条线索成本对照表（背下来）")
table(s, [
    ["档位", "价格", "线索", "单条成本", "销售怎么用"],
    ["专业版", "￥980", "3 条", "327 元", "入门，先验证真假"],
    ["认证版", [("￥1,980", True, ORANGE)], [("10 条", True, ORANGE)],
     [("198 元 ★", True, ORANGE)], [("全线最低！主推升级点", True, ORANGE)]],
    ["旗舰版", "￥9,800", "30 条", "327 元", "不讲线索单价，讲曝光位"],
    ["尊享版", "￥19,800", "50 条", "396 元", "不讲线索单价，讲首页与品牌"],
], Inches(0.9), y, [1.85, 1.7, 1.4, 1.8, 4.8], row_h=0.62, header_h=0.5,
    size=14, align_center_cols=[1, 2, 3])
rect(s, Inches(0.9), Inches(4.5), Inches(11.55), Inches(1.05), fill=ORANGE)
text(s, "黄金升级话术（边际成本只有 143 元/条）", Inches(1.2), Inches(4.62),
     Inches(10), Inches(0.3), size=12.5, color=RGBColor(0xFF, 0xE6, 0xD0),
     bold=True)
text(s, "“再加 1,000 块升到 1,980，线索变 10 条——多出来的 7 条一条才 143 块，认证和精装店铺白送。”",
     Inches(1.2), Inches(4.96), Inches(11.0), Inches(0.45), size=15.5,
     color=WHITE, bold=True)
paras(s, [
    [("纪律一：", True, NAVY), ("1,980 是全线性价比最高、最容易成交的升级——只要客户认可线索质量，这一步顺理成章。", False, DARKTXT)],
    [("纪律二：", True, REDNO), ("9,800 与 19,800 绝不能用线索单价销售（会被当场算出更贵）。它们卖的是位置、品牌、专属服务。", False, DARKTXT)],
], Inches(0.95), Inches(5.78), Inches(11.4), Inches(1.0), size=14, spacing=1.35)
footer(s)

# ================= 04 路径 =================
section_page("04", "销售路径", "从陌拜到 9,800", ["Day 0", "第 7—10 天", "第 30—60 天", "第 90—120 天"])

s = new_slide()
y = title_bar(s, "路径全景", "五个阶段：不要在第一次见面卖 980",
              "免费开通没有抗性，而 7 天体验期与线索条数会自动产生成交理由。")
steps = [("Day 0", "首次拜访", "免费开通 + 铺满资源\n+ 当场做一张海报 + 加微信", ORANGE),
         ("2—6 天", "轻触达", "帮他做当日行情海报\n带他看一次数据 · 不谈钱", NAVY),
         ("7—10 天", "成交 980", "双钩子：有线索用线索\n没线索用“工具到期 + 认证”", ORANGE),
         ("30—60 天", "升 1,980", "线索太少 + 买家不敢联系\n→ 143 元/条 的黄金话术", NAVY),
         ("90—120 天", "升 9,800", "排名不够 + 曝光不够\n→ 平台导流 + 专属顾问", NAVY)]
bw = Inches(2.19)
for i, (tag, name, desc, c) in enumerate(steps):
    l = Inches(0.9) + (bw + Inches(0.15)) * i
    rect(s, l, y, bw, Inches(2.5), fill=WHITE, line=LINE, lw=1.1)
    rect(s, l, y, bw, Inches(0.52), fill=c)
    text(s, tag, l, y + Inches(0.12), bw, Inches(0.3), size=14, color=WHITE,
         bold=True, align=PP_ALIGN.CENTER)
    text(s, name, l + Inches(0.16), y + Inches(0.7), bw - Inches(0.3),
         Inches(0.32), size=15.5, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    text(s, desc, l + Inches(0.16), y + Inches(1.15), bw - Inches(0.3),
         Inches(1.2), size=11.5, color=GRAY, spacing=1.35, align=PP_ALIGN.CENTER)
    if i < 4:
        text(s, "▶", l + bw + Inches(0.005), y + Inches(1.0), Inches(0.15),
             Inches(0.3), size=12, color=ORANGE, bold=True)
rect(s, Inches(0.9), Inches(5.15), Inches(11.55), Inches(0.85), fill=NAVY)
rect(s, Inches(0.9), Inches(5.15), Inches(0.07), Inches(0.85), fill=ORANGE)
text(s, "另有三类客户可直通 19,800：区域头部/有品牌的商家 · 已在其他平台大量投入的 · 有专职线上运营团队的。",
     Inches(1.2), Inches(5.38), Inches(11), Inches(0.42), size=14.5,
     color=WHITE)
text(s, "核心原则：销售不靠说服，靠等线索自己出现。", Inches(0.95), Inches(6.25),
     Inches(11), Inches(0.4), size=17, color=ORANGE, bold=True)
footer(s)

s = new_slide()
y = title_bar(s, "Day 0", "第一次拜访只做四个动作", "不讲五个版本、不报价、不讲功能表。")
acts = [("① 当场免费开通", "要营业执照、常备规格、联系电话。\n不当场做，就永远做不成。"),
        ("② 当场铺 5—10 个规格", "挂得越满，后面线索越多——\n这一步直接决定 7 天后能不能成交。"),
        ("③ 当场做一张“行情 + 现货”海报", "上半行情（平台免费提供）、下半他的现货，\n发到他手机上让他当场转到客户群。"),
        ("④ 加微信 + 埋钩子", "“我这七天帮您盯着，有线索我第一时间告诉您。”")]
for i, (h, d) in enumerate(acts):
    col, row = i % 2, i // 2
    l = Inches(0.9) + Inches(5.95) * col
    t = y + Inches(1.62) * row
    hot = (i == 2)
    rect(s, l, t, Inches(5.6), Inches(1.42), fill=WHITE, line=ORANGE if hot else LINE,
         lw=1.6 if hot else 1.1)
    rect(s, l, t, Inches(0.06), Inches(1.42), fill=ORANGE if hot else NAVY)
    text(s, h, l + Inches(0.25), t + Inches(0.18), Inches(5.2), Inches(0.32),
         size=16, color=ORANGE if hot else NAVY, bold=True)
    text(s, d, l + Inches(0.25), t + Inches(0.62), Inches(5.2), Inches(0.7),
         size=12.5, color=GRAY, spacing=1.35)
rect(s, Inches(0.9), Inches(5.35), Inches(11.55), Inches(1.05), fill=ORANGE)
text(s, "做海报时必须边做边说：", Inches(1.2), Inches(5.5), Inches(10), Inches(0.3),
     size=12.5, color=RGBColor(0xFF, 0xE6, 0xD0), bold=True)
text(s, "“您看，一半行情一半您的货——客户为了看价格点开，顺手就看到您的资源了。这个内容客户才愿意看。”",
     Inches(1.2), Inches(5.84), Inches(11), Inches(0.45), size=15,
     color=WHITE, bold=True)
footer(s)

s = new_slide()
y = title_bar(s, "第 7—10 天", "成交 980 的双钩子：没有线索时也能成交",
              "培训重点：一定要教会新人第二个钩子，否则平台早期销售会全线卡死。")
card(s, Inches(0.9), y, Inches(5.6), Inches(2.6), "钩子一 · 线索（首选）",
     [[("适用：", True, ORANGE), ("店里已经有线索", False, DARKTXT)],
      [("“王总，您店里现在有 5 条线索，是买家看了您资源之后留下的。但免费版只能看到条数、看不到是谁。980 开一下，这些人您就能直接联系。”",
        True, NAVY)]], head_fill=ORANGE, body_size=13.5)
card(s, Inches(6.85), y, Inches(5.6), Inches(2.6), "钩子二 · 工具到期 + 认证",
     [[("适用：", True, ORANGE), ("冷启动期，暂时还没有线索", False, DARKTXT)],
      [("“您海报工具今天到期了，这七天您发了 X 张、有 Y 个人扫过，停了就都没了。而且您店铺还没有认证标识，买家看到没验证过的商家，多数不会打电话。”",
        True, NAVY)]], body_size=13.5)
rect(s, Inches(0.9), Inches(5.35), Inches(11.55), Inches(1.0), fill=NAVY)
rect(s, Inches(0.9), Inches(5.35), Inches(0.07), Inches(1.0), fill=ORANGE)
text(s, "成交三前提（缺一个就不要开口要钱）：① 他已经用过（发过海报或看过看板）　② 他自己说出过痛点　③ 他看到过数字",
     Inches(1.2), Inches(5.68), Inches(11), Inches(0.42), size=14.5,
     color=WHITE, bold=True)
footer(s)

# ================= 05 沟通 =================
section_page("05", "痛点沟通", "三问法与三个算账", ["问出来的痛才是痛"])

s = new_slide()
y = title_bar(s, "三问法", "不要讲产品，先问这三个问题",
              "客户自己说出来的痛，才是痛。")
qs = [("问一", "您现在新客户主要从哪来？", "摸清获客现状。多数答案是“老客户介绍”“销售跑出来的”——线上等于零。", NAVY),
      ("问二", "您让销售发朋友圈、发社群吗？发了之后您怎么知道有没有用？",
       "全案最狠的一问。他答不出来（因为真的不知道）——980 的价值当场成立。", ORANGE),
      ("问三", "一个陌生买家在网上第一次看到您，凭什么敢给您打电话？",
       "挖出信任缺失——1,980 的价值当场成立。", NAVY)]
for i, (tag, q, why, c) in enumerate(qs):
    t = y + Inches(1.15) * i
    rect(s, Inches(0.9), t, Inches(11.55), Inches(1.0), fill=WHITE,
         line=ORANGE if c == ORANGE else LINE, lw=1.6 if c == ORANGE else 1.1)
    rect(s, Inches(0.9), t, Inches(1.15), Inches(1.0), fill=c)
    text(s, tag, Inches(0.9), t + Inches(0.32), Inches(1.15), Inches(0.35),
         size=16, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    text(s, q, Inches(2.25), t + Inches(0.16), Inches(10.0), Inches(0.38),
         size=16.5, color=NAVY, bold=True)
    text(s, why, Inches(2.25), t + Inches(0.58), Inches(10.0), Inches(0.34),
         size=12.5, color=GRAY)
rect(s, Inches(0.9), Inches(5.6), Inches(11.55), Inches(1.05), fill=NAVY)
text(s, "三问结束后不要急着报价，先总结他的话：", Inches(1.2), Inches(5.74),
     Inches(10), Inches(0.3), size=12.5, color=GOLD, bold=True)
text(s, "“所以您现在是——新客户全靠人跑，线上等于没有；销售天天发朋友圈，谁也说不清有没有用；陌生买家看到您也不敢联系。是这样吧？”",
     Inches(1.2), Inches(6.08), Inches(11), Inches(0.45), size=14,
     color=WHITE, bold=True)
footer(s)

quote_page("他点头，你就已经成交一半了。", tag="三问法的收口",
           note="客户自己承认的问题，不需要你再去说服。", bg=NAVY_D)

s = new_slide()
y = title_bar(s, "价值塑造", "三个算账：用他自己的数字算")
accs = [("算人力", "打 980", "“您有几个销售？每人每天发朋友圈半小时——X 个人一年就是几百个小时。这些时间您已经付了工资，但发的都是客户不想看的内容，等于白花。”"),
        ("算线索", "打 980 / 1,980", "“钢贸一单利润多少？就算几千块。3 条询盘成 1 条，980 就回来了；10 条成 2 条，1,980 也回来了。剩下工具和认证都是白得的。”"),
        ("算丢单", "打 1,980", "“过去半年有没有过”客户问了价、后来没影了“？有多少是因为他觉得不太放心？一单没成就是几千上万，认证一年才一千块。”")]
for i, (name, use, script) in enumerate(accs):
    t = y + Inches(1.35) * i
    rect(s, Inches(0.9), t, Inches(11.55), Inches(1.2), fill=LIGHT, line=LINE, lw=1.0)
    rect(s, Inches(0.9), t, Inches(0.06), Inches(1.2), fill=ORANGE)
    text(s, name, Inches(1.2), t + Inches(0.18), Inches(1.8), Inches(0.35),
         size=18, color=NAVY, bold=True)
    text(s, use, Inches(1.2), t + Inches(0.65), Inches(1.8), Inches(0.3),
         size=12, color=ORANGE, bold=True)
    text(s, script, Inches(3.15), t + Inches(0.2), Inches(9.1), Inches(0.85),
         size=13.5, color=DARKTXT, spacing=1.3)
footer(s)

# ================= 06 异议 =================
section_page("06", "抗拒解除", "17 个高频异议，抽考 8 条",
             ["第 1、2、4、5、8 条必考"])

objs = [
    ("你们不做交易，我上来干什么？", "最致命",
     "正因为我们不做交易，您才敢用。我们不抽成、不碰您的客户、不跟您抢生意。买家看到您直接打给您，价格您自己谈、钱您自己收。做交易的平台，您敢把客户名单和成本价放上去吗？", True),
    ("你们平台有人吗？没流量我上去也没用。", "第二致命",
     "三步答：① 给硬数据（2,100+ 商家、2 万用户、每天 1,000+ 搜索、150 万吨）② 说流量来源（重点讲 3,000 家老客户在找货）③ “您不用信我说的，免费开七天看数据”。", True),
    ("我先用免费的看看。", "拖延",
     "应该的，免费版永久有效。只提醒两点：海报工具只有 7 天，线索您只能看到条数看不到人。到时候看着有 5 条线索点不开，别着急就行。（留钩子，不硬推）", False),
    ("朋友圈我自己也能发，要你工具干嘛？", "必考",
     "两层答：① 内容不一样——您发的是“我有货”没人点，我们一半是行情，买家每天要看价格。您会做图，但您没有行情数据。② 我们的能收回数据：谁扫的、从哪来，全在看板上。", True),
    ("980 能给我带来多少生意？", "ROI",
     "我不能承诺成交，线索是询盘不是订单。但 3 条询盘成 1 条就回本，工具全年不限次用。我承诺的是让您看得见效果。另外每周有运营指导课，我们教您怎么把询盘跟成单。", True),
    ("线索是真的吗？会不会重复？", "质疑",
     "都是买家在平台上看您资源后留下的真实行为，重复的会去重。您先开 980 拿 3 条验证——这 3 条就是让您验真假的。", False),
    ("认证还要花钱？工商信息不是公开的？", "认证",
     "是公开的，但买家不会一家家去查。我们验过、标出来、按月复核。而且五维里的实景仓勘和生产厂家认证，工商信息里根本没有——那才是买家真想确认的。", False),
    ("我已经在找钢网了。", "竞品·必考",
     "继续用，不冲突。那是您的交易渠道，我们是您自己的线上门店和拓客工具。而且有一点可以对比：他们平台自己也有货源和交易员，我们不做交易。您的客户和报价放哪边更放心，您自己判断。", True),
    ("9,800 太贵了。", "价格",
     "一个月 800 多。这一档买的不是线索是位置——推荐位、Banner、置顶一天就那么几个坑。一次置顶接到一个大单，这一年就回来了。还配个顾问带着您做。", False),
    ("我考虑一下 / 跟老板商量。", "拖延",
     "应该商量。不过有件事今天就能定：免费开通不用商量。我现在给您开好、把资源挂上、海报发一张，七天后您拿着数据去跟老板说，比空口商量有用得多。（把大决策换成小决策）", True),
    ("能不能便宜点 / 打折？", "议价",
     "价格全国统一，我打不了折——打了对先买的客户不公平。但我能帮您两件事：把资源全挂满、店铺内容做好；再上门培训一次您的销售用海报工具。", False),
    ("我们规模小，用不上。", "自我否定",
     "规模小才更需要——大厂不上平台买家也知道他，您不上买家就真不知道您。而且您先用免费的，一分钱不花。", False),
    ("我不会用，买了也是白买。", "服务",
     "每周一场线上运营指导课，专门教怎么发海报、什么时间发、店铺怎么装。开通后我把您和销售拉进群，第一张海报我当面教会您。用不起来是我们的责任。", False),
    ("我能不能上你们那个行情直播？", "高意向信号",
     "可以——旗舰版和尊享版含上镜位，您真人出镜报现货，早盘那场买家最多。这条直播您还可以转给自己的采购客户看，等于我们帮您做了一场专属宣传。", False),
    ("你们平台会不会做两年就没了？", "风险",
     "光研发工程师就有近 40 位，公司近 100 人，平台每个月都在更新。而且您的客户资料在自己后台、导得出来——我们不做交易也不碰您的客户，这些本来就是您的。", False),
    ("你们是搞互联网的，不懂钢贸。", "不懂行",
     "我们创始人自己就是干钢贸出身的——做过销售、管过财务、当过总经理。平台上每个功能都是他当年被折磨过之后想出来的。您觉得哪儿不对，直接跟我说，我们改。", False),
    ("买了之后没人管我怎么办？", "服务",
     "三层服务：全员每周运营指导课；认证版起有在线人工客服；旗舰版配专属运营顾问、尊享版配 1 对 1 客户成功经理。而且教您的人是钢贸出身的。", False),
]
for chunk_i in range(0, len(objs), 5):
    chunk = objs[chunk_i:chunk_i + 5]
    s = new_slide()
    y = title_bar(s, "异议库", f"高频异议标准答案　{chunk_i + 1}—{chunk_i + len(chunk)} / 17")
    for k, (q, tag, a, must) in enumerate(chunk):
        t = y + Inches(1.1) * k
        rect(s, Inches(0.9), t, Inches(11.55), Inches(1.0), fill=WHITE,
             line=ORANGE if must else LINE, lw=1.5 if must else 1.0)
        rect(s, Inches(0.9), t, Inches(0.055), Inches(1.0),
             fill=ORANGE if must else NAVY)
        text(s, str(chunk_i + k + 1), Inches(1.06), t + Inches(0.12),
             Inches(0.42), Inches(0.3), size=15, color=ORANGE if must else NAVY,
             bold=True)
        text(s, [("Q：", True, GRAY), (q, True, NAVY),
                 ("　［" + tag + "］", False, ORANGE if must else GRAY)],
             Inches(1.58), t + Inches(0.1), Inches(10.7), Inches(0.32), size=13.5)
        text(s, "A：" + a, Inches(1.58), t + Inches(0.44), Inches(10.72),
             Inches(0.52), size=11.3, color=DARKTXT, spacing=1.24)
    text(s, "橙色边框 = 培训必考条目。总原则：只讲事实与分工，绝不贬低对手。",
         Inches(0.95), Inches(6.62), Inches(11), Inches(0.32), size=12.5,
         color=GRAY)
    footer(s)

# ================= 07 成交 =================
section_page("07", "成交动作", "三前提 · 三句话 · 成交后三件事")

s = new_slide()
y = title_bar(s, "成交", "三句成交话术（按情况选一句）")
closes = [("线索型（最强）", "“您店里现在有 X 条线索，都是看过您资源的买家。980 开通，今天就能联系上。我现在帮您开？”", ORANGE),
          ("到期型", "“海报工具今天到期，这七天您发了 X 张、Y 个人扫过。续上是 980 一年，停了这些就都没了。”", NAVY),
          ("升级型", "“加 1,000 块，线索从 3 条到 10 条，多的一条才 143 块，认证和精装店铺白送。这一档最划算，我给您升上去？”", ORANGE)]
for i, (name, sc, c) in enumerate(closes):
    t = y + Inches(1.18) * i
    rect(s, Inches(0.9), t, Inches(11.55), Inches(1.02), fill=WHITE, line=LINE, lw=1.1)
    rect(s, Inches(0.9), t, Inches(2.35), Inches(1.02), fill=c)
    text(s, name, Inches(0.9), t + Inches(0.33), Inches(2.35), Inches(0.35),
         size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    text(s, sc, Inches(3.5), t + Inches(0.2), Inches(8.8), Inches(0.65),
         size=14, color=NAVY, bold=True, spacing=1.28)
rect(s, Inches(0.9), Inches(5.35), Inches(11.55), Inches(1.4), fill=NAVY)
rect(s, Inches(0.9), Inches(5.35), Inches(0.07), Inches(1.4), fill=ORANGE)
text(s, "成交后必须做的三件事（决定续费与升级）", Inches(1.2), Inches(5.52),
     Inches(10), Inches(0.32), size=15, color=GOLD, bold=True)
paras(s, [
    [("① ", True, ORANGE), ("当场把店铺内容补全（认证资料、主营规格、企业介绍）——开通不等于开始用；", False, WHITE)],
    [("② ", True, ORANGE), ("拉群，教会至少 2 名销售用海报工具，并把他们拉进每周运营指导课；　", False, WHITE),
     ("③ ", True, ORANGE), ("约下次复盘：“下个月这个时候我来给您看一次数据。”", False, WHITE)],
], Inches(1.2), Inches(5.92), Inches(11.0), Inches(0.8), size=13, spacing=1.35,
    space_after=4)
footer(s)

# ================= 08 培训与红线 =================
section_page("08", "培训与红线", "3 天怎么练 · 9 条不许碰")

s = new_slide()
y = title_bar(s, "培训安排", "3 天大纲与考核")
days = [("Day 1", "认知与动手", ["三句地基逐字复述", "五档一句话 + 线索阶梯背诵",
                             "平台四个数据 + 六大流量来源", "海报机制必背话术",
                             "实操：注册、上架 10 个规格、做 3 张海报、看懂看板（限时 30 分钟）"]),
        ("Day 2", "痛点与异议", ["三问法角色演练：问出并总结", "三个算账",
                              "17 条异议抽考 8 条（1/2/4/5/8 必考）",
                              "980 → 1,980 黄金话术口试",
                              "信任状用法纪律 + 创始人两个用法"]),
        ("Day 3", "实战", ["完整路径通关演练", "陪访 2 家真实客户",
                          "至少 1 家当场免费开通", "当场铺满资源",
                          "当场做出 1 张行情海报"])]
for i, (d, name, items) in enumerate(days):
    l = Inches(0.9) + Inches(3.93) * i
    rect(s, l, y, Inches(3.66), Inches(3.6), fill=WHITE, line=LINE, lw=1.1)
    rect(s, l, y, Inches(3.66), Inches(0.72), fill=NAVY if i < 2 else ORANGE)
    text(s, d, l + Inches(0.22), y + Inches(0.1), Inches(3.2), Inches(0.32),
         size=17, color=WHITE, bold=True)
    text(s, name, l + Inches(0.22), y + Inches(0.42), Inches(3.2), Inches(0.26),
         size=12, color=RGBColor(0xC6, 0xD8, 0xE8))
    paras(s, [[("· ", True, ORANGE), (x, False, DARKTXT)] for x in items],
          l + Inches(0.22), y + Inches(0.92), Inches(3.24), Inches(2.5),
          size=11.8, spacing=1.3, space_after=7)
rect(s, Inches(0.9), Inches(6.0), Inches(11.55), Inches(0.72), fill=ORANGE)
text(s, "出师标准（四条全达标才可独立跑）：五句一句话零错误 · 平台四个数据零错误 · 三问法能自然问出并总结 · 三个动作能独立完成",
     Inches(1.2), Inches(6.19), Inches(11), Inches(0.4), size=14,
     color=WHITE, bold=True)
footer(s)

s = new_slide()
y = title_bar(s, "纪律", "9 条销售红线（违反即追责）")
reds = ["不许承诺线索一定成交 —— 统一口径：线索是询盘，不是订单。",
        "不许承诺具体成交金额、客户数或排名位次。",
        "不许打折、不许私下让价 —— 要优惠就给“帮做店铺内容 + 上门培训”。",
        "不许说“我们也能帮你成交/帮你卖货” —— 违反平台不做交易的定位。",
        "不许贬低竞对平台 —— 统一用“继续用，不冲突”的分工话术。",
        "不许代替客户填写认证资料或提供虚假认证材料。",
        "不许把某商家的买家线索透露给其他商家。",
        "不许自行估算或加码平台数据；不许把“用户数”说成“买家数”。",
        "不许承诺上镜位的具体时段与出镜效果（排期由运营统一安排）。"]
for i, r in enumerate(reds):
    col, row = i % 2, i // 2
    l = Inches(0.9) + Inches(5.95) * col
    t = y + Inches(0.72) * row
    rect(s, l, t, Inches(5.6), Inches(0.6), fill=LIGHT, line=LINE, lw=0.9)
    rect(s, l, t, Inches(0.05), Inches(0.6), fill=REDNO)
    text(s, str(i + 1), l + Inches(0.2), t + Inches(0.14), Inches(0.35),
         Inches(0.3), size=14, color=REDNO, bold=True)
    text(s, r, l + Inches(0.62), t + Inches(0.09), Inches(4.85), Inches(0.45),
         size=11.8, color=DARKTXT, spacing=1.22)
footer(s)

# ================= 随身卡 =================
s = new_slide(NAVY)
rect(s, Inches(0), Inches(0), Inches(0.22), H, fill=ORANGE)
text(s, "POCKET CARD", Inches(0.9), Inches(0.45), Inches(6), Inches(0.3),
     size=12, color=GOLD, bold=True)
text(s, "一页纸随身卡", Inches(0.86), Inches(0.75), Inches(8), Inches(0.6),
     size=30, color=WHITE, bold=True)
blocks = [
    ("三句地基", ["不做交易、不抽成、不抢客户 —— 只让买家看到你、相信你",
                "五档 = 有没有 / 有没有用 / 敢不敢 / 多不多 / 是不是第一眼",
                "免费版只见条数不见人 —— 等线索出现，它自己会成交"]),
    ("平台四个数据（每月更新）", ["2,100+ 商家 · 20,000 用户",
                          "每天 1,000+ 次搜索 · 日活跃库存 150 万吨+",
                          "讲数据必须带趋势：“每个月都在往上走”"]),
    ("五句一句话", ["0 元：先免费开个店，让买家能搜到你",
                "980：让销售发的朋友圈第一次能看见效果",
                "1,980：让陌生买家敢联系你",
                "9,800：让平台把流量推给你",
                "19,800：成为买家第一眼看到的那家"]),
    ("三问法", ["新客户主要从哪来？",
              "让销售发朋友圈吗？怎么知道有没有用？",
              "陌生买家凭什么敢给您打电话？"]),
    ("Day 0 四动作", ["当场开通 → 挂 5—10 个规格",
                  "当场做“行情 + 现货”海报（边做边讲）",
                  "加微信埋钩子"]),
    ("线索单价 · 黄金话术", ["980→327 元　1,980→198 元（最低）",
                      "9,800→327 元　19,800→396 元",
                      "加 1,000 块多 7 条，一条才 143 块"]),
]
for i, (h, items) in enumerate(blocks):
    col, row = i % 3, i // 3
    l = Inches(0.86) + Inches(4.02) * col
    t = Inches(1.62) + Inches(2.62) * row
    rect(s, l, t, Inches(3.75), Inches(2.32), fill=NAVY_D, line=NAVY_L, lw=1.0)
    rect(s, l, t, Inches(3.75), Pt(4), fill=ORANGE)
    text(s, h, l + Inches(0.2), t + Inches(0.2), Inches(3.4), Inches(0.3),
         size=14, color=GOLD, bold=True)
    paras(s, [[("· ", True, ORANGE), (x, False, RGBColor(0xDC, 0xE7, 0xF1))]
              for x in items],
          l + Inches(0.2), t + Inches(0.62), Inches(3.38), Inches(1.6),
          size=11, spacing=1.3, space_after=5)
rect(s, Inches(0.86), Inches(6.78), Inches(11.6), Inches(0.42), fill=ORANGE)
text(s, "最狠的一句：“您让销售天天发朋友圈——您不知道有没有用，他也不知道。好销售就是这么被磨走的。”",
     Inches(1.1), Inches(6.85), Inches(11.2), Inches(0.3), size=13,
     color=WHITE, bold=True)
footer(s)

# ================= 结束 =================
s = new_slide(NAVY)
rect(s, Inches(0), Inches(0), Inches(0.26), H, fill=ORANGE)
text(s, "现在开始做的三件事", Inches(1.1), Inches(1.55), Inches(9), Inches(0.5),
     size=16, color=GOLD, bold=True)
text(s, "背五句　问三问　做四动作", Inches(1.05), Inches(2.15), Inches(11),
     Inches(1.1), size=48, color=WHITE, bold=True)
rect(s, Inches(1.12), Inches(3.42), Inches(1.8), Pt(4), fill=ORANGE)
paras(s, [
    [("剩下的，让线索自己去成交。", True, ORANGE)],
    [("", False, WHITE)],
    [("完整文字版见《货袋子平台销售话术手册》", False, RGBColor(0x9F, 0xBA, 0xD2))],
    [("www.huodaizi.com", False, RGBColor(0x7E, 0x9C, 0xB8))],
], Inches(1.05), Inches(3.85), Inches(10), Inches(2.0), size=22, spacing=1.4,
    space_after=6)
footer(s)

out = "../货袋子平台销售培训.pptx"
prs.save(out)
print("OK ->", out, "共", len(prs.slides.__iter__.__self__._sldIdLst), "页")
