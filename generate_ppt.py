"""生成一季度运营分析会汇报 PPT（咨询类公司风格，主题色：中建蓝 #0070C0）。"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy
from lxml import etree

# ======== 主题配置 ========
PRIMARY = RGBColor(0x00, 0x70, 0xC0)        # 中建蓝
PRIMARY_DARK = RGBColor(0x00, 0x4F, 0x8B)   # 深蓝
PRIMARY_LIGHT = RGBColor(0xD9, 0xE9, 0xF6)  # 浅蓝
ACCENT = RGBColor(0xF5, 0xA6, 0x23)         # 强调橙
GRAY_900 = RGBColor(0x22, 0x2B, 0x38)
GRAY_700 = RGBColor(0x4A, 0x55, 0x68)
GRAY_500 = RGBColor(0x8C, 0x95, 0xA6)
GRAY_300 = RGBColor(0xD7, 0xDC, 0xE5)
GRAY_100 = RGBColor(0xF2, 0xF5, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC0, 0x39, 0x2B)
GREEN = RGBColor(0x2E, 0x8B, 0x57)

FONT_TITLE = "微软雅黑"
FONT_BODY = "微软雅黑"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]

# ======== 工具函数 ========

def set_run(run, text, *, font=FONT_BODY, size=14, bold=False, color=GRAY_900):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("eastAsia", "cs", "ascii"):
        existing = rPr.find(qn(f"a:{tag}"))
        if existing is not None:
            rPr.remove(existing)
        ele = etree.SubElement(rPr, qn(f"a:{tag}"))
        ele.set("typeface", font)


def add_text(slide, left, top, width, height, text, *, font=FONT_BODY,
             size=14, bold=False, color=GRAY_900, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, line_spacing=1.25):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    set_run(p.add_run(), text, font=font, size=size, bold=bold, color=color)
    return box


def add_multiline(slide, left, top, width, height, lines, *, font=FONT_BODY,
                  size=14, color=GRAY_700, align=PP_ALIGN.LEFT,
                  anchor=MSO_ANCHOR.TOP, line_spacing=1.35):
    """lines: list of (text, bold, size, color) tuples or plain str."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    for i, item in enumerate(lines):
        if isinstance(item, str):
            text, b, s, c = item, False, size, color
        else:
            text = item[0]
            b = item[1] if len(item) > 1 else False
            s = item[2] if len(item) > 2 else size
            c = item[3] if len(item) > 3 else color
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        set_run(p.add_run(), text, font=font, size=s, bold=b, color=c)
    return box


def add_rect(slide, left, top, width, height, *, fill=PRIMARY, line=None,
             shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    return shp


def add_line(slide, x1, y1, x2, y2, *, color=GRAY_300, weight=1.0):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = color
    line.line.width = Pt(weight)
    return line


def add_page_chrome(slide, page_no, total, section_label, title=None, subtitle=None):
    """添加咨询风格的页眉、页脚和页码。"""
    # 顶部细色块
    add_rect(slide, Emu(0), Emu(0), SLIDE_W, Inches(0.06), fill=PRIMARY)
    # 左上 logo 标识
    add_rect(slide, Inches(0.5), Inches(0.22), Inches(0.18), Inches(0.32), fill=PRIMARY)
    add_text(slide, Inches(0.75), Inches(0.18), Inches(4.5), Inches(0.4),
             "中建八局 · 一季度运营分析会", font=FONT_BODY, size=11,
             bold=True, color=PRIMARY_DARK, anchor=MSO_ANCHOR.MIDDLE)
    # 右上栏目标签
    add_text(slide, Inches(8.0), Inches(0.18), Inches(4.83), Inches(0.4),
             section_label, font=FONT_BODY, size=11, bold=False,
             color=GRAY_500, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # 标题区
    if title:
        add_text(slide, Inches(0.5), Inches(0.7), Inches(12.3), Inches(0.6),
                 title, font=FONT_TITLE, size=24, bold=True, color=GRAY_900,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 标题装饰
        add_rect(slide, Inches(0.5), Inches(1.3), Inches(0.4), Inches(0.05),
                 fill=PRIMARY)
        add_rect(slide, Inches(0.92), Inches(1.3), Inches(0.15), Inches(0.05),
                 fill=ACCENT)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(1.4), Inches(12.3), Inches(0.4),
                 subtitle, font=FONT_BODY, size=12, color=GRAY_500,
                 anchor=MSO_ANCHOR.TOP)

    # 底部
    add_line(slide, Inches(0.5), Inches(7.12), Inches(12.83), Inches(7.12),
             color=GRAY_300, weight=0.75)
    add_text(slide, Inches(0.5), Inches(7.18), Inches(6), Inches(0.3),
             "信息化运营管理部　|　季度运营分析专题汇报",
             font=FONT_BODY, size=9, color=GRAY_500)
    add_text(slide, Inches(7), Inches(7.18), Inches(5.83), Inches(0.3),
             f"P {page_no:02d} / {total:02d}",
             font=FONT_BODY, size=9, color=GRAY_500, align=PP_ALIGN.RIGHT)


def add_section_title(slide, number, title, subtitle):
    """章节封面页。"""
    # 整页底色块
    add_rect(slide, Emu(0), Emu(0), SLIDE_W, SLIDE_H, fill=GRAY_100)
    # 左侧蓝色块
    add_rect(slide, Emu(0), Emu(0), Inches(4.6), SLIDE_H, fill=PRIMARY)
    # 装饰斜线（用矩形堆叠模拟）
    add_rect(slide, Inches(4.6), Inches(0), Inches(0.08), SLIDE_H, fill=PRIMARY_DARK)
    add_rect(slide, Inches(4.7), Inches(0), Inches(0.04), SLIDE_H, fill=ACCENT)

    # 大编号
    add_text(slide, Inches(0.6), Inches(1.6), Inches(3.8), Inches(2.2),
             f"PART  {number}", font=FONT_TITLE, size=20, bold=False,
             color=RGBColor(0xB7, 0xD7, 0xF2))
    add_text(slide, Inches(0.5), Inches(2.0), Inches(4.0), Inches(3.0),
             f"0{number}", font=FONT_TITLE, size=140, bold=True, color=WHITE)

    # 右侧标题
    add_text(slide, Inches(5.2), Inches(2.6), Inches(7.8), Inches(0.6),
             "SECTION", font=FONT_BODY, size=11, color=PRIMARY)
    add_text(slide, Inches(5.2), Inches(3.0), Inches(7.8), Inches(1.0),
             title, font=FONT_TITLE, size=36, bold=True, color=GRAY_900)
    # 装饰条
    add_rect(slide, Inches(5.2), Inches(4.2), Inches(0.6), Inches(0.06),
             fill=PRIMARY)
    add_rect(slide, Inches(5.83), Inches(4.2), Inches(0.2), Inches(0.06),
             fill=ACCENT)
    add_text(slide, Inches(5.2), Inches(4.4), Inches(7.5), Inches(1.5),
             subtitle, font=FONT_BODY, size=14, color=GRAY_700,
             line_spacing=1.5)


# ======== 幻灯片内容生成 ========
TOTAL_PAGES = 19
slides_meta = []

# ---------- P1 封面 ----------
def slide_cover():
    s = prs.slides.add_slide(BLANK)
    # 整页背景
    add_rect(s, Emu(0), Emu(0), SLIDE_W, SLIDE_H, fill=WHITE)
    # 左侧色块
    add_rect(s, Emu(0), Emu(0), Inches(5.5), SLIDE_H, fill=PRIMARY)
    # 装饰几何
    add_rect(s, Inches(5.5), Emu(0), Inches(0.08), SLIDE_H, fill=PRIMARY_DARK)
    add_rect(s, Inches(5.6), Emu(0), Inches(0.04), SLIDE_H, fill=ACCENT)
    # 右下浅色块
    add_rect(s, Inches(5.6), Inches(5.5), Inches(7.7), Inches(2.0), fill=GRAY_100)

    # 左侧 logo 标识
    add_rect(s, Inches(0.7), Inches(0.6), Inches(0.22), Inches(0.4), fill=WHITE)
    add_text(s, Inches(1.0), Inches(0.55), Inches(4.5), Inches(0.5),
             "CSCEC · 第八工程局", font=FONT_BODY, size=12, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

    # 标题
    add_text(s, Inches(0.7), Inches(2.2), Inches(4.7), Inches(0.5),
             "QUARTERLY OPERATION REVIEW", font=FONT_BODY, size=12,
             color=RGBColor(0xB7, 0xD7, 0xF2))
    add_text(s, Inches(0.7), Inches(2.6), Inches(4.7), Inches(1.4),
             "一季度运营\n分析会汇报", font=FONT_TITLE, size=44, bold=True,
             color=WHITE, line_spacing=1.15)
    add_rect(s, Inches(0.7), Inches(4.5), Inches(0.6), Inches(0.06),
             fill=ACCENT)
    add_text(s, Inches(0.7), Inches(4.7), Inches(4.7), Inches(1.0),
             "信息化系统建设与运营专题", font=FONT_BODY, size=18,
             color=WHITE, line_spacing=1.4)

    # 右侧信息
    add_text(s, Inches(6.0), Inches(2.4), Inches(7.0), Inches(0.4),
             "REPORT　|　Q1", font=FONT_BODY, size=11, color=PRIMARY)
    add_text(s, Inches(6.0), Inches(2.7), Inches(7.0), Inches(0.6),
             "聚焦目标 · 复盘成效 · 谋划下阶段",
             font=FONT_TITLE, size=22, bold=True, color=GRAY_900)
    add_line(s, Inches(6.0), Inches(3.5), Inches(11.5), Inches(3.5),
             color=GRAY_300)

    # 关键数据展示（cover 上的 3 组）
    metrics = [
        ("313.8万元", "回款额（全部为现金）", PRIMARY),
        ("32249件", "信创OA累计发起流程数", PRIMARY_DARK),
        ("99.9%", "系统可用性目标", ACCENT),
    ]
    for i, (val, label, color) in enumerate(metrics):
        x = Inches(6.0 + i * 2.3)
        add_text(s, x, Inches(3.7), Inches(2.2), Inches(0.6), val,
                 font=FONT_TITLE, size=22, bold=True, color=color)
        add_text(s, x, Inches(4.3), Inches(2.2), Inches(0.4), label,
                 font=FONT_BODY, size=10, color=GRAY_700)

    # 右下角信息
    add_text(s, Inches(6.0), Inches(5.8), Inches(7.0), Inches(0.4),
             "汇报部门：信息化运营管理部", font=FONT_BODY, size=12,
             bold=True, color=GRAY_900)
    add_text(s, Inches(6.0), Inches(6.2), Inches(7.0), Inches(0.4),
             "汇报时间：二〇二六年四月", font=FONT_BODY, size=12,
             color=GRAY_700)
    add_text(s, Inches(6.0), Inches(6.6), Inches(7.0), Inches(0.4),
             "汇 报 人：×××", font=FONT_BODY, size=12, color=GRAY_700)


# ---------- P2 目录 ----------
def slide_toc():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 2, TOTAL_PAGES, "目录　CONTENTS",
                    title="目　　录", subtitle="CONTENTS")

    items = [
        ("01", "一季度总体指标完成情况", "Q1 KPI Performance"),
        ("02", "一季度工作进展", "Q1 Work Progress"),
        ("03", "存在的问题及不足", "Issues & Gaps"),
        ("04", "二季度工作计划", "Q2 Action Plan"),
    ]
    top = Inches(2.2)
    for i, (no, cn, en) in enumerate(items):
        row_top = top + Inches(i * 1.05)
        # 编号块
        add_rect(s, Inches(0.8), row_top, Inches(1.2), Inches(0.85), fill=PRIMARY)
        add_text(s, Inches(0.8), row_top, Inches(1.2), Inches(0.85),
                 no, font=FONT_TITLE, size=30, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 标题区底色
        add_rect(s, Inches(2.1), row_top, Inches(10.5), Inches(0.85),
                 fill=GRAY_100)
        # 装饰短条
        add_rect(s, Inches(2.1), row_top, Inches(0.06), Inches(0.85),
                 fill=ACCENT)
        # 中文标题
        add_text(s, Inches(2.4), row_top + Inches(0.08), Inches(8.0), Inches(0.45),
                 cn, font=FONT_TITLE, size=18, bold=True, color=GRAY_900,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 英文副标
        add_text(s, Inches(2.4), row_top + Inches(0.5), Inches(8.0), Inches(0.3),
                 en, font=FONT_BODY, size=10, color=GRAY_500)
        # 右侧箭头
        add_text(s, Inches(11.5), row_top, Inches(1.0), Inches(0.85),
                 "▶", font=FONT_BODY, size=14, color=PRIMARY,
                 align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


# ---------- P3 章节1封面 ----------
def slide_part1_cover():
    s = prs.slides.add_slide(BLANK)
    add_section_title(
        s, 1, "一季度总体指标完成情况",
        "聚焦合同、收入、利润、回款、负流等核心指标，"
        "全面复盘一季度经营运行态势，研判完成情况与质量。",
    )


# ---------- P4 KPI 概览 ----------
def slide_kpi_dashboard():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 4, TOTAL_PAGES, "01　一季度总体指标完成情况",
                    title="01　一季度核心指标一览",
                    subtitle="经营指标 · 现金质量 · 风险信号")

    cards = [
        ("合 同 额", "33.13", "万元", PRIMARY, "Q1 累计签约金额"),
        ("营业收入", "6.25", "万元", PRIMARY, "Q1 累计实现营收"),
        ("利 润 额", "243.67", "万元", GREEN, "经营效益保持稳健"),
        ("回 款 额", "313.8", "万元", PRIMARY_DARK, "现金回款 313.8 万元"),
        ("负流项目占比", "17", "%", ACCENT, "项目纠偏需重点关注"),
        ("总负流金额", "-85.3", "万元", RED, "现金流压力指标"),
    ]
    cols = 3
    card_w = Inches(4.0)
    card_h = Inches(1.85)
    gap_x = Inches(0.15)
    gap_y = Inches(0.25)
    start_x = Inches(0.5)
    start_y = Inches(2.0)
    for i, (label, val, unit, color, note) in enumerate(cards):
        r, c = divmod(i, cols)
        x = start_x + (card_w + gap_x) * c
        y = start_y + (card_h + gap_y) * r
        # 卡片底
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=GRAY_300)
        # 顶部色条
        add_rect(s, x, y, card_w, Inches(0.08), fill=color)
        # 标签
        add_text(s, x + Inches(0.25), y + Inches(0.18), card_w - Inches(0.5),
                 Inches(0.35), label, font=FONT_BODY, size=12, color=GRAY_500)
        # 数值
        add_text(s, x + Inches(0.25), y + Inches(0.55), card_w - Inches(0.5),
                 Inches(0.7), val, font=FONT_TITLE, size=34, bold=True,
                 color=color)
        # 单位
        add_text(s, x + Inches(2.7), y + Inches(0.85), Inches(1.0),
                 Inches(0.4), unit, font=FONT_BODY, size=12, color=GRAY_700)
        # 备注
        add_text(s, x + Inches(0.25), y + Inches(1.35), card_w - Inches(0.5),
                 Inches(0.4), note, font=FONT_BODY, size=10, color=GRAY_500)

    # 底部小结条
    add_rect(s, Inches(0.5), Inches(6.05), Inches(12.33), Inches(0.85),
             fill=PRIMARY_LIGHT)
    add_rect(s, Inches(0.5), Inches(6.05), Inches(0.08), Inches(0.85),
             fill=PRIMARY)
    add_text(s, Inches(0.75), Inches(6.05), Inches(12.0), Inches(0.85),
             "总体研判：回款全部为现金，资金质量较高；合同与营业收入受年初节奏影响规模偏小；"
             "负流占比 17%、总负流 -85.3 万元，需重点强化项目纠偏与现金流管控。",
             font=FONT_BODY, size=12, bold=True, color=PRIMARY_DARK,
             anchor=MSO_ANCHOR.MIDDLE)


# ---------- P5 指标解读 ----------
def slide_kpi_analysis():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 5, TOTAL_PAGES, "01　一季度总体指标完成情况",
                    title="01　核心指标质量分析",
                    subtitle="资金质量 · 经营效益 · 风险信号")

    # 左侧：回款结构
    add_rect(s, Inches(0.5), Inches(2.0), Inches(6.0), Inches(4.7),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(0.5), Inches(2.0), Inches(6.0), Inches(0.5),
             fill=PRIMARY)
    add_text(s, Inches(0.7), Inches(2.0), Inches(5.8), Inches(0.5),
             "▎一、回款结构分析", font=FONT_TITLE, size=14, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.8), Inches(2.7), Inches(5.6), Inches(0.5),
             "现金回款占比：100%", font=FONT_TITLE, size=20,
             bold=True, color=GREEN)
    # 进度条
    add_rect(s, Inches(0.8), Inches(3.4), Inches(5.4), Inches(0.35),
             fill=GRAY_100)
    add_rect(s, Inches(0.8), Inches(3.4), Inches(5.4), Inches(0.35),
             fill=GREEN)
    add_text(s, Inches(0.8), Inches(3.4), Inches(5.4), Inches(0.35),
             "313.8 / 313.8 万元", font=FONT_BODY, size=11, bold=True,
             color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_multiline(s, Inches(0.8), Inches(4.0), Inches(5.4), Inches(2.5), [
        ("现金部分：313.8 万元", True, 12, GRAY_900),
        ("非现金部分：0 万元", True, 12, GRAY_900),
        ("", False, 6, GRAY_900),
        ("研判：一季度回款全部为现金，资金质量较高；后续应继续坚持现金优先原则，"
         "兼顾回款规模与现金占比。", False, 11, GRAY_700),
    ])

    # 右侧：负流风险
    add_rect(s, Inches(6.83), Inches(2.0), Inches(6.0), Inches(4.7),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(6.83), Inches(2.0), Inches(6.0), Inches(0.5),
             fill=RED)
    add_text(s, Inches(7.03), Inches(2.0), Inches(5.8), Inches(0.5),
             "▎二、负流风险研判", font=FONT_TITLE, size=14, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

    add_text(s, Inches(7.13), Inches(2.7), Inches(2.7), Inches(0.5),
             "项目个数占比", font=FONT_BODY, size=11, color=GRAY_500)
    add_text(s, Inches(7.13), Inches(3.05), Inches(2.7), Inches(0.7),
             "17%", font=FONT_TITLE, size=30, bold=True, color=ACCENT)

    add_text(s, Inches(9.93), Inches(2.7), Inches(2.7), Inches(0.5),
             "总负流金额", font=FONT_BODY, size=11, color=GRAY_500)
    add_text(s, Inches(9.93), Inches(3.05), Inches(2.7), Inches(0.7),
             "-85.3 万元", font=FONT_TITLE, size=30, bold=True, color=RED)

    add_line(s, Inches(7.13), Inches(4.0), Inches(12.6), Inches(4.0),
             color=GRAY_300)
    add_multiline(s, Inches(7.13), Inches(4.1), Inches(5.5), Inches(2.5), [
        ("应对举措：", True, 12, GRAY_900),
        ("· 项目层面：开展负流项目专项盘点，逐项制定纠偏方案；", False, 11, GRAY_700),
        ("· 过程管控：强化产值、回款、成本三同步管理；", False, 11, GRAY_700),
        ("· 月度复盘：建立月度负流通报与预警机制，压实责任。", False, 11, GRAY_700),
    ])


# ---------- P6 章节2封面 ----------
def slide_part2_cover():
    s = prs.slides.add_slide(BLANK)
    add_section_title(
        s, 2, "一季度工作进展",
        "围绕信创OA系统、数智工会系统等核心信息化平台，"
        "统筹推进系统建设、流程优化与日常运维，全面复盘工作成效。",
    )


# ---------- P7 工作进展总览 ----------
def slide_progress_overview():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 7, TOTAL_PAGES, "02　一季度工作进展",
                    title="02　一季度工作进展总览",
                    subtitle="两条主线 · 多维突破")

    # 中部一句话总结条
    add_rect(s, Inches(0.5), Inches(2.0), Inches(12.33), Inches(0.7),
             fill=PRIMARY_LIGHT)
    add_text(s, Inches(0.5), Inches(2.0), Inches(12.33), Inches(0.7),
             "以业务需求为导向　|　以系统建设为牵引　|　以稳定运维为保障",
             font=FONT_TITLE, size=14, bold=True, color=PRIMARY_DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 两大主线
    blocks = [
        ("（一）系统建设情况", PRIMARY, [
            "信创OA系统建设：核心功能切换上线、26 家二级单位大家签适配",
            "信创OA系统建设：城市运营公司 104 条流程搭建，流程优化迭代",
            "数智工会系统：5 大模块 2 月 5 日全局上线，覆盖 2710 名干部",
            "数智工会系统：完成干部端 9 个功能模块开发",
        ]),
        ("（二）系统运维情况", PRIMARY_DARK, [
            "组织机构：保障中建科技合并、运营投资类公司整合",
            "八局通：处理工单 475 条，应用上线 8 个，缺陷修复 10/13",
            "信创OA上云：汉数科 5 类应用搭建上线，4 月 1 日切换审批流程",
            "八局云盘：版本升级，保障 8 万用户、69T 文件迁移",
        ]),
    ]
    top = Inches(3.0)
    for i, (title, color, items) in enumerate(blocks):
        x = Inches(0.5 + i * 6.33)
        # 标题
        add_rect(s, x, top, Inches(6.0), Inches(0.6), fill=color)
        add_text(s, x + Inches(0.3), top, Inches(5.7), Inches(0.6), title,
                 font=FONT_TITLE, size=16, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 内容卡
        add_rect(s, x, top + Inches(0.6), Inches(6.0), Inches(3.2),
                 fill=WHITE, line=GRAY_300)
        for j, line in enumerate(items):
            y = top + Inches(0.85 + j * 0.7)
            # 圆点
            add_rect(s, x + Inches(0.25), y + Inches(0.12), Inches(0.12),
                     Inches(0.12), fill=color, shape=MSO_SHAPE.OVAL)
            add_text(s, x + Inches(0.5), y, Inches(5.4), Inches(0.5), line,
                     font=FONT_BODY, size=12, color=GRAY_700,
                     anchor=MSO_ANCHOR.MIDDLE)


# ---------- P8 信创OA - 整体推进 ----------
def slide_oa_overview():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 8, TOTAL_PAGES, "02　一季度工作进展 / 系统建设",
                    title="02-1　信创 OA 系统｜整体推进情况",
                    subtitle="切换上线 · 运维保障 · 二级承接 · 优化迭代")

    items = [
        ("①", "核心功能\n切换上线", "局总部公文、新闻、公告、流程中心等核心功能全部切换上线，"
         "二级单位除基础功能外亦实现全面上线运行。"),
        ("②", "重点流程\n专项保障", "针对发起量较大的局总部公文、新闻、财务等关键流程开展专项保障，"
         "并对二级单位流程搭建提供答疑支持。"),
        ("③", "二级单位\n搭建承接", "城市运营公司完成 103 条流程搭建；南方公司完成 130+ 条需求梳理；"
         "浙江公司流程需求处于对接报价阶段。"),
        ("④", "局总部流程\n优化迭代", "局总部流程持续优化迭代，重点针对用印等高频流程进行优化完善"
         "（由致远提供相关支持）。"),
    ]
    top = Inches(2.0)
    card_w = Inches(2.95)
    card_h = Inches(4.7)
    gap = Inches(0.15)
    for i, (no, title, desc) in enumerate(items):
        x = Inches(0.5) + (card_w + gap) * i
        add_rect(s, x, top, card_w, card_h, fill=WHITE, line=GRAY_300)
        # 顶部蓝色块
        add_rect(s, x, top, card_w, Inches(1.5), fill=PRIMARY)
        # 大编号
        add_text(s, x, top + Inches(0.1), card_w, Inches(0.6), no,
                 font=FONT_TITLE, size=30, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)
        add_text(s, x, top + Inches(0.75), card_w, Inches(0.7), title,
                 font=FONT_TITLE, size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, line_spacing=1.2)
        # 描述
        add_text(s, x + Inches(0.25), top + Inches(1.7),
                 card_w - Inches(0.5), card_h - Inches(1.9), desc,
                 font=FONT_BODY, size=12, color=GRAY_700, line_spacing=1.5)


# ---------- P9 信创OA - 重点成效 ----------
def slide_oa_highlights():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 9, TOTAL_PAGES, "02　一季度工作进展 / 系统建设",
                    title="02-2　信创 OA 系统｜重点工作成效",
                    subtitle="功能适配 · 流程建设 · 数据对接 · 运营管控")

    blocks = [
        ("一", "聚焦功能适配与开发\n夯实系统应用基础",
         "完成 26 家二级单位『大家签』适配开发；外部主数据同步功能已上线，"
         "累计同步 28 家外部单位主数据、2402 条人员主数据。", PRIMARY),
        ("二", "强化流程建设与优化\n提升办公协同效能",
         "完成城市运营公司及下级单位 104 条流程上线；优化局总部行政用印、"
         "投资运营事业部、工程研究院相关流程；优化公文套红模板。", PRIMARY_DARK),
        ("三", "推进系统迭代与数据对接\n强化数据赋能能力",
         "完成文件中心三级单位页面开发上线，规范文件流转；"
         "完成集团应用使用情况看板指标对接，实现运营数据可视化。", PRIMARY),
        ("四", "抓实系统运营管控\n保障系统平稳运行",
         "系统累计发起流程 32249 件、办结 22249 件、完成 10257 件，"
         "未出现重大系统故障及业务卡顿问题，整体运行平稳。", PRIMARY_DARK),
    ]
    top = Inches(2.0)
    row_h = Inches(1.15)
    for i, (no, title, desc, color) in enumerate(blocks):
        y = top + (row_h + Inches(0.08)) * i
        # 编号
        add_rect(s, Inches(0.5), y, Inches(1.0), row_h, fill=color)
        add_text(s, Inches(0.5), y, Inches(1.0), row_h, no,
                 font=FONT_TITLE, size=36, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 标题
        add_rect(s, Inches(1.5), y, Inches(3.6), row_h, fill=PRIMARY_LIGHT)
        add_text(s, Inches(1.7), y, Inches(3.4), row_h, title,
                 font=FONT_TITLE, size=13, bold=True, color=PRIMARY_DARK,
                 anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.3)
        # 描述
        add_rect(s, Inches(5.1), y, Inches(7.73), row_h, fill=WHITE,
                 line=GRAY_300)
        add_text(s, Inches(5.3), y, Inches(7.5), row_h, desc,
                 font=FONT_BODY, size=12, color=GRAY_700,
                 anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.45)


# ---------- P10 信创OA - 运营数据 ----------
def slide_oa_data():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 10, TOTAL_PAGES, "02　一季度工作进展 / 系统建设",
                    title="02-3　信创 OA 系统｜运营数据",
                    subtitle="量化呈现 · 数据说话")

    # 三个流程数据卡
    cards = [
        ("32,249", "发起流程", "件", PRIMARY),
        ("22,249", "办结流程", "件", PRIMARY_DARK),
        ("10,257", "完成流程", "件", GREEN),
    ]
    for i, (val, label, unit, color) in enumerate(cards):
        x = Inches(0.5 + i * 4.28)
        add_rect(s, x, Inches(2.0), Inches(4.0), Inches(2.0), fill=WHITE,
                 line=GRAY_300)
        add_rect(s, x, Inches(2.0), Inches(0.08), Inches(2.0), fill=color)
        add_text(s, x + Inches(0.3), Inches(2.15), Inches(3.6), Inches(0.4),
                 label, font=FONT_BODY, size=12, color=GRAY_500)
        add_text(s, x + Inches(0.3), Inches(2.55), Inches(3.6), Inches(1.0),
                 val, font=FONT_TITLE, size=44, bold=True, color=color)
        add_text(s, x + Inches(0.3), Inches(3.55), Inches(3.6), Inches(0.4),
                 unit, font=FONT_BODY, size=12, color=GRAY_700)

    # 新老 OA 对比
    add_text(s, Inches(0.5), Inches(4.4), Inches(12.3), Inches(0.4),
             "▎新老 OA 系统流程审批数据对比", font=FONT_TITLE, size=14,
             bold=True, color=PRIMARY_DARK)
    add_line(s, Inches(0.5), Inches(4.85), Inches(12.83), Inches(4.85),
             color=PRIMARY)

    table_data = [
        ["系统", "流程审批量", "审批人次", "说明"],
        ["新信创 OA 系统（全局）", "3.557 万条", "35.289 万人次",
         "二级单位逐步切换，4 月 1 日完成审批切换"],
        ["原 OA 系统", "10.18 万条", "54.79 万人次",
         "并行运行期保障业务连续性"],
    ]
    col_w = [Inches(3.2), Inches(2.5), Inches(2.5), Inches(4.13)]
    row_h = Inches(0.6)
    y = Inches(5.0)
    for r, row in enumerate(table_data):
        x = Inches(0.5)
        for c, txt in enumerate(row):
            if r == 0:
                add_rect(s, x, y, col_w[c], row_h, fill=PRIMARY)
                color = WHITE
                bold = True
                size = 12
            else:
                fill = WHITE if r % 2 == 1 else GRAY_100
                add_rect(s, x, y, col_w[c], row_h, fill=fill, line=GRAY_300)
                color = GRAY_900 if c < 3 else GRAY_700
                bold = c < 3
                size = 12
            add_text(s, x + Inches(0.15), y, col_w[c] - Inches(0.3), row_h,
                     txt, font=FONT_BODY, size=size, bold=bold, color=color,
                     anchor=MSO_ANCHOR.MIDDLE,
                     align=PP_ALIGN.LEFT if c == 0 or c == 3 else PP_ALIGN.CENTER)
            x += col_w[c]
        y += row_h


# ---------- P11 数智工会-推广 ----------
def slide_union_launch():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 11, TOTAL_PAGES, "02　一季度工作进展 / 系统建设",
                    title="02-4　数智工会系统｜干部端全局推广上线",
                    subtitle="2 月 5 日 · 全局上线 · 全面覆盖")

    # 时间节点
    add_rect(s, Inches(0.5), Inches(2.0), Inches(12.33), Inches(0.7),
             fill=PRIMARY)
    add_text(s, Inches(0.5), Inches(2.0), Inches(12.33), Inches(0.7),
             "■  完成荣誉奖项、劳动竞赛、技能竞赛、服务阵地、困难职工 5 大模块开发测试，2026 年 2 月 5 日全局上线",
             font=FONT_TITLE, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 6 个数据卡
    metrics = [
        ("2,223", "工会组织（个）"),
        ("2,710", "覆盖工会干部（名）"),
        ("1,321", "荣誉奖项 · 集体（条）"),
        ("955", "荣誉奖项 · 个人（条）"),
        ("2,212", "劳模先进（条）"),
        ("157", "劳动竞赛局级立项（条）"),
    ]
    cols = 3
    card_w = Inches(4.0)
    card_h = Inches(1.4)
    gap_x = Inches(0.15)
    gap_y = Inches(0.2)
    start_x = Inches(0.5)
    start_y = Inches(3.0)
    for i, (val, lbl) in enumerate(metrics):
        r, c = divmod(i, cols)
        x = start_x + (card_w + gap_x) * c
        y = start_y + (card_h + gap_y) * r
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=GRAY_300)
        add_rect(s, x, y, Inches(0.08), card_h, fill=PRIMARY)
        add_text(s, x + Inches(0.3), y + Inches(0.1), card_w - Inches(0.4),
                 Inches(0.7), val, font=FONT_TITLE, size=28, bold=True,
                 color=PRIMARY_DARK)
        add_text(s, x + Inches(0.3), y + Inches(0.85), card_w - Inches(0.4),
                 Inches(0.4), lbl, font=FONT_BODY, size=11, color=GRAY_700)

    # 补充信息
    add_rect(s, Inches(0.5), Inches(6.3), Inches(12.33), Inches(0.6),
             fill=PRIMARY_LIGHT)
    add_text(s, Inches(0.7), Inches(6.3), Inches(12.0), Inches(0.6),
             "创新工作室录入 364 个，服务阵地覆盖 511 个，应用基础数据初具规模。",
             font=FONT_BODY, size=12, bold=True, color=PRIMARY_DARK,
             anchor=MSO_ANCHOR.MIDDLE)


# ---------- P12 数智工会-运维 ----------
def slide_union_dev():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 12, TOTAL_PAGES, "02　一季度工作进展 / 系统建设",
                    title="02-5　数智工会系统｜干部端运维与开发",
                    subtitle="9 大模块 · 完善服务体系")

    add_text(s, Inches(0.5), Inches(2.0), Inches(12.3), Inches(0.5),
             "围绕基层工会干部实际工作需求，完成 9 个核心模块的开发工作，进一步丰富数智工会系统功能体系。",
             font=FONT_BODY, size=13, color=GRAY_700)

    modules = [
        ("筑缘相亲交友", "服务关怀"),
        ("福利清单", "权益保障"),
        ("幸福清单", "权益保障"),
        ("积分管理", "运营激励"),
        ("会员码", "身份服务"),
        ("慰问申请", "服务关怀"),
        ("文体协会", "文体活动"),
        ("活动中心", "文体活动"),
        ("服务阵地", "阵地建设"),
    ]
    cols = 3
    card_w = Inches(4.0)
    card_h = Inches(1.15)
    gap_x = Inches(0.15)
    gap_y = Inches(0.18)
    start_x = Inches(0.5)
    start_y = Inches(2.8)
    for i, (name, tag) in enumerate(modules):
        r, c = divmod(i, cols)
        x = start_x + (card_w + gap_x) * c
        y = start_y + (card_h + gap_y) * r
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=GRAY_300)
        # 序号圆
        add_rect(s, x + Inches(0.2), y + Inches(0.3), Inches(0.55),
                 Inches(0.55), fill=PRIMARY, shape=MSO_SHAPE.OVAL)
        add_text(s, x + Inches(0.2), y + Inches(0.3), Inches(0.55),
                 Inches(0.55), f"{i+1:02d}", font=FONT_TITLE, size=12,
                 bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 名称
        add_text(s, x + Inches(0.9), y + Inches(0.2), Inches(2.9),
                 Inches(0.45), name, font=FONT_TITLE, size=14, bold=True,
                 color=GRAY_900)
        # 标签
        add_text(s, x + Inches(0.9), y + Inches(0.65), Inches(2.9),
                 Inches(0.4), f"# {tag}", font=FONT_BODY, size=10,
                 color=PRIMARY)


# ---------- P13 运维 - 组织机构 + 八局通 ----------
def slide_ops_1():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 13, TOTAL_PAGES, "02　一季度工作进展 / 系统运维",
                    title="02-6　系统运维｜组织机构 · 八局通",
                    subtitle="基础维护 · 工单闭环 · 缺陷管理")

    # 左：组织机构
    add_rect(s, Inches(0.5), Inches(2.0), Inches(6.0), Inches(4.7),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(0.5), Inches(2.0), Inches(6.0), Inches(0.5),
             fill=PRIMARY)
    add_text(s, Inches(0.7), Inches(2.0), Inches(5.8), Inches(0.5),
             "▎组织机构维护", font=FONT_TITLE, size=14, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    # 大数字
    add_text(s, Inches(0.7), Inches(2.7), Inches(5.6), Inches(0.4),
             "在册机构单位", font=FONT_BODY, size=11, color=GRAY_500)
    add_text(s, Inches(0.7), Inches(3.05), Inches(5.6), Inches(0.9),
             "913　家", font=FONT_TITLE, size=42, bold=True, color=PRIMARY)
    add_line(s, Inches(0.7), Inches(4.1), Inches(6.3), Inches(4.1),
             color=GRAY_300)
    add_multiline(s, Inches(0.7), Inches(4.2), Inches(5.6), Inches(2.5), [
        ("一季度重点保障：", True, 12, GRAY_900),
        ("· 完成三级单位组织调整事项的平稳落地；", False, 12, GRAY_700),
        ("· 保障中建科技公司合并；", False, 12, GRAY_700),
        ("· 保障运营投资类公司整合；", False, 12, GRAY_700),
        ("· 确保组织机构数据与实际管理架构同步。", False, 12, GRAY_700),
    ])

    # 右：八局通
    add_rect(s, Inches(6.83), Inches(2.0), Inches(6.0), Inches(4.7),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(6.83), Inches(2.0), Inches(6.0), Inches(0.5),
             fill=PRIMARY_DARK)
    add_text(s, Inches(7.03), Inches(2.0), Inches(5.8), Inches(0.5),
             "▎八局通运维支持", font=FONT_TITLE, size=14, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    # 4 个小指标
    items = [
        ("475", "工单处理（条）", "全部办结闭环"),
        ("8", "应用上线（个）", "助力业务上云"),
        ("0", "系统故障（起）", "运行稳定可控"),
        ("10/13", "缺陷修复进度", "剩 3 个待新版本验证"),
    ]
    for i, (val, lbl, note) in enumerate(items):
        r, c = divmod(i, 2)
        x = Inches(6.93) + Inches(c * 2.9)
        y = Inches(2.7) + Inches(r * 2.0)
        add_rect(s, x, y, Inches(2.8), Inches(1.85), fill=GRAY_100)
        add_text(s, x + Inches(0.2), y + Inches(0.15), Inches(2.6),
                 Inches(0.3), lbl, font=FONT_BODY, size=10, color=GRAY_500)
        add_text(s, x + Inches(0.2), y + Inches(0.45), Inches(2.6),
                 Inches(0.8), val, font=FONT_TITLE, size=26, bold=True,
                 color=PRIMARY_DARK)
        add_text(s, x + Inches(0.2), y + Inches(1.3), Inches(2.6),
                 Inches(0.4), note, font=FONT_BODY, size=10, color=GRAY_700)


# ---------- P14 运维 - 信创OA上云 + 八局云盘 ----------
def slide_ops_2():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 14, TOTAL_PAGES, "02　一季度工作进展 / 系统运维",
                    title="02-7　系统运维｜信创 OA 上云 · 八局云盘",
                    subtitle="云端切换 · 平稳运行 · 升级保障")

    # 上：信创OA上云汉数科
    add_rect(s, Inches(0.5), Inches(2.0), Inches(12.33), Inches(2.4),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(0.5), Inches(2.0), Inches(12.33), Inches(0.5),
             fill=PRIMARY)
    add_text(s, Inches(0.7), Inches(2.0), Inches(12.0), Inches(0.5),
             "▎信创 OA 上云汉数科：会议预约 / 发文 / 新闻 / 通知公告 / 审批流程  ——  全部完成搭建上线",
             font=FONT_TITLE, size=13, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    sub_metrics = [
        ("966", "会议申请（次）"),
        ("11", "文件发布（条）"),
        ("11", "通知公告（条）"),
        ("1", "新闻发布（条）"),
        ("4 月 1 日", "审批流程切换"),
    ]
    for i, (v, l) in enumerate(sub_metrics):
        x = Inches(0.7) + Inches(i * 2.4)
        add_text(s, x, Inches(2.65), Inches(2.3), Inches(0.7), v,
                 font=FONT_TITLE, size=24, bold=True, color=PRIMARY)
        add_text(s, x, Inches(3.4), Inches(2.3), Inches(0.4), l,
                 font=FONT_BODY, size=11, color=GRAY_500)
        if i < 4:
            add_line(s, x + Inches(2.3), Inches(2.7),
                     x + Inches(2.3), Inches(3.9), color=GRAY_300)
    add_text(s, Inches(0.7), Inches(3.95), Inches(12.0), Inches(0.4),
             "■ 全局新 OA 累计完成审批流程 3.557 万条 / 35.289 万人次；老 OA 完成 10.18 万条 / 54.79 万人次。",
             font=FONT_BODY, size=11, color=GRAY_700)

    # 下：八局云盘
    add_rect(s, Inches(0.5), Inches(4.6), Inches(12.33), Inches(2.3),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(0.5), Inches(4.6), Inches(12.33), Inches(0.5),
             fill=PRIMARY_DARK)
    add_text(s, Inches(0.7), Inches(4.6), Inches(12.0), Inches(0.5),
             "▎八局云盘版本升级：完成版本升级与平滑迁移",
             font=FONT_TITLE, size=13, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    items = [
        ("8 万", "用户规模", PRIMARY),
        ("69 T", "迁移文件量", PRIMARY_DARK),
        ("✓", "分享统计增强", GREEN),
        ("✓", "安全防护加固", GREEN),
    ]
    for i, (v, l, c) in enumerate(items):
        x = Inches(0.8) + Inches(i * 3.0)
        add_text(s, x, Inches(5.3), Inches(2.8), Inches(0.85), v,
                 font=FONT_TITLE, size=30, bold=True, color=c)
        add_text(s, x, Inches(6.2), Inches(2.8), Inches(0.4), l,
                 font=FONT_BODY, size=12, color=GRAY_700)


# ---------- P15 章节3封面 ----------
def slide_part3_cover():
    s = prs.slides.add_slide(BLANK)
    add_section_title(
        s, 3, "存在的问题及不足",
        "对照高质量发展要求，正视短板与不足，"
        "聚焦运维知识沉淀、推进进度、数据应用三方面深刻剖析。",
    )


# ---------- P16 问题与不足 ----------
def slide_issues():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 16, TOTAL_PAGES, "03　存在的问题及不足",
                    title="03　存在的问题及不足",
                    subtitle="正视短板 · 聚焦改进")

    issues = [
        ("01", "知识库沉淀不足",
         "现有运维知识体系尚不健全，相似问题缺乏标准化处置预案，"
         "运维人员遇到同类问题难以快速分析定位，影响响应效率与服务质量。",
         "▎ 改进方向：健全知识管理与经验复用机制，建立标准化运维知识库与典型问题处置手册。"),
        ("02", "二级单位推进进度不均衡",
         "各二级单位流程搭建与上线进度存在差异，部分单位需求梳理及报价对接仍需提速，"
         "整体推进节奏需进一步统筹协调。",
         "▎ 改进方向：建立二级单位推进周报机制，分层分类落实进度管控与协调督办。"),
        ("03", "系统数据应用深度有待提升",
         "信创 OA、数智工会等平台已积累一定规模运营数据，但在数据分析、决策支持等方面挖掘不充分，"
         "数据价值释放有待进一步加强。",
         "▎ 改进方向：完善运营看板指标体系，推动数据驱动管理与智能化决策应用。"),
    ]
    top = Inches(2.0)
    row_h = Inches(1.55)
    for i, (no, title, desc, advice) in enumerate(issues):
        y = top + (row_h + Inches(0.13)) * i
        # 编号 + 标题块
        add_rect(s, Inches(0.5), y, Inches(3.5), row_h, fill=PRIMARY)
        add_text(s, Inches(0.65), y + Inches(0.1), Inches(0.9), Inches(0.6),
                 no, font=FONT_TITLE, size=28, bold=True, color=ACCENT)
        add_text(s, Inches(1.6), y + Inches(0.2), Inches(2.3), Inches(0.45),
                 "ISSUE", font=FONT_BODY, size=10, color=PRIMARY_LIGHT)
        add_text(s, Inches(1.6), y + Inches(0.55), Inches(2.3), Inches(0.85),
                 title, font=FONT_TITLE, size=15, bold=True, color=WHITE,
                 line_spacing=1.2)
        # 描述区
        add_rect(s, Inches(4.0), y, Inches(8.83), row_h, fill=WHITE,
                 line=GRAY_300)
        add_text(s, Inches(4.2), y + Inches(0.1), Inches(8.5), Inches(0.7),
                 desc, font=FONT_BODY, size=12, color=GRAY_700,
                 line_spacing=1.4)
        add_text(s, Inches(4.2), y + Inches(0.95), Inches(8.5), Inches(0.55),
                 advice, font=FONT_BODY, size=11, bold=True, color=PRIMARY_DARK,
                 line_spacing=1.3)


# ---------- P17 章节4封面 ----------
def slide_part4_cover():
    s = prs.slides.add_slide(BLANK)
    add_section_title(
        s, 4, "二季度工作计划",
        "坚持目标导向、问题导向、结果导向，"
        "聚焦四项重点任务，确保高质量完成全年目标。",
    )


# ---------- P18 二季度工作计划 ----------
def slide_q2_plan():
    s = prs.slides.add_slide(BLANK)
    add_page_chrome(s, 18, TOTAL_PAGES, "04　二季度工作计划",
                    title="04　二季度重点工作计划",
                    subtitle="时间节点 · 责任明确 · 量化目标")

    # 时间轴
    timeline_y = Inches(2.3)
    add_line(s, Inches(0.8), timeline_y, Inches(12.5), timeline_y,
             color=PRIMARY, weight=2.5)
    nodes = [
        ("5 月底", Inches(2.5)),
        ("6 月底", Inches(6.6)),
        ("6 月 30 日", Inches(10.7)),
        ("持续推进", Inches(12.4)),
    ]
    for txt, x in nodes:
        add_rect(s, x - Inches(0.1), timeline_y - Inches(0.1), Inches(0.2),
                 Inches(0.2), fill=PRIMARY, shape=MSO_SHAPE.OVAL)
        add_text(s, x - Inches(0.7), timeline_y - Inches(0.55), Inches(1.4),
                 Inches(0.4), txt, font=FONT_TITLE, size=11, bold=True,
                 color=PRIMARY_DARK, align=PP_ALIGN.CENTER)

    plans = [
        ("01", "信创 OA 二级落地", "6 月 30 日前", PRIMARY,
         "完成信创 OA 系统二级单位剩余功能及流程的上线工作，"
         "实现二级单位全面应用、全员覆盖。"),
        ("02", "移动端功能开发", "6 月底前", PRIMARY_DARK,
         "完成移动端剩余 7 个功能模块的开发与上线，"
         "持续完善移动办公体系，提升用户体验。"),
        ("03", "机构系统功能优化", "5 月底前", ACCENT,
         "完成机构系统审批预选等相关功能的优化工作，"
         "进一步提升系统易用性与运行效率。"),
        ("04", "运维保障持续提升", "持续推进", GREEN,
         "确保各信息系统可用性达 99.9% 以上，"
         "SLA 达标率持续保持在 95% 以上。"),
    ]
    card_w = Inches(2.95)
    card_h = Inches(3.7)
    gap = Inches(0.15)
    for i, (no, title, deadline, color, desc) in enumerate(plans):
        x = Inches(0.5) + (card_w + gap) * i
        y = Inches(3.0)
        # 卡片
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=GRAY_300)
        # 顶部色条
        add_rect(s, x, y, card_w, Inches(0.1), fill=color)
        # 编号
        add_text(s, x + Inches(0.25), y + Inches(0.25), Inches(2.5),
                 Inches(0.6), no, font=FONT_TITLE, size=30, bold=True,
                 color=color)
        # 时间标签
        add_rect(s, x + Inches(0.25), y + Inches(0.95), Inches(2.0),
                 Inches(0.4), fill=GRAY_100)
        add_text(s, x + Inches(0.25), y + Inches(0.95), Inches(2.0),
                 Inches(0.4), f"⏱  {deadline}", font=FONT_BODY, size=11,
                 bold=True, color=color, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 标题
        add_text(s, x + Inches(0.25), y + Inches(1.5), card_w - Inches(0.5),
                 Inches(0.6), title, font=FONT_TITLE, size=16, bold=True,
                 color=GRAY_900, line_spacing=1.2)
        # 装饰
        add_rect(s, x + Inches(0.25), y + Inches(2.05), Inches(0.3),
                 Inches(0.04), fill=color)
        # 描述
        add_text(s, x + Inches(0.25), y + Inches(2.2), card_w - Inches(0.5),
                 Inches(1.4), desc, font=FONT_BODY, size=11.5, color=GRAY_700,
                 line_spacing=1.5)

    # 结语
    add_rect(s, Inches(0.5), Inches(6.85), Inches(12.33), Inches(0.25),
             fill=PRIMARY_LIGHT)
    add_text(s, Inches(0.5), Inches(7.0), Inches(12.33), Inches(0.4),
             "强化目标引领　|　压实工作责任　|　补齐短板弱项　|　高质量完成全年各项目标任务",
             font=FONT_TITLE, size=12, bold=True, color=PRIMARY_DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------- P19 结尾 ----------
def slide_end():
    s = prs.slides.add_slide(BLANK)
    add_rect(s, Emu(0), Emu(0), SLIDE_W, SLIDE_H, fill=PRIMARY)
    # 装饰
    add_rect(s, Emu(0), Inches(7.42), SLIDE_W, Inches(0.08), fill=ACCENT)
    add_rect(s, Inches(0), Inches(0), Inches(0.15), SLIDE_H, fill=PRIMARY_DARK)

    add_text(s, Inches(0.5), Inches(2.5), Inches(12.33), Inches(0.6),
             "THANKS  FOR  YOUR  ATTENTION",
             font=FONT_BODY, size=18, color=RGBColor(0xB7, 0xD7, 0xF2),
             align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.5), Inches(3.2), Inches(12.33), Inches(1.5),
             "感 谢 聆 听\n敬 请 批 评 指 正",
             font=FONT_TITLE, size=44, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, line_spacing=1.4)
    add_rect(s, Inches(6.27), Inches(5.8), Inches(0.8), Inches(0.06),
             fill=ACCENT)
    add_text(s, Inches(0.5), Inches(6.0), Inches(12.33), Inches(0.5),
             "信息化运营管理部　|　二〇二六年四月",
             font=FONT_BODY, size=14, color=RGBColor(0xB7, 0xD7, 0xF2),
             align=PP_ALIGN.CENTER)


# ======== 执行 ========
slide_cover()
slide_toc()
slide_part1_cover()
slide_kpi_dashboard()
slide_kpi_analysis()
slide_part2_cover()
slide_progress_overview()
slide_oa_overview()
slide_oa_highlights()
slide_oa_data()
slide_union_launch()
slide_union_dev()
slide_ops_1()
slide_ops_2()
slide_part3_cover()
slide_issues()
slide_part4_cover()
slide_q2_plan()
slide_end()

output = "一季度运营分析会汇报.pptx"
prs.save(output)
print(f"✓ 已生成 {output}，共 {len(prs.slides)} 页幻灯片。")
