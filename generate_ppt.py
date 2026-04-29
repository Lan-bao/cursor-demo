"""基于《附件2：2026年一季度运营分析报告PPT模板.pptx》生成一季度运营分析会汇报 PPT。

模板信息：
  - 主题色：中建蓝 #0070C0
  - 字体：微软雅黑
  - 版式：封面 / 目录页 / 节标题 / 2_~8_正文页 / 结束语
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

TEMPLATE = "附件2：2026年一季度运营分析报告PPT模板.pptx"
OUTPUT = "一季度运营分析会汇报.pptx"

# ======== 主题配置（取自模板） ========
PRIMARY = RGBColor(0x00, 0x70, 0xC0)
PRIMARY_2 = RGBColor(0x00, 0x80, 0xCB)
PRIMARY_DARK = RGBColor(0x00, 0x4F, 0x8B)
PRIMARY_LIGHT = RGBColor(0xD9, 0xE9, 0xF6)
PRIMARY_BG = RGBColor(0xEE, 0xF5, 0xFC)
ACCENT = RGBColor(0xF5, 0xA6, 0x23)
GRAY_900 = RGBColor(0x22, 0x2B, 0x38)
GRAY_700 = RGBColor(0x4A, 0x55, 0x68)
GRAY_500 = RGBColor(0x8C, 0x95, 0xA6)
GRAY_300 = RGBColor(0xD7, 0xDC, 0xE5)
GRAY_100 = RGBColor(0xF2, 0xF5, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC0, 0x39, 0x2B)
GREEN = RGBColor(0x2E, 0x8B, 0x57)

FONT = "微软雅黑"

prs = Presentation(TEMPLATE)
SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height
print(f"模板加载完成：{SLIDE_W/914400:.2f} x {SLIDE_H/914400:.2f} inches")


# ======== 工具函数 ========

def _set_run_font(run, font_name=FONT):
    """确保中文字体生效。"""
    rPr = run._r.get_or_add_rPr()
    for tag in ("ea", "cs", "latin"):
        existing = rPr.find(qn(f"a:{tag}"))
        if existing is not None:
            rPr.remove(existing)
        ele = etree.SubElement(rPr, qn(f"a:{tag}"))
        ele.set("typeface", font_name)


def set_run(run, text, *, font=FONT, size=14, bold=False, color=GRAY_900):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    _set_run_font(run, font)


def add_text(slide, left, top, width, height, text, *, font=FONT, size=14,
             bold=False, color=GRAY_900, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, line_spacing=1.3):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    set_run(p.add_run(), text, font=font, size=size, bold=bold, color=color)
    return box


def add_lines(slide, left, top, width, height, lines, *, font=FONT, size=12,
              color=GRAY_700, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
              line_spacing=1.4):
    """lines: list of (text, bold, size_override, color_override) tuples or str."""
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
            s = item[2] if len(item) > 2 and item[2] is not None else size
            c = item[3] if len(item) > 3 and item[3] is not None else color
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        set_run(p.add_run(), text, font=font, size=s, bold=b, color=c)
    return box


def add_shape(slide, shape_type, left, top, width, height, *, fill=PRIMARY,
              line=None):
    shp = slide.shapes.add_shape(shape_type, left, top, width, height)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    return shp


def add_rect(slide, left, top, width, height, *, fill=PRIMARY, line=None):
    return add_shape(slide, MSO_SHAPE.RECTANGLE, left, top, width, height,
                     fill=fill, line=line)


def add_line_conn(slide, x1, y1, x2, y2, *, color=GRAY_300, weight=1.0):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = color
    line.line.width = Pt(weight)
    return line


# ======== 幻灯片清理 ========
# 模板内含 11 张样例页。先删除所有，再按版式重新构建。
def _remove_all_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    rels = prs.part.rels
    ids = list(sldIdLst)
    for sld_id in ids:
        rId = sld_id.get(qn("r:id"))
        slide_part = rels[rId].target_part
        sldIdLst.remove(sld_id)
        prs.part.drop_rel(rId)


_remove_all_slides(prs)
print(f"已清空模板幻灯片，剩余: {len(prs.slides)}")


# ======== 获取 layout（按名称） ========
LAYOUTS = {l.name: l for l in prs.slide_masters[0].slide_layouts}
print(f"主母版可用版式：{list(LAYOUTS.keys())}")


def add_slide(layout_name):
    """从主母版按名称获取 layout 并新增幻灯片。"""
    layout = LAYOUTS[layout_name]
    return prs.slides.add_slide(layout)


_PLACEHOLDER_ALIASES = {
    "标题": ["标题", "Title", "PA_标题"],
    "副标题": ["副标题", "Subtitle"],
    "文本占位符": ["文本占位符", "Text Placeholder", "Content Placeholder"],
}


def find_placeholder_by_name(slide, name_keyword):
    keywords = _PLACEHOLDER_ALIASES.get(name_keyword, [name_keyword])
    for shp in slide.placeholders:
        for kw in keywords:
            if kw in shp.name:
                return shp
    return None


def set_placeholder_text(slide, name_keyword, text, *, size=None, bold=None,
                         color=None, align=None):
    shp = find_placeholder_by_name(slide, name_keyword)
    if shp is None:
        return None
    tf = shp.text_frame
    p = tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    if not p.runs:
        run = p.add_run()
    else:
        run = p.runs[0]
        for extra in p.runs[1:]:
            extra._r.getparent().remove(extra._r)
    run.text = text
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    run.font.name = FONT
    _set_run_font(run, FONT)
    return shp


# ============ P01 封面 ============
def slide_cover():
    s = add_slide("封面")
    # 隐藏占位符
    set_placeholder_text(s, "标题", " ", size=1, color=WHITE)
    set_placeholder_text(s, "副标题", " ", size=1, color=WHITE)
    # 主标题（部门 / 主标题）—— 与模板原版"行政办公项目组"位置对齐
    add_text(s, Inches(2.02), Inches(1.40), Inches(9.90), Inches(0.91),
             "一季度运营分析会汇报",
             font=FONT, size=48, bold=True, color=PRIMARY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 中部圆角条：年份 / 主题
    add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(2.98), Inches(2.44), Inches(7.61), Inches(0.66),
              fill=PRIMARY_2)
    add_text(s, Inches(2.98), Inches(2.44), Inches(7.61), Inches(0.66),
             "2026 年一季度运营分析会  ·  信息化系统建设与运营专题",
             font=FONT, size=20, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 汇报部门（对应模板 y=3.63）
    add_text(s, Inches(4.68), Inches(3.63), Inches(3.97), Inches(0.57),
             "汇报部门：信息化运营管理部",
             font=FONT, size=20, bold=True, color=GRAY_900,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 汇报时间（对应模板 y=4.40）
    add_text(s, Inches(4.68), Inches(4.40), Inches(3.97), Inches(0.57),
             "二〇二六年四月",
             font=FONT, size=20, bold=True, color=GRAY_900,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============ P02 目录页 ============
def slide_toc():
    s = add_slide("目录页")
    # 模板"添加目录"占位符
    sub = find_placeholder_by_name(s, "副标题")
    if sub is not None:
        # 隐藏占位符内容
        sub.text_frame.text = ""
    # 4 个目录项（参照模板布局，左右两列）
    items = [
        ("第一部分", "一季度总体指标完成情况"),
        ("第二部分", "一季度工作进展"),
        ("第三部分", "存在的问题及不足"),
        ("第四部分", "二季度工作计划"),
    ]
    base_x_no = Inches(4.28)
    base_x_title = Inches(6.08)
    base_y = Inches(2.03)
    row_h = Inches(0.70)
    gap = Inches(1.20)
    for i, (no, title) in enumerate(items):
        y = base_y + gap * i
        # 编号块（圆角矩形）
        add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, base_x_no, y,
                  Inches(2.05), Inches(0.48), fill=PRIMARY)
        add_text(s, base_x_no, y, Inches(2.05), Inches(0.48),
                 no, font=FONT, size=20, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 标题块（淡蓝底）
        add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, base_x_title, y,
                  Inches(5.95), Inches(0.48), fill=PRIMARY_BG,
                  line=PRIMARY_LIGHT)
        add_text(s, base_x_title + Inches(0.25), y,
                 Inches(5.7), Inches(0.48), title,
                 font=FONT, size=20, bold=True, color=PRIMARY_DARK,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)


# ============ 节标题页 ============
_CN_NUMS = {1: "一", 2: "二", 3: "三", 4: "四"}


def slide_section(part_no, title):
    s = add_slide("节标题")
    # 模板原版样式：第X部分 在上，章节标题在下
    title_ph = set_placeholder_text(
        s, "标题", title, size=44, bold=True, color=PRIMARY,
        align=PP_ALIGN.CENTER)
    if title_ph is not None:
        title_ph.left = Inches(2.43)
        title_ph.top = Inches(3.31)
        title_ph.width = Inches(8.50)
        title_ph.height = Inches(0.96)

    text_ph = set_placeholder_text(
        s, "文本占位符",
        f"第{_CN_NUMS.get(part_no, part_no)}部分",
        size=36, bold=True, color=PRIMARY, align=PP_ALIGN.CENTER)
    if text_ph is not None:
        text_ph.left = Inches(2.43)
        text_ph.top = Inches(2.26)
        text_ph.width = Inches(8.50)
        text_ph.height = Inches(1.05)
    return s


# ============ 正文页公共：标题条 ============
def add_subtitle(s, title, *, color=PRIMARY):
    """正文页正文区上方的小标题。"""
    add_rect(s, Inches(0.45), Inches(1.05), Inches(0.10), Inches(0.42),
             fill=color)
    add_text(s, Inches(0.65), Inches(1.05), Inches(12.5), Inches(0.42),
             title, font=FONT, size=20, bold=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE)


def add_subdesc(s, desc):
    add_text(s, Inches(0.45), Inches(1.50), Inches(12.5), Inches(0.32),
             desc, font=FONT, size=12, color=GRAY_500)


# ============ P04 一季度核心指标一览 ============
def slide_kpi_dashboard():
    s = add_slide("2_正文页")
    add_subtitle(s, "一、一季度核心指标完成情况")
    add_subdesc(s, "经营指标 · 现金质量 · 风险信号")

    # 6 张 KPI 卡片
    cards = [
        ("合 同 额", "33.13", "万元", PRIMARY, "Q1 累计签约金额"),
        ("营业收入", "6.25", "万元", PRIMARY, "Q1 累计实现营收"),
        ("利 润 额", "243.67", "万元", GREEN, "经营效益保持稳健"),
        ("回 款 额", "313.8", "万元", PRIMARY_DARK, "现金回款 313.8 万元"),
        ("负流项目占比", "17", "%", ACCENT, "项目纠偏需重点关注"),
        ("总负流金额", "-85.3", "万元", RED, "现金流压力指标"),
    ]
    cols = 3
    card_w = Inches(4.10)
    card_h = Inches(1.85)
    gap_x = Inches(0.10)
    gap_y = Inches(0.20)
    start_x = Inches(0.45)
    start_y = Inches(2.05)
    for i, (label, val, unit, color, note) in enumerate(cards):
        r, c = divmod(i, cols)
        x = start_x + (card_w + gap_x) * c
        y = start_y + (card_h + gap_y) * r
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=GRAY_300)
        add_rect(s, x, y, card_w, Inches(0.10), fill=color)
        add_text(s, x + Inches(0.25), y + Inches(0.20),
                 card_w - Inches(0.5), Inches(0.35), label,
                 font=FONT, size=12, color=GRAY_500)
        add_text(s, x + Inches(0.25), y + Inches(0.55),
                 card_w - Inches(2.0), Inches(0.7), val,
                 font=FONT, size=34, bold=True, color=color)
        add_text(s, x + Inches(2.7), y + Inches(0.85),
                 Inches(1.2), Inches(0.4), unit, font=FONT, size=12,
                 color=GRAY_700)
        add_text(s, x + Inches(0.25), y + Inches(1.40),
                 card_w - Inches(0.5), Inches(0.4), note,
                 font=FONT, size=10, color=GRAY_500)

    # 底部研判
    bottom_y = Inches(6.30)
    add_rect(s, Inches(0.45), bottom_y, Inches(12.40), Inches(0.65),
             fill=PRIMARY_LIGHT)
    add_rect(s, Inches(0.45), bottom_y, Inches(0.10), Inches(0.65),
             fill=PRIMARY)
    add_text(s, Inches(0.70), bottom_y, Inches(12.0), Inches(0.65),
             "总体研判：回款全部为现金，资金质量较高；合同与营业收入受年初节奏影响规模偏小；负流占比 17%、总负流 -85.3 万元，需重点强化项目纠偏与现金流管控。",
             font=FONT, size=12, bold=True, color=PRIMARY_DARK,
             anchor=MSO_ANCHOR.MIDDLE)


# ============ P05 指标质量分析 ============
def slide_kpi_analysis():
    s = add_slide("2_正文页")
    add_subtitle(s, "二、一季度核心指标质量分析")
    add_subdesc(s, "资金质量 · 经营效益 · 风险信号")

    # 左卡：回款结构
    add_rect(s, Inches(0.45), Inches(1.95), Inches(6.10), Inches(4.85),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(0.45), Inches(1.95), Inches(6.10), Inches(0.50),
             fill=PRIMARY)
    add_text(s, Inches(0.65), Inches(1.95), Inches(5.90), Inches(0.50),
             "▎一、回款结构分析", font=FONT, size=15, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.75), Inches(2.65), Inches(5.70), Inches(0.4),
             "现金回款占比", font=FONT, size=12, color=GRAY_500)
    add_text(s, Inches(0.75), Inches(3.00), Inches(5.70), Inches(0.7),
             "100%", font=FONT, size=42, bold=True, color=GREEN)
    # 进度条
    add_rect(s, Inches(0.75), Inches(3.85), Inches(5.50), Inches(0.30),
             fill=GRAY_100)
    add_rect(s, Inches(0.75), Inches(3.85), Inches(5.50), Inches(0.30),
             fill=GREEN)
    add_text(s, Inches(0.75), Inches(3.85), Inches(5.50), Inches(0.30),
             "313.8 / 313.8 万元", font=FONT, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, Inches(0.75), Inches(4.30), Inches(5.50), Inches(2.4), [
        ("· 现金部分：313.8 万元", True, 13, GRAY_900),
        ("· 非现金部分：0 万元", True, 13, GRAY_900),
        ("", False, 6, GRAY_900),
        ("研判：一季度回款全部为现金，资金质量较高；后续应继续坚持现金优先原则，兼顾回款规模与现金占比的协同提升。",
         False, 12, GRAY_700),
    ])

    # 右卡：负流风险
    add_rect(s, Inches(6.78), Inches(1.95), Inches(6.10), Inches(4.85),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(6.78), Inches(1.95), Inches(6.10), Inches(0.50),
             fill=RED)
    add_text(s, Inches(6.98), Inches(1.95), Inches(5.90), Inches(0.50),
             "▎二、负流风险研判", font=FONT, size=15, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    # 数据
    add_text(s, Inches(7.08), Inches(2.65), Inches(2.7), Inches(0.4),
             "项目个数占比", font=FONT, size=12, color=GRAY_500)
    add_text(s, Inches(7.08), Inches(3.00), Inches(2.7), Inches(0.7),
             "17%", font=FONT, size=36, bold=True, color=ACCENT)
    add_text(s, Inches(9.88), Inches(2.65), Inches(2.9), Inches(0.4),
             "总负流金额", font=FONT, size=12, color=GRAY_500)
    add_text(s, Inches(9.88), Inches(3.00), Inches(2.9), Inches(0.7),
             "-85.3 万元", font=FONT, size=30, bold=True, color=RED)

    add_line_conn(s, Inches(7.08), Inches(4.00),
                  Inches(12.65), Inches(4.00), color=GRAY_300)
    add_lines(s, Inches(7.08), Inches(4.10), Inches(5.55), Inches(2.6), [
        ("应对举措：", True, 13, GRAY_900),
        ("· 项目层面：开展负流项目专项盘点，逐项制定纠偏方案；",
         False, 12, GRAY_700),
        ("· 过程管控：强化产值、回款、成本三同步管理；",
         False, 12, GRAY_700),
        ("· 月度复盘：建立月度负流通报与预警机制，压实责任。",
         False, 12, GRAY_700),
    ])


# ============ P0X 工作进展总览 ============
def slide_progress_overview():
    s = add_slide("7_正文页")
    add_subtitle(s, "一、一季度工作进展总览")
    add_subdesc(s, "两条主线 · 多维突破")

    # 中部一句话总结条
    add_rect(s, Inches(0.45), Inches(2.00), Inches(12.45), Inches(0.65),
             fill=PRIMARY_LIGHT)
    add_text(s, Inches(0.45), Inches(2.00), Inches(12.45), Inches(0.65),
             "以业务需求为导向　|　以系统建设为牵引　|　以稳定运维为保障",
             font=FONT, size=15, bold=True, color=PRIMARY_DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    blocks = [
        ("（一）系统建设情况", PRIMARY, [
            "信创 OA 系统：核心功能切换上线、26 家二级单位大家签适配",
            "信创 OA 系统：城市运营公司 104 条流程搭建上线，流程优化迭代",
            "数智工会系统：5 大模块 2 月 5 日全局上线，覆盖 2710 名干部",
            "数智工会系统：完成干部端 9 个功能模块开发",
        ]),
        ("（二）系统运维情况", PRIMARY_DARK, [
            "组织机构：保障中建科技合并、运营投资类公司整合落地",
            "八局通：处理工单 475 条，应用上线 8 个，缺陷修复 10/13",
            "信创 OA 上云：汉数科 5 类应用搭建上线，4 月 1 日切换审批",
            "八局云盘：版本升级，保障 8 万用户、69T 文件平稳迁移",
        ]),
    ]
    top = Inches(2.95)
    for i, (title, color, items) in enumerate(blocks):
        x = Inches(0.45) + Inches(i * 6.30)
        # 标题
        add_rect(s, x, top, Inches(6.10), Inches(0.55), fill=color)
        add_text(s, x + Inches(0.30), top, Inches(5.80), Inches(0.55),
                 title, font=FONT, size=16, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 内容
        add_rect(s, x, top + Inches(0.55), Inches(6.10), Inches(3.5),
                 fill=WHITE, line=GRAY_300)
        for j, line in enumerate(items):
            y = top + Inches(0.85 + j * 0.75)
            add_shape(s, MSO_SHAPE.OVAL, x + Inches(0.30), y + Inches(0.16),
                      Inches(0.13), Inches(0.13), fill=color)
            add_text(s, x + Inches(0.55), y, Inches(5.45), Inches(0.50),
                     line, font=FONT, size=13, color=GRAY_700,
                     anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.3)


# ============ 信创 OA - 整体推进 ============
def slide_oa_overview():
    s = add_slide("7_正文页")
    add_subtitle(s, "二、信创 OA 系统｜整体推进情况")
    add_subdesc(s, "切换上线 · 运维保障 · 二级承接 · 优化迭代")

    items = [
        ("①", "核心功能\n切换上线",
         "局总部公文、新闻、公告、流程中心等核心功能全部切换上线，"
         "二级单位除基础功能外亦实现全面上线运行。"),
        ("②", "重点流程\n专项保障",
         "针对发起量较大的局总部公文、新闻、财务等关键流程开展专项保障，"
         "并对二级单位流程搭建提供答疑支持。"),
        ("③", "二级单位\n搭建承接",
         "城市运营公司完成 103 条流程搭建；南方公司完成 130+ 条需求梳理；"
         "浙江公司流程需求处于对接报价阶段。"),
        ("④", "局总部流程\n优化迭代",
         "局总部流程持续优化迭代，重点针对用印等高频流程进行优化完善"
         "（由致远提供相关支持）。"),
    ]
    top = Inches(2.00)
    card_w = Inches(3.05)
    card_h = Inches(4.85)
    gap = Inches(0.10)
    start_x = Inches(0.45)
    for i, (no, title, desc) in enumerate(items):
        x = start_x + (card_w + gap) * i
        add_rect(s, x, top, card_w, card_h, fill=WHITE, line=GRAY_300)
        add_rect(s, x, top, card_w, Inches(1.55), fill=PRIMARY)
        add_text(s, x, top + Inches(0.10), card_w, Inches(0.6), no,
                 font=FONT, size=30, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)
        add_text(s, x, top + Inches(0.78), card_w, Inches(0.7), title,
                 font=FONT, size=15, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, line_spacing=1.2)
        add_text(s, x + Inches(0.25), top + Inches(1.75),
                 card_w - Inches(0.5), card_h - Inches(2.0), desc,
                 font=FONT, size=13, color=GRAY_700, line_spacing=1.5)


# ============ 信创 OA - 重点工作成效 ============
def slide_oa_highlights():
    s = add_slide("7_正文页")
    add_subtitle(s, "三、信创 OA 系统｜重点工作成效")
    add_subdesc(s, "功能适配 · 流程建设 · 数据对接 · 运营管控")

    blocks = [
        ("一", "聚焦功能适配与开发\n夯实系统应用基础",
         "完成 26 家二级单位「大家签」适配开发；外部主数据同步功能已上线，"
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
    top = Inches(2.00)
    row_h = Inches(1.18)
    for i, (no, title, desc, color) in enumerate(blocks):
        y = top + (row_h + Inches(0.07)) * i
        add_rect(s, Inches(0.45), y, Inches(1.0), row_h, fill=color)
        add_text(s, Inches(0.45), y, Inches(1.0), row_h, no,
                 font=FONT, size=36, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, Inches(1.45), y, Inches(3.6), row_h, fill=PRIMARY_LIGHT)
        add_text(s, Inches(1.65), y, Inches(3.4), row_h, title,
                 font=FONT, size=13, bold=True, color=PRIMARY_DARK,
                 anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.3)
        add_rect(s, Inches(5.05), y, Inches(7.85), row_h,
                 fill=WHITE, line=GRAY_300)
        add_text(s, Inches(5.25), y, Inches(7.6), row_h, desc,
                 font=FONT, size=13, color=GRAY_700,
                 anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.45)


# ============ 信创 OA - 运营数据 ============
def slide_oa_data():
    s = add_slide("7_正文页")
    add_subtitle(s, "四、信创 OA 系统｜运营数据")
    add_subdesc(s, "量化呈现 · 数据说话")

    cards = [
        ("32,249", "发起流程", "件", PRIMARY),
        ("22,249", "办结流程", "件", PRIMARY_DARK),
        ("10,257", "完成流程", "件", GREEN),
    ]
    for i, (val, label, unit, color) in enumerate(cards):
        x = Inches(0.45 + i * 4.30)
        add_rect(s, x, Inches(2.05), Inches(4.05), Inches(2.0),
                 fill=WHITE, line=GRAY_300)
        add_rect(s, x, Inches(2.05), Inches(0.10), Inches(2.0), fill=color)
        add_text(s, x + Inches(0.30), Inches(2.20), Inches(3.6), Inches(0.4),
                 label, font=FONT, size=13, color=GRAY_500)
        add_text(s, x + Inches(0.30), Inches(2.60), Inches(3.6), Inches(1.0),
                 val, font=FONT, size=44, bold=True, color=color)
        add_text(s, x + Inches(0.30), Inches(3.60), Inches(3.6), Inches(0.4),
                 unit, font=FONT, size=13, color=GRAY_700)

    add_text(s, Inches(0.45), Inches(4.40), Inches(12.45), Inches(0.4),
             "▎新老 OA 系统流程审批数据对比", font=FONT, size=15,
             bold=True, color=PRIMARY_DARK)
    add_line_conn(s, Inches(0.45), Inches(4.85),
                  Inches(12.90), Inches(4.85),
                  color=PRIMARY, weight=1.5)

    table_data = [
        ["系统", "流程审批量", "审批人次", "说明"],
        ["新信创 OA 系统（全局）", "3.557 万条", "35.289 万人次",
         "二级单位逐步切换，4 月 1 日完成审批切换"],
        ["原 OA 系统", "10.18 万条", "54.79 万人次",
         "并行运行期保障业务连续性"],
    ]
    col_w = [Inches(3.30), Inches(2.50), Inches(2.50), Inches(4.15)]
    row_h = Inches(0.55)
    y = Inches(5.00)
    for r, row in enumerate(table_data):
        x = Inches(0.45)
        for c, txt in enumerate(row):
            if r == 0:
                add_rect(s, x, y, col_w[c], row_h, fill=PRIMARY)
                color = WHITE; bold = True; size = 13
            else:
                fill = WHITE if r % 2 == 1 else GRAY_100
                add_rect(s, x, y, col_w[c], row_h, fill=fill, line=GRAY_300)
                color = GRAY_900 if c < 3 else GRAY_700
                bold = c < 3; size = 12
            add_text(s, x + Inches(0.15), y, col_w[c] - Inches(0.3), row_h,
                     txt, font=FONT, size=size, bold=bold, color=color,
                     anchor=MSO_ANCHOR.MIDDLE,
                     align=PP_ALIGN.LEFT if c == 0 or c == 3
                     else PP_ALIGN.CENTER)
            x += col_w[c]
        y += row_h


# ============ 数智工会 - 推广 ============
def slide_union_launch():
    s = add_slide("7_正文页")
    add_subtitle(s, "五、数智工会系统｜干部端全局推广上线")
    add_subdesc(s, "2 月 5 日 · 全局上线 · 全面覆盖")

    add_rect(s, Inches(0.45), Inches(2.00), Inches(12.45), Inches(0.65),
             fill=PRIMARY)
    add_text(s, Inches(0.45), Inches(2.00), Inches(12.45), Inches(0.65),
             "■  完成荣誉奖项、劳动竞赛、技能竞赛、服务阵地、困难职工 5 大模块开发测试，2026 年 2 月 5 日全局上线",
             font=FONT, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    metrics = [
        ("2,223", "工会组织（个）"),
        ("2,710", "覆盖工会干部（名）"),
        ("1,321", "荣誉奖项 · 集体（条）"),
        ("955", "荣誉奖项 · 个人（条）"),
        ("2,212", "劳模先进（条）"),
        ("157", "劳动竞赛局级立项（条）"),
    ]
    cols = 3
    card_w = Inches(4.05)
    card_h = Inches(1.35)
    gap_x = Inches(0.10)
    gap_y = Inches(0.18)
    start_x = Inches(0.45)
    start_y = Inches(2.95)
    for i, (val, lbl) in enumerate(metrics):
        r, c = divmod(i, cols)
        x = start_x + (card_w + gap_x) * c
        y = start_y + (card_h + gap_y) * r
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=GRAY_300)
        add_rect(s, x, y, Inches(0.10), card_h, fill=PRIMARY)
        add_text(s, x + Inches(0.30), y + Inches(0.10),
                 card_w - Inches(0.4), Inches(0.7), val,
                 font=FONT, size=28, bold=True, color=PRIMARY_DARK)
        add_text(s, x + Inches(0.30), y + Inches(0.85),
                 card_w - Inches(0.4), Inches(0.4), lbl,
                 font=FONT, size=12, color=GRAY_700)

    add_rect(s, Inches(0.45), Inches(6.30), Inches(12.45), Inches(0.55),
             fill=PRIMARY_LIGHT)
    add_text(s, Inches(0.65), Inches(6.30), Inches(12.0), Inches(0.55),
             "创新工作室录入 364 个，服务阵地覆盖 511 个，应用基础数据初具规模。",
             font=FONT, size=13, bold=True, color=PRIMARY_DARK,
             anchor=MSO_ANCHOR.MIDDLE)


# ============ 数智工会 - 模块开发 ============
def slide_union_dev():
    s = add_slide("7_正文页")
    add_subtitle(s, "六、数智工会系统｜干部端运维与开发")
    add_subdesc(s, "9 大模块 · 完善服务体系")

    add_text(s, Inches(0.45), Inches(2.00), Inches(12.4), Inches(0.5),
             "围绕基层工会干部实际工作需求，完成 9 个核心模块开发，进一步丰富数智工会系统功能体系，提升用户体验。",
             font=FONT, size=13, color=GRAY_700)

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
    card_w = Inches(4.05)
    card_h = Inches(1.20)
    gap_x = Inches(0.10)
    gap_y = Inches(0.18)
    start_x = Inches(0.45)
    start_y = Inches(2.85)
    for i, (name, tag) in enumerate(modules):
        r, c = divmod(i, cols)
        x = start_x + (card_w + gap_x) * c
        y = start_y + (card_h + gap_y) * r
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=GRAY_300)
        add_shape(s, MSO_SHAPE.OVAL, x + Inches(0.20), y + Inches(0.30),
                  Inches(0.60), Inches(0.60), fill=PRIMARY)
        add_text(s, x + Inches(0.20), y + Inches(0.30),
                 Inches(0.60), Inches(0.60), f"{i+1:02d}",
                 font=FONT, size=12, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.95), y + Inches(0.20),
                 Inches(2.95), Inches(0.45), name,
                 font=FONT, size=15, bold=True, color=GRAY_900)
        add_text(s, x + Inches(0.95), y + Inches(0.68),
                 Inches(2.95), Inches(0.4), f"# {tag}",
                 font=FONT, size=11, color=PRIMARY)


# ============ 系统运维 - 组织 + 八局通 ============
def slide_ops_1():
    s = add_slide("7_正文页")
    add_subtitle(s, "七、系统运维｜组织机构 · 八局通")
    add_subdesc(s, "基础维护 · 工单闭环 · 缺陷管理")

    # 左：组织机构
    add_rect(s, Inches(0.45), Inches(2.00), Inches(6.10), Inches(4.85),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(0.45), Inches(2.00), Inches(6.10), Inches(0.50),
             fill=PRIMARY)
    add_text(s, Inches(0.65), Inches(2.00), Inches(5.90), Inches(0.50),
             "▎组织机构维护", font=FONT, size=15, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.65), Inches(2.65), Inches(5.70), Inches(0.4),
             "在册机构单位", font=FONT, size=12, color=GRAY_500)
    add_text(s, Inches(0.65), Inches(3.00), Inches(5.70), Inches(0.85),
             "913　家", font=FONT, size=42, bold=True, color=PRIMARY)
    add_line_conn(s, Inches(0.65), Inches(4.05),
                  Inches(6.40), Inches(4.05), color=GRAY_300)
    add_lines(s, Inches(0.65), Inches(4.20), Inches(5.70), Inches(2.6), [
        ("一季度重点保障：", True, 13, GRAY_900),
        ("· 完成三级单位组织调整事项的平稳落地；", False, 12, GRAY_700),
        ("· 保障中建科技公司合并；", False, 12, GRAY_700),
        ("· 保障运营投资类公司整合；", False, 12, GRAY_700),
        ("· 确保组织机构数据与实际管理架构同步。", False, 12, GRAY_700),
    ])

    # 右：八局通
    add_rect(s, Inches(6.78), Inches(2.00), Inches(6.10), Inches(4.85),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(6.78), Inches(2.00), Inches(6.10), Inches(0.50),
             fill=PRIMARY_DARK)
    add_text(s, Inches(6.98), Inches(2.00), Inches(5.90), Inches(0.50),
             "▎八局通运维支持", font=FONT, size=15, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    items = [
        ("475", "工单处理（条）", "全部办结闭环"),
        ("8", "应用上线（个）", "助力业务上云"),
        ("0", "系统故障（起）", "运行稳定可控"),
        ("10/13", "缺陷修复进度", "剩 3 个待新版本验证"),
    ]
    for i, (val, lbl, note) in enumerate(items):
        r, c = divmod(i, 2)
        x = Inches(6.88) + Inches(c * 2.95)
        y = Inches(2.65) + Inches(r * 2.05)
        add_rect(s, x, y, Inches(2.85), Inches(1.95), fill=GRAY_100)
        add_text(s, x + Inches(0.20), y + Inches(0.15),
                 Inches(2.65), Inches(0.30), lbl, font=FONT, size=11,
                 color=GRAY_500)
        add_text(s, x + Inches(0.20), y + Inches(0.45),
                 Inches(2.65), Inches(0.85), val, font=FONT, size=26,
                 bold=True, color=PRIMARY_DARK)
        add_text(s, x + Inches(0.20), y + Inches(1.40),
                 Inches(2.65), Inches(0.4), note, font=FONT, size=11,
                 color=GRAY_700)


# ============ 系统运维 - 上云 + 云盘 ============
def slide_ops_2():
    s = add_slide("7_正文页")
    add_subtitle(s, "八、系统运维｜信创 OA 上云 · 八局云盘")
    add_subdesc(s, "云端切换 · 平稳运行 · 升级保障")

    # 信创OA上云汉数科
    add_rect(s, Inches(0.45), Inches(2.00), Inches(12.45), Inches(2.40),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(0.45), Inches(2.00), Inches(12.45), Inches(0.50),
             fill=PRIMARY)
    add_text(s, Inches(0.65), Inches(2.00), Inches(12.0), Inches(0.50),
             "▎信创 OA 上云汉数科：会议预约 / 发文 / 新闻 / 通知公告 / 审批流程  ——  全部完成搭建上线",
             font=FONT, size=13, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    sub_metrics = [
        ("966", "会议申请（次）"),
        ("11", "文件发布（条）"),
        ("11", "通知公告（条）"),
        ("1", "新闻发布（条）"),
        ("4 月 1 日", "审批流程切换"),
    ]
    for i, (v, l) in enumerate(sub_metrics):
        x = Inches(0.65) + Inches(i * 2.42)
        add_text(s, x, Inches(2.65), Inches(2.30), Inches(0.7), v,
                 font=FONT, size=24, bold=True, color=PRIMARY)
        add_text(s, x, Inches(3.40), Inches(2.30), Inches(0.4), l,
                 font=FONT, size=12, color=GRAY_500)
        if i < 4:
            add_line_conn(s, x + Inches(2.30), Inches(2.70),
                          x + Inches(2.30), Inches(3.90), color=GRAY_300)
    add_text(s, Inches(0.65), Inches(3.95), Inches(12.0), Inches(0.4),
             "■ 全局新 OA 累计完成审批流程 3.557 万条 / 35.289 万人次；老 OA 完成 10.18 万条 / 54.79 万人次。",
             font=FONT, size=11, color=GRAY_700)

    # 八局云盘
    add_rect(s, Inches(0.45), Inches(4.55), Inches(12.45), Inches(2.30),
             fill=WHITE, line=GRAY_300)
    add_rect(s, Inches(0.45), Inches(4.55), Inches(12.45), Inches(0.50),
             fill=PRIMARY_DARK)
    add_text(s, Inches(0.65), Inches(4.55), Inches(12.0), Inches(0.50),
             "▎八局云盘版本升级：完成版本升级与平滑迁移",
             font=FONT, size=13, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    items = [
        ("8 万", "用户规模", PRIMARY),
        ("69 T", "迁移文件量", PRIMARY_DARK),
        ("✓", "分享统计增强", GREEN),
        ("✓", "安全防护加固", GREEN),
    ]
    for i, (v, l, c) in enumerate(items):
        x = Inches(0.80) + Inches(i * 3.0)
        add_text(s, x, Inches(5.25), Inches(2.85), Inches(0.85), v,
                 font=FONT, size=30, bold=True, color=c)
        add_text(s, x, Inches(6.20), Inches(2.85), Inches(0.4), l,
                 font=FONT, size=13, color=GRAY_700)


# ============ 存在的问题及不足 ============
def slide_issues():
    s = add_slide("6_正文页")
    add_subtitle(s, "存在的问题及不足")
    add_subdesc(s, "正视短板 · 聚焦改进")

    issues = [
        ("01", "知识库沉淀不足",
         "现有运维知识体系尚不健全，相似问题缺乏标准化处置预案，"
         "运维人员遇到同类问题难以快速分析定位，影响响应效率与服务质量。",
         "改进方向：健全知识管理与经验复用机制，建立标准化运维知识库与典型问题处置手册。"),
        ("02", "二级单位推进进度不均衡",
         "各二级单位流程搭建与上线进度存在差异，部分单位需求梳理及报价对接仍需提速，"
         "整体推进节奏需进一步统筹协调。",
         "改进方向：建立二级单位推进周报机制，分层分类落实进度管控与协调督办。"),
        ("03", "数据应用深度有待提升",
         "信创 OA、数智工会等平台已积累一定规模运营数据，但在数据分析、决策支持等方面挖掘不充分，"
         "数据价值释放有待进一步加强。",
         "改进方向：完善运营看板指标体系，推动数据驱动管理与智能化决策应用。"),
    ]
    top = Inches(2.00)
    row_h = Inches(1.55)
    for i, (no, title, desc, advice) in enumerate(issues):
        y = top + (row_h + Inches(0.13)) * i
        add_rect(s, Inches(0.45), y, Inches(3.55), row_h, fill=PRIMARY)
        add_text(s, Inches(0.60), y + Inches(0.10),
                 Inches(0.95), Inches(0.6), no,
                 font=FONT, size=28, bold=True, color=ACCENT)
        add_text(s, Inches(1.55), y + Inches(0.20),
                 Inches(2.40), Inches(0.45), "ISSUE",
                 font=FONT, size=10, color=PRIMARY_LIGHT)
        add_text(s, Inches(1.55), y + Inches(0.55),
                 Inches(2.40), Inches(0.85), title,
                 font=FONT, size=15, bold=True, color=WHITE,
                 line_spacing=1.2)
        add_rect(s, Inches(4.05), y, Inches(8.85), row_h,
                 fill=WHITE, line=GRAY_300)
        add_text(s, Inches(4.25), y + Inches(0.10),
                 Inches(8.55), Inches(0.7), desc,
                 font=FONT, size=12, color=GRAY_700, line_spacing=1.4)
        add_text(s, Inches(4.25), y + Inches(0.95),
                 Inches(8.55), Inches(0.55),
                 "▎ " + advice,
                 font=FONT, size=11.5, bold=True, color=PRIMARY_DARK,
                 line_spacing=1.3)


# ============ 二季度工作计划 ============
def slide_q2_plan():
    s = add_slide("3_正文页")
    add_subtitle(s, "二季度重点工作计划")
    add_subdesc(s, "时间节点 · 责任明确 · 量化目标")

    timeline_y = Inches(2.30)
    add_line_conn(s, Inches(0.80), timeline_y,
                  Inches(12.50), timeline_y, color=PRIMARY, weight=2.5)
    nodes = [
        ("5 月底", Inches(2.50)),
        ("6 月底", Inches(6.60)),
        ("6 月 30 日", Inches(10.70)),
        ("持续推进", Inches(12.40)),
    ]
    for txt, x in nodes:
        add_shape(s, MSO_SHAPE.OVAL, x - Inches(0.10),
                  timeline_y - Inches(0.10),
                  Inches(0.20), Inches(0.20), fill=PRIMARY)
        add_text(s, x - Inches(0.7), timeline_y - Inches(0.55),
                 Inches(1.4), Inches(0.4), txt,
                 font=FONT, size=11, bold=True, color=PRIMARY_DARK,
                 align=PP_ALIGN.CENTER)

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
    card_w = Inches(3.0)
    card_h = Inches(3.65)
    gap = Inches(0.13)
    start_x = Inches(0.45)
    for i, (no, title, deadline, color, desc) in enumerate(plans):
        x = start_x + (card_w + gap) * i
        y = Inches(3.00)
        add_rect(s, x, y, card_w, card_h, fill=WHITE, line=GRAY_300)
        add_rect(s, x, y, card_w, Inches(0.10), fill=color)
        add_text(s, x + Inches(0.25), y + Inches(0.25),
                 Inches(2.5), Inches(0.6), no,
                 font=FONT, size=30, bold=True, color=color)
        add_rect(s, x + Inches(0.25), y + Inches(0.95),
                 Inches(2.0), Inches(0.40), fill=GRAY_100)
        add_text(s, x + Inches(0.25), y + Inches(0.95),
                 Inches(2.0), Inches(0.40),
                 f"⏱  {deadline}", font=FONT, size=11, bold=True,
                 color=color, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.25), y + Inches(1.50),
                 card_w - Inches(0.5), Inches(0.6), title,
                 font=FONT, size=16, bold=True, color=GRAY_900,
                 line_spacing=1.2)
        add_rect(s, x + Inches(0.25), y + Inches(2.05),
                 Inches(0.30), Inches(0.04), fill=color)
        add_text(s, x + Inches(0.25), y + Inches(2.20),
                 card_w - Inches(0.5), Inches(1.4), desc,
                 font=FONT, size=11.5, color=GRAY_700, line_spacing=1.5)

    bottom_y = Inches(6.78)
    add_rect(s, Inches(0.45), bottom_y, Inches(12.45), Inches(0.45),
             fill=PRIMARY_LIGHT)
    add_text(s, Inches(0.45), bottom_y, Inches(12.45), Inches(0.45),
             "强化目标引领　|　压实工作责任　|　补齐短板弱项　|　高质量完成全年各项目标任务",
             font=FONT, size=13, bold=True, color=PRIMARY_DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============ 结束语 ============
def slide_end():
    s = add_slide("结束语")
    set_placeholder_text(s, "标题", "汇报结束　敬请批评指正",
                         size=44, bold=True, color=WHITE,
                         align=PP_ALIGN.CENTER)
    set_placeholder_text(s, "副标题",
                         "信息化运营管理部　|　二〇二六年四月",
                         size=20, bold=False, color=WHITE,
                         align=PP_ALIGN.CENTER)


# ======== 执行 ========
slide_cover()
slide_toc()

slide_section(1, "一季度总体指标完成情况")
slide_kpi_dashboard()
slide_kpi_analysis()

slide_section(2, "一季度工作进展")
slide_progress_overview()
slide_oa_overview()
slide_oa_highlights()
slide_oa_data()
slide_union_launch()
slide_union_dev()
slide_ops_1()
slide_ops_2()

slide_section(3, "存在的问题及不足")
slide_issues()

slide_section(4, "二季度工作计划")
slide_q2_plan()

slide_end()

prs.save(OUTPUT)
print(f"✓ 已生成 {OUTPUT}，共 {len(prs.slides)} 页幻灯片。")
