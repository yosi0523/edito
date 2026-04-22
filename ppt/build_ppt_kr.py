# -*- coding: utf-8 -*-
"""
Tenneco × Hyundai Project Presentation (Korean version)
For Tenneco US HQ internal seminar - 16:9
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ===== Brand colors =====
TEAL = RGBColor(0x1A, 0x6F, 0x8C)
TEAL_DEEP = RGBColor(0x0E, 0x43, 0x57)
TEAL_DARK = RGBColor(0x12, 0x55, 0x6E)
LIME = RGBColor(0xD7, 0xE5, 0x5A)
ACCENT = RGBColor(0xE8, 0xA8, 0x3E)  # orange accent
BG_LIGHT = RGBColor(0xF5, 0xF5, 0xF7)
BG_SOFT = RGBColor(0xE8, 0xF0, 0xF3)
TEXT_DARK = RGBColor(0x2B, 0x2B, 0x2B)
TEXT_MID = RGBColor(0x55, 0x55, 0x55)
TEXT_LIGHT = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC0, 0x39, 0x2B)
GREEN = RGBColor(0x2E, 0x8B, 0x57)
GRAY_LINE = RGBColor(0xD0, 0xD0, 0xD0)
GRAY_LIGHT = RGBColor(0xEE, 0xEE, 0xEE)

FONT_KR = "맑은 고딕"
FONT_EN = "Arial"

# ===== Presentation setup =====
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW = prs.slide_width
SH = prs.slide_height
BLANK = prs.slide_layouts[6]


# ===== Helpers =====
def add_slide():
    return prs.slides.add_slide(BLANK)


def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=0):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w) if line_w else Pt(0.75)
    return shp


def add_round_rect(slide, x, y, w, h, fill=None, line=None, radius=0.08):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    shp.adjustments[0] = radius
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    return shp


def add_text(slide, x, y, w, h, text, size=14, bold=False, color=TEXT_DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=FONT_KR,
             italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb


def add_multi_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT,
                   anchor=MSO_ANCHOR.TOP):
    """runs = list of list of dicts (paragraphs of runs)
       each run dict: text, size, bold, color, font"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for rd in para:
            r = p.add_run()
            r.text = rd.get("text", "")
            r.font.name = rd.get("font", FONT_KR)
            r.font.size = Pt(rd.get("size", 14))
            r.font.bold = rd.get("bold", False)
            r.font.italic = rd.get("italic", False)
            c = rd.get("color", TEXT_DARK)
            r.font.color.rgb = c
    return tb


def add_line(slide, x1, y1, x2, y2, color=GRAY_LINE, weight=1.0):
    ln = slide.shapes.add_connector(1, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln


def add_footer(slide, page_num=None, total=None):
    # Left footer
    add_text(slide, Inches(0.3), Inches(7.18), Inches(5), Inches(0.25),
             "General Business – Tenneco Confidential",
             size=9, color=TEXT_LIGHT, italic=True)
    # Right: Tenneco mark
    add_text(slide, Inches(11.2), Inches(7.18), Inches(2), Inches(0.25),
             "TENNECO × HYUNDAI", size=9, color=TEAL, bold=True,
             align=PP_ALIGN.RIGHT, font=FONT_EN)
    if page_num is not None:
        add_text(slide, Inches(12.7), Inches(7.18), Inches(0.55), Inches(0.25),
                 f"{page_num}", size=9, color=TEXT_MID,
                 align=PP_ALIGN.RIGHT, font=FONT_EN)


def add_top_bar(slide, section_num, section_name, title):
    """Thin top bar with section marker + title block for content slides."""
    # Top accent line
    add_rect(slide, 0, 0, SW, Inches(0.08), fill=TEAL_DEEP)
    # Section label tab
    add_rect(slide, Inches(0.5), Inches(0.22), Inches(1.1), Inches(0.38),
             fill=TEAL)
    add_text(slide, Inches(0.5), Inches(0.22), Inches(1.1), Inches(0.38),
             f"SECTION {section_num:02d}", size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=FONT_EN)
    # Section name
    add_text(slide, Inches(1.75), Inches(0.22), Inches(6), Inches(0.38),
             section_name, size=12, bold=False, color=TEXT_MID,
             anchor=MSO_ANCHOR.MIDDLE)
    # Main title
    add_text(slide, Inches(0.5), Inches(0.7), Inches(12.3), Inches(0.7),
             title, size=26, bold=True, color=TEAL_DEEP,
             anchor=MSO_ANCHOR.MIDDLE)
    # Divider under title
    add_line(slide, Inches(0.5), Inches(1.45), Inches(12.83), Inches(1.45),
             color=LIME, weight=2)


def add_placeholder(slide, x, y, w, h, label="이미지 삽입 자리"):
    """Dashed-looking placeholder block for images to be added later."""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    shp.fill.solid()
    shp.fill.fore_color.rgb = BG_SOFT
    shp.line.color.rgb = TEAL
    shp.line.width = Pt(1.0)
    # dashed
    lnL = shp.line._get_or_add_ln()
    prstDash = etree.SubElement(lnL, qn('a:prstDash'))
    prstDash.set('val', 'dash')
    # label text
    add_text(slide, x, y, w, h,
             f"📷  {label}",
             size=12, bold=True, color=TEAL,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return shp


def add_bullet_block(slide, x, y, w, h, items, size=12, color=TEXT_DARK,
                     bullet="• ", spacing=4):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_top = Inches(0.02)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        if i > 0:
            p.space_before = Pt(spacing)
        r = p.add_run()
        r.text = bullet + it
        r.font.name = FONT_KR
        r.font.size = Pt(size)
        r.font.color.rgb = color
    return tb


def make_table(slide, x, y, w, h, data, col_widths=None, header_fill=TEAL,
               header_color=WHITE, body_fill=WHITE, body_alt=BG_LIGHT,
               font_size=11, header_size=12, first_col_fill=None,
               first_col_color=None):
    rows = len(data)
    cols = len(data[0])
    tbl = slide.shapes.add_table(rows, cols, x, y, w, h).table
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = Emu(int(w * cw / total))
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            # fill
            if ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_fill
            elif first_col_fill is not None and ci == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = first_col_fill
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = body_alt if (ri % 2 == 0) else body_fill
            # text
            tf = cell.text_frame
            tf.margin_left = Inches(0.06)
            tf.margin_right = Inches(0.06)
            tf.margin_top = Inches(0.03)
            tf.margin_bottom = Inches(0.03)
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if ri == 0 else PP_ALIGN.LEFT
            p.text = ""
            r = p.add_run()
            r.text = str(val)
            r.font.name = FONT_KR
            if ri == 0:
                r.font.size = Pt(header_size)
                r.font.bold = True
                r.font.color.rgb = header_color
            else:
                r.font.size = Pt(font_size)
                if first_col_color is not None and ci == 0:
                    r.font.color.rgb = first_col_color
                    r.font.bold = True
                else:
                    r.font.color.rgb = TEXT_DARK
    return tbl


def add_kpi_card(slide, x, y, w, h, title, value, unit="", fill=TEAL,
                 title_color=WHITE, value_color=WHITE):
    add_round_rect(slide, x, y, w, h, fill=fill, radius=0.1)
    add_text(slide, x, y + Inches(0.15), w, Inches(0.35), title,
             size=11, bold=False, color=title_color, align=PP_ALIGN.CENTER)
    add_text(slide, x, y + Inches(0.48), w, Inches(0.6), value,
             size=28, bold=True, color=value_color, align=PP_ALIGN.CENTER,
             font=FONT_EN)
    if unit:
        add_text(slide, x, y + Inches(1.05), w, Inches(0.3), unit,
                 size=10, color=title_color, align=PP_ALIGN.CENTER)


def add_section_divider(slide, section_num, section_kr, section_en, page):
    # background full teal
    add_rect(slide, 0, 0, SW, SH, fill=TEAL_DEEP)
    # lime accent strip
    add_rect(slide, 0, Inches(3.4), SW, Inches(0.06), fill=LIME)
    # section num huge
    add_text(slide, Inches(0.8), Inches(2.1), Inches(4), Inches(1.8),
             f"{section_num:02d}", size=140, bold=True, color=LIME,
             font=FONT_EN, anchor=MSO_ANCHOR.MIDDLE)
    # Korean section name
    add_text(slide, Inches(4.8), Inches(2.6), Inches(8), Inches(0.8),
             section_kr, size=34, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.BOTTOM)
    # English subtitle
    add_text(slide, Inches(4.8), Inches(3.5), Inches(8), Inches(0.6),
             section_en, size=16, color=LIME, font=FONT_EN,
             anchor=MSO_ANCHOR.TOP)
    # Tenneco ribbon (bottom right)
    add_text(slide, Inches(10.5), Inches(6.8), Inches(2.5), Inches(0.4),
             "TENNECO × HYUNDAI", size=11, bold=True, color=LIME,
             align=PP_ALIGN.RIGHT, font=FONT_EN)
    # page
    add_text(slide, Inches(0.5), Inches(6.8), Inches(2), Inches(0.4),
             f"— {page} —", size=10, color=LIME, font=FONT_EN)


# =======================================================================
#                           S L I D E S
# =======================================================================

# -------------------- S1: Cover --------------------
def slide_cover():
    s = add_slide()
    # background teal
    add_rect(s, 0, 0, SW, SH, fill=TEAL_DEEP)
    # diagonal accent band (faux)
    accent = s.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM,
                                Inches(-1), Inches(5.8), Inches(16), Inches(1.2))
    accent.shadow.inherit = False
    accent.fill.solid()
    accent.fill.fore_color.rgb = TEAL
    accent.line.fill.background()
    accent.adjustments[0] = 0.6
    # Lime thin line
    add_rect(s, 0, Inches(5.65), SW, Inches(0.05), fill=LIME)
    # Top label
    add_text(s, Inches(0.8), Inches(0.8), Inches(12), Inches(0.4),
             "TENNECO INTERNAL SEMINAR — 2026", size=13, bold=True,
             color=LIME, font=FONT_EN)
    # Title Korean
    add_text(s, Inches(0.8), Inches(1.8), Inches(12), Inches(1.3),
             "Hyundai Motor Group 프로젝트 현황 및 공략 전략",
             size=40, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    # Title English
    add_text(s, Inches(0.8), Inches(3.1), Inches(12), Inches(0.7),
             "Tenneco × Hyundai : Current Business & Growth Strategy",
             size=20, color=LIME, font=FONT_EN, anchor=MSO_ANCHOR.MIDDLE)
    # Sub description
    add_text(s, Inches(0.8), Inches(4.0), Inches(12), Inches(0.5),
             "글로벌 시장 현황  ·  공급 프로그램  ·  경쟁사 비교  ·  CVSA2 확대  ·  운영 로드맵",
             size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    # Presenter block (bottom)
    add_text(s, Inches(0.8), Inches(6.2), Inches(6), Inches(0.35),
             "Tenneco Korea – Hyundai/Kia Account", size=13, bold=True,
             color=WHITE, font=FONT_EN)
    add_text(s, Inches(0.8), Inches(6.55), Inches(6), Inches(0.3),
             "대한민국 · Tenneco 본사 세미나 발표 자료",
             size=11, color=LIME)
    # Right side big "T"
    add_text(s, Inches(10.8), Inches(1.2), Inches(2.2), Inches(2.5),
             "T", size=260, bold=True, color=TEAL, font=FONT_EN,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# -------------------- S2: 목차 --------------------
def slide_toc():
    s = add_slide()
    add_rect(s, 0, 0, SW, SH, fill=WHITE)
    # Left teal column
    add_rect(s, 0, 0, Inches(4.3), SH, fill=TEAL_DEEP)
    add_text(s, Inches(0.5), Inches(0.6), Inches(3.6), Inches(0.4),
             "AGENDA", size=14, bold=True, color=LIME, font=FONT_EN)
    add_text(s, Inches(0.5), Inches(1.1), Inches(3.6), Inches(1.2),
             "목 차", size=66, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.TOP)
    add_text(s, Inches(0.5), Inches(2.7), Inches(3.6), Inches(0.6),
             "Hyundai Project\nPresentation", size=16, color=LIME,
             font=FONT_EN)
    add_text(s, Inches(0.5), Inches(6.7), Inches(3.6), Inches(0.5),
             "2026", size=22, bold=True, color=LIME, font=FONT_EN)

    items = [
        ("01", "현대 글로벌 시장 판매", "Hyundai Global Market Share"),
        ("02", "현대 북미 및 남미", "Hyundai N.America & S.America Market"),
        ("03", "현대에서의 Tenneco 점유율", "Tenneco Business Ratio in Hyundai"),
        ("04", "Tenneco의 현대 공급 프로그램", "Hyundai Supply Program (China · India · Brazil)"),
        ("05", "경쟁사 비교", "Competitor Analysis (Mando · ZF · Bilstein)"),
        ("06", "CVSA2 프로그램 확대 방안", "CVSA2 Program & Expansion Plan"),
        ("07", "현대 공략 방법", "Hyundai Winning Strategy"),
        ("08", "Tenneco Korea의 단계별 운영 방안", "Tenneco Korea Operating Roadmap"),
        ("09", "Q & A", "Discussion"),
    ]
    # Right content
    x0 = Inches(4.8)
    y0 = Inches(0.9)
    row_h = Inches(0.65)
    for i, (num, kr, en) in enumerate(items):
        y = y0 + row_h * i
        # number circle
        add_round_rect(s, x0, y + Inches(0.05), Inches(0.55), Inches(0.55),
                       fill=LIME, radius=0.5)
        add_text(s, x0, y + Inches(0.05), Inches(0.55), Inches(0.55),
                 num, size=16, bold=True, color=TEAL_DEEP,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 font=FONT_EN)
        # KR title
        add_text(s, x0 + Inches(0.8), y, Inches(7.5), Inches(0.38),
                 kr, size=17, bold=True, color=TEAL_DEEP)
        # EN subtitle
        add_text(s, x0 + Inches(0.8), y + Inches(0.38), Inches(7.5), Inches(0.28),
                 en, size=10, color=TEXT_MID, italic=True, font=FONT_EN)
        # divider line
        if i < len(items) - 1:
            add_line(s, x0 + Inches(0.8), y + row_h - Inches(0.03),
                     x0 + Inches(8.2), y + row_h - Inches(0.03),
                     color=GRAY_LINE, weight=0.5)
    add_footer(s, 2)


# -------------------- S3: Section 1 divider --------------------
def slide_section1():
    s = add_slide()
    add_section_divider(s, 1, "현대 글로벌 시장 판매",
                        "Hyundai Global Market Share · 2025", 3)


# -------------------- S4: Hyundai Global Market Share --------------------
def slide_global_market():
    s = add_slide()
    add_top_bar(s, 1, "현대 글로벌 시장 판매",
                "1. Hyundai Global Market Share  |  2025년 기준")
    # KPI cards row (top)
    add_kpi_card(s, Inches(0.5), Inches(1.7), Inches(2.8), Inches(1.3),
                 "2025 현대차 글로벌 판매", "414만", unit="대 (Total Sales)")
    add_kpi_card(s, Inches(3.45), Inches(1.7), Inches(2.8), Inches(1.3),
                 "최대 판매 지역", "아시아", unit="146만 대 (35.3%)",
                 fill=TEAL_DARK)
    add_kpi_card(s, Inches(6.4), Inches(1.7), Inches(2.8), Inches(1.3),
                 "핵심 전략 시장", "아메리카", unit="134만 대 (32.4%)",
                 fill=TEAL)
    add_kpi_card(s, Inches(9.35), Inches(1.7), Inches(2.8), Inches(1.3),
                 "기준 연도", "2025", unit="Year of Record",
                 fill=ACCENT, title_color=WHITE, value_color=WHITE)

    # Regional table (left)
    add_text(s, Inches(0.5), Inches(3.2), Inches(7.5), Inches(0.4),
             "■ 2025년 현대자동차 지역별 판매 현황", size=14, bold=True,
             color=TEAL_DEEP)
    table_data = [
        ["지역", "주요 국가 및 거점", "판매 수량 (추정)", "비중"],
        ["아메리카", "미국, 캐나다, 브라질(HMB), 멕시코", "약 134만 대", "32.4%"],
        ["아시아", "한국(내수), 인도, 베트남, 인도네시아", "약 146만 대", "35.3%"],
        ["유럽", "독일, 영국, 프랑스 등 EU 지역", "약 61만 대", "14.7%"],
        ["기타", "중동, 아프리카, 러시아, 호주 등", "약 73만 대", "17.6%"],
        ["합계", "전 세계 시장", "약 414만 대", "100%"],
    ]
    make_table(s, Inches(0.5), Inches(3.65), Inches(7.5), Inches(2.8),
               table_data, col_widths=[1.4, 3.5, 1.8, 1.0],
               first_col_fill=BG_SOFT, first_col_color=TEAL_DEEP)

    # Right column: Tenneco footprint
    add_rect(s, Inches(8.3), Inches(3.2), Inches(4.55), Inches(3.3),
             fill=TEAL_DEEP)
    add_text(s, Inches(8.45), Inches(3.3), Inches(4.3), Inches(0.4),
             "TENNECO GLOBAL FOOTPRINT", size=12, bold=True,
             color=LIME, font=FONT_EN)
    add_text(s, Inches(8.45), Inches(3.65), Inches(4.3), Inches(0.35),
             "전 세계 공급 역량", size=11, color=WHITE)

    kpis = [
        ("60,000", "Global Team Members"),
        ("184", "Manufacturing Plants"),
        ("23", "Distribution Centers"),
        ("40", "Engineering Facilities"),
    ]
    for i, (v, lbl) in enumerate(kpis):
        yy = Inches(4.1 + i * 0.55)
        add_text(s, Inches(8.45), yy, Inches(1.3), Inches(0.4),
                 v, size=22, bold=True, color=LIME,
                 font=FONT_EN, align=PP_ALIGN.RIGHT)
        add_text(s, Inches(9.85), yy + Inches(0.08), Inches(2.9),
                 Inches(0.4), lbl, size=11, color=WHITE, font=FONT_EN)

    # Bottom insight
    add_rect(s, Inches(0.5), Inches(6.55), Inches(12.33), Inches(0.5),
             fill=BG_SOFT)
    add_text(s, Inches(0.7), Inches(6.55), Inches(12.0), Inches(0.5),
             "Insight  ·  현대차는 지정학적 리스크와 경기 변동 속에서도 북미 최대 실적 및 인도 시장 성장을 통해 글로벌 판매 비중을 성공적으로 다변화했다.",
             size=11, color=TEAL_DEEP, anchor=MSO_ANCHOR.MIDDLE, italic=True)
    add_footer(s, 4)


# -------------------- S5: Section 2 --------------------
def slide_section2():
    s = add_slide()
    add_section_divider(s, 2, "현대 북미 및 남미",
                        "Hyundai N.America & S.America Market · 2025", 5)


# -------------------- S6: 2-1 America detail --------------------
def slide_america_detail():
    s = add_slide()
    add_top_bar(s, 2, "현대 북미 및 남미",
                "2-1. Hyundai N.America & S.America Market  |  국가별 판매 현황")

    # Left: table
    add_text(s, Inches(0.5), Inches(1.7), Inches(7), Inches(0.4),
             "■ 2025년 현대자동차 핵심 국가별 판매 현황 (글로벌 134만 대 기준)",
             size=13, bold=True, color=TEAL_DEEP)
    data = [
        ["지역", "핵심 국가", "판매 수량 (추정)", "글로벌 비중"],
        ["북미", "미국 (USA)", "약 90.2만 대", "21.8%"],
        ["북미", "캐나다 (Canada)", "약 14.1만 대", "3.4%"],
        ["북미", "멕시코 (Mexico)", "약 5.1만 대", "1.2%"],
        ["남미", "브라질 (HMB)", "약 18.5만 대", "4.5%"],
        ["기타", "기타 아메리카 국가", "약 6.1만 대", "1.5%"],
        ["합계", "아메리카 전체", "약 134.0만 대", "32.4%"],
    ]
    make_table(s, Inches(0.5), Inches(2.15), Inches(7.2), Inches(3.7), data,
               col_widths=[1.1, 2.9, 2.0, 1.2],
               first_col_fill=BG_SOFT, first_col_color=TEAL_DEEP,
               font_size=11)

    # Right: bar chart rendered with rectangles
    add_text(s, Inches(8.0), Inches(1.7), Inches(5), Inches(0.4),
             "■ 판매량 시각화 (만 대)", size=13, bold=True, color=TEAL_DEEP)
    bars = [
        ("미국", 90.2, TEAL_DEEP),
        ("브라질", 18.5, TEAL),
        ("캐나다", 14.1, TEAL),
        ("기타", 6.1, ACCENT),
        ("멕시코", 5.1, ACCENT),
    ]
    bar_x = Inches(9.2)
    bar_max_w = Inches(3.4)
    bar_max_val = 100.0
    by = Inches(2.2)
    for i, (lbl, val, col) in enumerate(bars):
        yy = by + Inches(0.6) * i
        add_text(s, Inches(8.0), yy, Inches(1.15), Inches(0.4),
                 lbl, size=11, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE,
                 align=PP_ALIGN.RIGHT)
        bw = Emu(int(bar_max_w * val / bar_max_val))
        add_rect(s, bar_x, yy + Inches(0.08), bw, Inches(0.28), fill=col)
        add_text(s, bar_x + bw + Inches(0.05), yy, Inches(1), Inches(0.4),
                 f"{val:.1f}만", size=10, color=TEXT_MID,
                 anchor=MSO_ANCHOR.MIDDLE, font=FONT_EN)

    # Bottom insight
    add_rect(s, Inches(0.5), Inches(6.15), Inches(12.33), Inches(0.9),
             fill=TEAL_DEEP)
    add_text(s, Inches(0.7), Inches(6.2), Inches(12.0), Inches(0.35),
             "KEY TAKEAWAY", size=10, bold=True, color=LIME, font=FONT_EN)
    add_text(s, Inches(0.7), Inches(6.5), Inches(12.0), Inches(0.5),
             "아메리카 판매의 67%가 미국에 집중 (90.2만 대). 브라질·멕시코의 현지 생산 기반(HMB)은 Tenneco 공급 확장의 핵심 레버리지.",
             size=12, color=WHITE, anchor=MSO_ANCHOR.TOP)
    add_footer(s, 6)


# -------------------- S7: 2-2 Competitor comparison --------------------
def slide_america_vs_competitor():
    s = add_slide()
    add_top_bar(s, 2, "현대 북미 및 남미",
                "2-2. GM · Toyota · Hyundai 판매량 비교  |  2025년 기준")

    # comparison cards (3 big cards)
    data = [
        {"name": "GM", "vol": "약 320 ~ 340만 대", "rank": "북미 1위",
         "desc": "대형 픽업트럭과 SUV 중심의 압도적 시장 지배력",
         "fill": TEAL_DEEP, "tag": "DOMINANT LEADER"},
        {"name": "TOYOTA", "vol": "약 280 ~ 300만 대", "rank": "하이브리드 1위",
         "desc": "하이브리드 절대 강자. 북미·남미 모두에서 높은 신뢰도 확보",
         "fill": TEAL, "tag": "HYBRID POWERHOUSE"},
        {"name": "HYUNDAI", "vol": "약 134만 대", "rank": "성장세 1위",
         "desc": "전기차 및 SUV 라인업 강화를 통해 추격 중. Tenneco의 주력 고객",
         "fill": ACCENT, "tag": "FASTEST GROWING"},
    ]
    card_w = Inches(4.1)
    gap = Inches(0.1)
    x0 = Inches(0.5)
    y0 = Inches(1.8)
    card_h = Inches(3.8)
    for i, d in enumerate(data):
        x = x0 + (card_w + gap) * i
        add_rect(s, x, y0, card_w, card_h, fill=d["fill"])
        # Tag
        add_text(s, x + Inches(0.2), y0 + Inches(0.2), card_w - Inches(0.4),
                 Inches(0.3), d["tag"], size=10, bold=True, color=LIME,
                 font=FONT_EN)
        # Name
        add_text(s, x + Inches(0.2), y0 + Inches(0.55), card_w - Inches(0.4),
                 Inches(0.8), d["name"], size=36, bold=True, color=WHITE,
                 font=FONT_EN)
        # Line
        add_rect(s, x + Inches(0.2), y0 + Inches(1.45),
                 Inches(0.8), Inches(0.04), fill=LIME)
        # Volume label
        add_text(s, x + Inches(0.2), y0 + Inches(1.6), card_w - Inches(0.4),
                 Inches(0.3), "아메리카 총 판매량", size=10, color=LIME)
        add_text(s, x + Inches(0.2), y0 + Inches(1.9), card_w - Inches(0.4),
                 Inches(0.45), d["vol"], size=20, bold=True, color=WHITE)
        # Rank
        add_text(s, x + Inches(0.2), y0 + Inches(2.45), card_w - Inches(0.4),
                 Inches(0.3), "시장 지위", size=10, color=LIME)
        add_text(s, x + Inches(0.2), y0 + Inches(2.7), card_w - Inches(0.4),
                 Inches(0.35), d["rank"], size=15, bold=True, color=WHITE)
        # Desc
        add_text(s, x + Inches(0.2), y0 + Inches(3.1), card_w - Inches(0.4),
                 Inches(0.7), d["desc"], size=10, color=WHITE)
    # bottom strip
    add_rect(s, Inches(0.5), Inches(5.8), Inches(12.33), Inches(1.2),
             fill=BG_SOFT)
    add_text(s, Inches(0.7), Inches(5.9), Inches(12), Inches(0.35),
             "■ 전략적 해석", size=12, bold=True, color=TEAL_DEEP)
    add_text(s, Inches(0.7), Inches(6.25), Inches(12), Inches(0.7),
             "현대차는 GM·Toyota 대비 절대 판매량은 낮지만 성장률·전동화 속도·신규 플랫폼(EV GMP) 기준으로 가장 공격적. "
             "Tenneco 입장에서 중장기 볼륨 확대 여지가 가장 큰 OEM.",
             size=11, color=TEXT_DARK)
    add_footer(s, 7)


# -------------------- S8: Section 3 --------------------
def slide_section3():
    s = add_slide()
    add_section_divider(s, 3, "현대에서의 Tenneco 점유율",
                        "Tenneco Business Ratio in Hyundai", 8)


# -------------------- S9: Tenneco share in Hyundai --------------------
def slide_tenneco_share():
    s = add_slide()
    add_top_bar(s, 3, "현대에서의 Tenneco 점유율",
                "3. Tenneco Business Ratio in Hyundai  |  공급사별 점유 현황")

    # Left: pie-chart-like visualization using rectangles (stacked)
    add_text(s, Inches(0.5), Inches(1.7), Inches(6), Inches(0.4),
             "■ 현대/기아 섀시 댐퍼 공급사 점유율 (추정)", size=13, bold=True,
             color=TEAL_DEEP)
    # Horizontal stacked bar
    bar_x = Inches(0.5)
    bar_y = Inches(2.3)
    bar_w = Inches(6.2)
    bar_h = Inches(0.6)
    segs = [
        ("Mando", 65, TEAL_DEEP),
        ("ZF", 18, TEAL),
        ("Bilstein", 10, ACCENT),
        ("Tenneco", 7, LIME),
    ]
    cur_x = bar_x
    for name, pct, col in segs:
        w = Emu(int(bar_w * pct / 100))
        add_rect(s, cur_x, bar_y, w, bar_h, fill=col)
        tc = TEAL_DEEP if col == LIME else WHITE
        add_text(s, cur_x, bar_y, w, bar_h,
                 f"{name}\n{pct}%", size=11, bold=True, color=tc,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cur_x = Emu(cur_x + w)
    # legend note
    add_text(s, Inches(0.5), Inches(3.05), Inches(6.2), Inches(0.35),
             "* 추정치: Mando 국내 현대/기아 기본 섀시 독점 + ZF(Genesis/Ioniq6 N) · Bilstein(High-end) · Tenneco(초기 진입)",
             size=9, italic=True, color=TEXT_LIGHT)

    # Supplier breakdown list under bar
    break_data = [
        ("Mando", "현대/기아 기본 섀시의 대부분 (ICE/EV 공용 Conventional + Semi-active)"),
        ("ZF", "Genesis Entry-Level, Ioniq6 N (Semi-active ECS)"),
        ("Bilstein", "Genesis High-end (Air shock + ECS) 전용 포지션"),
        ("Tenneco", "초기 진입 단계 — CVSA2 Demo Car 검증 단계"),
    ]
    add_text(s, Inches(0.5), Inches(3.5), Inches(6.2), Inches(0.35),
             "■ 공급 영역 Summary", size=13, bold=True, color=TEAL_DEEP)
    yy = Inches(3.95)
    for name, desc in break_data:
        col = {"Mando": TEAL_DEEP, "ZF": TEAL, "Bilstein": ACCENT,
               "Tenneco": LIME}[name]
        add_rect(s, Inches(0.5), yy, Inches(0.15), Inches(0.4), fill=col)
        add_text(s, Inches(0.75), yy, Inches(1.4), Inches(0.4),
                 name, size=11, bold=True, color=TEAL_DEEP,
                 anchor=MSO_ANCHOR.MIDDLE, font=FONT_EN)
        add_text(s, Inches(2.15), yy, Inches(4.7), Inches(0.4),
                 desc, size=10, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)
        yy = Emu(yy + Inches(0.45))

    # Right: Opportunity call-out
    add_rect(s, Inches(7.0), Inches(1.7), Inches(5.83), Inches(4.8),
             fill=TEAL_DEEP)
    add_text(s, Inches(7.2), Inches(1.85), Inches(5.5), Inches(0.35),
             "OPPORTUNITY", size=11, bold=True, color=LIME, font=FONT_EN)
    add_text(s, Inches(7.2), Inches(2.2), Inches(5.5), Inches(0.7),
             "Tenneco의 점유 확장 여지",
             size=22, bold=True, color=WHITE)
    add_line(s, Inches(7.2), Inches(3.0), Inches(9.0), Inches(3.0),
             color=LIME, weight=2)

    ops = [
        ("▲ 2-Valve MacPherson", "현대가 차기 제품에 요구하는 핵심 기술 — Tenneco만 보유"),
        ("▲ 가격 경쟁력", "ZF·Bilstein 대비 HMC가 합리적이라 평가"),
        ("▲ Golf 유럽 시승 호평", "2025년 4월 HMC 엔지니어 긍정 반응"),
        ("▲ VP Manfred Harrer", "유럽 공급사 선호 기조 확립"),
        ("▲ N → Genesis 수직 확산", "N Brand 검증 시 Genesis 자동 이관 구조"),
    ]
    yy = Inches(3.2)
    for t, d in ops:
        add_text(s, Inches(7.2), yy, Inches(5.5), Inches(0.3),
                 t, size=13, bold=True, color=LIME)
        add_text(s, Inches(7.35), yy + Inches(0.3), Inches(5.3), Inches(0.35),
                 d, size=10, color=WHITE)
        yy = Emu(yy + Inches(0.67))

    # bottom call
    add_rect(s, Inches(0.5), Inches(6.55), Inches(12.33), Inches(0.45),
             fill=LIME)
    add_text(s, Inches(0.7), Inches(6.55), Inches(12), Inches(0.45),
             "현재 Tenneco 점유율은 한 자릿수 초기 진입. CVSA2 Ioniq6 N RFQ 확정 시 두 자릿수 급상승 가능.",
             size=12, bold=True, color=TEAL_DEEP, anchor=MSO_ANCHOR.MIDDLE)
    add_footer(s, 9)


# ===== Run =====
if __name__ == "__main__":
    slide_cover()                       # 1
    slide_toc()                         # 2
    slide_section1()                    # 3
    slide_global_market()               # 4
    slide_section2()                    # 5
    slide_america_detail()              # 6
    slide_america_vs_competitor()       # 7
    slide_section3()                    # 8
    slide_tenneco_share()               # 9

    import os
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "Tenneco_Hyundai_KR_preview.pptx")
    prs.save(out_path)
    print(f"Saved: {out_path}  ({len(prs.slides)} slides)")

