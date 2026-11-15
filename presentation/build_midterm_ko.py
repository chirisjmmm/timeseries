"""Build the midterm progress deck from the proposal deck's template (same masters/layouts)."""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.oxml.ns import qn
from lxml import etree

SRC = "투수부상_조기경보_제안발표_KO.pptx"
OUT = "투수부상_조기경보_중간발표_KO.pptx"

BLACK = RGBColor(0, 0, 0)
GRAY = RGBColor(0x59, 0x59, 0x59)
MID = RGBColor(0x88, 0x88, 0x88)
LIGHT = RGBColor(0xD9, 0xD9, 0xD9)
BARGRAY = RGBColor(0x9C, 0xA3, 0xAF)
RED = RGBColor(0xC0, 0x50, 0x4D)
NAVY = RGBColor(0x1F, 0x3A, 0x6E)
BLUE = RGBColor(0x3B, 0x6E, 0xB5)
PALE = RGBColor(0xF2, 0xF4, 0xF8)

SECTIONS = ["피드백 반영", "진행 상황", "실험 결과", "향후 계획"]

prs = Presentation(SRC)
slides = prs.slides
cover_src, thanks_src = slides[0], slides[18]
SLDNUM = copy.deepcopy([s for s in slides[2].shapes if s.name.startswith("Slide Number")][0]._element)

# --- drop proposal content slides 2..18 -------------------------------------------------------
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst)[1:18]:
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)


# --- text helpers ------------------------------------------------------------------------------
def runs_from(text):
    """'plain **bold** plain' -> [(txt, bold)]"""
    out = []
    for i, part in enumerate(text.split("**")):
        if part:
            out.append((part, i % 2 == 1))
    return out


def fill_para(p, text, size=13, color=BLACK, bold=None, align=None, italic=False):
    if align is not None:
        p.alignment = align
    for txt, b in runs_from(text):
        r = p.add_run()
        r.text = txt
        r.font.size = Pt(size)
        r.font.bold = bold if bold is not None else b
        r.font.italic = italic
        r.font.color.rgb = color


def no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    for el in pPr.findall(qn("a:buNone")):
        pPr.remove(el)
    etree.SubElement(pPr, qn("a:buNone"))


def textbox(slide, x, y, w, h, lines, size=13, color=BLACK, align=PP_ALIGN.LEFT,
            anchor=MSO_ANCHOR.TOP, border=None, border_w=0.75, fill=None, pad=0.0,
            spacing=0, name="Text"):
    """lines: list of str (supports **bold**) or (str, dict(size, color, bold, align))."""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.name = name
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if border is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = border
        shp.line.width = Pt(border_w)
    tf = shp.text_frame
    tf.word_wrap = True
    m = Inches(pad)
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = m
    tf.vertical_anchor = anchor
    first = True
    for ln in lines:
        opts = {}
        if isinstance(ln, tuple):
            ln, opts = ln
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        no_bullet(p)
        if spacing and not opts.get("nospace"):
            p.space_before = Pt(spacing)
        if "space" in opts:
            p.space_before = Pt(opts["space"])
        fill_para(p, ln, size=opts.get("size", size), color=opts.get("color", color),
                  bold=opts.get("bold"), align=opts.get("align", align))
    return shp


def bullets(slide, x, y, w, h, items, size=13, spacing=4, name="Bullets"):
    """items: str -> level-1 ❑ bullet; ('→', str) -> follow-on arrow line; ('-', str) -> sub bullet."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        kind, txt = ("box", it) if isinstance(it, str) else it
        pPr = p._p.get_or_add_pPr()
        if kind == "box":
            pPr.set("marL", "279400"); pPr.set("indent", "-279400")
            etree.SubElement(pPr, qn("a:buSzPct")).set("val", "100000")
            etree.SubElement(pPr, qn("a:buChar")).set("char", "❏")
            p.space_before = Pt(spacing if i else 0)
        elif kind == "-":
            pPr.set("marL", "558800"); pPr.set("indent", "-177800")
            etree.SubElement(pPr, qn("a:buChar")).set("char", "•")
            p.space_before = Pt(2)
        else:  # arrow continuation
            pPr.set("marL", "279400"); pPr.set("indent", "0")
            etree.SubElement(pPr, qn("a:buNone"))
            p.space_before = Pt(1)
            txt = "→ " + txt
        fill_para(p, txt, size=size if kind != "-" else size - 1)
    return tb


def rect(slide, x, y, w, h, color=BLACK, weight=0.75, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if color is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = color; s.line.width = Pt(weight)
    return s


def arrow(slide, x, y, w=0.36, h=0.36, down=False):
    s = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW if down else MSO_SHAPE.RIGHT_ARROW,
                               Inches(x), Inches(y), Inches(w), Inches(h))
    s.name = "Arrow"
    s.shadow.inherit = False
    s.fill.solid(); s.fill.fore_color.rgb = GRAY
    s.line.fill.background()
    return s


def hline(slide, x, y, w, color=BLACK, weight=0.75):
    c = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    c.line.color.rgb = color; c.line.width = Pt(weight)
    return c


def vline(slide, x, y, h, color=BLACK, weight=1.5):
    c = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x), Inches(y + h))
    c.line.color.rgb = color; c.line.width = Pt(weight)
    return c


def _border(tcPr, tag, w, color):
    el = tcPr.find(qn(tag))
    if el is not None:
        tcPr.remove(el)
    ln = etree.SubElement(tcPr, qn(tag))
    if w == 0:
        ln.set("w", "0"); etree.SubElement(ln, qn("a:noFill"))
    else:
        ln.set("w", str(int(w * 12700))); ln.set("cmpd", "sng")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
    return ln


def table(slide, x, y, colw, rows, rowh=0.3, size=11, head_size=11.5, align=None,
          bold_first_col=True, heights=None, group_rows=(), name="Table", highlight_rows=()):
    """rows[0] is header. group_rows: indexes of full-width section label rows (left aligned, bold)."""
    nr, nc = len(rows), len(colw)
    gf = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(colw)), Inches(rowh * nr))
    gf.name = name
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for a in ("firstRow", "bandRow"):
        tblPr.set(a, "0")
    st = tblPr.find(qn("a:tableStyleId"))
    if st is not None:
        tblPr.remove(st)
    for j, w in enumerate(colw):
        tbl.columns[j].width = Inches(w)
    for i in range(nr):
        tbl.rows[i].height = Inches(heights[i] if heights else rowh)
    align = align or [PP_ALIGN.CENTER] * nc
    for i, row in enumerate(rows):
        if i in group_rows:
            c0 = tbl.cell(i, 0)
            c0.merge(tbl.cell(i, nc - 1))
        for j in range(nc):
            cell = tbl.cell(i, j)
            txt = row[j] if j < len(row) else ""
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i in highlight_rows:
                cell.fill.solid(); cell.fill.fore_color.rgb = PALE
            else:
                cell.fill.background()
            tf = cell.text_frame
            tf.word_wrap = True
            lines = txt.split("\n") if txt else [""]
            for k, ln in enumerate(lines):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                no_bullet(p)
                if i == 0:
                    fill_para(p, ln, size=head_size, bold=True, align=PP_ALIGN.CENTER)
                elif i in group_rows:
                    fill_para(p, ln, size=size + 0.5, bold=True, align=PP_ALIGN.LEFT)
                else:
                    b = True if (j == 0 and bold_first_col) else None
                    fill_para(p, ln, size=size, bold=b, align=align[j])
            tcPr = cell._tc.get_or_add_tcPr()
            _border(tcPr, "a:lnL", 0, None)
            _border(tcPr, "a:lnR", 0, None)
            top = 1.25 if i == 0 else (0.75 if i == 1 else 0.5)
            topc = "000000" if i <= 1 else "BFBFBF"
            _border(tcPr, "a:lnT", top, topc)
            bot = 1.25 if i == nr - 1 else 0.5
            botc = "000000" if i == nr - 1 or i == 0 else "BFBFBF"
            if i == 0:
                bot = 0.75
            _border(tcPr, "a:lnB", bot, botc)
            # schema order: lnL lnR lnT lnB ... fill must come after borders
            fills = [e for e in tcPr if e.tag in (qn("a:solidFill"), qn("a:noFill"))]
            for f in fills:
                tcPr.remove(f); tcPr.append(f)
    return gf


def bar_chart(slide, x, y, w, h, cats, vals, colors, title=None, maxv=None, fmt="0", size=10,
              horizontal=False, gap=60):
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("v", vals)
    kind = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gfr = slide.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gfr.chart
    ch.has_legend = False
    ch.font.size = Pt(size)
    ch.font.color.rgb = BLACK
    if title:
        ch.has_title = True
        ch.chart_title.text_frame.text = title
        r = ch.chart_title.text_frame.paragraphs[0].runs[0]
        r.font.size = Pt(size + 1); r.font.bold = False
    else:
        ch.has_title = False
    plot = ch.plots[0]
    plot.gap_width = gap
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = fmt; dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(size)
    ser = plot.series[0]
    for i, c in enumerate(colors):
        pt = ser.points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = c
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    if maxv:
        va.maximum_scale = maxv
    va.minimum_scale = 0
    ca = ch.category_axis
    ca.tick_labels.font.size = Pt(size)
    ca.format.line.color.rgb = MID
    ca.has_major_gridlines = False
    if horizontal:
        ca.reverse_order = True
    return gfr


# --- slide factories -------------------------------------------------------------------------
L_CONTENT = [l for l in prs.slide_layouts if l.name == "CONTENT"][0]


def content(title, subtitle, active=None):
    s = slides.add_slide(L_CONTENT)
    for ph in s.placeholders:
        if ph.placeholder_format.idx == 103:
            ph.text_frame.text = title
        elif ph.placeholder_format.idx == 104:
            ph.text_frame.text = subtitle
    if active is not None:
        x0, w, step = 6.16, 0.88, 0.92
        for k, name in enumerate(SECTIONS):
            on = name == active
            t = textbox(s, x0 + k * step, 0.15, w, 0.27, [name], size=8.5,
                        color=BLACK if on else RGBColor(0x9A, 0x9A, 0x9A), align=PP_ALIGN.CENTER,
                        anchor=MSO_ANCHOR.MIDDLE, border=BLACK if on else LIGHT,
                        border_w=1.0 if on else 0.5, name=f"Section tracker {k + 1}")
            if on:
                for r in t.text_frame.paragraphs[0].runs:
                    r.font.bold = True
    s.shapes._spTree.append(copy.deepcopy(SLDNUM))
    return s


def caption(s, y, text, x=0.3, w=9.4):
    return textbox(s, x, y, w, 0.22, [text], size=9, color=MID, align=PP_ALIGN.CENTER)


# =================================================================================================
# Cover (reuse Korean proposal cover, update text)
for ph in cover_src.placeholders:
    if ph.placeholder_format.idx == 106:
        paras = ph.text_frame.paragraphs
        paras[0].runs[0].text = "중간 발표 · MLB Statcast 2015–2026 정규시즌"
        paras[2].runs[0].text = "2026.11"

# ---- Role -------------------------------------------------------------------------------------
s = content("역할", "팀 구성 및 역할 분담")
members = [("김소연 (팀장)", "2023035269"), ("정민서", "2022059961"),
           ("강동한", "2021055178"), ("최민석", "2022005496")]
rows = [["이름", "학번", "역할 및 기여"]] + [[m, n, ""] for m, n in members]
gf = table(s, 0.46, 1.16, [1.72, 1.6, 5.98], rows, heights=[0.32, 0.88, 0.88, 0.88, 0.88],
           size=12, head_size=12, align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT],
           bold_first_col=False)
for j in range(3):
    p = gf.table.cell(0, j).text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for r in p.runs:
        r.font.bold = False

# ---- Feedback 1 -------------------------------------------------------------------------------
s = content("피드백 반영", "(1) 경보의 의미와 활용", "피드백 반영")
rows = [["", "무엇을 하는가", "연구 범위"],
        ["조기 경보 (Early warning)", "부상이 IL로 드러나기 전에 높아진 위험을 알림", "포함 (주 목적)"],
        ["감지 (Detection)", "이미 있지만 아직 보고되지 않은 부상을 발견", "일부 포함"],
        ["설명 (Explanation)", "부상이 왜 생기는지 설명 (인과)", "제외"]]
table(s, 0.3, 1.12, [2.3, 5.6, 1.5], rows, heights=[0.32, 0.34, 0.34, 0.34], size=11.5, head_size=11.5,
      align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
bx_y, bx_h = 2.86, 1.9
textbox(s, 0.3, bx_y, 4.45, bx_h,
        [("경보 = 모델의 출력", {"bold": True, "size": 13}),
         ("투수 **자신의 과거**로 예측한 구속과 실제 구속을 비교합니다.", {"space": 8}),
         ("**7일 안에 팔 부상 IL 등재** 위험이 높으면 → 경보", {"space": 6})],
        size=12, border=BLACK, pad=0.16, anchor=MSO_ANCHOR.MIDDLE)
arrow(s, 4.82, bx_y + bx_h / 2 - 0.18, 0.36, 0.36)
textbox(s, 5.25, bx_y, 4.45, bx_h,
        [("후속 조치 = 사람의 결정", {"bold": True, "size": 13}),
         ("트레이너를 위한 **짧은 확인 목록**: 검진, 추가 휴식, 투구량 조절", {"space": 8}),
         ("유용한 선행 시간: **다음 등판 전** (1–5일)", {"space": 6})],
        size=12, border=BLACK, pad=0.16, anchor=MSO_ANCHOR.MIDDLE, fill=PALE)

# ---- Feedback 2 -------------------------------------------------------------------------------
s = content("피드백 반영", "(2) 정답 시점, 교란 요인, 데이터 누수", "피드백 반영")
rows = [["피드백", "반영 방법"],
        ["IL 등재일 vs 실제 발생", "IL 등재일 = 팀이 부상을 인지하고 조치한 날; 이를 정답으로 사용 (한계로 명시)"],
        ["부상 유형", "팔 부상 IL만 사용; 어깨 · 팔꿈치 · 전완을 따로 모델링"],
        ["투구량", "최근 7 / 28일 투구 수, 당일 투구 수, 휴식일"],
        ["개인차", "투수별 개인 기준선 (EWMA, 칼만 필터)"],
        ["회복", "연구 범위 밖; 과거 팔 IL 이력을 피처로 시험"]]
table(s, 0.3, 1.12, [2.2, 7.2], rows, heights=[0.32] + [0.36] * 5, size=11.5, head_size=11.5,
      align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT])
ty = 3.62
textbox(s, 0.3, ty, 1.6, 0.4, ["**누수 방지**"], size=12, anchor=MSO_ANCHOR.MIDDLE)
segs = [("학습  2015–2023", 4.2), ("보정  2024", 1.3), ("임계값  2025", 1.3), ("평가  2026", 1.2)]
xx = 1.9
for i, (lab, w) in enumerate(segs):
    textbox(s, xx, ty, w, 0.4, [f"**{lab}**"], size=11, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            border=BLACK, border_w=1.0, fill=PALE if i == 3 else None)
    xx += w
bullets(s, 0.3, 4.24, 9.4, 0.6,
        ["피처는 각 등판 **이전** 데이터만 사용; 모든 선택은 2025년까지로 하고 **2026년은 최종 평가 전까지 미열람**"],
        size=12)

# ---- Progress: data + features ------------------------------------------------------------------
s = content("진행 상황", "데이터 정제와 피처 엔지니어링", "진행 상황")
kpis = [("2015–2026", "정규시즌\n(기존 2021–2026)"), ("2,620", "팔 관련 IL\n(전체 5,516건 중)"),
        ("FF / SI / FC", "투수·시즌당\n직구 구종 1개"), ("1행", "투수 ×\n등판 단위")]
kw = 2.2
for i, (big, small) in enumerate(kpis):
    textbox(s, 0.3 + i * (kw + 0.2), 1.1, kw, 0.92,
            [(f"**{big}**", {"size": 17, "align": PP_ALIGN.CENTER}), (small, {"size": 10.5, "color": GRAY, "align": PP_ALIGN.CENTER, "space": 2})],
            border=BLACK, border_w=1.0, pad=0.06, anchor=MSO_ANCHOR.MIDDLE)
rows = [["구분", "피처 (해당 등판 이전 데이터로 계산)", "실험"],
        ["구속", "개인 기준 대비 감소율 (EWMA) · 최근 1–3등판 기울기 · 연속 하락 횟수", "1, 3, 4"],
        ["상태공간", "칼만 필터 구속 예측오차 z", "7"],
        ["회전수", "회전수 ÷ 구속: 감소와 기울기", "2"],
        ["투구량", "최근 7 / 28일 투구 수, 당일 투구 수, 휴식일", "8"],
        ["투수 특성", "키 × 구속 집단 위험 · 과거 팔 IL 이력", "5, 9"]]
table(s, 0.3, 2.24, [1.5, 7.1, 0.8], rows, heights=[0.32] + [0.34] * 5, size=11.5, head_size=11.5,
      align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
caption(s, 4.38, "분류기: balanced 로지스틱 회귀 → sigmoid 확률 보정 (2024) → 경보 임계값 (2025)")

# ---- Worked example -------------------------------------------------------------------------------
s = content("진행 상황", "예시: 한 투수에게 경보가 울리는 과정", "진행 상황")
apps = list(range(1, 13))
obs = [95.2, 94.9, 95.4, 95.1, 95.3, 94.8, 95.0, 95.2, 94.9, 94.4, 94.0, 93.3]
fc = [95.1, 95.1, 95.1, 95.2, 95.2, 95.2, 95.1, 95.1, 95.1, 95.1, 95.0, 94.9]
cd = CategoryChartData()
cd.categories = [str(a) for a in apps]
cd.add_series("과거 기반 예측", fc)
cd.add_series("95% 구간 상한", [f + 0.95 for f in fc])
cd.add_series("95% 구간 하한", [f - 0.95 for f in fc])
cd.add_series("실제 구속", obs)
gfr = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.3), Inches(1.06), Inches(5.3), Inches(3.3), cd)
ch = gfr.chart
ch.font.size = Pt(9.5)
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.TOP
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(9)
va = ch.value_axis
va.minimum_scale, va.maximum_scale, va.major_unit = 92.5, 96.5, 1.0
va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = RGBColor(0xE5, 0xE5, 0xE5)
va.tick_labels.font.size = Pt(9)
va.has_title = True
va.axis_title.text_frame.text = "직구 구속 (mph)"
va.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
va.axis_title.text_frame.paragraphs[0].runs[0].font.bold = False
va.format.line.fill.background()
ca = ch.category_axis
ca.has_title = True
ca.axis_title.text_frame.text = "시즌 내 등판 순서"
ca.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
ca.axis_title.text_frame.paragraphs[0].runs[0].font.bold = False
ca.tick_labels.font.size = Pt(9)
ca.format.line.color.rgb = MID
styles = [(NAVY, 1.75, False, None), (BARGRAY, 1.0, True, None), (BARGRAY, 1.0, True, None), (RED, 2.0, False, RED)]
from pptx.enum.chart import XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
for ser, (c, w, dash, mk) in zip(ch.plots[0].series, styles):
    ser.smooth = False
    ser.format.line.color.rgb = c
    ser.format.line.width = Pt(w)
    if dash:
        ser.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    if mk is None:
        ser.marker.style = XL_MARKER_STYLE.NONE
    else:
        ser.marker.style = XL_MARKER_STYLE.CIRCLE
        ser.marker.size = 5
        ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = mk
        ser.marker.format.line.color.rgb = mk
caption(s, 4.4, "가상의 투수 한 명에 대한 예시 수치 (실제 사례 아님)", x=0.3, w=5.3)
steps = [("1  축 읽기", "x = 이번 시즌 등판 순서, y = 직구 구속(mph)\n파랑 = 과거 기반 예측, 점선 = 95% 구간, 빨강 = 실제 구속"),
         ("2  비교", "12번째 등판: 예측 94.9, 실제 93.3 mph\n→ z = −3.3 (평소보다 크게 낮음)"),
         ("3  점수화", "z, z의 EWMA, 기울기로 로지스틱 모델 → 보정된 7일 내 팔 IL 위험"),
         ("4  결정", "위험이 2025년 임계값을 넘음 → **경보**: 13번째 등판 전 트레이너 확인")]
sy = 1.1
for i, (h, d) in enumerate(steps):
    hgt = 0.82 if i == 0 else 0.74
    textbox(s, 5.8, sy, 3.9, hgt, [(h, {"bold": True, "size": 11.5}), (d, {"size": 10, "space": 2})],
            border=BLACK if i == 3 else LIGHT, border_w=1.0, pad=0.08, anchor=MSO_ANCHOR.MIDDLE,
            fill=PALE if i == 3 else None)
    sy += hgt + 0.1

# ---- Results 1: baselines ---------------------------------------------------------------------
s = content("실험 결과", "(1) 2026년 평가 시즌의 기준 모델", "실험 결과")
rows = [["모델 (정답)", "양성", "TP", "FP", "정밀도", "재현율", "F1"],
        ["구속 EWMA + 기울기 (전체 IL)", "228", "30", "594", "4.8%", "13.2%", "0.070"],
        ["추세 제거 구속 + 회전 비율 (전체 IL)", "183", "27", "455", "5.6%", "14.8%", "0.081"],
        ["추세 제거 구속 (팔 IL)", "97", "16", "323", "4.7%", "16.5%", "0.073"],
        ["상태공간 칼만 (팔 IL)", "–", "17", "237", "6.7%", "–", "–"]]
table(s, 0.3, 1.12, [3.6, 1.0, 0.7, 0.8, 1.1, 1.1, 1.1], rows, heights=[0.4] + [0.44] * 4, size=12,
      head_size=12, align=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 6, highlight_rows=(3,))
bullets(s, 0.3, 3.6, 9.4, 1.2,
        ["기저율: 양성 등판은 **1.2–2.3%** 뿐인데, 경보의 적중률은 **약 5%** (**무작위의 2–4배**)",
         "칼만 기준선은 구속을 잘 예측하고(**MAE 0.61 mph**) 오경보를 줄이지만, 탐지도 함께 줄어듦"],
        size=13, spacing=10)

# ---- Results 2: adding features ----------------------------------------------------------------
s = content("실험 결과", "(2) 부정적 결과: 피처 추가는 도움이 되지 않음", "실험 결과")
caption(s, 1.02, "막대 = 각 모델이 같은 수의 부상을 탐지할 때의 2025년 오경보 수.  낮을수록 좋음.  회색 = 기준 모델",
        x=0.3, w=9.4)
bar_chart(s, 0.3, 1.3, 4.6, 2.45,
          ["구속 EWMA\n(기준)", "+ 지속성", "+ 집단\n사전 위험", "+ 키·\n평속", "+ 사전 위험\n+ 지속성"],
          [287, 261, 456, 363, 407], [BARGRAY, NAVY, RED, RED, RED],
          title="구속 EWMA 대비 (탐지 11건)", maxv=560)
bar_chart(s, 5.1, 1.3, 4.6, 2.45,
          ["상태공간\n(기준)", "투구량만", "상태공간\n+ 투구량", "+ z × 투구량"],
          [392, 577, 438, 445], [BARGRAY, RED, RED, RED],
          title="상태공간 대비 (탐지 20건)", maxv=700)
bullets(s, 0.3, 3.95, 9.4, 0.9,
        ["집단 위험, 키, 투구량, IL 이력은 모두 오경보를 **늘림**; 지속성은 2026년에 탐지도 함께 줄어듦",
         "부상 세트별 7일 모델: 어깨는 26건 중 4건 탐지, 팔꿈치는 **49건 중 1건**"],
        size=12, spacing=6)

# ---- Results 3: why so low ---------------------------------------------------------------------
s = content("실험 결과", "(3) 왜 수치가 낮은가?", "실험 결과")
textbox(s, 0.3, 1.1, 9.4, 0.46, ["**현재 점수는 구속 신호 자체가 허용하는 상한에 가깝다**"],
        size=13, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, border=BLACK, border_w=1.0, fill=PALE)
rows = [["팔 IL, 2026", "기대치*", "실제"],
        ["정밀도", "3.1–3.7%", "4.7%"],
        ["재현율", "11–13%", "16.5%"],
        ["F1", "0.05–0.06", "0.073"],
        ["전부 경보", "F1 0.024", ""]]
table(s, 0.3, 1.78, [1.6, 1.3, 1.2], rows, heights=[0.38] * 5, size=11.5, head_size=11.5)
textbox(s, 0.3, 3.8, 4.1, 0.9,
        ["*EDA에서 본 IL 직전 하락(−0.37 mph, 등판 간 잡음 약 0.76 mph ≈ 0.5 SD)이 유일한 신호이고, "
         "양성 1.2%, 등판의 4.3%에 경보를 낸다고 가정한 근사치"], size=9, color=MID)
bullets(s, 4.75, 1.85, 4.95, 2.9,
        ["**작은 신호, 드문 사건**: 유병률 1.2%에서 0.5 SD 하락이면 경보 대부분이 오경보",
         "**늦은 신호**: 하락은 주로 마지막 등판에서만 나타나고, 팔꿈치 파열은 갑자기 오는 경우가 많음 (49건 중 1건)",
         "**적은 평가 사건**: 양성 97건, 재현율 95% CI 10–25% → TP 16 vs 13은 잡음 수준",
         "**고정 피처**(키, 이력)는 위험한 시점이 아니라 위험한 선수를 고름"],
        size=12.5, spacing=12)

# ---- Next steps -------------------------------------------------------------------------------
s = content("향후 계획", "평가 계획, 남은 단계, 기대 결과", "향후 계획")
colx, colw, coly, colh = [0.3, 3.45, 6.6], 3.1, 1.2, 3.3
heads = ["평가 계획", "남은 단계", "기대 결과"]
bodies = [
    ["**선행 시간**: 첫 경보부터 IL 등재까지 일수",
     "**투수·시즌당 오경보 수**, 탐지한 부상 수",
     "**PR-AUC와 기저율 대비 배수**, 부트스트랩 신뢰구간"],
    ["더 강한 신호: **릴리스 포인트, 구속 변동성** (MEWMA)",
     "**CUSUM / ARIMAX** 잔차 모니터링",
     "ML 모델 비교; 선발 vs 불펜"],
    ["진단이 아닌 선별 목록",
     "탐지 1건당 필요한 확인 수의 trade-off 곡선",
     "부정적 결과도 한계와 함께 보고"]]
for i in range(3):
    textbox(s, colx[i], coly, colw, 0.42, [f"**{heads[i]}**"], size=13, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE, border=BLACK, border_w=1.0, fill=PALE)
    rect(s, colx[i], coly + 0.42, colw, colh - 0.42, color=BLACK, weight=1.0)
    bullets(s, colx[i] + 0.14, coly + 0.62, colw - 0.28, colh - 0.75, bodies[i], size=13.5, spacing=16)

# ---- move "Thank you" (original slide 19, now index 1) to the end ------------------------------
ids = list(sldIdLst)
thanks = ids[1]
sldIdLst.remove(thanks)
sldIdLst.append(thanks)

prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides")
