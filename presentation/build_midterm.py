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

SRC = "Pitcher_Injury_Early_Warning_Proposal_EN.pptx"
OUT = "Pitcher_Injury_Early_Warning_Midterm_EN.pptx"

BLACK = RGBColor(0, 0, 0)
GRAY = RGBColor(0x59, 0x59, 0x59)
MID = RGBColor(0x88, 0x88, 0x88)
LIGHT = RGBColor(0xD9, 0xD9, 0xD9)
BARGRAY = RGBColor(0x9C, 0xA3, 0xAF)
RED = RGBColor(0xC0, 0x50, 0x4D)
NAVY = RGBColor(0x1F, 0x3A, 0x6E)
BLUE = RGBColor(0x3B, 0x6E, 0xB5)
PALE = RGBColor(0xF2, 0xF4, 0xF8)

SECTIONS = ["Feedback", "Progress", "Results", "Next steps"]

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
# Cover (reuse proposal cover, update text)
for ph in cover_src.placeholders:
    if ph.placeholder_format.idx == 106:
        paras = ph.text_frame.paragraphs
        paras[0].runs[0].text = "Midterm progress · MLB Statcast, 2015–2026 regular seasons"
        paras[2].runs[0].text = "2026.11"

# ---- Role -------------------------------------------------------------------------------------
s = content("Role", "Team composition and contribution")
members = [("Soyeon Kim (Lead)", "2023035269"), ("Minseo Jeong", "2022059961"),
           ("Donghan Kang", "2021055178"), ("Minseok Choi", "2022005496")]
rows = [["Member", "Student number", "Role and contribution"]] + [[m, n, ""] for m, n in members]
gf = table(s, 0.46, 1.16, [1.72, 1.6, 5.98], rows, heights=[0.32, 0.88, 0.88, 0.88, 0.88],
           size=12, head_size=12, align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT],
           bold_first_col=False)
for j in range(3):  # header left aligned, plain weight like the proposal
    p = gf.table.cell(0, j).text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for r in p.runs:
        r.font.bold = False


# ---- Feedback 1 -------------------------------------------------------------------------------
s = content("Reflection on feedback", "(1) What an alert means, and what it is for", "Feedback")
rows = [["", "What it does", "Our scope"],
        ["Early warning", "Flags rising risk before the injury reaches the IL", "Yes (main)"],
        ["Detection", "Finds an injury that exists but is not yet reported", "Yes (partly)"],
        ["Explanation", "Explains why the injury happens (causal)", "No"]]
table(s, 0.3, 1.12, [1.6, 6.3, 1.5], rows, heights=[0.32, 0.34, 0.34, 0.34], size=11.5, head_size=11.5,
      align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
bx_y, bx_h = 2.86, 1.9
textbox(s, 0.3, bx_y, 4.45, bx_h,
        [("Alert = model output", {"bold": True, "size": 13}),
         ("Velocity is compared with a forecast from the pitcher's **own past**.", {"space": 8}),
         ("High risk of an **arm-IL placement within 7 days** → flag.", {"space": 6})],
        size=12, border=BLACK, pad=0.16, anchor=MSO_ANCHOR.MIDDLE)
arrow(s, 4.82, bx_y + bx_h / 2 - 0.18, 0.36, 0.36)
textbox(s, 5.25, bx_y, 4.45, bx_h,
        [("Follow-up = human decision", {"bold": True, "size": 13}),
         ("A **short check list** for trainers: exam, extra rest or lighter load.", {"space": 8}),
         ("Useful lead time: **before the next appearance** (1–5 days).", {"space": 6})],
        size=12, border=BLACK, pad=0.16, anchor=MSO_ANCHOR.MIDDLE, fill=PALE)

# ---- Feedback 2 -------------------------------------------------------------------------------
s = content("Reflection on feedback", "(2) Labels, confounders, leakage", "Feedback")
rows = [["Comment", "How we handled it"],
        ["IL date vs. onset", "IL date = day the team acted; used as the label (a stated limit)"],
        ["Injury type", "Arm IL only; shoulder · elbow · forearm modeled separately"],
        ["Workload", "7 / 28-day pitch counts, pitches today, rest days"],
        ["Individual differences", "Personal baseline per pitcher (EWMA, Kalman filter)"],
        ["Recovery", "Out of scope; past arm-IL history tested as a feature"]]
table(s, 0.3, 1.12, [2.2, 7.2], rows, heights=[0.32] + [0.36] * 5, size=11.5, head_size=11.5,
      align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT])
ty = 3.62
textbox(s, 0.3, ty, 1.6, 0.4, ["**No leakage**"], size=12, anchor=MSO_ANCHOR.MIDDLE)
segs = [("Train  2015–2023", 4.2), ("Calibrate  2024", 1.3), ("Threshold  2025", 1.3), ("Test  2026", 1.2)]
xx = 1.9
for i, (lab, w) in enumerate(segs):
    textbox(s, xx, ty, w, 0.4, [f"**{lab}**"], size=11, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            border=BLACK, border_w=1.0, fill=PALE if i == 3 else None)
    xx += w
bullets(s, 0.3, 4.24, 9.4, 0.6,
        ["Features use only data **before** each appearance; all choices made on ≤ 2025, **2026 untouched** until the final test"],
        size=12)

# ---- Progress: data + features ------------------------------------------------------------------
s = content("Progress", "Data cleaning and feature engineering", "Progress")
kpis = [("2015–2026", "regular seasons\n(was 2021–2026)"), ("2,620", "arm-related IL\nof 5,516 placements"),
        ("FF / SI / FC", "one fastball type per\npitcher-season"), ("1 row", "per pitcher ×\nappearance")]
kw = 2.2
for i, (big, small) in enumerate(kpis):
    textbox(s, 0.3 + i * (kw + 0.2), 1.1, kw, 0.92,
            [(f"**{big}**", {"size": 17, "align": PP_ALIGN.CENTER}), (small, {"size": 10.5, "color": GRAY, "align": PP_ALIGN.CENTER, "space": 2})],
            border=BLACK, border_w=1.0, pad=0.06, anchor=MSO_ANCHOR.MIDDLE)
rows = [["Group", "Feature (from data before the appearance)", "Exp"],
        ["Velocity", "% drop vs. own baseline (EWMA) · slope over last 1–3 games · consecutive drops", "1, 3, 4"],
        ["State-space", "Kalman forecast error z of velocity", "7"],
        ["Spin", "Spin ÷ velocity: drop and slope", "2"],
        ["Workload", "Pitches in last 7 / 28 days, pitches today, rest days", "8"],
        ["Pitcher", "Height × velocity group risk · past arm-IL history", "5, 9"]]
table(s, 0.3, 2.24, [1.5, 7.1, 0.8], rows, heights=[0.32] + [0.34] * 5, size=11.5, head_size=11.5,
      align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
caption(s, 4.38, "Classifier: balanced logistic regression → sigmoid calibration (2024) → alert threshold (2025)")

# ---- Worked example -------------------------------------------------------------------------------
s = content("Progress", "Worked example: how one pitcher gets an alert", "Progress")
apps = list(range(1, 13))
obs = [95.2, 94.9, 95.4, 95.1, 95.3, 94.8, 95.0, 95.2, 94.9, 94.4, 94.0, 93.3]
fc = [95.1, 95.1, 95.1, 95.2, 95.2, 95.2, 95.1, 95.1, 95.1, 95.1, 95.0, 94.9]
cd = CategoryChartData()
cd.categories = [str(a) for a in apps]
cd.add_series("Forecast from his own past", fc)
cd.add_series("Upper 95% band", [f + 0.95 for f in fc])
cd.add_series("Lower 95% band", [f - 0.95 for f in fc])
cd.add_series("Observed velocity", obs)
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
va.axis_title.text_frame.text = "Fastball velocity (mph)"
va.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
va.axis_title.text_frame.paragraphs[0].runs[0].font.bold = False
va.format.line.fill.background()
ca = ch.category_axis
ca.has_title = True
ca.axis_title.text_frame.text = "Appearance number in the season"
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
# remove the 'Lower band' legend entry duplication: keep, it's informative
caption(s, 4.4, "Illustrative numbers for one hypothetical pitcher (not a real case)", x=0.3, w=5.3)
steps = [("1  Read the axes", "x = his appearances this season; y = fastball mph.\nBlue = forecast from his past, dashed = 95% band, red = what he threw."),
         ("2  Compare", "Game 12: forecast 94.9, observed 93.3 mph\n→ z = −3.3 (far below his own normal)"),
         ("3  Score", "Logistic model on z, its EWMA and slope → calibrated 7-day arm-IL risk"),
         ("4  Decide", "Risk above the 2025 threshold → **alert**: trainer checks him before game 13")]
sy = 1.1
for i, (h, d) in enumerate(steps):
    hgt = 0.82 if i == 0 else 0.74
    textbox(s, 5.8, sy, 3.9, hgt, [(h, {"bold": True, "size": 11.5}), (d, {"size": 10, "space": 2})],
            border=BLACK if i == 3 else LIGHT, border_w=1.0, pad=0.08, anchor=MSO_ANCHOR.MIDDLE,
            fill=PALE if i == 3 else None)
    sy += hgt + 0.1


# ---- Results 1: baselines ---------------------------------------------------------------------
s = content("Results", "(1) Baselines on the 2026 test season", "Results")
rows = [["Model (target)", "Positives", "TP", "FP", "Precision", "Recall", "F1"],
        ["Velocity EWMA + slope (all IL)", "228", "30", "594", "4.8%", "13.2%", "0.070"],
        ["Detrended velocity + spin ratio (all IL)", "183", "27", "455", "5.6%", "14.8%", "0.081"],
        ["Detrended velocity (arm IL)", "97", "16", "323", "4.7%", "16.5%", "0.073"],
        ["State-space Kalman (arm IL)", "–", "17", "237", "6.7%", "–", "–"]]
table(s, 0.3, 1.12, [3.6, 1.0, 0.7, 0.8, 1.1, 1.1, 1.1], rows, heights=[0.4] + [0.44] * 4, size=12,
      head_size=12, align=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 6, highlight_rows=(3,))
bullets(s, 0.3, 3.6, 9.4, 1.2,
        ["Base rate: only **1.2–2.3%** of appearances are positive; our alerts are right **~5%** of the time (**2–4× chance**)",
         "Kalman baseline forecasts velocity well (**MAE 0.61 mph**) and cuts false alarms, but also hits"],
        size=13, spacing=10)

# ---- Results 2: adding features ----------------------------------------------------------------
s = content("Results", "(2) Negative result: extra features did not help", "Results")
caption(s, 1.02, "Bar = false alarms in 2025 when each model catches the same number of injuries.  Lower is better.  "
        "Gray = reference", x=0.3, w=9.4)
bar_chart(s, 0.3, 1.3, 4.6, 2.45,
          ["Velocity\nEWMA (ref)", "+ Persist-\nence", "+ Group\nprior", "+ Height\n& velo", "+ Prior &\npersist."],
          [287, 261, 456, 363, 407], [BARGRAY, NAVY, RED, RED, RED],
          title="vs. velocity EWMA (11 detections)", maxv=560)
bar_chart(s, 5.1, 1.3, 4.6, 2.45,
          ["State-space\n(ref)", "Workload\nonly", "State-space\n+ workload", "+ z × load"],
          [392, 577, 438, 445], [BARGRAY, RED, RED, RED],
          title="vs. state-space (20 detections)", maxv=700)
bullets(s, 0.3, 3.95, 9.4, 0.9,
        ["Group risk, height, workload and IL history all **added** false alarms; persistence also lost hits in 2026",
         "7-day models by injury set: shoulder 4 of 26 injuries caught, elbow only **1 of 49**"],
        size=12, spacing=6)

# ---- Results 3: why so low ---------------------------------------------------------------------
s = content("Results", "(3) Why are the numbers so low?", "Results")
textbox(s, 0.3, 1.1, 9.4, 0.46, ["**The scores are close to the ceiling that the velocity signal itself allows**"],
        size=13, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, border=BLACK, border_w=1.0, fill=PALE)
rows = [["Arm IL, 2026", "Expected*", "Observed"],
        ["Precision", "3.1–3.7%", "4.7%"],
        ["Recall", "11–13%", "16.5%"],
        ["F1", "0.05–0.06", "0.073"],
        ["Flag everyone", "F1 0.024", ""]]
table(s, 0.3, 1.78, [1.6, 1.3, 1.2], rows, heights=[0.38] * 5, size=11.5, head_size=11.5)
textbox(s, 0.3, 3.8, 4.1, 0.9,
        ["*If the only signal were the pre-IL drop seen in EDA (−0.37 mph vs. ~0.76 mph appearance noise ≈ 0.5 SD), "
         "with 1.2% positives and 4.3% of appearances flagged"], size=9, color=MID)
bullets(s, 4.75, 1.85, 4.95, 2.9,
        ["**Small shift, rare event**: a 0.5 SD drop at 1.2% prevalence means most flags are false",
         "**Late signal**: the drop shows mostly in the last appearance; elbow tears are often sudden (1 of 49)",
         "**Few test events**: 97 positives, recall 95% CI 10–25%, so TP 16 vs. 13 is noise",
         "**Static features** (height, history) rank risky pitchers, not risky moments"],
        size=12.5, spacing=12)

# ---- Next steps -------------------------------------------------------------------------------
s = content("Next steps", "Evaluation plan, remaining steps, expected outcomes", "Next steps")
colx, colw, coly, colh = [0.3, 3.45, 6.6], 3.1, 1.2, 3.3
heads = ["Evaluation plan", "Remaining steps", "Expected outcomes"]
bodies = [
    ["**Lead time**: days from first alert to IL",
     "**False alarms per pitcher-season**, injuries caught",
     "**PR-AUC and lift** over the base rate, bootstrap CIs"],
    ["Stronger signal: **release point, velocity variability** (MEWMA)",
     "**CUSUM / ARIMAX** residual monitors",
     "ML comparison; starter vs. reliever"],
    ["A screening list, not a diagnosis",
     "A trade-off curve: checks per injury caught",
     "Negative results reported with their limits"]]
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
