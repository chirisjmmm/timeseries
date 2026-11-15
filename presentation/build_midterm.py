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

# ---- Contents ---------------------------------------------------------------------------------
s = content("Contents", "Outline")
toc = [("01", "Team roles", "Who did what"),
       ("02", "Feedback", "What an alert means, IL dates, leakage, a worked example"),
       ("03", "Progress", "Data cleaning, features, modeling pipeline"),
       ("04", "Results", "Baselines, added features, negative results and why"),
       ("05", "Next steps", "Evaluation plan, remaining models, expected outcomes")]
y0, gapy = 1.12, 0.8
for i, (n, h, d) in enumerate(toc):
    y = y0 + i * gapy
    t = textbox(s, 0.66, y, 8.6, 0.42, [f"**{n}**"], size=20, anchor=MSO_ANCHOR.MIDDLE)
    t2 = s.shapes.add_textbox(Inches(1.42), Inches(y), Inches(7.9), Inches(0.42))
    tf = t2.text_frame; tf.margin_left = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    from pptx.enum.text import MSO_AUTO_SIZE
    tf.auto_size = MSO_AUTO_SIZE.NONE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    fill_para(p, f"**{h}**", size=20)
    fill_para(p, f"   {d}", size=13, color=GRAY)
    hline(s, 0.66, y + 0.58, 8.74, color=RGBColor(0xBF, 0xBF, 0xBF), weight=1)

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

# ---- Feedback 1: what an alert means -------------------------------------------------------------
s = content("Reflection on feedback", "(1) What an alert means, and what it is for", "Feedback")
textbox(s, 0.3, 1.06, 9.4, 0.44,
        ["**Feedback**  State exactly what an alert means and what action it supports; keep early warning, "
         "detection and explanation distinct; say how much warning time is useful"],
        size=11.5, border=BLACK, pad=0.1, anchor=MSO_ANCHOR.MIDDLE)
rows = [["", "What it does", "Assumes", "Our scope"],
        ["Early warning", "Flags a rising risk before damage shows on the IL", "Prevention is possible", "Yes (main)"],
        ["Detection", "Finds an injury that already exists but is not yet reported", "Worsening can be stopped", "Yes (partly)"],
        ["Explanation", "Explains why the injury happens", "Causes can be identified", "No"]]
table(s, 0.3, 1.66, [1.35, 3.6, 2.15, 1.1], rows, heights=[0.3, 0.32, 0.32, 0.32], size=10.5,
      head_size=11, align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
# right-hand: alert definition + action, two boxes
bx_y, bx_h = 3.12, 1.62
textbox(s, 0.3, bx_y, 4.6, bx_h,
        [("Alert = model output", {"bold": True, "size": 12.5}),
         ("After each appearance, the pitcher's velocity is compared with a forecast from **his own past**.", {"space": 5}),
         ("If the calibrated risk of an **arm-IL placement in the next 1–7 days** passes a threshold fixed on 2025, he is flagged.", {"space": 4})],
        size=11, border=BLACK, pad=0.12)
textbox(s, 5.1, bx_y, 4.6, bx_h,
        [("Follow-up = human decision", {"bold": True, "size": 12.5}),
         ("Trainers cannot examine every pitcher every day: the alert gives a **short check list**.", {"space": 5}),
         ("Possible actions: a check-up, an extra rest day, a lighter workload, or a medical exam.", {"space": 4}),
         ("Useful lead time: **before the next appearance** (about 1–5 days), so one decision fits in.", {"space": 4})],
        size=11, border=BLACK, pad=0.12)
arrow(s, 4.93, bx_y + bx_h / 2 - 0.16, 0.16, 0.32)

# ---- Feedback 2: IL dates, injury type, leakage --------------------------------------------------
s = content("Reflection on feedback", "(2) Label timing, confounders, leakage", "Feedback")
rows = [["Comment", "How we handled it"],
        ["IL date vs. onset", "IL date = the day the team recognized the injury and acted. Treated as the label, a stated limit;\nappearances on the IL date itself are dropped"],
        ["Injury type", "Arm-only target; three sets: shoulder · elbow/UCL · forearm/biceps/triceps (separate models)"],
        ["Workload", "7- and 28-day pitch counts, pitches in the latest appearance, rest days (Exp 8)"],
        ["Recovery", "Out of scope (we predict, not rehabilitate); past arm-IL count and days since last IL (Exp 9)"],
        ["Individual differences", "Personal baseline per pitcher (EWMA, Kalman filter); height × velocity group risk (Exp 5)"],
        ["Test-period leakage", "Time-ordered split; every feature uses only data strictly before the appearance;\nthresholds chosen on 2025, never on 2026"]]
table(s, 0.3, 1.1, [1.95, 7.45], rows, heights=[0.3, 0.5, 0.34, 0.34, 0.34, 0.34, 0.5], size=10.5,
      head_size=11, align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT])
# split timeline
ty = 4.0
textbox(s, 0.3, ty, 1.55, 0.34, ["**Time-ordered split**"], size=11, anchor=MSO_ANCHOR.MIDDLE)
segs = [("Train  2015–2023", 4.3), ("Calibrate  2024", 1.25), ("Threshold  2025", 1.25), ("Test  2026", 1.2)]
xx = 1.9
for i, (lab, w) in enumerate(segs):
    textbox(s, xx, ty, w, 0.34, [f"**{lab}**"], size=10.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            border=BLACK, border_w=1.0, fill=PALE if i == 3 else None)
    xx += w
bullets(s, 0.3, 4.48, 9.4, 0.45,
        ["Honest caveat: 2026 was looked at during EDA, so it is reported as a **held-out test, not an unseen one**; "
         "two early leaks (pitch-type table, league trend fitted on 2015–2023) were found and flagged"], size=10.5)

# ---- Progress: data cleaning ----------------------------------------------------------------------
s = content("Progress", "Data cleaning: what changed since the proposal", "Progress")
rows = [["Item", "Proposal", "Now"],
        ["Period", "2021–2026 (6 seasons)", "2015–2026 regular seasons (12): more injuries to learn from"],
        ["IL events", "3,340 placements (1,683 arm)", "5,516 new placements (player × date), **2,620 arm-related**"],
        ["Pitch type", "Four-seam only", "One of FF / SI / FC per pitcher-season, picked from the previous season"],
        ["Exclusions", "–", "Double-headers, appearance on the IL date, short history, incomplete tracking"],
        ["Body data", "–", "Height and weight for 2,778 players (MLB API, current values)"],
        ["League trend", "Observed +1.08 mph", "Velocity trend by hand: LHP +0.097, RHP +0.140 mph per year (2015–2023)"]]
table(s, 0.3, 1.1, [1.45, 2.4, 5.55], rows, heights=[0.3] + [0.36] * 6, size=10.5, head_size=11,
      align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT])
bullets(s, 0.3, 3.86, 9.4, 1.0,
        ["Pitch-type rule: **FF** if the pitcher threw ≥ 1 four-seamer last season; otherwise his most-used SI/FC; FF if he has no history",
         "Non-arm and unlabeled IL events are kept in the event file, so later appearances can be **censored** when another injury comes first",
         "EDA on height and season velocity (2,423 pitchers, 9,278 pitcher-seasons): K-means groups had similar IL rates (**37.6 / 41.7 / 38.3%**)"],
        size=11)

# ---- Progress: feature engineering --------------------------------------------------------------
s = content("Progress", "Feature engineering: one row per pitcher × appearance", "Progress")
rows = [["Group", "Feature", "Definition (all use data before the prediction time)", "Exp"],
        ["Velocity", "Drop EWMA", "EWMA (α = 0.3) of % drop vs. personal baseline: first 3 games last season, then this season so far", "1"],
        ["", "Drop slope", "Velocity slope over the last 1 / 2 / 3 appearances (also on league-detrended velocity)", "1, 3"],
        ["", "Persistence", "Drops among the last two changes; consecutive drops (0–2)", "4"],
        ["", "State-space z", "Standardized forecast error from a per-pitcher local linear trend Kalman filter", "7"],
        ["Spin", "Spin / velocity", "rpm ÷ mph per appearance; its drop EWMA and slope", "2"],
        ["Workload", "Load · rest", "Pitches in last 7 / 28 days, pitches today, rest days; z × 7-day load", "8"],
        ["Pitcher", "Group prior", "Past arm-IL rate of the height (5 cm) × velocity (3 mph) cell, shrunk to the mean", "5"],
        ["", "IL history", "Past arm-IL count (all / same set), days since the last one, no-history flag", "9"]]
table(s, 0.3, 1.1, [1.05, 1.45, 6.3, 0.6], rows, heights=[0.3] + [0.385] * 8, size=10.5, head_size=11,
      align=[PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
caption(s, 4.62, "Exp = experiment number in the results; the same pitcher's appearances form an irregularly spaced series")

# ---- Progress: modeling pipeline ---------------------------------------------------------------
s = content("Progress", "Initial modeling: from a personal baseline to an alert", "Progress")
boxes = [("Appearance series", "Per-pitcher velocity, spin, load by appearance index"),
         ("Personal baseline", "EWMA of the drop\nor Kalman forecast\n(local linear trend)"),
         ("Classifier", "Balanced logistic regression on baseline features"),
         ("Calibrate · threshold", "Sigmoid fit on 2024;\nthreshold set on 2025\n(best F1 / equal detections)"),
         ("Alert", "Flag the pitcher\nfor a check")]
bw, by, bh, gap = 1.56, 1.2, 1.9, 0.4
for i, (h, d) in enumerate(boxes):
    bx = 0.3 + i * (bw + gap)
    textbox(s, bx, by, bw, bh, [(h, {"bold": True, "size": 12.5}), (d, {"size": 10.5, "space": 6})],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, border=BLACK, border_w=1.0, pad=0.08,
            fill=PALE if i == 4 else None)
    if i < 4:
        arrow(s, bx + bw + 0.06, by + bh / 2 - 0.14, 0.28, 0.28)
rows = [["Target", "Positive appearance", "Used in"],
        ["Next appearance", "First IL placement before the pitcher's next appearance (all IL, then arm-only)", "Exp 1–5, 7–8"],
        ["7-day, by injury set", "First arm IL 1–7 days after the appearance; one model per set (shoulder / elbow / forearm)", "Exp 6, 9"]]
table(s, 0.3, 3.45, [1.75, 6.25, 1.4], rows, heights=[0.3, 0.4, 0.4], size=10.5, head_size=11,
      align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
caption(s, 4.72, "Metrics: TP / FP / FN / TN at the chosen threshold, F1, and false alarms at an equal number of detections; "
        "for 7-day targets also unique injuries detected", x=0.3, w=9.4)

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
         ("3  Score", "z, its EWMA and recent slope → calibrated risk of an arm IL in the next 7 days"),
         ("4  Decide", "Risk above the 2025 threshold → **alert**: trainer checks him before game 13")]
sy = 1.1
for i, (h, d) in enumerate(steps):
    hgt = 0.82 if i == 0 else 0.74
    textbox(s, 5.8, sy, 3.9, hgt, [(h, {"bold": True, "size": 11.5}), (d, {"size": 10, "space": 2})],
            border=BLACK if i == 3 else LIGHT, border_w=1.0, pad=0.08, anchor=MSO_ANCHOR.MIDDLE,
            fill=PALE if i == 3 else None)
    sy += hgt + 0.1

# ---- Results 1: baselines -------------------------------------------------------------------------
s = content("Results", "(1) Baselines on the 2026 test season", "Results")
rows = [["Model (target)", "Appear-\nances", "Positives", "TP", "FP", "FN", "Precision", "Recall", "F1"],
        ["Velocity EWMA + slope n=2 (all IL)", "11,287", "228", "30", "594", "198", "4.8%", "13.2%", "0.070"],
        ["+ spin/velocity ratio, n=3 (all IL)", "11,257", "228", "–", "–", "–", "–", "–", "0.071"],
        ["Detrended velocity + ratio, n=1 (all IL)", "7,938", "183", "27", "455", "156", "5.6%", "14.8%", "0.081"],
        ["Detrended velocity, n=3 (arm IL)", "7,852", "97", "16", "323", "81", "4.7%", "16.5%", "0.073"],
        ["State-space Kalman (arm IL)", "matched", "–", "17", "237", "–", "6.7%", "–", "–"],
        ["  vs. EWMA on that sample", "", "", "21", "334", "", "5.9%", "", ""]]
table(s, 0.3, 1.1, [3.05, 0.95, 0.85, 0.55, 0.6, 0.55, 0.95, 0.85, 1.05], rows,
      heights=[0.4] + [0.32] * 6, size=10.5, head_size=10.5,
      align=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 8, highlight_rows=(4,))
bullets(s, 0.3, 3.75, 9.4, 1.1,
        ["How to read it: a random alert would be right as often as the base rate (**1.2–2.3%** of appearances); "
         "our alerts are right **4.7–6.7%** of the time, about **2–4× chance**",
         "Best arm-IL model catches **16 of 97** injured appearances while raising 323 false alarms",
         "Kalman baseline: velocity forecast **MAE 0.61 mph**, 95% band covers **92.9%**; fewer false alarms than EWMA, but also fewer hits"],
        size=11)

# ---- Results 2: adding features (negative) -------------------------------------------------------
s = content("Results", "(2) Adding features: does it cut false alarms?", "Results")
caption(s, 1.02, "Bar = false alarms in 2025 when each model catches the same number of injuries.  Lower is better.  "
        "Gray = reference model", x=0.3, w=9.4)
bar_chart(s, 0.3, 1.3, 4.6, 2.35,
          ["Velocity\nEWMA (ref)", "+ Persist-\nence", "+ Group\nprior", "+ Height\n& velo", "+ Prior &\npersist."],
          [287, 261, 456, 363, 407], [BARGRAY, NAVY, RED, RED, RED],
          title="vs. velocity EWMA (11 detections)", maxv=560)
bar_chart(s, 5.1, 1.3, 4.6, 2.35,
          ["State-space\n(ref)", "Workload\nonly", "State-space\n+ workload", "+ z × load"],
          [392, 577, 438, 445], [BARGRAY, RED, RED, RED],
          title="vs. state-space (20 detections)", maxv=700)
bullets(s, 0.3, 3.82, 9.4, 1.05,
        ["Only **persistence** lowered 2025 false alarms (287 → 261), but in 2026 it also lost hits (TP 16 → 13, FP 313 → 279)",
         "Group prior, height/velocity, workload and their combinations all **added** false alarms",
         "Same in 2026: group prior TP 17 · FP 488 vs. reference TP 16 · FP 313"],
        size=11)

# ---- Results 3: 7-day set models ------------------------------------------------------------------
s = content("Results", "(3) 7-day warning by injury set (2026)", "Results")
rows = [["Injury set", "Model (chosen on 2025)", "TP", "FP", "FN", "Precision", "Recall", "Injuries caught"],
        ["Shoulder · rotator cuff", "Velocity + ratio (Exp 6)", "4", "108", "31", "3.6%", "11.4%", "4 / 26"],
        ["", "State-space A (Exp 9)", "5", "145", "30", "3.3%", "14.3%", "–"],
        ["Elbow · UCL · Tommy John", "Velocity + ratio (Exp 6)", "2", "96", "77", "2.0%", "2.5%", "1 / 49"],
        ["", "State-space + load C (Exp 9)", "4", "579", "75", "0.7%", "5.1%", "–"],
        ["Forearm · biceps · triceps", "Velocity + ratio (Exp 6)", "3", "400", "32", "0.7%", "8.6%", "3 / 21"],
        ["", "State-space A (Exp 9)", "5", "393", "30", "1.3%", "14.3%", "–"]]
table(s, 0.3, 1.1, [2.1, 2.35, 0.55, 0.6, 0.55, 1.0, 0.95, 1.3], rows, heights=[0.32] + [0.33] * 6,
      size=10.5, head_size=10.5, align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 6)
bullets(s, 0.3, 3.62, 9.4, 1.25,
        ["Shoulder is the most detectable set; elbow injuries, the most costly, are almost **never caught** (1 of 49)",
         "Adding past arm-IL history (B, D) **raised** 2025 false alarms in every set (e.g. shoulder 140 → 398)",
         "Elbow C looked better on 2025 (FP 561 → 543) but the gain **did not hold** on 2026"],
        size=11)

# ---- Negative results: why ----------------------------------------------------------------------
s = content("Results", "Why the negative results? Our current reading", "Results")
cards = [("Rare, noisy label", "Only 1–2% of appearances are positive, and the IL date is when the team acted, not the onset: "
          "a few days of shift moves an event in or out of the window"),
         ("Velocity drop is not specific", "Fatigue, role changes, weather or minor soreness also lower velocity, so many drops never end on the IL"),
         ("Group risk is flat", "Height × velocity cells differ little (EDA: 37.6–41.7% IL rate), so the prior adds noise, not signal"),
         ("Workload is already managed", "Teams cut the load of tired pitchers, so load and velocity carry overlapping, confounded information"),
         ("Too few validation hits", "Thresholds are set with only 11–20 detections in 2025, so small gains (elbow C) do not survive a new season")]
cw, ch_, gx = 2.98, 1.5, 0.23
for i, (h, d) in enumerate(cards):
    r, c = divmod(i, 3)
    if r == 1:
        cx = 0.3 + (c + 0.5) * (cw + gx)
    else:
        cx = 0.3 + c * (cw + gx)
    cy = 1.1 + r * (ch_ + 0.18)
    textbox(s, cx, cy, cw, ch_, [(h, {"bold": True, "size": 13}), (d, {"size": 11.5, "space": 6})],
            border=BLACK, border_w=1.0, pad=0.14, anchor=MSO_ANCHOR.MIDDLE)
caption(s, 4.5, "What did work: the Kalman baseline forecasts velocity well and cuts false alarms; the bottleneck is turning anomalies into injury-specific alerts")

# ---- Next steps -----------------------------------------------------------------------------------
s = content("Next steps", "Evaluation plan, remaining steps, expected outcomes", "Next steps")
colx, colw, coly, colh = [0.3, 3.45, 6.6], 3.1, 1.1, 3.48
heads = ["Evaluation plan", "Remaining steps", "Expected outcomes"]
bodies = [
    ["**Lead time**: days from the first alert to the IL date (median, distribution)",
     "**False alarms per pitcher-season** and injuries caught (event-level recall)",
     "Appearance-level **AUPRC** with pitcher-cluster bootstrap CIs",
     "Rolling-origin check (2022–2025) so one season does not decide the model; 2026 reported once"],
    ["Residual monitors from the proposal: **CUSUM, ARIMAX** with rest days and load",
     "**MEWMA** on velocity + release point + spin together",
     "ML comparison (gradient boosting) on the same features and split",
     "Pitcher types (starter / reliever, hard throwers: RQ4) and a **real** worked example"],
    ["A screening list, not a diagnosis: precision around **2–4× the base rate**",
     "A clear trade-off curve: how many checks per injury caught",
     "Honest answer to RQ3: workload and history did not help so far",
     "Report negative results as findings, with the label limits"]]
for i in range(3):
    textbox(s, colx[i], coly, colw, 0.38, [f"**{heads[i]}**"], size=12.5, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE, border=BLACK, border_w=1.0, fill=PALE)
    rect(s, colx[i], coly + 0.38, colw, colh - 0.38, color=BLACK, weight=1.0)
    bullets(s, colx[i] + 0.12, coly + 0.52, colw - 0.24, colh - 0.6, bodies[i], size=11.5, spacing=9)

# ---- move "Thank you" (original slide 19, now index 1) to the end ------------------------------
ids = list(sldIdLst)
thanks = ids[1]
sldIdLst.remove(thanks)
sldIdLst.append(thanks)

prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides")
