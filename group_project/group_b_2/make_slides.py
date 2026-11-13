#!/usr/bin/env python3
"""Build the Study B QC comparison deck (group_b_2) as a PowerPoint file."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------- palette ----------
NAVY   = RGBColor(0x10, 0x23, 0x3F)   # dark background
NAVY2  = RGBColor(0x1B, 0x35, 0x5C)   # dark card
INK    = RGBColor(0x22, 0x30, 0x47)   # body text on light
MUTED  = RGBColor(0x5C, 0x6B, 0x7A)
PAPER  = RGBColor(0xF7, 0xF9, 0xFC)   # light slide bg
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
TEAL   = RGBColor(0x17, 0xBE, 0xBB)
CORAL  = RGBColor(0xE7, 0x6F, 0x51)
AMBER  = RGBColor(0xE9, 0xC4, 0x6A)
GREEN  = RGBColor(0x2A, 0x9D, 0x8F)
TEAL_T  = RGBColor(0xDD, 0xF3, 0xF2)  # tints for table cells
GREEN_T = RGBColor(0xDE, 0xF2, 0xEC)
AMBER_T = RGBColor(0xFB, 0xF0, 0xD9)
CORAL_T = RGBColor(0xFB, 0xE3, 0xDB)
LINE    = RGBColor(0xD8, 0xE0, 0xEA)
FAINT   = RGBColor(0xBF, 0xD0, 0xE4)  # light text on navy

FONT = "Segoe UI"
FONT_B = "Segoe UI Semibold"

SW, SH = Inches(13.333), Inches(7.5)

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
BLANK = prs.slide_layouts[6]

# ---------- helpers ----------
def slide_new(dark=False):
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY if dark else PAPER
    bg.line.fill.background()
    bg.shadow.inherit = False
    return s

def box(s, x, y, w, h, fill=None, line=None, round_=False, radius=0.08):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if round_:
        try: shp.adjustments[0] = radius
        except Exception: pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp

def txt(s, x, y, w, h, runs, size=16, color=INK, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, leading=1.0, space_after=4, font=FONT):
    """runs: str, or list of paragraphs; a paragraph is str or list of (text, dict) runs."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(runs, str):
        runs = [runs]
    elif runs and isinstance(runs[0], tuple):  # single paragraph of runs
        runs = [runs]
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = leading
        p.space_after = Pt(space_after)
        if isinstance(para, str):
            para = [(para, {})]
        for text, style in para:
            r = p.add_run(); r.text = text
            f = r.font
            f.name = style.get("font", font)
            f.size = Pt(style.get("size", size))
            f.bold = style.get("bold", bold)
            f.italic = style.get("italic", False)
            f.color.rgb = style.get("color", color)
    return tb

def header(s, kicker, title, num=None):
    box(s, 0, 0, 13.333, 0.14, fill=TEAL)
    box(s, 9.0, 0, 4.333, 0.14, fill=CORAL)
    txt(s, 0.55, 0.32, 11.5, 0.32, kicker.upper(), size=12, color=TEAL, bold=True)
    txt(s, 0.55, 0.60, 12.2, 0.85, title, size=28, color=NAVY, bold=True, font=FONT_B)
    if num:
        c = box(s, 12.45, 0.42, 0.5, 0.5, fill=NAVY, round_=True, radius=0.5)
        tf = c.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = str(num)
        r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = WHITE
        r.font.name = FONT_B
    footer(s, num)

def footer(s, num):
    txt(s, 0.55, 7.08, 9.0, 0.3, "Study B · group_b_contraception · QC compared",
        size=10, color=MUTED)
    if num:
        txt(s, 12.3, 7.08, 0.5, 0.3, f"{num:02d}", size=10, color=MUTED, align=PP_ALIGN.RIGHT)

def pic(s, path, x, y, w, h):
    """place image centred inside box, preserving 1344x960 aspect."""
    ratio = 1344 / 960
    if w / h > ratio:
        ih = h; iw = h * ratio
    else:
        iw = w; ih = w / ratio
    ix = x + (w - iw) / 2; iy = y + (h - ih) / 2
    return s.shapes.add_picture(path, Inches(ix), Inches(iy), Inches(iw), Inches(ih))

def cell_set(cell, text, fill=None, color=INK, size=12, bold=False, align=PP_ALIGN.CENTER):
    cell.margin_left = cell.margin_right = Inches(0.08)
    cell.margin_top = cell.margin_bottom = Inches(0.03)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is not None:
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
    tf = cell.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = color; r.font.name = FONT

# ============================================================
# SLIDE 1 — title
# ============================================================
s = slide_new(dark=True)
box(s, 0, 0, 0.22, 7.5, fill=TEAL)
box(s, 0.22, 0, 0.06, 7.5, fill=CORAL)
txt(s, 0.9, 1.05, 11.6, 0.4, "DATA SCIENCE FOR BIOLOGY · GROUP B · WEEK 2", size=13, color=TEAL, bold=True)
txt(s, 0.9, 1.55, 11.8, 1.9, [
    [("Quality control, compared", {"size": 44, "bold": True, "color": WHITE, "font": FONT_B})],
    [("Three QC reports for Study B (group_b_contraception), read side by side",
      {"size": 20, "color": FAINT})],
])
chips = [("93", "samples, one visit each"), ("220", "species"),
         ("14,226", "genes"), ("3", "contraceptive arms")]
cx = 0.9
for v, lab in chips:
    box(s, cx, 3.55, 2.6, 1.05, fill=NAVY2, round_=True, radius=0.15)
    txt(s, cx + 0.25, 3.68, 2.2, 0.45, v, size=24, color=TEAL, bold=True, font=FONT_B)
    txt(s, cx + 0.25, 4.14, 2.2, 0.35, lab, size=11, color=FAINT)
    cx += 2.85
txt(s, 0.9, 5.35, 11.5, 0.4, "Reports compared:  01-qc Priv  ·  01-qc Samwl  ·  StudyB_QC_Fathy",
    size=15, color=WHITE, bold=True)
txt(s, 0.9, 5.85, 11.5, 0.6,
    "Five checks asked of every report: sample IDs · library sizes in both count tables · "
    "missing metadata · batch × arm · baseline balance.", size=13, color=RGBColor(0x9F, 0xB4, 0xCC))

# ============================================================
# SLIDE 2 — the three reports at a glance
# ============================================================
s = slide_new()
header(s, "The line-up", "Three reports, three styles, one dataset", 2)
cards = [
    ("01-qc Priv", CORAL, "Quarto, checks 1–7 + verdict",
     ["IDs: helper-guaranteed + cytokines checked by hand",
      "Extra: bv ↔ Nugent consistency audit (all 93 agree)",
      "Extra: Bray–Curtis PCoA outliers, by arm and batch",
      "Extra: caught B018 & B029 — all 15 cytokines NA",
      "Not done: gene-table library sizes, baseline table"]),
    ("01-qc Samwl", TEAL, "Quarto, section per question",
     ["IDs: same set AND same order, all four tables",
      "Library sizes for BOTH count tables (10k / 1M flags)",
      "Baseline: pH, Nugent, BV, STI per arm",
      "Shannon + Kruskal–Wallis (p = 0.47)",
      "Not done: ordination outlier check"]),
    ("StudyB_QC_Fathy", AMBER, "Hand-written HTML summary",
     ["Choices made + batch × arm cross-tab",
      "Library range quoted (min 5,247 — “sufficient”)",
      "Alpha diversity described, no test reported",
      "Recommends poscounts size factors",
      "Not shown: ID checks, gene table, baseline"]),
]
cx = 0.55
for name, col, style, items in cards:
    box(s, cx, 1.75, 4.05, 4.9, fill=WHITE, line=LINE, round_=True, radius=0.05)
    box(s, cx, 1.75, 4.05, 0.14, fill=col)
    txt(s, cx + 0.25, 2.02, 3.6, 0.4, name, size=17, color=NAVY, bold=True, font=FONT_B)
    txt(s, cx + 0.25, 2.42, 3.6, 0.3, [(style, {"italic": True})], size=11, color=MUTED)
    paras = [[("▸ ", {"color": col, "bold": True}), (it, {})] for it in items]
    txt(s, cx + 0.25, 2.85, 3.6, 3.6, paras, size=11.5, space_after=7, leading=1.05)
    cx += 4.15
footer(s, 2)

# ============================================================
# SLIDE 3 — scorecard
# ============================================================
s = slide_new()
header(s, "The five checks", "Who checked what — the scorecard", 3)
rows = [
    ("Sample IDs match across all tables",
     ("✓", GREEN, "helper + cytokines by hand"),
     ("✓✓", GREEN, "set AND order, 4 tables"),
     ("✗", CORAL, "not shown — helper trusted")),
    ("Library sizes, both count tables",
     ("✓/✗", AMBER, "microbiome only"),
     ("✓", GREEN, "species + genes, 10k / 1M"),
     ("✗", CORAL, "microbiome only")),
    ("Which samples have very few reads",
     ("✓", GREEN, "B006 flagged at <10k"),
     ("✓", GREEN, "B006 flagged at <10k"),
     ("✗", CORAL, "min quoted, no flag")),
    ("Missing metadata",
     ("✓", GREEN, "0 NA + cytokine NAs found"),
     ("✓", GREEN, "0 NA, all 8 columns"),
     ("✓", GREEN, "0 NA stated")),
    ("rnaseq_batch × arm",
     ("✓", GREEN, "with % in B1 (51/50/50)"),
     ("✓", GREEN, "row % + Study A aside"),
     ("✓", GREEN, "cross-tab shown")),
    ("Arms differed at baseline?",
     ("✗", CORAL, "not attempted"),
     ("✓", GREEN, "pH, Nugent, BV, STI"),
     ("✗", CORAL, "not attempted")),
]
tblshape = s.shapes.add_table(len(rows) + 1, 4, Inches(0.55), Inches(1.7),
                              Inches(12.25), Inches(4.9))
tbl = tblshape.table
tbl.columns[0].width = Inches(4.15)
for i in range(1, 4):
    tbl.columns[i].width = Inches(2.7)
for j, htx in enumerate(["Check", "Priv", "Samwl", "Fathy"]):
    cell_set(tbl.cell(0, j), htx, fill=NAVY, color=WHITE, size=14, bold=True,
             align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
for i, (chk, *cols) in enumerate(rows, start=1):
    cell_set(tbl.cell(i, 0), chk, fill=WHITE, color=INK, size=12.5, bold=True,
             align=PP_ALIGN.LEFT)
    for j, (sym, col, note) in enumerate(cols, start=1):
        tint = {GREEN: GREEN_T, AMBER: AMBER_T, CORAL: CORAL_T}[col]
        c = tbl.cell(i, j)
        cell_set(c, "", fill=tint)
        tf = c.text_frame
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = sym + "  "
        r.font.bold = True; r.font.size = Pt(14)
        r.font.color.rgb = col; r.font.name = FONT_B
        r2 = p.add_run(); r2.text = note
        r2.font.size = Pt(10.5); r2.font.color.rgb = INK; r2.font.name = FONT
txt(s, 0.55, 6.75, 12.2, 0.4,
    [("Verdict: ", {"bold": True, "color": NAVY}),
     ("Samwl is the most complete report; Fathy is the thinnest; Priv is the only one that hunted outliers. "
      "No report caught an error the others would call fatal.", {})],
    size=12.5, color=INK)

# ============================================================
# SLIDE 4 — IDs + missing metadata details
# ============================================================
s = slide_new()
header(s, "Checks 1 & 3 — passed everywhere", "IDs match, metadata complete — but two extra catches", 4)
box(s, 0.55, 1.7, 6.0, 2.5, fill=WHITE, line=LINE, round_=True, radius=0.05)
txt(s, 0.8, 1.9, 5.5, 0.4, [("Sample IDs", {"bold": True, "size": 16, "color": NAVY, "font": FONT_B})])
txt(s, 0.8, 2.35, 5.55, 1.8, [
    [("All three reports: pass. ", {"bold": True}),
     ("read_counts() stops on any mismatch, and Priv additionally verified the cytokine file by hand "
      "(setequal + match → TRUE).", {})],
    [("Samwl went furthest: ", {"bold": True}),
     ("identical() checks the column ORDER of species, genes and cytokines against the metadata — "
      "what DESeq2 silently relies on.", {})],
], size=12.5, space_after=8, leading=1.05)
box(s, 6.8, 1.7, 6.0, 2.5, fill=WHITE, line=LINE, round_=True, radius=0.05)
txt(s, 7.05, 1.9, 5.5, 0.4, [("Missing metadata", {"bold": True, "size": 16, "color": NAVY, "font": FONT_B})])
txt(s, 7.05, 2.35, 5.55, 1.8, [
    [("All three reports: 0 NA ", {"bold": True}),
     ("across all 8 metadata columns (sample_id, participant_id, arm, bv, nugent, ph, any_sti, rnaseq_batch).", {})],
    [("Priv's extra catch: ", {"bold": True, "color": CORAL}),
     ("B018 and B029 have NA for all 15 cytokines — fine for microbiome/RNA-seq, fatal for cytokine work.", {})],
], size=12.5, space_after=8, leading=1.05)
box(s, 0.55, 4.45, 12.25, 2.25, fill=TEAL_T, round_=True, radius=0.05)
txt(s, 0.8, 4.62, 11.8, 0.35, [("Only Priv checked — does the metadata contradict itself?",
                                {"bold": True, "size": 15, "color": NAVY, "font": FONT_B})])
txt(s, 0.8, 5.05, 11.9, 1.5, [
    [("bv vs Nugent: ", {"bold": True}),
     ("Nugent 0–3 should mean “negative”, 4–6 “intermediate”, 7–10 “positive”. Cross-tab: 51/51, 3/3, 39/39 agree — "
      "the two clinical columns never disagree. Priv's choice of cut points (0–3 / 4–6 / 7–10) is a listed choice; "
      "alternative was trusting the bv column.", {})],
], size=12.5, leading=1.1)
footer(s, 4)

# ============================================================
# SLIDE 5 — library sizes
# ============================================================
s = slide_new()
header(s, "Check 2 — the only gap between reports", "Library sizes: Samwl read both tables, the others read one", 5)
txt(s, 0.55, 1.65, 5.6, 2.2, [
    [("Species table: ", {"bold": True, "color": NAVY}),
     ("min 5,247 · median 29,212 · max 77,353 reads.", {})],
    [("Gene table: ", {"bold": True, "color": NAVY}),
     ("min 5.86M · median 12.86M · max 24.0M reads.", {})],
    [("Samwl's thresholds: ", {"bold": True}),
     ("<10k (species) and <1M (genes). Result: only B006 flagged; gene table clean (min is ~6× the threshold).", {})],
], size=13, space_after=9, leading=1.1)
box(s, 0.55, 4.05, 5.6, 2.6, fill=CORAL_T, round_=True, radius=0.06)
txt(s, 0.8, 4.25, 5.1, 0.4, [("The one small sample: B006", {"bold": True, "size": 15, "color": NAVY, "font": FONT_B})])
txt(s, 0.8, 4.7, 5.15, 1.85, [
    [("5,247 reads ≈ 18% of the median — small but not pathological (the practice data had a 438-read sample).", {})],
    [("Priv and Samwl both keep it, pending one question → open-questions slide.", {})],
    [("Fathy quotes the same minimum but calls it “sufficient” without flagging the sample.", {})],
], size=12, space_after=7, leading=1.05)
pic(s, "slides_assets/samwl_1.png", 6.45, 1.55, 6.35, 3.2)
pic(s, "slides_assets/priv_1.png", 6.45, 4.85, 6.35, 1.95)
txt(s, 6.45, 6.9, 6.3, 0.25, "Top: both tables (Samwl) · bottom: 10 smallest libraries (Priv)",
    size=10, color=MUTED)
footer(s, 5)

# ============================================================
# SLIDE 6 — batch × arm
# ============================================================
s = slide_new()
header(s, "Check 4 — unanimous", "rnaseq_batch × arm: balanced, all three agree", 6)
tblshape = s.shapes.add_table(4, 4, Inches(0.55), Inches(1.8), Inches(6.2), Inches(2.4))
tbl = tblshape.table
tbl.columns[0].width = Inches(2.0)
for i in range(1, 4): tbl.columns[i].width = Inches(1.4)
for j, htx in enumerate(["arm", "B1", "B2", "% in B1"]):
    cell_set(tbl.cell(0, j), htx, fill=NAVY, color=WHITE, size=13, bold=True,
             align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
data = [("injectable", 20, 19, "51.3%"), ("pill", 14, 14, "50.0%"), ("ring", 13, 13, "50.0%")]
for i, row in enumerate(data, start=1):
    for j, v in enumerate(row):
        cell_set(tbl.cell(i, j), str(v),
                 fill=WHITE if i % 2 else PAPER, size=12.5,
                 bold=(j == 0 or j == 3),
                 align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
txt(s, 0.55, 4.55, 6.2, 2.0, [
    [("Every arm sits almost exactly half B1, half B2. ", {"bold": True}),
     ("A batch effect cannot easily masquerade as a method effect.", {})],
    [("All three reports cross-tabulated this — the one check with perfect three-way agreement.", {})],
], size=13, space_after=8, leading=1.1)
box(s, 7.1, 1.8, 5.7, 1.75, fill=WHITE, line=LINE, round_=True, radius=0.06)
txt(s, 7.35, 2.0, 5.2, 0.35, [("Samwl's bonus: Study A cross-check", {"bold": True, "size": 14, "color": NAVY, "font": FONT_B})])
txt(s, 7.35, 2.4, 5.25, 1.0,
    "Study A (probiotic trial): pre/post also balanced 86/90 per batch — the lab randomises well.",
    size=12, leading=1.1)
box(s, 7.1, 3.8, 5.7, 1.55, fill=TEAL_T, round_=True, radius=0.06)
txt(s, 7.35, 4.0, 5.2, 0.35, [("Decision recorded", {"bold": True, "size": 14, "color": NAVY, "font": FONT_B})])
txt(s, 7.35, 4.4, 5.25, 0.9,
    "Fathy: “we will not apply batch correction for downstream analysis.” Priv's PCoA by batch (next slides) shows no batch cluster.",
    size=12, leading=1.1)
txt(s, 0.55, 6.55, 12.2, 0.4,
    [("Why this check matters: ", {"bold": True, "color": NAVY}),
     ("Samwl's report is explicit — an unbalanced batch would look exactly like a treatment effect. "
      "Skipping the check and “hoping the lab randomised well” is the listed alternative.", {})],
    size=12, color=INK, leading=1.1)
footer(s, 6)

# ============================================================
# SLIDE 7 — baseline
# ============================================================
s = slide_new()
header(s, "Check 5 — one report only", "Did the arms differ at baseline? Only Samwl asked", 7)
tblshape = s.shapes.add_table(4, 6, Inches(0.55), Inches(1.8), Inches(7.1), Inches(2.3))
tbl = tblshape.table
tbl.columns[0].width = Inches(1.7)
for i in range(1, 6): tbl.columns[i].width = Inches(1.08)
for j, htx in enumerate(["arm", "n", "mean pH", "mean Nugent", "BV+", "STI+"]):
    cell_set(tbl.cell(0, j), htx, fill=NAVY, color=WHITE, size=12.5, bold=True,
             align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
data = [("injectable", 39, "4.86", "4.6", 19, 19), ("pill", 28, "4.59", "3.9", 10, 12),
        ("ring", 26, "4.97", "3.8", 10, 17)]
for i, row in enumerate(data, start=1):
    for j, v in enumerate(row):
        cell_set(tbl.cell(i, j), str(v), fill=WHITE if i % 2 else PAPER, size=12,
                 bold=(j == 0), align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
txt(s, 0.55, 4.45, 7.1, 2.2, [
    [("The study was not randomised ", {"bold": True}),
     ("— participants chose their method — so baseline differences are expected, not a data error.", {})],
    [("The arms start different ", {"bold": True, "color": CORAL}),
     ("(e.g. 19 vs 10 BV-positive) → later models should adjust for BV status and pH rather than trust the raw arm contrast.", {})],
    [("Priv and Fathy never ran this check — it is the biggest single coverage gap across the three reports.", {})],
], size=13, space_after=8, leading=1.1)
box(s, 8.0, 1.8, 4.8, 2.3, fill=WHITE, line=LINE, round_=True, radius=0.06)
txt(s, 8.25, 2.0, 4.3, 0.4, [("How to read it", {"bold": True, "size": 14, "color": NAVY, "font": FONT_B})])
txt(s, 8.25, 2.42, 4.35, 1.6, [
    [("pH & Nugent ", {"bold": True}), ("are shifted toward dysbiosis in the injectable arm (4.86 / 4.6).", {})],
    [("STI+ ", {"bold": True}), ("is 19/39 injectable vs ~half that elsewhere.", {})],
], size=12, space_after=7, leading=1.05)
box(s, 8.0, 4.35, 4.8, 2.0, fill=AMBER_T, round_=True, radius=0.06)
txt(s, 8.25, 4.55, 4.3, 0.35, [("Caveat Samwl flags", {"bold": True, "size": 14, "color": NAVY, "font": FONT_B})])
txt(s, 8.25, 4.95, 4.35, 1.3,
    "These are raw means, no formal test — imbalance is documented, not yet tested. Study A's baseline table (probiotic vs placebo) shows the same pattern of near-balance on age and pH.",
    size=11.5, leading=1.1)
footer(s, 7)

# ============================================================
# SLIDE 8 — exploratory: Shannon + PCoA
# ============================================================
s = slide_new()
header(s, "Bonus checks", "Exploratory pictures: diversity and outliers", 8)
pic(s, "slides_assets/samwl_3.png", 0.55, 1.65, 4.7, 3.6)
pic(s, "slides_assets/priv_2.png", 5.45, 1.65, 3.85, 2.8)
pic(s, "slides_assets/priv_3.png", 9.4, 1.65, 3.85, 2.8)
txt(s, 0.55, 5.35, 4.7, 0.3, "Shannon by arm (Samwl)", size=10, color=MUTED, align=PP_ALIGN.CENTER)
txt(s, 5.45, 4.55, 3.85, 0.3, "PCoA by arm (Priv)", size=10, color=MUTED, align=PP_ALIGN.CENTER)
txt(s, 9.4, 4.55, 3.85, 0.3, "PCoA by batch (Priv)", size=10, color=MUTED, align=PP_ALIGN.CENTER)
box(s, 5.45, 5.0, 7.35, 1.75, fill=WHITE, line=LINE, round_=True, radius=0.06)
txt(s, 5.7, 5.18, 6.9, 1.5, [
    [("Shannon: ", {"bold": True}), ("medians 1.74 / 0.90 / 1.59 (injectable / pill / ring); Kruskal–Wallis p = 0.47 — no significant arm difference. "
     "Pill arm looks bimodal (some near-0, Lactobacillus-only).", {})],
    [("PCoA: ", {"bold": True}), ("no batch cluster in the ordination → visual confirmation of the balanced cross-tab.", {})],
], size=12, space_after=7, leading=1.08)
txt(s, 0.55, 5.7, 4.7, 1.0,
    [("Fathy describes this same figure as a biological effect (“known effect of oral contraceptives”) — "
      "without reporting a test.", {})], size=11.5, color=CORAL, leading=1.05)
footer(s, 8)

# ============================================================
# SLIDE 9 — choices made, three columns
# ============================================================
s = slide_new()
header(s, "Choices made — side by side", "All three reports, choice by choice", 9)
cols9 = [
    (0.55, CORAL, "01-qc Priv — 5 choices", [
        "read_counts() helper instead of six manual read_csv() calls",
        "Small-library flag at 10,000 reads (alt: 5k flags nothing, 20k flags more)",
        "bv ↔ Nugent cut points 0–3 / 4–6 / 7–10 (alt: trust bv)",
        "Outlier ordination: Bray–Curtis PCoA on untransformed counts (alt: CLR + Euclidean)",
        "Keep all 93 samples, incl. 2 with missing cytokines (alt: drop them everywhere)",
    ]),
    (4.70, TEAL, "01-qc Samwl — 7 choices", [
        "QC on raw counts, before any transform",
        "“Established” species = ≥10 of 93 samples (alt: min read-count rule)",
        "Arm summaries as means (alt: medians)",
        "Cross-tab batch vs arm — never assume balance",
        "QC on all 220 taxa; analysis docs filter to 62",
        "Keep all 93; B006 flagged, kept pending ordination",
        "Shannon by hand + Kruskal–Wallis (alt: ANOVA on rarefied tables)",
    ]),
    (8.85, AMBER, "StudyB_QC_Fathy — 3 choices", [
        "Keep all 93 samples — min library 5,247 reads is “sufficient”; nothing dropped",
        "Keep all 220 taxa for the initial QC",
        "Batch highly balanced → no batch correction downstream",
    ]),
]
for bx, col, ttl, lst in cols9:
    box(s, bx, 1.7, 3.95, 5.0, fill=WHITE, line=col, round_=True, radius=0.04)
    box(s, bx, 1.7, 3.95, 0.5, fill=col)
    txt(s, bx + 0.2, 1.79, 3.6, 0.35, [(ttl, {"bold": True, "size": 13, "color": WHITE, "font": FONT_B})])
    paras = []
    for i, ch_ in enumerate(lst, 1):
        paras.append([(f"{i}. ", {"bold": True, "color": col}), (ch_, {})])
    txt(s, bx + 0.2, 2.38, 3.6, 4.15, paras, size=10.5, space_after=8, leading=1.08)
footer(s, 9)

# ============================================================
# SLIDE 10 — three-way differences: can either say why?
# ============================================================
s = slide_new()
header(s, "Choices made — the three-way differences", "For each difference: can any of us say why?", 10)
rows = [
    ("Species filter: all-zero → 134 (Priv)  vs  ≥10 samples → 62 (Samwl, Fathy >10%)",
     "YES", GREEN,
     "Priv: don't drag all-zero rows through every distance. Samwl: singletons can't support a 3-arm comparison; Fathy: a strict filter. All explained — but 134 ≠ 62, so the group must pick one."),
    ("Outlier check: PCoA (Priv)  vs  none (Samwl, Fathy)",
     "NO", CORAL,
     "Samwl's and Fathy's reports never mention ordination — silence, not a stated decision. → open-questions slide."),
    ("Alpha diversity: tested, p = 0.47 (Samwl)  vs  described, untested (Fathy)  vs  absent (Priv)",
     "NO", CORAL,
     "Fathy calls it “a known biological effect” with no test; Samwl tested and found nothing. → open-questions slide."),
    ("Sample B006: keep pending ordination (Priv, Samwl)  vs  “sufficient”, no flag (Fathy)",
     "PARTLY", AMBER,
     "Fathy states a depth-based reason but never engages the ordination edge test the other two defer to. → feeds open question 2."),
    ("Batch correction: explicitly skipped (Fathy)  vs  decision unstated (Priv, Samwl)",
     "YES", GREEN,
     "Fathy records the decision and says why (balanced cross-tab); the others only show the evidence. → adopt: no batch correction."),
    ("Threshold, samples & taxa: 10k flag, keep all 93, QC on all 220 (all three)",
     "SAME", TEAL,
     "Same choice, same reasoning — no conflict to resolve. Agreed."),
]
tblshape = s.shapes.add_table(len(rows) + 1, 3, Inches(0.55), Inches(1.75), Inches(12.25), Inches(4.7))
tbl = tblshape.table
tbl.columns[0].width = Inches(4.3)
tbl.columns[1].width = Inches(0.9)
tbl.columns[2].width = Inches(7.05)
for j, htx in enumerate(["Difference between the three reports", "Why?", "Can any of us explain it?"]):
    cell_set(tbl.cell(0, j), htx, fill=NAVY, color=WHITE, size=13, bold=True,
             align=PP_ALIGN.LEFT)
for i, (diff, verdict, col, why) in enumerate(rows, start=1):
    cell_set(tbl.cell(i, 0), diff, fill=WHITE if i % 2 else PAPER, size=11, bold=True, align=PP_ALIGN.LEFT)
    cell_set(tbl.cell(i, 1), verdict, fill={GREEN: GREEN_T, CORAL: CORAL_T, TEAL: TEAL_T, AMBER: AMBER_T}[col],
             color=col, size=12, bold=True)
    cell_set(tbl.cell(i, 2), why, fill=WHITE if i % 2 else PAPER, size=10.5, align=PP_ALIGN.LEFT)
txt(s, 0.55, 6.7, 12.2, 0.45,
    [("Rule used: ", {"bold": True, "color": NAVY}),
     ("if any report states its reason, the difference is discussable; where the other reports are silent, "
      "the item moves to the group's open-questions slide.", {})],
    size=12, color=INK)
footer(s, 10)

# ============================================================
# SLIDE 11 — open questions compared
# ============================================================
s = slide_new()
header(s, "Open questions — three reports", "What each report leaves open, and where they converge", 11)
cols = [
    ("Priv (verdict)", CORAL, [
        "B006: rerun the analysis without it if it sits at the edge of the ordination — “see whether the answer moves”",
        "86 all-zero species: filter before the main analyses",
        "Cytokine NAs (B018, B029): cytokine analyses must remember they are missing",
    ]),
    ("Samwl", TEAL, [
        "DESeq2 size factors: smallest nonzero count is 2 → use poscounts",
        "Filter species present in ≥10 of 93 samples (62 qualify)",
        "Alpha diversity: drop the follow-up unless the BV-adjusted analysis changes the picture",
    ]),
    ("Fathy", AMBER, [
        "Use poscounts size factors for DESeq2",
        "Strict low-abundance filter: keep taxa present in >10% of samples",
        "(no baseline, ID or outlier questions raised)",
    ]),
]
cx = 0.55
for name, col, items in cols:
    box(s, cx, 1.7, 4.05, 3.5, fill=WHITE, line=LINE, round_=True, radius=0.05)
    box(s, cx, 1.7, 4.05, 0.42, fill=col)
    txt(s, cx + 0.22, 1.77, 3.6, 0.3, [(name, {"bold": True, "size": 13, "color": WHITE, "font": FONT_B})])
    paras = [[("▸ ", {"color": col, "bold": True}), (it, {})] for it in items]
    txt(s, cx + 0.22, 2.3, 3.65, 2.8, paras, size=11.5, space_after=8, leading=1.08)
    cx += 4.15
box(s, 0.55, 5.45, 12.25, 1.2, fill=GREEN_T, round_=True, radius=0.08)
txt(s, 0.85, 5.62, 11.7, 0.35, [("Where the three converge", {"bold": True, "size": 14, "color": NAVY, "font": FONT_B})])
txt(s, 0.85, 6.0, 11.7, 0.55,
    [("poscounts size factors", {"bold": True}), (" — all three · ", {}),
     ("low-abundance filter", {"bold": True}), (" — Samwl (≥10 samples) ≈ Fathy (>10%): effectively the same rule · ", {}),
     ("keep all 93 samples", {"bold": True}), (" — all three (Fathy implicitly)", {})],
    size=12.5, color=INK)
footer(s, 11)

# ============================================================
# SLIDE 12 — the group's open questions (dark, deliverable)
# ============================================================
s = slide_new(dark=True)
box(s, 0, 0, 13.333, 0.14, fill=TEAL)
box(s, 9.0, 0, 4.333, 0.14, fill=CORAL)
txt(s, 0.55, 0.4, 11.5, 0.35, "THE DELIVERABLE", size=12, color=TEAL, bold=True)
txt(s, 0.55, 0.72, 12.2, 0.7, "The group's open-questions slide", size=28, color=WHITE, bold=True, font=FONT_B)
oq = [
    ("1", CORAL, "Which species filter do we standardise on?",
     "Priv drops only the 86 all-zero rows (134 species); Samwl and Fathy want a prevalence rule (≥10 samples ≈ >10%). Pick one before any per-species test."),
    ("2", CORAL, "Does B006 (5,247 reads) sit at the edge of the ordination?",
     "Priv and Samwl keep it “pending”; Fathy calls the depth “sufficient” without the ordination test. Rerun without B006 and see whether the answer moves."),
    ("3", TEAL, "Alpha diversity: signal or noise?",
     "Samwl's Kruskal–Wallis says no arm difference (p = 0.47); Fathy's report reads the pattern as biology. Which framing goes in the final report?"),
    ("4", TEAL, "Who owns the outlier check?",
     "Priv ran the only ordination (Bray–Curtis, untransformed). Does the group adopt it, or rerun under CLR + Euclidean — Samwl's stated alternative?"),
    ("5", AMBER, "Cytokines: confirm the B018/B029 rule",
     "Priv proposes dropping them from cytokine analyses only. Fathy and Samwl are silent — needs one group decision."),
]
y = 1.75
for num, col, q, detail in oq:
    c = box(s, 0.55, y, 0.5, 0.5, fill=col, round_=True, radius=0.5)
    tf = c.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = num; r.font.bold = True; r.font.size = Pt(16)
    r.font.color.rgb = NAVY if col is AMBER else WHITE; r.font.name = FONT_B
    txt(s, 1.25, y - 0.02, 11.6, 0.35, [(q, {"bold": True, "size": 15, "color": WHITE, "font": FONT_B})])
    txt(s, 1.25, y + 0.33, 11.6, 0.55, detail, size=12, color=FAINT, leading=1.0)
    y += 0.98
box(s, 0.55, 6.75, 12.25, 0.55, fill=NAVY2, round_=True, radius=0.15)
txt(s, 0.85, 6.86, 11.7, 0.35,
    [("Verdict: ", {"bold": True, "color": TEAL}),
     ("data are fit for analysis — IDs match, metadata complete, batch balanced, no sample must be dropped. "
      "Settle the five questions above before the differential tests.", {"color": WHITE})],
    size=12.5)

prs.save("StudyB_QC_slides.pptx")
print("saved StudyB_QC_slides.pptx with", len(prs.slides._sldIdLst), "slides")
