#!/usr/bin/env python3
"""Professional deck for the Group B1 overall QC (source: 'Group B1 overall.html')."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

NAVY   = RGBColor(0x10, 0x23, 0x3F)
NAVY2  = RGBColor(0x1B, 0x35, 0x5C)
INK    = RGBColor(0x22, 0x30, 0x47)
MUTED  = RGBColor(0x5C, 0x6B, 0x7A)
PAPER  = RGBColor(0xF7, 0xF9, 0xFC)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
TEAL   = RGBColor(0x17, 0xBE, 0xBB)
CORAL  = RGBColor(0xE7, 0x6F, 0x51)
AMBER  = RGBColor(0xE9, 0xC4, 0x6A)
GREEN  = RGBColor(0x2A, 0x9D, 0x8F)
GREEN_T = RGBColor(0xDE, 0xF2, 0xEC)
AMBER_T = RGBColor(0xFB, 0xF0, 0xD9)
CORAL_T = RGBColor(0xFB, 0xE3, 0xDB)
TEAL_T  = RGBColor(0xDD, 0xF3, 0xF2)
LINE    = RGBColor(0xD8, 0xE0, 0xEA)
FAINT   = RGBColor(0xBF, 0xD0, 0xE4)

FONT, FONT_B = "Segoe UI", "Segoe UI Semibold"
SW, SH = Inches(13.333), Inches(7.5)

prs = Presentation()
prs.slide_width, prs.slide_height = SW, SH
BLANK = prs.slide_layouts[6]

def slide_new(dark=False):
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid(); bg.fill.fore_color.rgb = NAVY if dark else PAPER
    bg.line.fill.background(); bg.shadow.inherit = False
    return s

def box(s, x, y, w, h, fill=None, line=None, round_=False, radius=0.08):
    shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,
                             Inches(x), Inches(y), Inches(w), Inches(h))
    if round_:
        try: shp.adjustments[0] = radius
        except Exception: pass
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb = line; shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp

def txt(s, x, y, w, h, runs, size=16, color=INK, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, leading=1.0, space_after=4, font=FONT):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(runs, str): runs = [runs]
    elif runs and isinstance(runs[0], tuple): runs = [runs]
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align; p.line_spacing = leading; p.space_after = Pt(space_after)
        if isinstance(para, str): para = [(para, {})]
        for text, style in para:
            r = p.add_run(); r.text = text
            f = r.font
            f.name = style.get("font", font); f.size = Pt(style.get("size", size))
            f.bold = style.get("bold", bold); f.italic = style.get("italic", False)
            f.color.rgb = style.get("color", color)
    return tb

def header(s, kicker, title, num=None, pass_label=None, pcolor=GREEN):
    box(s, 0, 0, 13.333, 0.14, fill=TEAL)
    box(s, 9.0, 0, 4.333, 0.14, fill=CORAL)
    txt(s, 0.55, 0.32, 9.6, 0.32, kicker.upper(), size=12, color=TEAL, bold=True)
    txt(s, 0.55, 0.60, 10.6, 0.85, title, size=27, color=NAVY, bold=True, font=FONT_B)
    if pass_label:
        c = box(s, 11.0, 0.55, 1.75, 0.62, fill=pcolor, round_=True, radius=0.25)
        tf = c.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = pass_label
        r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_B
    footer(s, num)

def footer(s, num):
    txt(s, 0.55, 7.08, 9.0, 0.3, "Study B · Group B1 overall QC · group_b_contraception",
        size=10, color=MUTED)
    if num:
        txt(s, 12.3, 7.08, 0.5, 0.3, f"{num:02d}", size=10, color=MUTED, align=PP_ALIGN.RIGHT)

def pic(s, path, x, y, w, h, frame=True):
    ratio = 1344 / 960
    if w / h > ratio: ih = h; iw = h * ratio
    else: iw = w; ih = w / ratio
    ix = x + (w - iw) / 2; iy = y + (h - ih) / 2
    p = s.shapes.add_picture(path, Inches(ix), Inches(iy), Inches(iw), Inches(ih))
    if frame:
        p.line.color.rgb = LINE; p.line.width = Pt(0.75)
    return p

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

def verdict_bar(s, y, verdict, color, note):
    box(s, 0.55, y, 12.25, 0.62, fill={GREEN: GREEN_T, AMBER: AMBER_T, CORAL: CORAL_T,
                                       TEAL: TEAL_T}[color], round_=True, radius=0.12)
    c = box(s, 0.75, y + 0.13, 1.25, 0.36, fill=color, round_=True, radius=0.5)
    tf = c.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = verdict
    r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_B
    txt(s, 2.2, y + 0.09, 10.4, 0.5, note, size=12.5, color=INK, leading=1.0,
        anchor=MSO_ANCHOR.MIDDLE)

# ============================================================
# SLIDE 1 — title
# ============================================================
s = slide_new(dark=True)
box(s, 0, 0, 0.22, 7.5, fill=TEAL)
box(s, 0.22, 0, 0.06, 7.5, fill=CORAL)
txt(s, 0.9, 0.95, 11.6, 0.4, "GROUP PROJECT · STEP 1 · QUALITY CONTROL", size=13, color=TEAL, bold=True)
txt(s, 0.9, 1.45, 11.8, 1.9, [
    [("Study B — Group B1 overall QC", {"size": 40, "bold": True, "color": WHITE, "font": FONT_B})],
    [("Five checks, re-run live on the data by the group's combined document",
      {"size": 20, "color": FAINT})],
])
chips = [("93", "samples, one visit each"), ("220", "species"), ("14,226", "genes"),
         ("15", "cytokine markers"), ("6", "analyst reports merged")]
cx = 0.9
for v, lab in chips:
    box(s, cx, 3.45, 2.2, 1.05, fill=NAVY2, round_=True, radius=0.15)
    txt(s, cx + 0.2, 3.58, 1.9, 0.45, v, size=22, color=TEAL, bold=True, font=FONT_B)
    txt(s, cx + 0.2, 4.04, 1.9, 0.4, lab, size=10.5, color=FAINT, leading=0.95)
    cx += 2.32
txt(s, 0.9, 5.35, 11.5, 0.4,
    "Every number in this deck was computed from data/2-workshop2/group_b_contraception — not quoted from any report.",
    size=14, color=WHITE, bold=True)
txt(s, 0.9, 5.85, 11.5, 0.6,
    "The five workshop checks: sample IDs · library sizes in both count tables · missing metadata · "
    "rnaseq_batch × arm · baseline balance.", size=13, color=RGBColor(0x9F, 0xB4, 0xCC))

# ============================================================
# SLIDE 2 — executive scorecard
# ============================================================
s = slide_new()
header(s, "Executive summary", "The five checks — results at a glance", 2)
rows = [
    ("1", "Sample IDs match across all tables", "PASS", GREEN,
     "Same IDs in the same ORDER, all four tables (metadata · species · genes · cytokines)"),
    ("2", "Library sizes in both count tables", "PASS", GREEN,
     "Species: 5,247–77,353 (median 29,212) · Genes: 5.86M–24.0M (median 12.86M) — both depth-checked"),
    ("2b", "Which samples have very few reads?", "ONE FLAG", AMBER,
     "B006 at 5,247 reads (18% of median); low tail B083, B065; gene table clean"),
    ("3", "Missing metadata", "PASS", GREEN,
     "0 NA in all 8 columns × 93 rows — but B018 & B029 miss all 15 cytokines"),
    ("4", "rnaseq_batch × arm", "BALANCED", GREEN,
     "51.3% / 50.0% / 50.0% in B1 — but B1 ran deeper (13.7M vs 12.0M): keep batch in the design"),
    ("5", "Arms differed at baseline?", "DIRECTIONAL", AMBER,
     "BV+ 19/39 vs 10/28 and 10/26; Fisher p = 0.454 — not significant, still adjust for bv/pH"),
]
y = 1.62
for num, chk, verdict, col, note in rows:
    box(s, 0.55, y, 12.25, 0.78, fill=WHITE, line=LINE, round_=True, radius=0.10)
    c = box(s, 0.75, y + 0.17, 0.44, 0.44, fill=NAVY, round_=True, radius=0.5)
    tf = c.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = num; r.font.size = Pt(13); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = FONT_B
    txt(s, 1.35, y + 0.10, 4.15, 0.6, chk, size=13.5, color=NAVY, bold=True,
        anchor=MSO_ANCHOR.MIDDLE, leading=0.95)
    c2 = box(s, 5.55, y + 0.22, 1.45, 0.36, fill=col, round_=True, radius=0.5)
    tf = c2.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = verdict
    r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_B
    txt(s, 7.2, y + 0.08, 5.45, 0.66, note, size=10.5, color=INK, leading=0.98,
        anchor=MSO_ANCHOR.MIDDLE)
    y += 0.86
footer(s, 2)

# ============================================================
# SLIDE 3 — Check 1: IDs
# ============================================================
s = slide_new()
header(s, "Check 1 of 5", "Do sample IDs match across all the tables?", 3,
       pass_label="PASS")
txt(s, 0.55, 1.6, 12.2, 0.4,
    "Why it matters: DESeq2 pairs count-matrix columns with metadata rows by position only — a mismatch fails silently.",
    size=13, color=MUTED)
box(s, 0.55, 2.15, 6.0, 2.9, fill=WHITE, line=LINE, round_=True, radius=0.05)
txt(s, 0.8, 2.35, 5.5, 0.35, [("The test that was run", {"bold": True, "size": 15, "color": NAVY, "font": FONT_B})])
txt(s, 0.8, 2.78, 5.55, 2.2, [
    [("setequal() ", {"bold": True, "color": TEAL}), ("— same set of IDs in every table → ", {}), ("TRUE", {"bold": True})],
    [("identical() ", {"bold": True, "color": TEAL}), ("— same ORDER in species, genes, cytokines vs metadata → ", {}), ("TRUE ×3", {"bold": True})],
    [("Coverage: ", {"bold": True}), ("all four tables, including the cytokine file — the one the helper does not auto-check, and the one two analyst reports skipped.", {})],
], size=12.5, space_after=8, leading=1.05)
box(s, 6.8, 2.15, 6.0, 2.9, fill=WHITE, line=LINE, round_=True, radius=0.05)
txt(s, 7.05, 2.35, 5.5, 0.35, [("What the six reports did", {"bold": True, "size": 15, "color": NAVY, "font": FONT_B})])
rows31 = [("Munya", "stopifnot(identical(...)) — hard stop, all 3 tables", GREEN),
          ("Priv", "setequal + hand-check of cytokines", GREEN),
          ("Samwl", "set AND order across all four tables", GREEN),
          ("Sam / Anthony", "all.equal() on species + genes only", AMBER),
          ("Fathy", "no explicit check shown — helper trusted", CORAL)]
y3 = 2.78
for who, what, col in rows31:
    box(s, 7.05, y3 + 0.06, 0.18, 0.18, fill=col, round_=True, radius=0.5)
    txt(s, 7.35, y3, 5.3, 0.3, [(who + " — ", {"bold": True}), (what, {})], size=11.5)
    y3 += 0.42
verdict_bar(s, 5.45, "PASS", GREEN,
    "Same IDs, same order, all four tables. Adopt the hard-stop pattern (Munya's stopifnot) in every analysis script.")
footer(s, 3)

# ============================================================
# SLIDE 4 — Check 2a: library sizes both tables
# ============================================================
s = slide_new()
header(s, "Check 2 of 5", "Library sizes in BOTH count tables", 4, pass_label="PASS")
tblshape = s.shapes.add_table(3, 5, Inches(0.55), Inches(1.7), Inches(7.4), Inches(1.5))
tbl = tblshape.table
tbl.columns[0].width = Inches(2.2)
for i in range(1, 5): tbl.columns[i].width = Inches(1.3)
for j, htx in enumerate(["Count table", "Min", "Median", "Mean", "Max"]):
    cell_set(tbl.cell(0, j), htx, fill=NAVY, color=WHITE, size=12.5, bold=True,
             align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
data = [("Species (microbiome)", "5,247", "29,212", "32,365", "77,353"),
        ("Genes (RNA-seq)", "5.86M", "12.86M", "13.11M", "24.03M")]
for i, row in enumerate(data, start=1):
    for j, v in enumerate(row):
        cell_set(tbl.cell(i, j), v, fill=WHITE if i % 2 else PAPER, size=12,
                 bold=(j == 0), align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
txt(s, 0.55, 3.5, 7.4, 1.4, [
    [("Flag rules applied: ", {"bold": True}),
     ("<10,000 reads (species) and <1,000,000 (genes) — the rules both Quarto reports used.", {})],
    [("Result: ", {"bold": True, "color": CORAL}), ("one flagged sample, B006 — see next slide. Gene table clean: min is ~6× the threshold.", {})],
], size=13, space_after=7, leading=1.05)
box(s, 8.15, 1.7, 4.65, 3.2, fill=AMBER_T, round_=True, radius=0.06)
txt(s, 8.4, 1.9, 4.2, 0.35, [("Who actually checked both?", {"bold": True, "size": 14, "color": NAVY, "font": FONT_B})])
txt(s, 8.4, 2.32, 4.2, 2.5, [
    [("Both tables: ", {"bold": True}), ("Samwl", {})],
    [("Species only: ", {"bold": True}), ("Priv, Fathy", {})],
    [("Plots but unnamed samples: ", {"bold": True}), ("Munya", {})],
    [("Numbers only: ", {"bold": True}), ("Sam, Anthony", {})],
    [("", {})],
    [("The combined document re-ran the check on both tables and names every low-read sample — the gap all six reports left open.", {})],
], size=11.5, space_after=6, leading=1.02)
pic(s, "slides_assets/samwl_1.png", 0.55, 4.75, 7.4, 2.05)
txt(s, 8.15, 5.5, 4.6, 1.3,
    [("Reading the plot: ", {"bold": True}),
     ("every one of 93 samples per table, sorted. The gene panel's spread is moderate (min = 45% of median) — "
      "DESeq2's size factors absorb it.", {})], size=11, color=MUTED, leading=1.05)
footer(s, 4)

# ============================================================
# SLIDE 5 — Check 2b: very few reads
# ============================================================
s = slide_new()
header(s, "Check 2 of 5 — the detail", "Which samples have very few reads?", 5,
       pass_label="1 FLAG", pcolor=AMBER)
tblshape = s.shapes.add_table(4, 4, Inches(0.55), Inches(1.7), Inches(6.6), Inches(1.9))
tbl = tblshape.table
tbl.columns[0].width = Inches(1.1); tbl.columns[1].width = Inches(1.5)
tbl.columns[2].width = Inches(2.2); tbl.columns[3].width = Inches(1.8)
for j, htx in enumerate(["Sample", "Reads", "vs median", "Verdict"]):
    cell_set(tbl.cell(0, j), htx, fill=NAVY, color=WHITE, size=12.5, bold=True)
rows5 = [("B006", "5,247", "18% of median", "clean outlier — the flag", CORAL),
         ("B083", "10,756", "37% of median", "low tail, not a gap", AMBER),
         ("B065", "11,570", "40% of median", "low tail, not a gap", AMBER)]
for i, (sm, rd, vs, vd, col) in enumerate(rows5, start=1):
    vals = [sm, rd, vs, vd]
    for j, v in enumerate(vals):
        cell_set(tbl.cell(i, j), v, fill=WHITE, size=12, bold=(j == 0),
                 color=col if j == 3 else INK)
txt(s, 0.55, 3.85, 6.6, 1.5, [
    [("Gene table: ", {"bold": True}),
     ("lowest is B042 at 5.86M — still 45% of the median (then B071 6.64M, B008 7.45M). "
      "No failure-scale outlier; nothing flagged under 1M.", {})],
], size=12.5, leading=1.08)
box(s, 0.55, 5.45, 6.6, 1.25, fill=TEAL_T, round_=True, radius=0.08)
txt(s, 0.8, 5.6, 6.1, 1.0,
    [("Two drop rules compared: ", {"bold": True}),
     ("10%-of-median cut (2,921) drops NOTHING; visual-gap rule drops B006 only. "
      "B006 stays either way — the rule choice is a recorded open question (Sam).", {})],
    size=11.5, leading=1.05)
pic(s, "slides_assets/priv_1.png", 7.45, 1.7, 5.35, 3.3)
box(s, 7.45, 5.25, 5.35, 1.45, fill=CORAL_T, round_=True, radius=0.06)
txt(s, 7.7, 5.42, 4.9, 1.15,
    [("Why one small library matters: ", {"bold": True}),
     ("percentages from a 5,247-read sample are noisy and can pretend to be biology. "
      "Priv's ordination shows B006 inside the cloud — keep, sensitivity-check later.", {})],
    size=11.5, leading=1.05)
footer(s, 5)

# ============================================================
# SLIDE 6 — Check 3: missing metadata
# ============================================================
s = slide_new()
header(s, "Check 3 of 5", "Missing metadata — and the cytokine gap", 6, pass_label="PASS")
box(s, 0.55, 1.7, 6.0, 2.75, fill=WHITE, line=LINE, round_=True, radius=0.05)
txt(s, 0.8, 1.9, 5.5, 0.35, [("Metadata: 0 missing, 93 × 8", {"bold": True, "size": 15, "color": NAVY, "font": FONT_B})])
txt(s, 0.8, 2.32, 5.55, 2.0, [
    [("All eight columns complete: ", {"bold": True}),
     ("sample_id, participant_id, arm, bv, nugent, ph, any_sti, rnaseq_batch.", {})],
    [("Six reports, one habit: ", {"bold": True}),
     ("only Priv's and Munya's reports actually ran the is.na() sweep. Sam, Anthony, Samwl and Fathy described the metadata without checking it — we got lucky.", {})],
], size=12.5, space_after=8, leading=1.05)
box(s, 6.8, 1.7, 6.0, 2.75, fill=WHITE, line=LINE, round_=True, radius=0.05)
txt(s, 7.05, 1.9, 5.5, 0.35, [("The 4th table almost slipped through", {"bold": True, "size": 15, "color": NAVY, "font": FONT_B})])
txt(s, 7.05, 2.32, 5.55, 2.0, [
    [("B018 and B029: ", {"bold": True, "color": CORAL}),
     ("NA for all 15 cytokine markers — an all-or-nothing gap (2.2% of samples), consistent with a failed lab plate.", {})],
    [("Every other sample complete; ", {"bold": True}),
     ("no negative values anywhere (minimum 0.005).", {})],
    [("Decision owed: ", {"bold": True}),
     ("drop them from cytokine analyses only — open question 7.", {})],
], size=12.5, space_after=7, leading=1.05)
box(s, 0.55, 4.7, 12.25, 1.95, fill=TEAL_T, round_=True, radius=0.06)
txt(s, 0.8, 4.88, 11.8, 0.35, [("Bonus check run live: does the metadata contradict itself?", {"bold": True, "size": 14, "color": NAVY, "font": FONT_B})])
txt(s, 0.8, 5.28, 11.8, 1.25,
    "bv ↔ Nugent cross-tab: 51/51 negative in Nugent 0–3 · 3/3 intermediate in 4–6 · 39/39 positive in 7–10. "
    "The two clinical columns never disagree (checked by Priv and Munya independently) — no data-entry surprises.",
    size=12.5, leading=1.1)
footer(s, 6)

# ============================================================
# SLIDE 7 — Check 4: batch × arm
# ============================================================
s = slide_new()
header(s, "Check 4 of 5", "rnaseq_batch × arm — balanced, but not interchangeable", 7, pass_label="BALANCED")
tblshape = s.shapes.add_table(4, 4, Inches(0.55), Inches(1.75), Inches(5.9), Inches(2.2))
tbl = tblshape.table
tbl.columns[0].width = Inches(1.9)
for i in range(1, 4): tbl.columns[i].width = Inches(1.33)
for j, htx in enumerate(["arm", "B1", "B2", "% in B1"]):
    cell_set(tbl.cell(0, j), htx, fill=NAVY, color=WHITE, size=13, bold=True,
             align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
for i, row in enumerate([("injectable", 20, 19, "51.3%"), ("pill", 14, 14, "50.0%"),
                         ("ring", 13, 13, "50.0%")], start=1):
    for j, v in enumerate(row):
        cell_set(tbl.cell(i, j), str(v), fill=WHITE if i % 2 else PAPER, size=12.5,
                 bold=(j == 0 or j == 3), align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
txt(s, 0.55, 4.2, 5.9, 1.2,
    [("The one unanimous check: ", {"bold": True}),
     ("all six reports produced the same cross-tab. A batch effect cannot masquerade as a method effect by assignment.", {})],
    size=12.5, leading=1.08)
box(s, 0.55, 5.55, 5.9, 1.15, fill=AMBER_T, round_=True, radius=0.08)
txt(s, 0.8, 5.7, 5.45, 0.9,
    [("But balanced ≠ similar: ", {"bold": True}),
     ("B1's RNA-seq libraries run deeper — median 13.7M vs 12.0M reads. "
      "A depth difference is a batch effect even when the design is balanced.", {})],
    size=11.5, leading=1.05)
box(s, 6.8, 1.75, 6.0, 4.95, fill=WHITE, line=LINE, round_=True, radius=0.05)
pic(s, "slides_assets/priv_3.png", 7.0, 1.95, 5.6, 3.6, frame=False)
txt(s, 7.0, 5.65, 5.6, 0.9,
    [("PCoA coloured by batch: ", {"bold": True}),
     ("no batch cluster in the ordination — the depth difference does not rearrange the communities. "
      "Decision: no batch correction, but rnaseq_batch stays in every DESeq2 design.", {})],
    size=11.5, leading=1.08)
footer(s, 7)

# ============================================================
# SLIDE 8 — Check 5: baseline
# ============================================================
s = slide_new()
header(s, "Check 5 of 5", "Did the arms differ at baseline? Now formally tested", 8,
       pass_label="DIRECTIONAL", pcolor=AMBER)
tblshape = s.shapes.add_table(4, 6, Inches(0.55), Inches(1.7), Inches(7.6), Inches(2.2))
tbl = tblshape.table
tbl.columns[0].width = Inches(1.6)
for i in range(1, 6): tbl.columns[i].width = Inches(1.2)
for j, htx in enumerate(["arm", "n", "pH", "Nugent†", "BV+", "STI+"]):
    cell_set(tbl.cell(0, j), htx, fill=NAVY, color=WHITE, size=12, bold=True,
             align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
data = [("injectable", 39, "4.86", "4.6 / 2.0", 19, 19),
        ("pill", 28, "4.59", "3.9 / 2.0", 10, 12),
        ("ring", 26, "4.97", "3.8 / 2.5", 10, 17)]
for i, row in enumerate(data, start=1):
    for j, v in enumerate(row):
        cell_set(tbl.cell(i, j), str(v), fill=WHITE if i % 2 else PAPER, size=11.5,
                 bold=(j == 0), align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
txt(s, 0.55, 4.0, 7.6, 0.35, "† mean / median — medians sit at ~2 in every arm; means are pulled up by the positive arm",
    size=10, color=MUTED)
txt(s, 0.55, 4.5, 7.6, 2.1, [
    [("The study was not randomised. ", {"bold": True}),
     ("BV leans positive in the injectable arm (19/39 vs 10/28 and 10/26) — the imbalance every report eyeballed.", {})],
    [("Now tested: ", {"bold": True, "color": TEAL}),
     ("Fisher's exact test on BV × arm gives p = 0.454 — not significant at n = 93. "
      "Adjust for bv/pH anyway: it is pre-planned, costs nothing, and may matter at higher power.", {})],
], size=12.5, space_after=8, leading=1.08)
box(s, 8.4, 1.7, 4.4, 2.6, fill=WHITE, line=LINE, round_=True, radius=0.06)
txt(s, 8.65, 1.9, 3.9, 0.35, [("Sanity check: pH tracks BV", {"bold": True, "size": 13.5, "color": NAVY, "font": FONT_B})])
for i, (bv, ph) in enumerate([("negative", "4.4"), ("intermediate", "4.7"), ("positive", "5.0")]):
    yy = 2.4 + i * 0.55
    txt(s, 8.65, yy, 1.9, 0.35, bv, size=12, color=MUTED)
    txt(s, 10.6, yy, 1.4, 0.35, "pH " + ph, size=12, bold=True, color=NAVY)
txt(s, 8.65, 3.95, 3.95, 0.3, "Biologically expected — verified.", size=10.5, color=MUTED)
box(s, 8.4, 4.55, 4.4, 2.05, fill=CORAL_T, round_=True, radius=0.06)
txt(s, 8.65, 4.72, 3.9, 1.75,
    [("Consequence for Step 2: ", {"bold": True}),
     ("arm contrasts in DESeq2 and per-species tests should carry bv (and pH) as covariates rather than trust the raw arm difference.", {})],
    size=11.5, leading=1.05)
footer(s, 8)

# ============================================================
# SLIDE 9 — supporting evidence
# ============================================================
s = slide_new()
header(s, "Supporting evidence", "Ordination and diversity — the checks behind the verdict", 9)
pic(s, "slides_assets/priv_2.png", 0.55, 1.7, 6.0, 4.0)
pic(s, "slides_assets/samwl_3.png", 6.85, 1.7, 5.95, 4.0)
txt(s, 0.55, 5.85, 6.0, 0.3, "Bray–Curtis PCoA by arm (Priv, re-run)", size=10.5, color=MUTED, align=PP_ALIGN.CENTER)
txt(s, 6.85, 5.85, 5.95, 0.3, "Shannon diversity by arm (Samwl, re-run)", size=10.5, color=MUTED, align=PP_ALIGN.CENTER)
box(s, 0.55, 6.25, 12.25, 0.75, fill=WHITE, line=LINE, round_=True, radius=0.10)
txt(s, 0.8, 6.4, 11.8, 0.5,
    [("Left: ", {"bold": True, "color": TEAL}),
     ("arms overlap across the ordination; B006 is not an edge case — evidence for keeping it.   ", {}),
     ("Right: ", {"bold": True, "color": TEAL}),
     ("medians 1.74 / 0.90 / 1.59 but Kruskal–Wallis p = 0.47 — no significant arm difference.", {})],
    size=12, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 9)

# ============================================================
# SLIDE 10 — open questions
# ============================================================
s = slide_new()
header(s, "What remains open", "Open questions — owners and status", 10)
oq = [
    ("Species filter: all-zero drop (134) vs ≥10-sample prevalence (62)?", "group", "OPEN — pick before per-species tests", CORAL),
    ("Low-read rule: 10%-of-median (drops none) vs visual gap (drops B006)?", "Sam", "OPEN — B006 stays either way", CORAL),
    ("rnaseq_batch in every DESeq2 design?", "Anthony", "ADOPTED — yes; B1 runs 1.7M reads deeper", GREEN),
    ("B018/B029 dropped from cytokine analyses only?", "Anthony · Munya", "OPEN — confirm before cytokine models", CORAL),
    ("bv at 3 levels, or merge 'intermediate' (3 people)?", "Sam", "OPEN — keep 3 levels for now", CORAL),
    ("any_sti as a covariate in arm comparisons?", "Sam", "OPEN — decided with the baseline model", CORAL),
    ("Formal baseline-balance test", "Munya", "DONE — Fisher p = 0.454 reported here", GREEN),
    ("Alpha diversity framing (test says noise, Fathy says biology)", "group", "ANSWERED — p = 0.47; treat as noise", GREEN),
    ("Anthony re-renders the duplicated report", "all", "PROCESS — outside this document", AMBER),
]
y = 1.62
for q, owner, status, col in oq:
    box(s, 0.55, y, 12.25, 0.55, fill=WHITE, line=LINE, round_=True, radius=0.12)
    box(s, 0.72, y + 0.13, 0.28, 0.28, fill=col, round_=True, radius=0.5)
    txt(s, 1.15, y + 0.05, 6.55, 0.45, q, size=11.5, color=INK, anchor=MSO_ANCHOR.MIDDLE, leading=0.9)
    txt(s, 7.8, y + 0.05, 1.6, 0.45, owner, size=10.5, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
    c = box(s, 9.5, y + 0.11, 3.1, 0.33, fill={CORAL: CORAL_T, GREEN: GREEN_T, AMBER: AMBER_T}[col],
            round_=True, radius=0.5)
    tf = c.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = status
    r.font.size = Pt(9.5); r.font.bold = True
    r.font.color.rgb = {CORAL: CORAL, GREEN: GREEN, AMBER: AMBER}[col]; r.font.name = FONT
    y += 0.60
footer(s, 10)

# ============================================================
# SLIDE 11 — verdict (dark close)
# ============================================================
s = slide_new(dark=True)
box(s, 0, 0, 13.333, 0.14, fill=TEAL)
box(s, 9.0, 0, 4.333, 0.14, fill=CORAL)
txt(s, 0.55, 0.5, 11.5, 0.35, "VERDICT", size=13, color=TEAL, bold=True)
txt(s, 0.55, 0.88, 12.2, 0.9, "Safe to analyse — with decisions taken first",
    size=32, color=WHITE, bold=True, font=FONT_B)
pts = [
    ("All five checks pass on the data.", TEAL,
     "IDs match in order, metadata is complete, both tables depth-checked, batch balanced, baseline tested."),
    ("Nothing must be dropped.", TEAL,
     "B006 sits inside the ordination cloud; the cytokine gap affects only the cytokine analyses."),
    ("Two practice changes adopted.", CORAL,
     "Keep rnaseq_batch in every DESeq2 design (B1 runs 1.7M reads deeper) · adjust for bv in arm contrasts."),
    ("Settle OQ1 and OQ5 before Step 2.", AMBER,
     "Species filter (134 vs 62) and the low-read rule — both affect every per-species test that follows."),
]
y = 2.1
for head, col, detail in pts:
    c = box(s, 0.55, y, 0.16, 0.95, fill=col, round_=True, radius=0.5)
    txt(s, 0.95, y - 0.04, 11.8, 0.4, head, size=17, color=WHITE, bold=True, font=FONT_B)
    txt(s, 0.95, y + 0.38, 11.8, 0.6, detail, size=13, color=FAINT, leading=1.0)
    y += 1.12
box(s, 0.55, 6.7, 12.25, 0.55, fill=NAVY2, round_=True, radius=0.15)
txt(s, 0.85, 6.81, 11.7, 0.35,
    [("Source: ", {"bold": True, "color": TEAL}),
     ("Group B1 overall.html — every number re-runnable via its .qmd; rebuild with quarto render.", {"color": WHITE})],
    size=11.5)

prs.save("Group_B1_overall.pptx")
print("saved Group_B1_overall.pptx with", len(prs.slides._sldIdLst), "slides")
