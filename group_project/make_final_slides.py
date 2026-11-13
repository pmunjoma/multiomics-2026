#!/usr/bin/env python3
"""Professional deck for Group B final answers (source: 'Group_B_Final_Answers.html')."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

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
ASSETS = "slides_assets"

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

def header(s, kicker, title, num=None):
    box(s, 0, 0, 13.333, 0.14, fill=TEAL)
    box(s, 9.0, 0, 4.333, 0.14, fill=CORAL)
    txt(s, 0.55, 0.32, 9.6, 0.32, kicker.upper(), size=12, color=TEAL, bold=True)
    txt(s, 0.55, 0.60, 11.6, 0.85, title, size=26, color=NAVY, bold=True, font=FONT_B)
    footer(s, num)

def footer(s, num):
    txt(s, 0.55, 7.08, 9.0, 0.3, "Study B · Group B final answers · group_b_contraception",
        size=10, color=MUTED)
    if num:
        txt(s, 12.3, 7.08, 0.5, 0.3, f"{num:02d}", size=10, color=MUTED, align=PP_ALIGN.RIGHT)

def pic_fit(s, path, x, y, max_w, max_h, frame=True):
    """Insert picture centred inside the max_w x max_h box, preserving aspect."""
    iw, ih = Image.open(path).size
    ratio = iw / ih
    if max_w / max_h > ratio: h = max_h; w = h * ratio
    else: w = max_w; h = w / ratio
    ix = x + (max_w - w) / 2; iy = y + (max_h - h) / 2
    p = s.shapes.add_picture(path, Inches(ix), Inches(iy), Inches(w), Inches(h))
    if frame:
        p.line.color.rgb = LINE; p.line.width = Pt(0.75)
    return p

def verdict_bar(s, y, verdict, color, note):
    box(s, 0.55, y, 12.25, 0.62, fill={GREEN: GREEN_T, AMBER: AMBER_T, CORAL: CORAL_T,
                                       TEAL: TEAL_T}[color], round_=True, radius=0.12)
    c = box(s, 0.75, y + 0.13, 1.35, 0.36, fill=color, round_=True, radius=0.5)
    tf = c.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = verdict
    r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_B
    txt(s, 2.3, y + 0.06, 10.3, 0.54, note, size=12.5, color=INK, leading=1.0,
        anchor=MSO_ANCHOR.MIDDLE)

def caption_box(s, x, y, w, h, readout, size=11.5):
    box(s, x, y, w, h, fill=TEAL_T, line=TEAL, round_=True, radius=0.09)
    txt(s, x + 0.22, y + 0.08, w - 0.44, 0.3, "CAPTION — READ THIS OUT", size=10.5,
        color=TEAL, bold=True)
    txt(s, x + 0.22, y + 0.34, w - 0.44, h - 0.42, readout, size=size, color=INK, leading=1.06)

# ============================================================
# SLIDE 1 — title
# ============================================================
s = slide_new(dark=True)
box(s, 0, 0, 0.22, 7.5, fill=TEAL)
box(s, 0.22, 0, 0.06, 7.5, fill=CORAL)
txt(s, 0.9, 0.95, 11.6, 0.4, "GROUP PROJECT · STEP 2 · FINAL ANALYSIS", size=13, color=TEAL, bold=True)
txt(s, 0.9, 1.45, 11.8, 1.9, [
    [("Group B — Final answers", {"size": 40, "bold": True, "color": WHITE, "font": FONT_B})],
    [("Contraception method, vaginal communities, and host gene expression — "
      "the merged best-of-two-drafts analysis, re-run with a fixed seed",
      {"size": 19, "color": FAINT})],
])
chips = [("93", "participants, one visit each"), ("3", "arms: pill · injectable · ring"),
         ("32/220", "species kept (95.3% of reads)"), ("13,760", "genes tested of 14,226")]
cx = 0.9
for v, lab in chips:
    box(s, cx, 3.45, 2.75, 1.05, fill=NAVY2, round_=True, radius=0.15)
    txt(s, cx + 0.2, 3.58, 2.4, 0.45, v, size=21, color=TEAL, bold=True, font=FONT_B)
    txt(s, cx + 0.2, 4.04, 2.4, 0.4, lab, size=10.5, color=FAINT, leading=0.95)
    cx += 2.87
txt(s, 0.9, 5.35, 11.5, 0.4,
    "Every number in this deck was computed from data/2-workshop2/group_b_contraception — "
    "reported in Group_B_Final_Answers.html.", size=14, color=WHITE, bold=True)
txt(s, 0.9, 5.85, 11.5, 0.6,
    "Module 3 (community composition) · Module 4 (differential expression) · "
    "PERMANOVA with betadisper checks · DESeq2 · one shared head-to-head model.",
    size=13, color=RGBColor(0x9F, 0xB4, 0xCC))

# ============================================================
# SLIDE 2 — the question, in one sentence
# ============================================================
s = slide_new()
header(s, "The question", "What we settled on — in one sentence", 2)
box(s, 0.55, 1.65, 12.25, 1.55, fill=TEAL_T, line=TEAL, round_=True, radius=0.09)
txt(s, 0.85, 1.85, 11.7, 1.2, [
    [("“Do the three contraception methods differ in vaginal bacterial community "
      "composition — and does the locally-acting ring change host gene expression "
      "compared with the systemically-acting methods — or is any apparent pattern "
      "better explained by bacterial vaginosis status?”",
      {"size": 17.5, "color": NAVY, "bold": True})],
], leading=1.15)
rows = [
    ("Q1", "Community composition — do pill, injectable and ring users harbour "
           "different bacterial communities?", TEAL),
    ("Q2", "Host gene expression — does the ring (local hormone) change vaginal "
           "tissue gene expression vs pill + injectable (systemic)?", CORAL),
    ("Q3", "Head-to-head — which explains host expression better: the woman's "
           "community type or her contraceptive method?", AMBER),
]
y = 3.55
for tag, q, col in rows:
    box(s, 0.55, y, 12.25, 0.92, fill=WHITE, line=LINE, round_=True, radius=0.10)
    c = box(s, 0.75, y + 0.19, 0.62, 0.55, fill=col, round_=True, radius=0.35)
    tf = c.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = tag
    r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_B
    txt(s, 1.6, y + 0.10, 11.0, 0.75, q, size=13.5, color=INK, anchor=MSO_ANCHOR.MIDDLE,
        leading=1.05)
    y += 1.08

# ============================================================
# SLIDE 3 — what we did, including filters
# ============================================================
s = slide_new()
header(s, "Methods", "What we did — and what we filtered, and why", 3)
steps = [
    ("1", "QC inherited", "93 samples × 1 visit; column↔row order checked before any test.", TEAL),
    ("2", "Filter", "Species: ≥10 counts in ≥20% of samples → 32 of 220 taxa, 95.3% of reads kept. "
                    "Genes: ≥10 counts in ≥23 samples (Module 4's quarter rule) → 13,760 of 14,226.", CORAL),
    ("3", "Look", "Bray–Curtis PCoA (microbiome) and CLR + Euclidean PCoA (genes); axis % reported.", TEAL),
    ("4", "Test globally", "PERMANOVA (999 permutations, seed fixed) with arm + bv + STI + batch "
                           "(`by = \"margin\"`); every PERMANOVA paired with a betadisper spread check.", CORAL),
    ("5", "Test gene-by-gene", "DESeq2 on raw counts, design ~ batch + bv + delivery; FDR < 5%; "
                               "robustness: drop bv, and ring vs each systemic arm separately.", TEAL),
]
y = 1.62
for tag, t, d_, col in steps:
    box(s, 0.55, y, 12.25, 0.78, fill=WHITE, line=LINE, round_=True, radius=0.12)
    c = box(s, 0.73, y + 0.17, 0.44, 0.44, fill=col, round_=True, radius=0.5)
    tf = c.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = tag
    r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_B
    txt(s, 1.35, y + 0.06, 2.35, 0.66, t, size=13, color=NAVY, bold=True,
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, 3.8, y + 0.05, 8.85, 0.68, d_, size=11.5, color=INK, anchor=MSO_ANCHOR.MIDDLE,
        leading=1.02)
    y += 0.90
box(s, 0.55, 6.20, 12.25, 0.72, fill=AMBER_T, round_=True, radius=0.10)
txt(s, 0.8, 6.28, 11.8, 0.6,
    [("Why filter at all:  ", {"bold": True, "color": NAVY}),
     ("a species seen in only 2 of 93 samples adds large distance to the few pairs it is in "
      "but carries no information — it distorts Bray–Curtis for no gain. Rare genes are "
      "untestable zeros. Raw counts (never CPM) went into DESeq2, which does its own normalisation.",
      {"color": INK})], size=11.5, leading=1.05)

# ============================================================
# SLIDE 4 — Q1 figure: ordination
# ============================================================
s = slide_new()
header(s, "Sub-question 1 · Community composition", "Bray–Curtis PCoA: the arms overlap — BV owns the map", 4)
pic_fit(s, f"{ASSETS}/fig1_ordination.png", 0.75, 1.50, 11.85, 3.55)
caption_box(s, 0.55, 5.15, 12.25, 1.15,
    "The same ordination coloured two ways. By contraception arm (left), pill, injectable and "
    "ring samples are thoroughly interleaved — no arm forms its own cloud. By BV status "
    "(right), BV-positive samples sweep to one side of Axis 1 (30.7% of variation), a "
    "Lactobacillus-vs-anaerobe gradient (r = −0.96). PERMANOVA agrees: arm explains ~3% "
    "(p = 0.09), BV 23% (p = 0.001).")
verdict_bar(s, 6.40, "Q1 · NO", CORAL,
    "No convincing difference in community composition between the three methods.")

# ============================================================
# SLIDE 5 — Q2 figure: volcano + boxplots
# ============================================================
s = slide_new()
header(s, "Sub-question 2 · Host gene expression", "DESeq2: the ring changes 139 host genes at FDR 5%", 5)
pic_fit(s, f"{ASSETS}/fig2_volcano.png", 0.55, 1.50, 12.25, 3.55)
caption_box(s, 0.55, 5.15, 12.25, 1.15,
    "Left: volcano — 139 genes at FDR 5% (73 up, 66 down in ring users); BPIFA1 is the extreme "
    "positive point (+5.8 log2FC), keratins (KRT4, KRT1) down. Right: four top genes as "
    "DESeq2-normalised counts — every box separates in the direction of its fold change. The "
    "signal survives dropping BV from the design, and all 139 genes point the same way against "
    "pill and injectable separately.")
verdict_bar(s, 6.40, "Q2 · YES", GREEN,
    "Yes — a modest but robust host-expression shift with local (ring) delivery.")

# ============================================================
# SLIDE 6 — Q3 figure: head-to-head
# ============================================================
s = slide_new()
header(s, "Sub-question 3 · Head-to-head", "One shared model: community type beats method", 6)
pic_fit(s, f"{ASSETS}/fig3_headtohead.png", 0.55, 1.50, 7.3, 3.55)
txt(s, 8.5, 1.60, 4.3, 3.5, [
    [("The fair comparison", {"size": 15, "bold": True, "color": NAVY, "font": FONT_B})],
    [("Both explanations in one DESeq2 design (~ batch + delivery + community): "
      "each term gets credit only for what it explains beyond the other.", {"size": 12.5})],
    [("220 vs 142 genes at FDR 5% — and the two gene lists share only 25 genes: "
      "different expression programs, not one signal seen twice.", {"size": 12.5})],
    [("PERMANOVA (whole transcriptome): a tie (R² 0.014 vs 0.015) — the community "
      "advantage is gene-specific, which is exactly why the gene-by-gene test settles it.",
      {"size": 12.5})],
], leading=1.1, space_after=10)
caption_box(s, 0.55, 5.15, 12.25, 1.15,
    "Community type — Lactobacillus- vs anaerobe-dominant, defined from the microbiome table "
    "and tested on the gene table (no circularity) — explains more host genes than the method: "
    "220 vs 142 at FDR 5% from one shared model (raw p < 0.05: 2,084 vs 1,648). The effect is "
    "directional: 192 of the 220 genes are down in anaerobe-dominant samples.")
verdict_bar(s, 6.40, "Q3 · YES", AMBER,
    "Community type — i.e. dysbiosis — explains host expression better than the method does.")

# ============================================================
# SLIDE 7 — choice we made: the two routes
# ============================================================
s = slide_new()
header(s, "Choice we made", "The two routes from the QC document — which we kept, and why", 7)
txt(s, 0.55, 1.42, 12.2, 0.4,
    "Group_B1_overall.html (QC) left one route open: which species filter to use before "
    "analysis — open question #1.", size=13, color=MUTED)
# Route A
box(s, 0.55, 1.95, 5.95, 2.75, fill=WHITE, line=LINE, round_=True, radius=0.08)
box(s, 0.55, 1.95, 5.95, 0.5, fill=MUTED, round_=True, radius=0.08)
txt(s, 0.8, 2.02, 5.5, 0.36, "ROUTE A — drop all-zero species only", size=13.5, color=WHITE,
    bold=True, font=FONT_B)
txt(s, 0.8, 2.6, 5.5, 2.0, [
    [("Keep every species detected at least once: 134 of 220 taxa.", {"size": 13})],
    [("The other 86 species are all-zero noise; but the remaining 134 still include "
      "species seen in only a handful of women.", {"size": 12.5, "color": MUTED})],
    [("Consequence: Bray–Curtis distances are inflated by rare, uninformative species — "
      "more ordination distortion, no extra signal.", {"size": 12.5, "color": MUTED})],
], leading=1.1, space_after=8)
# Route B
box(s, 6.85, 1.95, 5.95, 2.75, fill=TEAL_T, line=TEAL, round_=True, radius=0.08)
box(s, 6.85, 1.95, 5.95, 0.5, fill=TEAL, round_=True, radius=0.08)
txt(s, 7.1, 2.02, 5.5, 0.36, "ROUTE B — prevalence rule  ✓ KEPT", size=13.5, color=WHITE,
    bold=True, font=FONT_B)
txt(s, 7.1, 2.6, 5.5, 2.0, [
    [("Keep only species seen above a prevalence threshold; we tightened the QC's "
      "\u226510-sample version to \u226510 counts in \u226520% of samples.", {"size": 13})],
    [("Result: 32 of 220 taxa kept — carrying 95.3% of all reads, so the dropped 188 "
      "species were the rare tail, not the biomass.", {"size": 12.5, "color": INK})],
    [("Consequence: Bray–Curtis measures real abundance differences, not sampling noise.",
      {"size": 12.5, "color": INK})],
], leading=1.1, space_after=8)
box(s, 0.55, 4.95, 12.25, 0.95, fill=GREEN_T, round_=True, radius=0.10)
txt(s, 0.8, 5.03, 11.8, 0.8,
    [("Why we kept the prevalence route:  ", {"bold": True, "color": NAVY}),
     ("Bray–Curtis counts up differences between every pair of samples, so rare species add "
      "distance without adding information; the filter is the same cutoff the workshop's "
      "filter grid picked, and the ordination and PERMANOVA results are stable under it "
      "(the CLR sensitivity check agrees).", {"color": INK})], size=12.5, leading=1.08)
txt(s, 0.55, 6.15, 12.25, 0.75,
    [("Second route from the same QC list (open question #5), also settled:  ",
      {"bold": True, "color": NAVY, "size": 12}),
     ("low-read samples — the 10%-of-median rule drops nobody, the visual-gap rule drops B006 "
      "(5,247 reads); we kept all 93 samples, because dropping B006 moves nothing material.",
      {"color": INK, "size": 12})], leading=1.1)

# ============================================================
# SLIDE 8 — source we opened and read
# ============================================================
s = slide_new()
header(s, "Source", "One source we opened and read ourselves", 8)
box(s, 0.55, 1.7, 12.25, 1.35, fill=WHITE, line=LINE, round_=True, radius=0.08)
txt(s, 0.85, 1.88, 11.7, 1.05, [
    [("Love MI, Huber W, Anders S (2014). ", {"size": 15, "bold": True, "color": NAVY}),
     ("Moderated estimation of fold change and dispersion for RNA-seq data with DESeq2. ",
      {"size": 15, "color": INK}),
     ("Genome Biology 15:550.", {"size": 15, "italic": True, "color": MUTED})],
    [("genomebiology.biomedcentral.com/articles/10.1186/s13059-014-0550-8 — opened and read "
      "(abstract, model & normalisation, and shrinkage sections).", {"size": 11.5, "color": TEAL})],
], leading=1.15)
box(s, 0.55, 3.25, 12.25, 1.5, fill=TEAL_T, line=TEAL, round_=True, radius=0.09)
txt(s, 0.85, 3.42, 11.7, 1.2, [
    [("What we took from it:  ", {"size": 13, "bold": True, "color": TEAL}),
     ("“DESeq2 integrates methodological advances with several novel features to facilitate a "
      "more quantitative analysis of comparative RNA-seq data using shrinkage estimators for "
      "dispersion and fold change.”", {"size": 13.5, "italic": True, "color": NAVY, "bold": True})],
], leading=1.18)
txt(s, 0.85, 4.90, 11.7, 0.9,
    "This is why Part 2 feeds DESeq2 raw counts (never CPM or logged): DESeq2 models counts "
    "with a negative binomial GLM, estimates size factors itself, and shrinks noisy fold "
    "changes of low-count genes toward zero — so our 139-gene list is ranked by stable "
    "effects, not low-count noise.",
    size=13, color=INK, leading=1.15)
verdict_bar(s, 6.15, "WHY IT MATTERS", TEAL,
    "The paper's own caveat: a fixed pseudocount “only reacts to one of the sources of "
    "uncertainty, low counts” — the mistake our pipeline avoids.")

# ============================================================
# SLIDE 9 — the answer, plainly
# ============================================================
s = slide_new(dark=True)
box(s, 0, 0, 13.333, 0.14, fill=TEAL)
box(s, 9.0, 0, 4.333, 0.14, fill=CORAL)
txt(s, 0.9, 0.45, 11.5, 0.35, "THE ANSWER", size=13, color=TEAL, bold=True)
txt(s, 0.9, 0.78, 11.8, 0.7, "Stated as plainly as the data allow", size=27, color=WHITE,
    bold=True, font=FONT_B)
answers = [
    ("Q1 · Community composition", CORAL,
     "No. The three methods leave the vaginal community unchanged. Arm explains ~3% of the "
     "variation (p = 0.09; 0.06 adjusted) and the one borderline pair fails both the Holm "
     "correction and the compositional-distance check. BV status, not the method, owns the map."),
    ("Q2 · Host gene expression", GREEN,
     "Yes, modestly. The ring changes 139 host genes at FDR 5% — protective secretions up "
     "(BPIFA1), keratinisation down, prostaglandin pathway up (PTGS2) — and the signal "
     "survives every robustness check we ran."),
    ("Q3 · Head-to-head", AMBER,
     "Community type explains more than the method: 220 vs 142 genes at FDR 5% in one shared "
     "model. Read plainly: bacterial dysbiosis moves host expression; a well-tolerated "
     "contraceptive mostly doesn't."),
]
y = 1.75
for title, col, body in answers:
    box(s, 0.9, y, 11.55, 1.42, fill=NAVY2, round_=True, radius=0.10)
    box(s, 0.9, y, 0.12, 1.42, fill=col)
    txt(s, 1.25, y + 0.10, 10.9, 0.35, title, size=14, color=col, bold=True, font=FONT_B)
    txt(s, 1.25, y + 0.45, 10.9, 0.9, body, size=13, color=WHITE, leading=1.08)
    y += 1.57
txt(s, 0.9, 6.55, 11.8, 0.5,
    "Caveat we stand behind: BV/dysbiosis is the dominant variable in this cohort (R² ≈ 0.23); "
    "with 93 participants, subtle method effects are hard to detect above it.",
    size=12.5, color=FAINT, leading=1.1)

prs.save("Group_B_Final_Answers.pptx")
print("saved Group_B_Final_Answers.pptx with", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
