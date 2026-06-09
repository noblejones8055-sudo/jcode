from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pptx
import pptx.oxml.ns as nsmap
from lxml import etree
import copy

# ── Colour Palette ──────────────────────────────────────────────────────────
C_DARK       = RGBColor(0x0D, 0x1F, 0x16)   # near-black forest
C_MID        = RGBColor(0x1B, 0x46, 0x32)   # deep hunter green
C_GREEN      = RGBColor(0x2D, 0x8C, 0x5F)   # vivid emerald
C_LIME       = RGBColor(0x52, 0xB7, 0x88)   # bright mint
C_SAGE       = RGBColor(0xB7, 0xE4, 0xC7)   # pale sage
C_GOLD       = RGBColor(0xE9, 0xC4, 0x6A)   # warm gold accent
C_CORAL      = RGBColor(0xE7, 0x6F, 0x51)   # danger / highlight
C_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT_GREY = RGBColor(0xF0, 0xF4, 0xF1)
C_DARK_GREY  = RGBColor(0x3A, 0x3A, 0x3A)
C_MID_GREY   = RGBColor(0x8A, 0x8A, 0x8A)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

blank_layout = prs.slide_layouts[6]  # completely blank

# ── Helpers ─────────────────────────────────────────────────────────────────

def rgb_hex(r):
    return "{:02X}{:02X}{:02X}".format(r.r, r.g, r.b)

def add_rect(slide, x, y, w, h, fill_color=None, line_color=None, line_w=None):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(1, x, y, w, h)  # 1 = MSO_SHAPE_TYPE.RECTANGLE
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_w:
            shape.line.width = line_w
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, text, x, y, w, h,
                font_size=18, bold=False, italic=False,
                color=C_WHITE, align=PP_ALIGN.LEFT,
                word_wrap=True, font_name="Calibri"):
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame
    tf.word_wrap = word_wrap
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size      = Pt(font_size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = color
    run.font.name      = font_name
    return txb

def add_multiline(slide, lines, x, y, w, h,
                  font_size=14, color=C_WHITE,
                  align=PP_ALIGN.LEFT, line_spacing=1.1,
                  font_name="Calibri"):
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame
    tf.word_wrap = True
    first = True
    for (txt, sz, bold, col) in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = txt
        run.font.size      = Pt(sz or font_size)
        run.font.bold      = bold
        run.font.color.rgb = col or color
        run.font.name      = font_name
    return txb

def add_bg(slide, color=C_DARK):
    bg = add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_color=color)
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)
    return bg

def accent_bar(slide, x, y, w, h=Inches(0.05), color=C_GOLD):
    add_rect(slide, x, y, w, h, fill_color=color)

def slide_number(slide, n, total=15):
    add_textbox(slide, f"{n} / {total}",
                SLIDE_W - Inches(1.1), SLIDE_H - Inches(0.4),
                Inches(0.9), Inches(0.3),
                font_size=9, color=C_MID_GREY, align=PP_ALIGN.RIGHT)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1 – TITLE
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)

# Full-bleed dark green panel left 55 %
add_rect(s, 0, 0, Inches(7.3), SLIDE_H, fill_color=C_MID)

# Gold diagonal accent stripe (thin rect rotated)
accent_bar(s, 0, Inches(5.5), Inches(7.3), Inches(0.07), C_GOLD)
accent_bar(s, 0, Inches(5.58), Inches(4.0), Inches(0.03), C_LIME)

# ── Left panel text
add_textbox(s, "NATURE IS THE NEXT FRONTIER",
            Inches(0.5), Inches(0.7), Inches(6.5), Inches(0.5),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")

add_textbox(s, "Biodiversity as a\nBusiness Imperative",
            Inches(0.5), Inches(1.3), Inches(6.4), Inches(2.2),
            font_size=44, bold=True, color=C_WHITE, font_name="Calibri")

add_textbox(s, "GRI 101: Biodiversity 2024 — Strategic Compliance\nand Nature-Positive Corporate Leadership",
            Inches(0.5), Inches(3.6), Inches(6.4), Inches(1.0),
            font_size=16, italic=True, color=C_SAGE, font_name="Calibri")

# Key stats bottom left
add_textbox(s, "Effective 1 January 2026  |  8 Mandatory Disclosures  |  Aligned with Kunming-Montreal GBF",
            Inches(0.5), Inches(5.8), Inches(6.5), Inches(0.4),
            font_size=10, color=C_GOLD, font_name="Calibri")

# ── Right panel — decorative quote block
add_rect(s, Inches(7.5), Inches(1.5), Inches(5.3), Inches(3.8),
         fill_color=RGBColor(0x1E, 0x50, 0x38))

add_textbox(s, "“",
            Inches(7.6), Inches(1.4), Inches(0.8), Inches(1.0),
            font_size=60, bold=True, color=C_GOLD, font_name="Georgia")

add_textbox(s,
    "More than half of global GDP — $44 trillion — is moderately or highly dependent on nature and its services.",
    Inches(7.8), Inches(2.1), Inches(4.8), Inches(1.8),
    font_size=18, italic=True, color=C_WHITE, font_name="Georgia")

add_textbox(s, "— World Economic Forum, 2023",
            Inches(8.2), Inches(4.0), Inches(4.2), Inches(0.4),
            font_size=11, color=C_LIME, font_name="Calibri")

# Bottom right — firm / date
add_textbox(s, "Strategic Advisory  |  June 2026",
            Inches(7.5), Inches(6.8), Inches(5.3), Inches(0.4),
            font_size=11, color=C_MID_GREY, align=PP_ALIGN.RIGHT, font_name="Calibri")

slide_number(s, 1)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2 – AGENDA / TABLE OF CONTENTS
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, Inches(0.18), SLIDE_H, fill_color=C_GREEN)
accent_bar(s, Inches(0.3), Inches(1.35), Inches(2.5), Inches(0.04), C_GOLD)

add_textbox(s, "AGENDA", Inches(0.35), Inches(0.3), Inches(4), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "What We Cover Today",
            Inches(0.35), Inches(0.65), Inches(7), Inches(0.6),
            font_size=30, bold=True, color=C_WHITE, font_name="Calibri")

items = [
    ("01", "The Burning Platform — Why Biodiversity, Why Now",    C_GOLD),
    ("02", "Regulatory Landscape 2026 — GRI 101 & Beyond",        C_LIME),
    ("03", "GRI 101 Architecture — The 8 Disclosures",            C_GOLD),
    ("04", "Deep Dive: Each Disclosure & Key Reporting Demands",   C_LIME),
    ("05", "Materiality, Methodology & Data Quality",              C_GOLD),
    ("06", "Linking Biodiversity to Climate & Social Risk",        C_LIME),
    ("07", "Implementation Roadmap — From Gap to Green",           C_GOLD),
    ("08", "Business Case — Risks, Opportunities & Investor Signal",C_LIME),
    ("09", "Strategic Recommendations — 5 Priorities to Act Now",  C_GOLD),
    ("10", "Q&A & Next Steps",                                     C_LIME),
]

col_x = [Inches(0.35), Inches(7.0)]
col_y_start = Inches(1.55)
row_h = Inches(0.5)

for i, (num, title, accent) in enumerate(items):
    col = i % 2
    row = i // 2
    x = col_x[col]
    y = col_y_start + row * row_h
    # number bubble
    add_rect(s, x, y + Inches(0.04), Inches(0.42), Inches(0.36),
             fill_color=C_MID if col == 0 else RGBColor(0x1E,0x50,0x38))
    add_textbox(s, num, x + Inches(0.04), y, Inches(0.4), Inches(0.42),
                font_size=12, bold=True, color=accent, align=PP_ALIGN.CENTER,
                font_name="Calibri")
    add_textbox(s, title, x + Inches(0.5), y + Inches(0.06),
                Inches(5.8), Inches(0.35),
                font_size=13, color=C_WHITE, font_name="Calibri")

slide_number(s, 2)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3 – THE BURNING PLATFORM
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
# top banner
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "01  |  THE BURNING PLATFORM",
            Inches(0.4), Inches(0.15), Inches(8), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "Why Biodiversity. Why Now. Why It Cannot Wait.",
            Inches(0.4), Inches(0.55), Inches(10), Inches(0.6),
            font_size=26, bold=True, color=C_WHITE, font_name="Calibri")

# 4 stat cards
cards = [
    (C_CORAL,  "1 MILLION",  "species threatened\nwith extinction",   "IPBES 2019"),
    (C_GOLD,   "$44 TRILLION","GDP dependent on\nnature & ecosystem\nservices", "WEF 2023"),
    (C_LIME,   "68%",         "decline in wildlife\npopulations since\n1970",   "WWF 2022"),
    (C_GREEN,  "2030",        "deadline to protect\n30% of land & ocean\n(30x30 Target)", "KM-GBF 2022"),
]
card_w = Inches(3.0)
card_h = Inches(4.5)
gap    = Inches(0.22)
start_x = Inches(0.35)

for i, (clr, big, desc, src) in enumerate(cards):
    cx = start_x + i * (card_w + gap)
    cy = Inches(1.5)
    add_rect(s, cx, cy, card_w, card_h,
             fill_color=RGBColor(0x18, 0x3A, 0x28))
    accent_bar(s, cx, cy, card_w, Inches(0.08), clr)
    add_textbox(s, big, cx + Inches(0.15), cy + Inches(0.2),
                card_w - Inches(0.3), Inches(1.0),
                font_size=32, bold=True, color=clr,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    add_textbox(s, desc, cx + Inches(0.1), cy + Inches(1.2),
                card_w - Inches(0.2), Inches(2.4),
                font_size=15, color=C_WHITE,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    add_textbox(s, src, cx + Inches(0.1), cy + Inches(4.1),
                card_w - Inches(0.2), Inches(0.3),
                font_size=10, color=C_MID_GREY,
                align=PP_ALIGN.CENTER, font_name="Calibri")

# bottom insight
add_rect(s, Inches(0.35), Inches(6.2), Inches(12.6), Inches(0.9),
         fill_color=C_MID)
add_textbox(s,
    "KEY INSIGHT:  Biodiversity loss is now classified as one of the top five global risks over the next decade by the World Economic Forum — alongside climate change, weapons of mass destruction and geoeconomic confrontation.",
    Inches(0.5), Inches(6.25), Inches(12.3), Inches(0.75),
    font_size=12, color=C_WHITE, font_name="Calibri")
slide_number(s, 3)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 4 – REGULATORY LANDSCAPE
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "02  |  REGULATORY LANDSCAPE 2026",
            Inches(0.4), Inches(0.15), Inches(9), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "Converging Frameworks — A New Disclosure Ecosystem",
            Inches(0.4), Inches(0.55), Inches(11), Inches(0.6),
            font_size=26, bold=True, color=C_WHITE, font_name="Calibri")

regs = [
    ("GRI 101\nBiodiversity 2024",
     "Effective Jan 2026. 8 mandatory disclosures covering policies, management, locations, drivers, ecosystem state & services. The primary reporting standard addressed in this engagement.",
     C_LIME, Inches(0.35)),
    ("Kunming-Montreal\nGlobal Biodiversity\nFramework (KM-GBF)",
     "Adopted Dec 2022 (COP15). 2050 Vision: 'Living in harmony with nature'. 23 action targets for 2030 including 30×30 land/ocean protection and $700B in biodiversity finance.",
     C_GOLD, Inches(3.65)),
    ("TNFD — Taskforce on\nNature-related Financial\nDisclosures",
     "Beta v1.0 launched 2023. LEAP approach (Locate, Evaluate, Assess, Prepare). Already adopted voluntarily by 320+ organisations globally. Closely aligned with GRI 101.",
     C_CORAL, Inches(6.95)),
    ("EU CSRD &\nEuropean\nSustainability Standards",
     "ESRS E4 (Biodiversity) requires disclosure on sites in/near protected areas, species affected and ecosystem services. Applies from 2025–2028 depending on company size.",
     C_SAGE, Inches(10.25)),
]

for (title, body, clr, x) in regs:
    bw = Inches(3.05)
    by = Inches(1.5)
    bh = Inches(5.5)
    add_rect(s, x, by, bw, bh, fill_color=RGBColor(0x18,0x3A,0x28))
    accent_bar(s, x, by, bw, Inches(0.07), clr)
    add_textbox(s, title, x+Inches(0.12), by+Inches(0.15),
                bw-Inches(0.2), Inches(1.0),
                font_size=14, bold=True, color=clr, font_name="Calibri")
    add_textbox(s, body, x+Inches(0.12), by+Inches(1.2),
                bw-Inches(0.2), Inches(3.8),
                font_size=12, color=RGBColor(0xD0,0xE8,0xDA), font_name="Calibri")

slide_number(s, 4)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 5 – GRI 101 ARCHITECTURE OVERVIEW
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "03  |  GRI 101 ARCHITECTURE",
            Inches(0.4), Inches(0.15), Inches(8), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "Eight Disclosures — One Integrated Framework",
            Inches(0.4), Inches(0.55), Inches(11), Inches(0.6),
            font_size=26, bold=True, color=C_WHITE, font_name="Calibri")

# Central label
add_rect(s, Inches(5.2), Inches(3.0), Inches(2.9), Inches(1.5),
         fill_color=C_GREEN)
add_textbox(s, "GRI 101\nBiodiversity\n2024",
            Inches(5.3), Inches(3.05), Inches(2.7), Inches(1.4),
            font_size=15, bold=True, color=C_WHITE,
            align=PP_ALIGN.CENTER, font_name="Calibri")

disclosures = [
    ("101-1", "Policies to halt &\nreverse biodiversity loss",  C_GOLD,
     Inches(0.3), Inches(1.55)),
    ("101-2", "Management of\nbiodiversity impacts",            C_LIME,
     Inches(0.3), Inches(3.15)),
    ("101-3", "Access &\nbenefit-sharing (ABS)",                C_CORAL,
     Inches(0.3), Inches(4.75)),
    ("101-4", "Identification of\nbiodiversity impacts",        C_SAGE,
     Inches(0.3), Inches(6.15)),
    ("101-5", "Locations with\nbiodiversity impacts",           C_GOLD,
     Inches(9.8), Inches(1.55)),
    ("101-6", "Direct drivers of\nbiodiversity loss",           C_LIME,
     Inches(9.8), Inches(3.15)),
    ("101-7", "Changes to the\nstate of biodiversity",          C_CORAL,
     Inches(9.8), Inches(4.75)),
    ("101-8", "Ecosystem\nservices",                            C_SAGE,
     Inches(9.8), Inches(6.15)),
]

for (code, title, clr, x, y) in disclosures:
    bw = Inches(3.2)
    bh = Inches(1.25)
    add_rect(s, x, y, bw, bh, fill_color=RGBColor(0x18,0x3A,0x28))
    accent_bar(s, x, y, bw, Inches(0.055), clr)
    add_textbox(s, code, x+Inches(0.1), y+Inches(0.1),
                Inches(0.8), Inches(0.4),
                font_size=13, bold=True, color=clr, font_name="Calibri")
    add_textbox(s, title, x+Inches(0.1), y+Inches(0.48),
                bw-Inches(0.2), Inches(0.7),
                font_size=12, color=C_WHITE, font_name="Calibri")

# connecting lines label
add_textbox(s, "STRATEGIC\nPOLICY",
            Inches(3.65), Inches(2.0), Inches(1.5), Inches(0.7),
            font_size=10, bold=True, color=C_MID_GREY,
            align=PP_ALIGN.CENTER, font_name="Calibri")
add_textbox(s, "SITE-LEVEL\nOPERATIONS",
            Inches(8.4), Inches(2.0), Inches(1.5), Inches(0.7),
            font_size=10, bold=True, color=C_MID_GREY,
            align=PP_ALIGN.CENTER, font_name="Calibri")

slide_number(s, 5)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 6 – DISCLOSURES 101-1 & 101-2
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "04  |  DEEP DIVE — POLICY & MANAGEMENT",
            Inches(0.4), Inches(0.15), Inches(9), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "101-1: Policies  |  101-2: Management & Mitigation Hierarchy",
            Inches(0.4), Inches(0.55), Inches(11), Inches(0.6),
            font_size=24, bold=True, color=C_WHITE, font_name="Calibri")

# Left — 101-1
add_rect(s, Inches(0.35), Inches(1.5), Inches(6.1), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(0.35), Inches(1.5), Inches(6.1), Inches(0.08), C_GOLD)
add_textbox(s, "101-1  Policies to Halt & Reverse Biodiversity Loss",
            Inches(0.5), Inches(1.62), Inches(5.8), Inches(0.5),
            font_size=15, bold=True, color=C_GOLD, font_name="Calibri")

p1_items = [
    "► Describe commitments / policies aligned with KM-GBF 2050 Goals & 2030 Targets",
    "► Scope must cover both OWN ACTIVITIES and BUSINESS RELATIONSHIPS (value chain)",
    "► Quantified goals & targets with base year and scientific basis",
    "► Progress indicators: e.g. Ha restored, species monitored, % supply chain traced",
    "► Critical link: cite specific KM-GBF Targets (e.g. Target 3 for 30×30, Target 15 for business)",
    "► Absence of a policy is reportable — state reason for omission",
]
for i, txt in enumerate(p1_items):
    add_textbox(s, txt,
                Inches(0.5), Inches(2.25) + i * Inches(0.72),
                Inches(5.7), Inches(0.65),
                font_size=12, color=C_LIGHT_GREY, font_name="Calibri")

# Right — Mitigation Hierarchy
add_rect(s, Inches(6.75), Inches(1.5), Inches(6.25), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(6.75), Inches(1.5), Inches(6.25), Inches(0.08), C_LIME)
add_textbox(s, "101-2  The Mitigation Hierarchy",
            Inches(6.9), Inches(1.62), Inches(5.8), Inches(0.5),
            font_size=15, bold=True, color=C_LIME, font_name="Calibri")

hier = [
    (C_CORAL,  "1. AVOID",     "Proactive design and siting to prevent\nnew negative biodiversity impacts",     "Highest Priority"),
    (C_GOLD,   "2. MINIMIZE",  "Reduce duration, intensity & extent\nof unavoidable impacts",                   ""),
    (C_LIME,   "3. RESTORE",   "Rehabilitate degraded ecosystems at\nor near sites of operation",               ""),
    (C_GREEN,  "4. OFFSET",    "Compensate for residual impacts using\ncertified, in-kind biodiversity offsets", ""),
    (C_SAGE,   "5. TRANSFORM", "Systemic change & additional conservation\nactions beyond footprint",            "Lowest Priority"),
]
for i, (clr, lbl, desc, note) in enumerate(hier):
    hy = Inches(2.15) + i * Inches(0.92)
    add_rect(s, Inches(6.85), hy, Inches(1.3), Inches(0.78),
             fill_color=clr)
    add_textbox(s, lbl, Inches(6.87), hy + Inches(0.18),
                Inches(1.26), Inches(0.5),
                font_size=10, bold=True, color=C_DARK,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    add_textbox(s, desc, Inches(8.3), hy + Inches(0.12),
                Inches(4.45), Inches(0.65),
                font_size=11.5, color=C_WHITE, font_name="Calibri")
    if note:
        add_textbox(s, f"◄ {note}", Inches(8.3), hy + Inches(0.55),
                    Inches(3.5), Inches(0.28),
                    font_size=9.5, italic=True, color=C_GOLD, font_name="Calibri")

slide_number(s, 6)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 7 – DISCLOSURES 101-3 & 101-4
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "04  |  DEEP DIVE — ABS & IMPACT IDENTIFICATION",
            Inches(0.4), Inches(0.15), Inches(9), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "101-3: Access & Benefit-Sharing  |  101-4: Identification Methodology",
            Inches(0.4), Inches(0.55), Inches(12), Inches(0.6),
            font_size=24, bold=True, color=C_WHITE, font_name="Calibri")

# 101-3 card
add_rect(s, Inches(0.35), Inches(1.5), Inches(6.1), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(0.35), Inches(1.5), Inches(6.1), Inches(0.08), C_CORAL)
add_textbox(s, "101-3  Access & Benefit-Sharing (ABS)",
            Inches(0.5), Inches(1.62), Inches(5.8), Inches(0.5),
            font_size=15, bold=True, color=C_CORAL, font_name="Calibri")

abs_items = [
    ("What is ABS?",
     "Legal framework under the Nagoya Protocol ensuring equitable sharing of benefits arising from use of genetic resources (plants, animals, microbes)."),
    ("Disclosure Requirements",
     "Describe compliance process with ABS laws. Report voluntary actions taken beyond legal minimum. Disclose activities on the High Seas (Areas Beyond National Jurisdiction)."),
    ("Why It Matters",
     "Non-compliance carries legal and reputational risk — especially for pharma, biotech, food & beverage, cosmetics sectors. Indigenous communities are key stakeholders."),
]
for i, (hdr, body) in enumerate(abs_items):
    add_textbox(s, hdr,
                Inches(0.5), Inches(2.25) + i * Inches(1.45),
                Inches(5.7), Inches(0.4),
                font_size=13, bold=True, color=C_CORAL, font_name="Calibri")
    add_textbox(s, body,
                Inches(0.5), Inches(2.7) + i * Inches(1.45),
                Inches(5.7), Inches(0.9),
                font_size=12, color=C_LIGHT_GREY, font_name="Calibri")

# 101-4 card
add_rect(s, Inches(6.75), Inches(1.5), Inches(6.25), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(6.75), Inches(1.5), Inches(6.25), Inches(0.08), C_SAGE)
add_textbox(s, "101-4  Identification of Biodiversity Impacts",
            Inches(6.9), Inches(1.62), Inches(5.8), Inches(0.5),
            font_size=15, bold=True, color=C_SAGE, font_name="Calibri")

id_blocks = [
    ("Methodology",
     "Explain HOW the organisation determined which sites and supply-chain products have the most significant impacts. Define thresholds used."),
    ("Leading Tools & Data Sources",
     "• ENCORE (Natural Capital Finance Alliance)\n• SBTN (Science Based Targets for Nature)\n• IUCN Red List & CITES databases\n• WWF Biodiversity Risk Filter\n• TNFD LEAP Framework\n• Satellite / geospatial overlays"),
    ("Data Quality Disclosure",
     "Distinguish primary vs secondary/modelled data. Report plans to improve data accuracy over time — a key analyst and investor question."),
]
for i, (hdr, body) in enumerate(id_blocks):
    add_textbox(s, hdr,
                Inches(6.9), Inches(2.25) + i * Inches(1.6),
                Inches(5.9), Inches(0.4),
                font_size=13, bold=True, color=C_SAGE, font_name="Calibri")
    add_textbox(s, body,
                Inches(6.9), Inches(2.68) + i * Inches(1.6),
                Inches(5.9), Inches(1.05),
                font_size=12, color=C_LIGHT_GREY, font_name="Calibri")

slide_number(s, 7)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 8 – DISCLOSURES 101-5 & 101-6
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "04  |  DEEP DIVE — LOCATIONS & DIRECT DRIVERS",
            Inches(0.4), Inches(0.15), Inches(9), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "101-5: Location Impacts  |  101-6: The Five Direct Drivers of Loss",
            Inches(0.4), Inches(0.55), Inches(12), Inches(0.6),
            font_size=24, bold=True, color=C_WHITE, font_name="Calibri")

# 101-5
add_rect(s, Inches(0.35), Inches(1.5), Inches(4.6), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(0.35), Inches(1.5), Inches(4.6), Inches(0.07), C_GOLD)
add_textbox(s, "101-5  Locations",
            Inches(0.5), Inches(1.62), Inches(4.3), Inches(0.45),
            font_size=15, bold=True, color=C_GOLD, font_name="Calibri")
loc_rows = [
    ("In / Near ecologically sensitive areas?", "Classify per site"),
    ("Distance to protected area (km)",         "If site is 'Near'"),
    ("Areas of biodiversity importance",         "e.g. KBA, IBA"),
    ("High ecosystem integrity zones",           "Primary habitat"),
    ("Rapidly declining ecosystem integrity",    "Degradation hotspots"),
    ("High physical water risk areas",           "WRI Aqueduct link"),
    ("Services to IP & local communities",       "Free, Prior, Informed"),
    ("Size of eco-sensitive area (Ha)",          "Quantified boundary"),
]
add_textbox(s, "Per-site data required:",
            Inches(0.5), Inches(2.15), Inches(4.2), Inches(0.35),
            font_size=11, bold=True, color=C_MID_GREY, font_name="Calibri")
for i, (field, note) in enumerate(loc_rows):
    add_textbox(s, f"• {field}",
                Inches(0.5), Inches(2.55) + i * Inches(0.38),
                Inches(3.3), Inches(0.35),
                font_size=11, color=C_LIGHT_GREY, font_name="Calibri")
    add_textbox(s, note,
                Inches(3.8), Inches(2.55) + i * Inches(0.38),
                Inches(1.0), Inches(0.35),
                font_size=9, italic=True, color=C_SAGE, font_name="Calibri")

# 101-6 — Five Drivers
add_rect(s, Inches(5.2), Inches(1.5), Inches(7.8), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(5.2), Inches(1.5), Inches(7.8), Inches(0.07), C_LIME)
add_textbox(s, "101-6  Five Direct Drivers of Biodiversity Loss",
            Inches(5.35), Inches(1.62), Inches(7.5), Inches(0.45),
            font_size=15, bold=True, color=C_LIME, font_name="Calibri")

drivers = [
    (C_CORAL, "Land / Sea Use Change",
     "Natural ecosystem converted (Ha), cut-off date, ecosystem type before & after. Intensive use conversion also reported."),
    (C_GOLD,  "Resource Over-exploitation",
     "Wild species harvested — type, quantity, extinction risk classification (IUCN Red List). Include fisheries and timber."),
    (C_LIME,  "Climate Change",
     "Cross-reference to GHG data (GRI 305-7 air pollutants). Report water withdrawal/consumption (ML) per site."),
    (C_SAGE,  "Pollution",
     "Pollutants by type and quantity. Cross-reference GRI 303 (water) and GRI 306 (spills). Nutrient loading key metric."),
    (C_CORAL, "Invasive Alien Species",
     "How IAS were introduced (pathway analysis). Report prevention and control measures taken."),
]
for i, (clr, drv, desc) in enumerate(drivers):
    dy = Inches(2.2) + i * Inches(0.96)
    add_rect(s, Inches(5.3), dy, Inches(1.5), Inches(0.82), fill_color=clr)
    add_textbox(s, f"#{i+1}", Inches(5.35), dy + Inches(0.2),
                Inches(0.6), Inches(0.45),
                font_size=18, bold=True, color=C_DARK,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    add_textbox(s, drv, Inches(7.0), dy + Inches(0.05),
                Inches(2.8), Inches(0.35),
                font_size=12, bold=True, color=clr, font_name="Calibri")
    add_textbox(s, desc, Inches(7.0), dy + Inches(0.42),
                Inches(5.75), Inches(0.45),
                font_size=10.5, color=C_LIGHT_GREY, font_name="Calibri")

slide_number(s, 8)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 9 – DISCLOSURES 101-7 & 101-8
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "04  |  DEEP DIVE — STATE OF BIODIVERSITY & ECOSYSTEM SERVICES",
            Inches(0.4), Inches(0.15), Inches(11), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "101-7: State Changes  |  101-8: Ecosystem Services",
            Inches(0.4), Inches(0.55), Inches(12), Inches(0.6),
            font_size=24, bold=True, color=C_WHITE, font_name="Calibri")

# 101-7
add_rect(s, Inches(0.35), Inches(1.5), Inches(6.1), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(0.35), Inches(1.5), Inches(6.1), Inches(0.07), C_CORAL)
add_textbox(s, "101-7  Changes to the State of Biodiversity",
            Inches(0.5), Inches(1.62), Inches(5.7), Inches(0.45),
            font_size=15, bold=True, color=C_CORAL, font_name="Calibri")

state_content = [
    ("What Must Be Reported",
     "For each site: ecosystem type, size (Ha), and condition — comparing BASE YEAR vs CURRENT PERIOD. Multiple ecosystem types per site require separate rows."),
    ("Ecosystem Size",
     "Ha of affected ecosystem. Enables tracking shrinkage or recovery over time. Use satellite imagery, field surveys or modelled estimates with appropriate caveats."),
    ("Ecosystem Condition",
     "Qualitative or quantitative assessment (e.g. Habitat Condition Score, BHM, IBAT). Must state the method used. Trajectory (improving / stable / declining) is key."),
    ("Data Transparency",
     "Specify whether data is primary (field) or secondary (modelled/national databases). Report known limitations. Auditor comfort requires methodology documentation."),
]
for i, (hdr, body) in enumerate(state_content):
    add_textbox(s, hdr,
                Inches(0.5), Inches(2.2) + i * Inches(1.2),
                Inches(5.7), Inches(0.38),
                font_size=13, bold=True, color=C_CORAL, font_name="Calibri")
    add_textbox(s, body,
                Inches(0.5), Inches(2.62) + i * Inches(1.2),
                Inches(5.7), Inches(0.75),
                font_size=11.5, color=C_LIGHT_GREY, font_name="Calibri")

# 101-8
add_rect(s, Inches(6.75), Inches(1.5), Inches(6.25), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(6.75), Inches(1.5), Inches(6.25), Inches(0.07), C_SAGE)
add_textbox(s, "101-8  Ecosystem Services",
            Inches(6.9), Inches(1.62), Inches(5.9), Inches(0.45),
            font_size=15, bold=True, color=C_SAGE, font_name="Calibri")

es_cats = [
    (C_LIME,  "Provisioning Services",
     "Food, water, raw materials, genetic resources, medicinal plants. Directly tied to supply chain and ABS."),
    (C_GOLD,  "Regulating & Maintenance",
     "Climate regulation, water purification, flood protection, pollination, carbon sequestration, disease control."),
    (C_CORAL, "Cultural Services",
     "Recreation, tourism, spiritual, educational. Often critical for local community stakeholder engagement."),
]
add_textbox(s, "Three Service Categories:",
            Inches(6.9), Inches(2.15), Inches(5.9), Inches(0.35),
            font_size=12, bold=True, color=C_MID_GREY, font_name="Calibri")
for i, (clr, cat, desc) in enumerate(es_cats):
    cy = Inches(2.6) + i * Inches(1.25)
    add_rect(s, Inches(6.85), cy, Inches(0.25), Inches(0.85),
             fill_color=clr)
    add_textbox(s, cat, Inches(7.25), cy + Inches(0.05),
                Inches(5.5), Inches(0.35),
                font_size=12, bold=True, color=clr, font_name="Calibri")
    add_textbox(s, desc, Inches(7.25), cy + Inches(0.42),
                Inches(5.5), Inches(0.55),
                font_size=11.5, color=C_LIGHT_GREY, font_name="Calibri")

add_textbox(s, "Also Required: Identify BENEFICIARIES (and numbers where known) and explain HOW the service and beneficiaries are/could be affected. This is the highest-disclosure-maturity element.",
            Inches(6.9), Inches(6.55), Inches(6.0), Inches(0.8),
            font_size=11, italic=True, color=C_GOLD, font_name="Calibri")

slide_number(s, 9)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 10 – MATERIALITY & DATA QUALITY
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "05  |  MATERIALITY & DATA QUALITY",
            Inches(0.4), Inches(0.15), Inches(8), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "The Rigour Gap — Where Most Organisations Fall Short",
            Inches(0.4), Inches(0.55), Inches(12), Inches(0.6),
            font_size=26, bold=True, color=C_WHITE, font_name="Calibri")

# Table header
hy = Inches(1.55)
col_xs = [Inches(0.35), Inches(3.0), Inches(6.2), Inches(9.6)]
col_ws = [Inches(2.55), Inches(3.1), Inches(3.3), Inches(3.4)]
hdrs   = ["ELEMENT", "COMMON PITFALL", "BEST PRACTICE", "MATURITY LEVEL"]
hdr_cols = [C_GOLD, C_CORAL, C_LIME, C_SAGE]

for cx, cw, hdr, hc in zip(col_xs, col_ws, hdrs, hdr_cols):
    add_rect(s, cx, hy, cw, Inches(0.45), fill_color=C_MID)
    add_textbox(s, hdr, cx + Inches(0.08), hy + Inches(0.08),
                cw - Inches(0.15), Inches(0.35),
                font_size=11, bold=True, color=hc, font_name="Calibri")

rows_data = [
    ("Materiality\nAssessment",     "Binary yes/no\ndetermination",
     "Quantified scoring against proximity to protected areas, sector dependencies & regulatory exposure",
     "★★★★★"),
    ("Site Boundary\nDefinition",   "Operational fence-line only",
     "TNFD LEAP 'Locate' phase — expand to area of influence, upstream supply sheds, downstream water users",
     "★★★★☆"),
    ("Ecosystem\nCondition",        "Qualitative narrative\n('good', 'fair', 'poor')",
     "Scored metric (e.g. BHM, HII) with temporal baseline. Must state measurement methodology.",
     "★★★★★"),
    ("Supply Chain\nTraceability",  "Tier-1 only",
     "Identify high-impact commodities (soy, palm, beef, timber, cocoa) to country-of-origin minimum",
     "★★★☆☆"),
    ("Data\nVerification",          "No third-party check",
     "Independent verification against SBTN, IUCN or TNFD. Limited assurance at minimum for high-material sites.",
     "★★★★☆"),
]

for r, (elem, pitfall, best, stars) in enumerate(rows_data):
    ry = Inches(2.05) + r * Inches(0.95)
    bg = RGBColor(0x16, 0x36, 0x26) if r % 2 == 0 else RGBColor(0x13, 0x2E, 0x20)
    for cx, cw in zip(col_xs, col_ws):
        add_rect(s, cx, ry, cw, Inches(0.9), fill_color=bg)
    add_textbox(s, elem,     col_xs[0]+Inches(0.08), ry+Inches(0.1),
                col_ws[0]-Inches(0.15), Inches(0.75),
                font_size=11, bold=True, color=C_WHITE, font_name="Calibri")
    add_textbox(s, pitfall,  col_xs[1]+Inches(0.08), ry+Inches(0.1),
                col_ws[1]-Inches(0.15), Inches(0.75),
                font_size=11, color=C_CORAL, font_name="Calibri")
    add_textbox(s, best,     col_xs[2]+Inches(0.08), ry+Inches(0.1),
                col_ws[2]-Inches(0.15), Inches(0.75),
                font_size=10.5, color=C_LIGHT_GREY, font_name="Calibri")
    add_textbox(s, stars,    col_xs[3]+Inches(0.08), ry+Inches(0.22),
                col_ws[3]-Inches(0.15), Inches(0.45),
                font_size=16, color=C_GOLD, font_name="Calibri")

slide_number(s, 10)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 11 – BIODIVERSITY × CLIMATE × SOCIAL NEXUS
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "06  |  THE NATURE-CLIMATE-PEOPLE NEXUS",
            Inches(0.4), Inches(0.15), Inches(10), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "Biodiversity Does Not Exist in Isolation — Integrating GRI 101 with Your ESG Stack",
            Inches(0.4), Inches(0.55), Inches(12.5), Inches(0.6),
            font_size=22, bold=True, color=C_WHITE, font_name="Calibri")

# Venn-style 3-circle representation using overlapping rectangles with transparency labels
circles = [
    (Inches(1.0), Inches(2.1), C_GREEN,  "BIODIVERSITY\n(GRI 101)"),
    (Inches(4.8), Inches(2.1), C_CORAL,  "CLIMATE\n(GRI 305 / TCFD)"),
    (Inches(2.9), Inches(4.4), C_GOLD,   "SOCIAL\n(GRI 411 / UNGP)"),
]
for (x, y, clr, lbl) in circles:
    add_rect(s, x, y, Inches(3.5), Inches(2.5), fill_color=clr)
    add_textbox(s, lbl, x+Inches(0.3), y+Inches(0.7),
                Inches(2.9), Inches(1.0),
                font_size=16, bold=True, color=C_WHITE,
                align=PP_ALIGN.CENTER, font_name="Calibri")

# Nexus labels
add_textbox(s, "Nature-based\nClimate Solutions\n(NbS)",
            Inches(3.0), Inches(2.3), Inches(2.2), Inches(0.9),
            font_size=10, bold=True, color=C_DARK,
            align=PP_ALIGN.CENTER, font_name="Calibri")
add_textbox(s, "Indigenous\nRights & ABS",
            Inches(1.6), Inches(4.2), Inches(2.0), Inches(0.7),
            font_size=10, bold=True, color=C_DARK,
            align=PP_ALIGN.CENTER, font_name="Calibri")
add_textbox(s, "Water Stress\n& Ecosystem\nServices",
            Inches(4.8), Inches(4.2), Inches(2.0), Inches(0.9),
            font_size=10, bold=True, color=C_DARK,
            align=PP_ALIGN.CENTER, font_name="Calibri")
add_textbox(s, "TRIPLE\nNEXUS",
            Inches(3.1), Inches(4.0), Inches(1.8), Inches(0.7),
            font_size=12, bold=True, color=C_DARK,
            align=PP_ALIGN.CENTER, font_name="Calibri")

# Right — Integration guidance
rx = Inches(8.5)
add_rect(s, rx, Inches(1.5), Inches(4.5), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, rx, Inches(1.5), Inches(4.5), Inches(0.07), C_LIME)
add_textbox(s, "Integration Imperatives",
            rx+Inches(0.15), Inches(1.65), Inches(4.2), Inches(0.45),
            font_size=15, bold=True, color=C_LIME, font_name="Calibri")

integ = [
    ("Cross-reference GRI 305",
     "Report air pollutants (305-7) per site in your 101-6 drivers — same data point, dual disclosure value."),
    ("Cross-reference GRI 303 & 306",
     "Water withdrawal/consumption and spills are direct biodiversity drivers. Eliminate double-handling."),
    ("Cross-reference GRI 411",
     "Rights of Indigenous Peoples apply directly to ABS compliance and ecosystem service beneficiary identification."),
    ("TCFD / TNFD Alignment",
     "Physical and transition risks map to both climate and biodiversity. Combined scenario analysis is best practice."),
]
for i, (hdr, body) in enumerate(integ):
    iy = Inches(2.2) + i * Inches(1.15)
    add_textbox(s, f"▸  {hdr}", rx+Inches(0.15), iy,
                Inches(4.15), Inches(0.38),
                font_size=12, bold=True, color=C_GOLD, font_name="Calibri")
    add_textbox(s, body, rx+Inches(0.15), iy+Inches(0.38),
                Inches(4.15), Inches(0.65),
                font_size=11, color=C_LIGHT_GREY, font_name="Calibri")

slide_number(s, 11)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 12 – IMPLEMENTATION ROADMAP
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "07  |  IMPLEMENTATION ROADMAP",
            Inches(0.4), Inches(0.15), Inches(8), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "From Compliance Gap to Nature-Positive — A 3-Phase Journey",
            Inches(0.4), Inches(0.55), Inches(12), Inches(0.6),
            font_size=26, bold=True, color=C_WHITE, font_name="Calibri")

phases = [
    (C_GOLD,  "PHASE 1\nFoundation",  "Months 1–3",
     ["Conduct biodiversity materiality assessment",
      "Map all operational sites vs protected areas (IBAT / TNFD LEAP)",
      "Audit existing policies for KM-GBF alignment",
      "Establish data governance and reporting team",
      "Define site list for 101-5 through 101-8",
      "Baseline ecosystem condition per site"]),
    (C_LIME,  "PHASE 2\nBuild",       "Months 4–8",
     ["Develop / update biodiversity policy (101-1)",
      "Implement mitigation hierarchy per site (101-2)",
      "Conduct ABS compliance review (101-3)",
      "Engage supply chain — trace high-impact commodities",
      "Collect direct driver data per site (101-6)",
      "Engage ecosystem service beneficiaries (101-8)"]),
    (C_CORAL, "PHASE 3\nReport & Improve", "Months 9–12+",
     ["Compile GRI 101 disclosure narrative",
      "Obtain limited assurance on high-material sites",
      "Set science-based targets (SBTN framework)",
      "Integrate into Annual / Sustainability Report",
      "Establish year-on-year tracking dashboard",
      "Engage investors & ratings agencies proactively"]),
]

pw = Inches(4.0)
gap = Inches(0.25)
start_x = Inches(0.35)

for i, (clr, title, timing, steps) in enumerate(phases):
    px = start_x + i * (pw + gap)
    ph = Inches(5.7)
    add_rect(s, px, Inches(1.5), pw, ph,
             fill_color=RGBColor(0x18,0x3A,0x28))
    add_rect(s, px, Inches(1.5), pw, Inches(1.1), fill_color=clr)
    add_textbox(s, title, px+Inches(0.15), Inches(1.55),
                pw-Inches(0.3), Inches(0.65),
                font_size=18, bold=True, color=C_DARK,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    add_textbox(s, timing, px+Inches(0.15), Inches(2.45),
                pw-Inches(0.3), Inches(0.35),
                font_size=12, italic=True, color=clr,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    for j, step in enumerate(steps):
        add_textbox(s, f"◉  {step}",
                    px+Inches(0.15), Inches(2.85) + j * Inches(0.7),
                    pw-Inches(0.3), Inches(0.65),
                    font_size=11, color=C_LIGHT_GREY, font_name="Calibri")

# Arrow connectors
add_textbox(s, "→", Inches(4.45), Inches(3.7), Inches(0.2), Inches(0.5),
            font_size=24, bold=True, color=C_GOLD,
            align=PP_ALIGN.CENTER, font_name="Calibri")
add_textbox(s, "→", Inches(8.9), Inches(3.7), Inches(0.2), Inches(0.5),
            font_size=24, bold=True, color=C_GOLD,
            align=PP_ALIGN.CENTER, font_name="Calibri")

slide_number(s, 12)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 13 – BUSINESS CASE
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "08  |  THE BUSINESS CASE",
            Inches(0.4), Inches(0.15), Inches(8), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "Risks, Opportunities & the Investor Signal",
            Inches(0.4), Inches(0.55), Inches(12), Inches(0.6),
            font_size=26, bold=True, color=C_WHITE, font_name="Calibri")

# Risks column
add_rect(s, Inches(0.35), Inches(1.5), Inches(4.05), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(0.35), Inches(1.5), Inches(4.05), Inches(0.07), C_CORAL)
add_textbox(s, "RISKS OF INACTION",
            Inches(0.5), Inches(1.62), Inches(3.8), Inches(0.45),
            font_size=14, bold=True, color=C_CORAL, font_name="Calibri")

risks = [
    ("Regulatory Risk", "GRI 101 non-compliance; EU CSRD penalties; stranded assets in nature-dependent sectors"),
    ("Physical Risk", "Supply chain disruption, water insecurity, crop failure from ecosystem collapse"),
    ("Reputational Risk", "Greenwashing litigation, activist investor pressure, consumer backlash"),
    ("Financial Risk", "TNFD adoption by banks: biodiversity risk embedded in loan covenants by 2027"),
    ("Legal Risk", "Nagoya Protocol ABS violations; indigenous community legal challenges"),
]
for i, (hdr, body) in enumerate(risks):
    ry = Inches(2.2) + i * Inches(0.95)
    add_textbox(s, f"⚠  {hdr}", Inches(0.5), ry,
                Inches(3.75), Inches(0.35),
                font_size=12, bold=True, color=C_CORAL, font_name="Calibri")
    add_textbox(s, body, Inches(0.5), ry+Inches(0.37),
                Inches(3.75), Inches(0.52),
                font_size=10.5, color=C_LIGHT_GREY, font_name="Calibri")

# Opportunities column
add_rect(s, Inches(4.65), Inches(1.5), Inches(4.05), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(4.65), Inches(1.5), Inches(4.05), Inches(0.07), C_LIME)
add_textbox(s, "OPPORTUNITIES FOR LEADERS",
            Inches(4.8), Inches(1.62), Inches(3.8), Inches(0.45),
            font_size=14, bold=True, color=C_LIME, font_name="Calibri")

opps = [
    ("Green Financing", "$700B annual biodiversity finance gap = enormous opportunity for NbS-linked bonds and blended finance"),
    ("Nature-based Products", "Carbon + biodiversity credits (stacked), green commodities premiums, eco-certification"),
    ("Cost Reduction", "Ecosystem restoration reduces water treatment costs, disaster risk and input volatility"),
    ("Investor Premium", "ESG-oriented AUM now >$35T; TNFD-aligned companies attract lower cost of capital"),
    ("Competitive Moat", "Early movers shape regulation, set industry standards and secure secure market access"),
]
for i, (hdr, body) in enumerate(opps):
    ry = Inches(2.2) + i * Inches(0.95)
    add_textbox(s, f"✦  {hdr}", Inches(4.8), ry,
                Inches(3.75), Inches(0.35),
                font_size=12, bold=True, color=C_LIME, font_name="Calibri")
    add_textbox(s, body, Inches(4.8), ry+Inches(0.37),
                Inches(3.75), Inches(0.52),
                font_size=10.5, color=C_LIGHT_GREY, font_name="Calibri")

# Investor signal column
add_rect(s, Inches(9.0), Inches(1.5), Inches(4.0), Inches(5.6),
         fill_color=RGBColor(0x18,0x3A,0x28))
accent_bar(s, Inches(9.0), Inches(1.5), Inches(4.0), Inches(0.07), C_GOLD)
add_textbox(s, "INVESTOR SIGNAL",
            Inches(9.15), Inches(1.62), Inches(3.7), Inches(0.45),
            font_size=14, bold=True, color=C_GOLD, font_name="Calibri")

signals = [
    ("320+ organisations", "already adopting TNFD voluntary disclosures as of 2024"),
    ("MSCI & Sustainalytics", "now scoring nature-related risks in company ratings"),
    ("CDP Nature Module", "live since 2023 — aligns directly with GRI 101"),
    ("Science Based Targets\nfor Nature (SBTN)", "investor-grade targets expected as standard by 2027"),
    ("EU Taxonomy\nNature Screen", "biodiversity DNSH criteria now applicable — affects green bond issuance"),
]
for i, (hdr, body) in enumerate(signals):
    ry = Inches(2.2) + i * Inches(0.95)
    add_textbox(s, hdr, Inches(9.15), ry,
                Inches(3.7), Inches(0.4),
                font_size=12, bold=True, color=C_GOLD, font_name="Calibri")
    add_textbox(s, body, Inches(9.15), ry+Inches(0.38),
                Inches(3.7), Inches(0.52),
                font_size=10.5, color=C_LIGHT_GREY, font_name="Calibri")

slide_number(s, 13)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 14 – STRATEGIC RECOMMENDATIONS
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(1.3), fill_color=C_MID)
add_textbox(s, "09  |  STRATEGIC RECOMMENDATIONS",
            Inches(0.4), Inches(0.15), Inches(8), Inches(0.4),
            font_size=11, bold=True, color=C_LIME, font_name="Calibri")
add_textbox(s, "Five Priorities to Build a Nature-Positive Organisation",
            Inches(0.4), Inches(0.55), Inches(12), Inches(0.6),
            font_size=26, bold=True, color=C_WHITE, font_name="Calibri")

recs = [
    (C_GOLD, "01", "Commit at Board Level — Now",
     "Biodiversity must be a Board-level agenda item by 2026. Appoint a Nature Officer or integrate into the CRO / CSO remit. Set a public, science-aligned nature-positive commitment with a 2030 milestone.",
     "URGENCY: Immediate",
     "GRI 101-1 | KM-GBF Target 15"),

    (C_LIME, "02", "Map & Baseline Your Footprint",
     "Run TNFD LEAP for all operations and top-50 supply chain partners. Produce a heatmap of nature risk by site. This is the indispensable foundation — without it, all other disclosures are guesswork.",
     "URGENCY: Q1 2026",
     "GRI 101-4, 101-5, 101-6"),

    (C_CORAL, "03", "Embed the Mitigation Hierarchy in Capital Allocation",
     "Make 'Avoid first, offset last' a binding requirement for any CapEx affecting land, water or marine ecosystems. Establish a biodiversity impact screening gate equivalent to your TCFD climate risk gate.",
     "URGENCY: Q2 2026",
     "GRI 101-2 | TNFD LEAP"),

    (C_SAGE, "04", "Trace and Transform Your Supply Chain",
     "Prioritise high-impact commodities (soy, palm, beef, timber, cocoa, aquaculture). Achieve country-of-origin traceability minimum. Engage top-20 suppliers in joint nature-positive transition plans.",
     "URGENCY: Q3 2026",
     "GRI 101-6-e | Table 4"),

    (C_GOLD, "05", "Report With Rigour & Seek Assurance",
     "Publish your first GRI 101-aligned disclosure with limited third-party assurance on high-material sites. Integrate with CDP Nature, TNFD and ESRS E4. Set year-on-year improvement targets publicly.",
     "URGENCY: Annual Report 2026",
     "GRI 101-1 to 101-8"),
]

for i, (clr, num, title, body, urgency, ref) in enumerate(recs):
    row = i // 3
    col = i % 3
    if i < 3:
        rx = Inches(0.35) + col * Inches(4.35)
        ry = Inches(1.55)
    else:
        rx = Inches(0.35) + (i - 3) * Inches(6.55)
        ry = Inches(4.3)

    rw = Inches(4.05) if i < 3 else Inches(6.15)
    rh = Inches(2.55) if i < 3 else Inches(2.85)

    add_rect(s, rx, ry, rw, rh, fill_color=RGBColor(0x18,0x3A,0x28))
    accent_bar(s, rx, ry, rw, Inches(0.06), clr)
    add_rect(s, rx + Inches(0.12), ry + Inches(0.12),
             Inches(0.55), Inches(0.55), fill_color=clr)
    add_textbox(s, num, rx+Inches(0.13), ry+Inches(0.14),
                Inches(0.5), Inches(0.4),
                font_size=16, bold=True, color=C_DARK,
                align=PP_ALIGN.CENTER, font_name="Calibri")
    add_textbox(s, title, rx+Inches(0.82), ry+Inches(0.18),
                rw-Inches(1.0), Inches(0.45),
                font_size=13, bold=True, color=clr, font_name="Calibri")
    add_textbox(s, body, rx+Inches(0.15), ry+Inches(0.72),
                rw-Inches(0.3), Inches(1.4),
                font_size=10.5, color=C_LIGHT_GREY, font_name="Calibri")
    add_textbox(s, urgency, rx+Inches(0.15), ry+rh-Inches(0.55),
                rw*0.5, Inches(0.38),
                font_size=10, bold=True, italic=True, color=C_CORAL, font_name="Calibri")
    add_textbox(s, ref, rx+rw*0.5, ry+rh-Inches(0.55),
                rw*0.5, Inches(0.38),
                font_size=10, italic=True, color=C_MID_GREY, font_name="Calibri")

slide_number(s, 14)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 15 – CLOSING / Q&A
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_bg(s, C_MID)

# Large decorative background text
add_textbox(s, "NATURE\nPOSITIVE",
            Inches(-0.3), Inches(1.8), Inches(14),
            Inches(5.0),
            font_size=120, bold=True,
            color=RGBColor(0x25, 0x5C, 0x3D),
            align=PP_ALIGN.CENTER, font_name="Calibri")

accent_bar(s, Inches(3.0), Inches(3.35), Inches(7.33), Inches(0.08), C_GOLD)

add_textbox(s, "The question is no longer whether to act.\nThe question is whether you act first — or last.",
            Inches(1.5), Inches(2.9), Inches(10.33), Inches(1.1),
            font_size=22, italic=True, color=C_WHITE,
            align=PP_ALIGN.CENTER, font_name="Georgia")

add_textbox(s, "QUESTIONS & DISCUSSION",
            Inches(3.5), Inches(4.1), Inches(6.33), Inches(0.6),
            font_size=28, bold=True, color=C_GOLD,
            align=PP_ALIGN.CENTER, font_name="Calibri")

add_textbox(s, "GRI 101: Biodiversity 2024  |  Effective 1 January 2026\nStrategic Advisory Engagement  |  June 2026",
            Inches(2.5), Inches(5.6), Inches(8.33), Inches(0.7),
            font_size=13, color=C_SAGE,
            align=PP_ALIGN.CENTER, font_name="Calibri")

add_textbox(s, "This presentation is prepared for discussion purposes. All data sourced from GRI 101: Biodiversity 2024,\nWEF Global Risks Report 2024, IPBES 2019, WWF Living Planet Report 2022, TNFD v1.0.",
            Inches(1.0), Inches(6.75), Inches(11.33), Inches(0.55),
            font_size=9, italic=True, color=C_MID_GREY,
            align=PP_ALIGN.CENTER, font_name="Calibri")

slide_number(s, 15)

# ── Save ─────────────────────────────────────────────────────────────────────
out_path = "/home/user/jcode/GRI_101_Biodiversity_Strategy_Presentation.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
