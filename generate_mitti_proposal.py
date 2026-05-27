# -*- coding: utf-8 -*-
"""Generates the 'Mitti Se Mitti' designed CSR proposal for Eicher Group Foundation."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── PALETTE (earthy "mitti" + forest green) ───────────────────────────────────
FOREST     = RGBColor(0x1B, 0x5E, 0x20)   # deep forest green (primary)
FOREST_HEX = '1B5E20'
LEAF       = RGBColor(0x37, 0x8C, 0x3F)   # mid leaf green
LEAF_HEX   = '378C3F'
TERRA      = RGBColor(0xB5, 0x5A, 0x2A)   # terracotta / mitti accent
TERRA_HEX  = 'B55A2A'
SAND_HEX   = 'F4EEE3'                       # warm sand panel
SAGE_HEX   = 'E8F1E5'                       # soft sage panel
CREAM_HEX  = 'FBF8F2'
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
INK        = RGBColor(0x2B, 0x2B, 0x2B)    # body text
GREYTXT    = RGBColor(0x55, 0x55, 0x55)

FONT = 'Calibri'

doc = Document()

# default font
style = doc.styles['Normal']
style.font.name = FONT
style.font.size = Pt(10.5)
style.font.color.rgb = INK

for section in doc.sections:
    section.top_margin    = Cm(1.9)
    section.bottom_margin = Cm(1.9)
    section.left_margin   = Cm(2.2)
    section.right_margin  = Cm(2.2)
    section.page_width     = Cm(21.0)
    section.page_height    = Cm(29.7)

# ── schema-ordered insertion (keeps Word happy) ──────────────────────────────
TCPR_ORDER = ['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd','noWrap',
              'tcMar','textDirection','tcFit','vAlign','hideMark','headers']
RPR_ORDER = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike',
             'outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden',
             'color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr',
             'shd','fitText','vertAlign','rtl','cs','em','lang']

def _insert_ordered(parent, child, order):
    local = child.tag.split('}')[-1]
    if local not in order:
        parent.append(child); return
    pos = order.index(local)
    for existing in parent:
        ex_local = existing.tag.split('}')[-1]
        if ex_local in order and order.index(ex_local) > pos:
            existing.addprevious(child); return
    parent.append(child)

# ── low-level helpers ─────────────────────────────────────────────────────────
def shade(el, hex_fill):
    pr = el.get_or_add_tcPr() if el.tag.endswith('}tc') else el.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_fill)
    _insert_ordered(pr, shd, TCPR_ORDER)

def shade_cell(cell, hex_fill): shade(cell._tc, hex_fill)
def shade_para(p, hex_fill):    shade(p._p, hex_fill)

def no_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'none'); e.set(qn('w:sz'), '0')
        borders.append(e)
    tblPr.append(borders)

def cell_borders(cell, color='D8D2C4', sz='4'):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for edge in ('top','left','bottom','right'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), sz); e.set(qn('w:color'), color)
        b.append(e)
    _insert_ordered(tcPr, b, TCPR_ORDER)

def set_cell_margins(cell, top=60, bottom=60, left=110, right=110):
    tcPr = cell._tc.get_or_add_tcPr()
    m = OxmlElement('w:tcMar')
    for tag, val in (('top',top),('bottom',bottom),('start',left),('end',right)):
        e = OxmlElement(f'w:{tag}')
        e.set(qn('w:w'), str(val)); e.set(qn('w:type'), 'dxa')
        m.append(e)
    _insert_ordered(tcPr, m, TCPR_ORDER)

def vcenter(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def run(p, text, bold=False, italic=False, size=10.5, color=INK, font=FONT, caps=False):
    r = p.add_run(text)
    r.bold = bold; r.italic = italic
    r.font.size = Pt(size); r.font.name = font
    r.font.color.rgb = color
    if caps:
        rPr = r._r.get_or_add_rPr()
        c = OxmlElement('w:caps'); c.set(qn('w:val'), 'true')
        _insert_ordered(rPr, c, RPR_ORDER)
    return r

def para(align=None, before=0, after=4, line=None):
    p = doc.add_paragraph()
    if align is not None: p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after)
    if line: pf.line_spacing = line
    return p

def spacer(pts=6):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(pts); p.paragraph_format.space_before = Pt(0)
    return p

def body(text, after=6, size=10.5, color=INK, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = para(align=align, after=after, line=1.18)
    run(p, text, size=size, color=color)
    return p

# Section banner (full-width colored bar with white title)
def banner(title, hexcol=FOREST_HEX, icon=''):
    t = doc.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    no_borders(t)
    cell = t.rows[0].cells[0]
    cell.width = Cm(16.6)
    shade_cell(cell, hexcol)
    set_cell_margins(cell, top=90, bottom=90, left=160, right=160)
    p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
    run(p, (icon + '  ' if icon else '') + title, bold=True, size=13, color=WHITE, caps=True)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

# Sub-heading (colored, no bar)
def subhead(text, color=TERRA):
    p = para(before=8, after=3)
    run(p, text, bold=True, size=11.5, color=color)
    return p

# Bullet
def bullet(text, color=INK, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Inches(0.28)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.12
    if bold_lead:
        run(p, bold_lead, bold=True, size=10.5, color=FOREST)
    run(p, text, size=10.5, color=color)
    return p

# Row of stat callout boxes
def stat_row(items, fill=SAGE_HEX, numcol=FOREST, labelcol=GREYTXT):
    """items: list of (number, label)"""
    n = len(items)
    t = doc.add_table(rows=1, cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(t)
    for i, (num, label) in enumerate(items):
        cell = t.rows[i].cells[i] if False else t.rows[0].cells[i]
        shade_cell(cell, fill)
        set_cell_margins(cell, top=130, bottom=130, left=80, right=80)
        vcenter(cell)
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        run(p, num, bold=True, size=20, color=numcol)
        p2 = cell.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0); p2.paragraph_format.space_after = Pt(0)
        run(p2, label, size=8.5, color=labelcol, caps=True)
    # spacing between boxes via cell margins; add small gap row after
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

# ══════════════════════════════════════════════════════════════════════════════
#  COVER PAGE
# ══════════════════════════════════════════════════════════════════════════════
# Top thin accent line
def thin_band(hexcol, height_pts=6):
    t = doc.add_table(rows=1, cols=1); no_borders(t)
    c = t.rows[0].cells[0]; c.width = Cm(16.6)
    shade_cell(c, hexcol)
    p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
    r = p.add_run(' '); r.font.size = Pt(height_pts)

spacer(10)
p = para(align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
run(p, 'PROJECT PROPOSAL', bold=True, size=11, color=TERRA, caps=True)
p = para(align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
run(p, 'FY 2026–27', size=10, color=GREYTXT)

spacer(14)
# Big title block
t = doc.add_table(rows=1, cols=1); no_borders(t)
c = t.rows[0].cells[0]; c.width = Cm(16.6)
shade_cell(c, FOREST_HEX)
set_cell_margins(c, top=420, bottom=300, left=200, right=200)
vcenter(c)
p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(2)
run(p, 'MITTI SE MITTI', bold=True, size=40, color=WHITE)
p2 = c.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER; p2.paragraph_format.space_after = Pt(0); p2.paragraph_format.space_before=Pt(2)
run(p2, 'कचरे से खाद, खाद से समृद्धि', italic=True, size=12, color=RGBColor(0xD8,0xE8,0xD4))

# tagline strip
t = doc.add_table(rows=1, cols=1); no_borders(t)
c = t.rows[0].cells[0]; c.width = Cm(16.6)
shade_cell(c, TERRA_HEX)
set_cell_margins(c, top=120, bottom=120, left=200, right=200)
p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
run(p, 'From Soil to Soil — Composting Wet Waste into Livelihoods on the Edge of Corbett Tiger Reserve',
    bold=True, size=12, color=WHITE)

spacer(18)
# Location chip
p = para(align=WD_ALIGN_PARAGRAPH.CENTER, after=3)
run(p, 'Himmatpur Dotiyal  •  Eastern Periphery of Corbett Tiger Reserve', size=11, color=FOREST, bold=True)
p = para(align=WD_ALIGN_PARAGRAPH.CENTER, after=3)
run(p, 'Ramnagar, District Nainital, Uttarakhand', size=10.5, color=GREYTXT)

spacer(30)
# Submitted to / by block
t = doc.add_table(rows=1, cols=2); no_borders(t)
left, right = t.rows[0].cells
for cc in (left, right):
    set_cell_margins(cc, top=120, bottom=120, left=160, right=160)
shade_cell(left, SAND_HEX); shade_cell(right, SAGE_HEX)
# submitted to
p = left.paragraphs[0]; p.paragraph_format.space_after=Pt(2)
run(p, 'SUBMITTED TO', bold=True, size=9, color=TERRA, caps=True)
p = left.add_paragraph(); p.paragraph_format.space_after=Pt(0)
run(p, 'Eicher Group Foundation', bold=True, size=12, color=INK)
# submitted by
p = right.paragraphs[0]; p.paragraph_format.space_after=Pt(2)
run(p, 'SUBMITTED BY', bold=True, size=9, color=FOREST, caps=True)
p = right.add_paragraph(); p.paragraph_format.space_after=Pt(0)
run(p, 'Waste Warriors Society', bold=True, size=12, color=INK)
p = right.add_paragraph(); p.paragraph_format.space_after=Pt(0)
run(p, 'www.wastewarriors.org', size=9, color=GREYTXT)

spacer(8)
p = para(align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
run(p, 'Project Duration: 3 Months   |   Total Outlay: ₹2,75,420', size=10, color=GREYTXT, italic=True)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  1. EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
banner('Executive Summary')
body("On the eastern periphery of Corbett Tiger Reserve, one of India’s most celebrated wildlife "
     "landscapes, a quiet crisis is unfolding. The region now generates nearly 42 metric tonnes of "
     "waste every single day — driven by booming tourism, changing rural lifestyles, and rising "
     "consumption of plastic and packaged goods. A large share of this is wet (organic) waste that, "
     "left unmanaged, rots in the open, pollutes rivers, attracts wildlife into conflict, and erodes "
     "the very ecosystem that makes Corbett extraordinary.")

body("Waste Warriors Society proposes “Mitti Se Mitti” — a 3-month, women-led initiative that "
     "closes the loop on wet waste by converting it into compost at our Material Recovery Facility "
     "in Himmatpur Dotiyal, while building awareness, strengthening source segregation, and expanding "
     "the livelihoods of the Paryavaran Sakhis — the 25 women entrepreneurs who lead waste management "
     "across the region. The name says it all: from the soil it came, to the soil it returns.")

subhead('What this project will deliver in 3 months', color=FOREST)
stat_row([
    ('18,000 kg', 'Wet Waste Processed into Compost'),
    ('6', 'Paryavaran Sakhis Trained'),
    ('18', 'Awareness & Stakeholder Engagements'),
], fill=SAGE_HEX, numcol=FOREST)

body("With a focused investment of ₹2,75,420, the Eicher Group Foundation can help turn an "
     "environmental liability into a circular, dignified, and replicable model of rural waste "
     "management — protecting Corbett’s forests and rivers while creating sustainable incomes for "
     "local women.", after=4)

# ══════════════════════════════════════════════════════════════════════════════
#  2. ABOUT WASTE WARRIORS
# ══════════════════════════════════════════════════════════════════════════════
banner('About Waste Warriors Society')
body("Waste Warriors Society (WWS) is a non-profit organisation catalysing systemic change to solve "
     "the waste management crisis of the Indian Himalayan Region. Our work lies at the intersection of "
     "Climate Change, Biodiversity & Habitat Conservation, and Livelihood Generation. We enable better "
     "governance, address infrastructure and policy gaps, and activate communities to co-create "
     "lasting solutions. In November 2025, Waste Warriors marked 13 years of service to the mountains.")

subhead('Our Impact Since 2012', color=FOREST)
stat_row([
    ('4,50,000+', 'People Engaged on Waste Awareness'),
    ('8,500+ MT', 'Waste Collected from Himalayan Ecosystems'),
    ('1,200+', 'Waste Workers Onboarded & Trained'),
], fill=SAND_HEX, numcol=TERRA)

subhead('Recent Recognition', color=FOREST)
bullet('volunteers mobilised across 101 locations to remove 18,600 kg of waste in just one hour on World Environment Day 2025 — a feat certified as a World Record by the World Records Union, USA.', bold_lead='Largest Mountain Cleanup 2025: 8,492 ')
bullet('with the Rural Development Department of Himachal Pradesh as State Partner for Solid Waste Management for the next 3 years.', bold_lead='Strategic MoU ')
bullet('including the Climate Action & Sustainability Awards and two awards at the ISWA World Congress.', bold_lead='Multiple honours, ')
body("Waste Warriors Society is registered under the Indian Societies Act and holds valid 12A, 80G, "
     "CSR-1, and FCRA registrations — ensuring full compliance and transparency for institutional partners.",
     after=4)

# ══════════════════════════════════════════════════════════════════════════════
#  3. THE CONTEXT — CORBETT
# ══════════════════════════════════════════════════════════════════════════════
banner('The Context: Corbett Tiger Reserve, Nainital')
body("Nestled in the foothills of the Himalayas in Uttarakhand, Corbett Tiger Reserve is one of "
     "India’s most celebrated wildlife landscapes, known for its rich biodiversity and ecological "
     "significance. Spread across dense sal forests, grasslands, river belts, and hilly terrain, the "
     "reserve is home to the majestic Bengal tiger, Asian elephant, leopard, numerous deer species, "
     "reptiles, and over 600 species of birds. This delicate natural balance makes the reserve a "
     "critical habitat for wildlife conservation — and a landscape worth protecting at every cost.")

subhead('A Landscape Under Pressure', color=TERRA)
body("In recent years, tourism in and around Corbett has witnessed remarkable growth, bringing new "
     "opportunities through hospitality, transport, guiding services, and small businesses. But "
     "alongside this growth comes an equally significant challenge: waste. The rising volume is driven "
     "not only by tourism, but also by changing lifestyles and increasing consumption of plastic and "
     "packaged products within local communities.")

# highlight box — the 42 MT stat
t = doc.add_table(rows=1, cols=1); no_borders(t)
c = t.rows[0].cells[0]; c.width = Cm(16.6)
shade_cell(c, TERRA_HEX); set_cell_margins(c, top=140, bottom=140, left=200, right=200)
p = c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(2)
run(p, '≈42 metric tonnes of waste generated every single day', bold=True, size=15, color=WHITE)
p2 = c.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER; p2.paragraph_format.space_after=Pt(0)
run(p2, 'across the Corbett region — a burden steadily pressing on its forests, rivers, and villages',
    size=9.5, color=RGBColor(0xF6,0xE6,0xDA), italic=True)
spacer(6)

body("A substantial portion of this is wet, organic waste. When mixed with dry and plastic waste and "
     "left unmanaged, it spoils otherwise recyclable material, generates foul leachate that "
     "contaminates soil and water, emits methane, and draws wildlife into human settlements — "
     "heightening the risk of human–animal conflict on the reserve’s fragile edge.", after=4)

# ══════════════════════════════════════════════════════════════════════════════
#  4. NEED FOR INTERVENTION
# ══════════════════════════════════════════════════════════════════════════════
banner('Why Wet Waste, Why Now')
body("Historically, organic waste in these villages was absorbed at the household level — fed to "
     "livestock or composted on open land. But rapid, semi-urban development has changed everything. "
     "With shrinking open spaces, changing consumption patterns, and reduced dependency on livestock, "
     "the demand for organised wet waste collection and processing has risen sharply — making a "
     "dedicated wet waste solution both timely and necessary.")

body("Until recently, the Paryavaran Sakhis were primarily engaged in collecting and processing dry "
     "waste, which limited the services they could offer and the income they could earn. The "
     "establishment of a wet waste processing facility now allows them to serve bulk waste generators "
     "— hotels, resorts, and hospitality establishments — unlocking a more consistent, year-round, "
     "and sustainable revenue stream. This project builds directly on that opportunity.", after=4)

# ══════════════════════════════════════════════════════════════════════════════
#  5. THE PROVEN MODEL
# ══════════════════════════════════════════════════════════════════════════════
banner('Our Proven Foundation: The Paryavaran Sakhi Model')
body("“Paryavaran Sakhi” — friend of the environment — is Waste Warriors’ flagship community "
     "entrepreneurship model. Started formally in 2021, it addresses the region’s waste crisis through "
     "community empowerment, gender equality, and inclusive livelihoods. Today it engages 25 women "
     "from diverse backgrounds, equipping them with dignified, entrepreneurial roles in waste "
     "management. They are not beneficiaries — they are professional environmental service providers.")

subhead('Infrastructure Already in Place', color=FOREST)
stat_row([
    ('480 sq. ft.', 'Wet Waste Processing Shed'),
    ('400 kg/day', 'Processing Capacity'),
    ('1', 'Industrial Shredder Machine'),
], fill=SAGE_HEX, numcol=FOREST)

body("This facility — established with the support of Shearwater Geoservices India Pvt. Ltd. — gives "
     "the project a running start. The hard infrastructure exists; what “Mitti Se Mitti” now funds is "
     "the operations, awareness, training, and consumables needed to run it at scale and embed it in "
     "the community.")

# livelihood callout
t = doc.add_table(rows=1, cols=1); no_borders(t)
c = t.rows[0].cells[0]; c.width = Cm(16.6)
shade_cell(c, SAND_HEX); set_cell_margins(c, top=120, bottom=120, left=200, right=200)
p = c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
run(p, 'Up to +120 additional livelihood days every month', bold=True, size=14, color=TERRA)
p2 = c.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER; p2.paragraph_format.space_after=Pt(0)
run(p2, 'created for the Paryavaran Sakhis through year-round wet waste operations',
    size=9.5, color=GREYTXT, italic=True)
spacer(6)

# ══════════════════════════════════════════════════════════════════════════════
#  6. THE PROPOSED PROJECT
# ══════════════════════════════════════════════════════════════════════════════
banner('The Project: “Mitti Se Mitti”')

subhead('Project Goal', color=TERRA)
body("To establish a sustainable, community-led wet waste management system on the eastern periphery "
     "of Corbett Tiger Reserve — converting organic waste into compost, strengthening source "
     "segregation, and expanding the livelihoods of the Paryavaran Sakhis — thereby protecting the "
     "reserve’s fragile ecosystem while closing the loop from waste to wealth.")

subhead('Project Objectives', color=TERRA)
bullet('Process 6,000 kg of wet waste per month (18,000 kg over the project) into quality compost at the Material Recovery Facility in Himmatpur Dotiyal.', bold_lead='Convert waste to compost — ')
bullet('Train 6 Paryavaran Sakhis to lead the full wet waste value chain — from awareness and collection to transportation and on-site processing.', bold_lead='Build women’s capacity — ')
bullet('Conduct 12 community awareness sessions to drive source segregation at the household level.', bold_lead='Shift behaviour — ')
bullet('Convene 6 stakeholder meetings with hotels and resorts (bulk waste generators) to formalise source segregation and divert their wet waste from dumping.', bold_lead='Engage bulk generators — ')
bullet('Build a recognisable identity and market for the Sakhis’ ready-to-use compost, turning a waste stream into a revenue stream.', bold_lead='Create a market — ')

subhead('Core Activities', color=TERRA)
bullet('through chai-pe-charcha sessions, focus group discussions, and exposure visits to the processing site, positioning it as a local knowledge hub.', bold_lead='Awareness & IEC — ')
bullet('processing of 6,000 kg/month of wet waste into compost, supported by collection drums, trolleys, and safety equipment for the Sakhis.', bold_lead='Wet waste operations — ')
bullet('promoting home composting through earthen pots, step-by-step guides, and certificates of recognition for participating households (25 each).', bold_lead='Home composting drive — ')
bullet('direction and recognition signboards across the region to promote the processing site and celebrate active composting households.', bold_lead='Visibility — ')
bullet('designing a brand identity — logo and name — for the compost produced by the Paryavaran Sakhis.', bold_lead='Branding — ')

# ══════════════════════════════════════════════════════════════════════════════
#  7. EXPECTED OUTCOMES
# ══════════════════════════════════════════════════════════════════════════════
banner('Expected Outcomes & Impact')

outcomes = [
    ('Environmental', '18,000 kg of wet waste diverted from open dumping and converted to compost over 3 months, reducing methane emissions, leachate pollution, and wildlife attraction on Corbett’s periphery.'),
    ('Social / Gender', '6 Paryavaran Sakhis equipped with end-to-end operational skills and up to 120 additional livelihood days per month — advancing women’s economic independence and dignity.'),
    ('Behavioural', 'Measurable improvement in household source segregation across target villages, reinforced by 12 awareness sessions and a home-composting movement.'),
    ('Institutional', '6 hotels/resorts onboarded to segregate at source and route wet waste responsibly — a replicable template for the wider hospitality sector around Corbett.'),
    ('Economic', 'A branded, ready-to-use compost product that turns waste into a sustainable revenue stream for the Sakhis — building long-term financial viability beyond the grant period.'),
]
tb = doc.add_table(rows=len(outcomes), cols=2); no_borders(tb)
for i, (head, desc) in enumerate(outcomes):
    lc, rc = tb.rows[i].cells
    lc.width = Cm(3.4); rc.width = Cm(13.2)
    fill = SAGE_HEX if i % 2 == 0 else CREAM_HEX
    shade_cell(lc, FOREST_HEX); shade_cell(rc, fill)
    set_cell_margins(lc, top=90, bottom=90, left=120, right=120)
    set_cell_margins(rc, top=90, bottom=90, left=140, right=140)
    vcenter(lc); vcenter(rc)
    p = lc.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    run(p, head, bold=True, size=10.5, color=WHITE)
    p = rc.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    run(p, desc, size=10, color=INK)
spacer(6)

# ══════════════════════════════════════════════════════════════════════════════
#  8. BUDGET
# ══════════════════════════════════════════════════════════════════════════════
banner('Detailed Budget')
body("The total outlay for the 3-month project is ₹2,75,420 (Two Lakh Seventy-Five Thousand Four "
     "Hundred and Twenty Only). The full head-wise, activity-wise break-up is given below.", after=6)

# Budget table: columns S.No | Head/Details | Target | Unit Cost | Total
bud_headers = ['Head / Activity', 'Detail', 'Target', 'Unit Cost (₹)', 'Total (₹)']
# rows: tuple(code_or_blank, name, detail, target, unit, total, kind)
# kind: 'section', 'item', 'subtotal', 'grandtotal'
bud = [
    ('A. AWARENESS & IEC', '', '', '', '1,47,920', 'section'),
    ('Awareness Sessions', 'Chai-pe-charcha, focus group discussions & exposure visits on source segregation', '12', '660', '7,920', 'item'),
    ("Stakeholders’ Meet", 'Meetings with bulk waste generators (hotels & resorts) to drive segregation & wet waste processing', '6', '5,000', '30,000', 'item'),
    ('IEC Collaterals', 'On-site IEC positioning the facility as a knowledge hub for communities & tourists', '–', '–', '50,000', 'item'),
    ('Signboards', 'Direction signboards for the processing site & recognition boards for active composting households', '4', '15,000', '60,000', 'item'),

    ('B. OPERATIONS', '', '', '', '1,07,500', 'section'),
    ('Waste Management Consumables', 'Gloves, aprons, shoes, masks, collection drums & trolley; earthen pots, shovels, guides & certificates; compost brand identity', '–', '–', '92,500', 'item'),
    ('Training & Capacity Building', 'Training of 6 Paryavaran Sakhis to lead the full wet waste value chain', '3', '5,000', '15,000', 'item'),

    ('C. OFFICE EXPENSES', '', '', '', '20,000', 'section'),
    ('Printing & Stationery', "Stakeholders’ reports and other office expenses", '–', '–', '10,000', 'item'),
    ('Local Travel', 'Field travel for collection, monitoring & coordination', '–', '–', '10,000', 'item'),

    ('TOTAL PROJECT COST', '', '', '', '2,75,420', 'grandtotal'),
]

tbl = doc.add_table(rows=1+len(bud), cols=5)
tbl.style = 'Table Grid'
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
widths = [Cm(3.9), Cm(6.7), Cm(1.5), Cm(2.0), Cm(2.0)]
# header
for j, h in enumerate(bud_headers):
    cell = tbl.rows[0].cells[j]
    shade_cell(cell, FOREST_HEX); cell_borders(cell, color=FOREST_HEX)
    set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
    p = cell.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    if j >= 2: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, h, bold=True, size=9, color=WHITE)
# rows
for i, rowdata in enumerate(bud):
    name, detail, target, unit, total, kind = rowdata
    row = tbl.rows[i+1]
    cells = row.cells
    for j in range(5): set_cell_margins(cells[j], top=55, bottom=55, left=100, right=100)
    if kind == 'section':
        # merge name across detail col for clean section header look
        shade_for = TERRA_HEX
        for j in range(5):
            shade_cell(cells[j], shade_for); cell_borders(cells[j], color=shade_for)
        p = cells[0].paragraphs[0]; p.paragraph_format.space_after=Pt(0)
        run(p, name, bold=True, size=9.5, color=WHITE)
        # total in last col
        p = cells[4].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(0)
        run(p, total, bold=True, size=9.5, color=WHITE)
        for j in (1,2,3):
            cells[j].paragraphs[0].paragraph_format.space_after=Pt(0)
    elif kind == 'grandtotal':
        for j in range(5):
            shade_cell(cells[j], FOREST_HEX); cell_borders(cells[j], color=FOREST_HEX)
        p = cells[0].paragraphs[0]; p.paragraph_format.space_after=Pt(0)
        run(p, name, bold=True, size=10.5, color=WHITE)
        p = cells[4].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(0)
        run(p, total, bold=True, size=11, color=WHITE)
    else:
        fill = CREAM_HEX if i % 2 == 0 else 'FFFFFF'
        for j in range(5):
            shade_cell(cells[j], fill); cell_borders(cells[j])
        p = cells[0].paragraphs[0]; p.paragraph_format.space_after=Pt(0)
        run(p, name, bold=True, size=9, color=FOREST)
        p = cells[1].paragraphs[0]; p.paragraph_format.space_after=Pt(0)
        run(p, detail, size=8.5, color=GREYTXT)
        p = cells[2].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
        run(p, target, size=9, color=INK)
        p = cells[3].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
        run(p, unit, size=9, color=INK)
        p = cells[4].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(0)
        run(p, total, size=9, color=INK)
spacer(8)

# ══════════════════════════════════════════════════════════════════════════════
#  9. TIMELINE
# ══════════════════════════════════════════════════════════════════════════════
banner('Implementation Timeline (3 Months)')
tl = [
    ('Month 1', 'Setup & Mobilisation', 'Train 6 Sakhis; procure consumables & safety equipment; design compost brand identity; launch first awareness sessions; install initial signboards; begin wet waste collection & processing.'),
    ('Month 2', 'Scale & Engage', 'Full-scale wet waste processing (6,000 kg/month); ongoing awareness sessions; stakeholder meetings with hotels & resorts; roll out home-composting kits to households.'),
    ('Month 3', 'Consolidate & Sustain', 'Sustained processing; final stakeholder & awareness rounds; recognition of active composting households; compost branding & market linkage; documentation and reporting.'),
]
tt = doc.add_table(rows=len(tl), cols=3); no_borders(tt)
for i, (m, phase, desc) in enumerate(tl):
    c0, c1, c2 = tt.rows[i].cells
    c0.width = Cm(2.2); c1.width = Cm(4.0); c2.width = Cm(10.4)
    shade_cell(c0, FOREST_HEX); shade_cell(c1, SAGE_HEX); shade_cell(c2, CREAM_HEX)
    for cc in (c0,c1,c2): set_cell_margins(cc, top=90, bottom=90, left=130, right=130); vcenter(cc)
    p=c0.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
    run(p, m, bold=True, size=10.5, color=WHITE)
    p=c1.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    run(p, phase, bold=True, size=10, color=FOREST)
    p=c2.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    run(p, desc, size=9.5, color=INK)
spacer(6)

# ══════════════════════════════════════════════════════════════════════════════
#  10. SUSTAINABILITY
# ══════════════════════════════════════════════════════════════════════════════
banner('Sustainability Beyond the Grant')
body("“Mitti Se Mitti” is designed to outlast its funding. The infrastructure is already in place; "
     "this grant activates the people and systems that make it self-sustaining:")
bullet('Selling branded, ready-to-use compost to nurseries, farmers, hotels, and tourists creates a recurring income stream.', bold_lead='Revenue from compost — ')
bullet('Formal source-segregation arrangements with hotels and resorts secure a steady, paid supply of wet waste for processing.', bold_lead='Bulk-generator contracts — ')
bullet('Trained Paryavaran Sakhis become permanent, professional service providers — not dependent on any single grant cycle.', bold_lead='Empowered Sakhis — ')
bullet('Awareness and home-composting embed segregation as a community habit, reducing the waste burden at source over time.', bold_lead='Behaviour change — ')
spacer(4)

# ══════════════════════════════════════════════════════════════════════════════
#  11. LOOKING AHEAD (4-wheeler hint)
# ══════════════════════════════════════════════════════════════════════════════
banner('Looking Ahead: Scaling the Model', hexcol=TERRA_HEX)
body("This proposal focuses on getting the wet waste system running and embedded in the community. As "
     "operations scale across more villages and bulk generators on Corbett’s periphery, reliable "
     "collection and transportation will become the single biggest constraint on growth.")
# hint box
t = doc.add_table(rows=1, cols=1); no_borders(t)
c = t.rows[0].cells[0]; c.width = Cm(16.6)
shade_cell(c, SAND_HEX); set_cell_margins(c, top=120, bottom=120, left=200, right=200)
p = c.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
run(p, 'A note on what comes next:  ', bold=True, size=10.5, color=TERRA)
run(p, "to conduct the awareness, operations, and collection activities described here at full scale "
       "— and to expand them across the region — a dedicated four-wheeler vehicle will be essential "
       "for waste collection, transport to the processing facility, and field mobility for the Sakhis. "
       "We look forward to sharing a detailed follow-on proposal covering this requirement, so that "
       "the Eicher Group Foundation’s support can translate into lasting, scalable impact on the "
       "ground.", size=10.5, color=INK)
spacer(8)

# ══════════════════════════════════════════════════════════════════════════════
#  12. CONTACT / THANK YOU
# ══════════════════════════════════════════════════════════════════════════════
t = doc.add_table(rows=1, cols=1); no_borders(t)
c = t.rows[0].cells[0]; c.width = Cm(16.6)
shade_cell(c, FOREST_HEX); set_cell_margins(c, top=200, bottom=200, left=220, right=220)
p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(4)
run(p, 'Thank you for considering our proposal.', bold=True, size=14, color=WHITE)
p=c.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
run(p, 'Together, we can ensure that on the edge of Corbett, waste returns to the soil — as compost, '
       'as dignity, and as livelihoods.', italic=True, size=10.5, color=RGBColor(0xDD,0xEC,0xDA))
spacer(10)

subhead('Get in Touch', color=FOREST)
ct = doc.add_table(rows=2, cols=2); no_borders(ct)
contacts = [
    ('Waste Warriors Society', 'Registered under the Indian Societies Act\n12A | 80G | CSR-1 | FCRA'),
    ('Address', '136/2/2, Shivam Vihar, Jakhan, Rajpur Road,\nDehradun, Uttarakhand – 248001'),
    ('Email', 'partnerships@wastewarriors.org'),
    ('Website', 'www.wastewarriors.org'),
]
for i, (k, v) in enumerate(contacts):
    cell = ct.rows[i//2].cells[i%2]
    shade_cell(cell, SAGE_HEX if i%2==0 else CREAM_HEX)
    set_cell_margins(cell, top=90, bottom=90, left=140, right=140)
    p=cell.paragraphs[0]; p.paragraph_format.space_after=Pt(1)
    run(p, k, bold=True, size=9.5, color=FOREST, caps=True)
    for li, line in enumerate(v.split('\n')):
        pp = cell.add_paragraph() if li>0 else cell.add_paragraph()
        pp.paragraph_format.space_after=Pt(0)
        run(pp, line, size=9.5, color=INK)

out = '/home/user/jcode/Mitti_Se_Mitti_Corbett_Proposal_EGF.docx'
doc.save(out)
print('Saved:', out)
