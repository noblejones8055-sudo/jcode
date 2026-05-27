# -*- coding: utf-8 -*-
"""Mitti Se Mitti — v2 in Trek-and-Trails style (clean, photo-forward, minimal text)"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Palette (light, natural — matches Trek & Trails aesthetic) ────────────────
GREEN      = RGBColor(0x2E, 0x7D, 0x32)   # accent green (for headings, bold numbers)
GREEN_HEX  = '2E7D32'
LAVENDER   = RGBColor(0xE8, 0xE8, 0xF4)   # light lavender callout (like Trek pdf)
LAV_HEX    = 'E8E8F4'
SAGE_HEX   = 'EAF4EA'                       # light sage for stat boxes
SAND_HEX   = 'FDF6EC'                       # warm sand for photo placeholders
DARK       = RGBColor(0x1A, 0x1A, 0x1A)   # near-black body
GREY       = RGBColor(0x55, 0x55, 0x55)
LTGREY_HEX = 'E0E0E0'
PHOTO_HEX  = 'D4D4D4'                       # placeholder fill
TERRA      = RGBColor(0xBF, 0x60, 0x20)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
FONT = 'Calibri'

doc = Document()
style = doc.styles['Normal']
style.font.name = FONT
style.font.size = Pt(10.5)
style.font.color.rgb = DARK

for section in doc.sections:
    section.top_margin    = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── OOXML helpers ─────────────────────────────────────────────────────────────
TCPR_ORDER = ['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd',
              'noWrap','tcMar','textDirection','tcFit','vAlign','hideMark','headers']

def _ins(parent, child, order):
    local = child.tag.split('}')[-1]
    if local in order:
        pos = order.index(local)
        for ex in parent:
            el = ex.tag.split('}')[-1]
            if el in order and order.index(el) > pos:
                ex.addprevious(child); return
    parent.append(child)

def _shade_tc(tc, h):
    pr = tc.get_or_add_tcPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),h)
    _ins(pr, s, TCPR_ORDER)

def _shade_p(p, h):
    pr = p._p.get_or_add_pPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),h)
    pr.append(s)

def sc(cell, h): _shade_tc(cell._tc, h)

def no_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for e in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement(f'w:{e}')
        el.set(qn('w:val'),'none'); el.set(qn('w:sz'),'0')
        b.append(el)
    tblPr.append(b)

def grid_border(table, color='BBBBBB', sz='4'):
    tbl = table._tbl
    tblPr = tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for e in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement(f'w:{e}')
        el.set(qn('w:val'),'single'); el.set(qn('w:sz'),sz); el.set(qn('w:color'),color)
        b.append(el)
    tblPr.append(b)

def cm_(cell, top=80, bot=80, l=120, r=120):
    pr = cell._tc.get_or_add_tcPr()
    m = OxmlElement('w:tcMar')
    for tag, v in (('top',top),('bottom',bot),('start',l),('end',r)):
        e = OxmlElement(f'w:{tag}')
        e.set(qn('w:w'),str(v)); e.set(qn('w:type'),'dxa')
        m.append(e)
    _ins(pr, m, TCPR_ORDER)

def vc(cell): cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def R(p, text, bold=False, italic=False, sz=10.5, col=DARK, underline=False):
    r = p.add_run(text)
    r.bold=bold; r.italic=italic; r.underline=underline
    r.font.name=FONT; r.font.size=Pt(sz); r.font.color.rgb=col
    return r

def P(align=None, before=0, after=5):
    p = doc.add_paragraph()
    if align: p.alignment=align
    p.paragraph_format.space_before=Pt(before)
    p.paragraph_format.space_after=Pt(after)
    return p

def sp(pts=4):
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(pts); p.paragraph_format.space_before=Pt(0)

def body(txt, after=6, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p=P(align=align, after=after)
    p.paragraph_format.line_spacing=1.15
    R(p, txt, sz=10.5, col=DARK)
    return p

def italic_body(txt, after=6):
    p=P(after=after)
    R(p, txt, italic=True, sz=10.5, col=GREY)
    return p

def heading1(txt):
    p=P(before=14, after=4)
    R(p, txt, bold=True, sz=15, col=DARK)
    return p

def heading2(txt, color=DARK, underline=True):
    p=P(before=10, after=3)
    r=R(p, txt, bold=True, sz=11, col=color, underline=underline)
    return p

def bullet_item(lead_bold, rest, col=DARK):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent=Inches(0.25)
    p.paragraph_format.space_after=Pt(3)
    p.paragraph_format.line_spacing=1.12
    if lead_bold:
        R(p, lead_bold, bold=True, sz=10.5, col=DARK)
    R(p, rest, sz=10.5, col=col)

# ── Photo placeholder ─────────────────────────────────────────────────────────
def photo_placeholder(caption='[ Photo ]', height_cm=5.5, cols=1, captions=None):
    """Renders a grey placeholder box (or row of boxes if cols>1)."""
    n = cols
    t = doc.add_table(rows=1, cols=n); no_border(t)
    for i in range(n):
        cell = t.rows[0].cells[i]
        cap = (captions[i] if captions else caption)
        cell.width = Cm(16.0/n - 0.2)
        sc(cell, PHOTO_HEX)
        cm_(cell, top=int(height_cm*360), bot=int(height_cm*360), l=120, r=120)
        vc(cell)
        p = cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after=Pt(0)
        R(p, cap, sz=9, col=GREY, italic=True)
    sp(6)

# ── Callout box (lavender, like Trek & Trails highlight) ──────────────────────
def callout(text, fill=LAV_HEX, sz=10.5, bold=False, italic=False, color=DARK):
    t = doc.add_table(rows=1, cols=1); no_border(t)
    cell = t.rows[0].cells[0]; cell.width=Cm(16.0)
    sc(cell, fill); cm_(cell, top=140, bot=140, l=160, r=160)
    p = cell.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    R(p, text, bold=bold, italic=italic, sz=sz, col=color)
    sp(6)

# ── Stat row (like Trek & Trails impact box) ──────────────────────────────────
def stat_box(items, fill=SAGE_HEX):
    """items = [(number_str, label_str), ...]"""
    n = len(items)
    t = doc.add_table(rows=1, cols=n); no_border(t)
    outer = OxmlElement('w:tblBorders')
    for e in ('top','left','bottom','right'):
        el = OxmlElement(f'w:{e}')
        el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'6')
        el.set(qn('w:color'),'AAAAAA')
        outer.append(el)
    t._tbl.tblPr.append(outer)
    for i,(num,lbl) in enumerate(items):
        c = t.rows[0].cells[i]
        sc(c, fill); cm_(c,top=180,bot=160,l=80,r=80); vc(c)
        p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(2)
        R(p, num, bold=True, sz=22, col=GREEN)
        p2=c.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(0)
        R(p2, lbl, sz=8.5, col=GREY)
    sp(8)

# ── Horizontal rule ───────────────────────────────────────────────────────────
def hrule(color='CCCCCC'):
    p=doc.add_paragraph(); pf=p.paragraph_format
    pf.space_before=Pt(2); pf.space_after=Pt(4)
    pPr=p._p.get_or_add_pPr()
    pBdr=OxmlElement('w:pBdr')
    bot=OxmlElement('w:bottom')
    bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'4')
    bot.set(qn('w:space'),'1'); bot.set(qn('w:color'),color)
    pBdr.append(bot); pPr.append(pBdr)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 1 — COVER
# ══════════════════════════════════════════════════════════════════════════════
# Full-bleed photo placeholder at top
photo_placeholder('[ Cover photo — Paryavaran Sakhis at the Himmatpur Dotiyal facility / Corbett landscape ]',
                  height_cm=6.5)

sp(8)
p=P(align=WD_ALIGN_PARAGRAPH.LEFT, after=2)
R(p, 'Mitti Se Mitti', bold=True, sz=28, col=DARK)
p=P(align=WD_ALIGN_PARAGRAPH.LEFT, after=3)
R(p, 'A Wet Waste-to-Compost Initiative by the Paryavaran Sakhis of Corbett', sz=13, col=GREY, italic=True)

sp(4)
p=P(align=WD_ALIGN_PARAGRAPH.LEFT, after=3)
R(p, 'Proposal  |  FY 2026–27', sz=10.5, col=GREY)

sp(10)
p=P(align=WD_ALIGN_PARAGRAPH.LEFT, after=2)
R(p, 'Shared by', sz=9, col=GREY)
p=P(align=WD_ALIGN_PARAGRAPH.LEFT, after=0)
R(p, 'Waste Warriors Society', bold=True, sz=13, col=GREEN)
p=P(align=WD_ALIGN_PARAGRAPH.LEFT, after=0)
R(p, 'www.wastewarriors.org', sz=9.5, col=GREY)

sp(10)
p=P(align=WD_ALIGN_PARAGRAPH.LEFT, after=2)
R(p, 'Submitted to', sz=9, col=GREY)
p=P(align=WD_ALIGN_PARAGRAPH.LEFT, after=0)
R(p, 'Eicher Group Foundation', bold=True, sz=13, col=DARK)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 2 — OPENING NARRATIVE
# ══════════════════════════════════════════════════════════════════════════════
heading1('From the Edge of Corbett')
heading2('A moment in the village', underline=True)

italic_body(
    'Just beyond the boundary of Corbett Tiger Reserve, morning light filters through the sal trees. '
    'Women carry baskets to work. Children leave for school. And every day, a quiet but '
    'growing problem festers: organic waste, mixed, dumped, and left to rot.')

italic_body(
    'Hotels and resorts that feed from the tiger reserve\'s fame generate tonnes of kitchen waste. '
    'Households, no longer dependent on livestock, have nowhere to send their wet organic refuse. '
    'The rivers that flow through Corbett carry the consequences downstream.')

italic_body(
    'The question before us is simple: Can the communities on Corbett\'s edge turn this waste '
    'back into something that gives — back to the soil, back to the family, back to the forest?')

p=P(after=4)
R(p, 'We believe they can.', bold=True, sz=11, col=DARK)

sp(8)
heading2('About Waste Warriors', underline=True)
body('Waste Warriors Society is a non-profit organisation catalysing systemic change to solve the '
     'waste management crisis of the Indian Himalayan Region. Our work lies at the intersection of '
     'Climate Change, Biodiversity Habitat Conservation, and Informal Livelihoods through bringing '
     'behavioural change in the community. With a team of 240+ Warriors across Uttarakhand and '
     'Himachal Pradesh, we focus on enabling better governance, addressing infrastructure and policy '
     'gaps, and activating communities to co-create solutions.')

# Three-column quick facts
t=doc.add_table(rows=1,cols=3); no_border(t)
for txt,col in [('Zero Waste Programs','Shimla · Manali · Dharamshala · Corbett and more'),
                ('1,200+ Waste Workers','Engaged, onboarded, or trained across the region'),
                ('13 Years','Building waste solutions in the Indian Himalayas')]:
    c=t.rows[0].cells[['Zero Waste Programs','1,200+ Waste Workers','13 Years'].index(txt)]
    sc(c, SAGE_HEX); cm_(c,top=100,bot=100,l=120,r=120); vc(c)
    p=c.paragraphs[0]; p.paragraph_format.space_after=Pt(2)
    R(p, txt.split(' — ')[0], bold=True, sz=10, col=GREEN)
    p2=c.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(0)
    R(p2, col, sz=8.5, col=GREY)
sp(6)

# Photo placeholder
photo_placeholder(
    '[ Photo — Waste Warriors team / field activity in Corbett or Himachal Pradesh ]',
    height_cm=4.5)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 3 — THE PROBLEM
# ══════════════════════════════════════════════════════════════════════════════
heading1('The Waste Crisis at Corbett\'s Edge')

# Two-column layout: text left, photo right
t=doc.add_table(rows=1,cols=2); no_border(t)
lc,rc=t.rows[0].cells[0],t.rows[0].cells[1]
lc.width=Cm(9.4); rc.width=Cm(6.4)
cm_(lc,top=40,bot=40,l=0,r=200); cm_(rc,top=0,bot=0,l=120,r=0)

p=lc.paragraphs[0]; p.paragraph_format.space_after=Pt(5)
R(p,'Corbett Tiger Reserve is home to over 600 species of birds, the Bengal tiger, Asian elephant, '
    'and leopard. It is one of India\'s most ecologically significant landscapes.', sz=10.5, col=DARK)
p2=lc.add_paragraph(); p2.paragraph_format.space_after=Pt(5)
R(p2,'In recent years, tourism around the reserve has witnessed remarkable growth — bringing '
    'livelihoods, but also waste. The increasing volume is driven by booming hospitality, '
    'changing rural lifestyles, and rising consumption of packaged goods.', sz=10.5, col=DARK)
p3=lc.add_paragraph(); p3.paragraph_format.space_after=Pt(0)
R(p3,'Much of this is wet, organic waste — left unmanaged, it rots in the open, leaches into '
    'rivers, and draws wildlife towards human settlements, increasing conflict.', sz=10.5, col=DARK)

sc(rc, PHOTO_HEX); cm_(rc,top=280,bot=280,l=80,r=80); vc(rc)
pp=rc.paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; pp.paragraph_format.space_after=Pt(0)
R(pp,'[ Photo —\nwaste near\nvillage/river ]', sz=8.5, col=GREY, italic=True)

sp(10)

# Hero stat callout
callout(
    '≈ 42 metric tonnes of waste generated every single day\n'
    'in the Corbett region — a burden steadily pressing on its forests, rivers, and villages.',
    fill=LAV_HEX, sz=12, bold=False, italic=True, color=DARK)

body('According to Waste Warriors\' annual waste assessment of the Corbett region. Most of this '
     'is unsegregated at source. Without a formal system, wet waste is either dumped in the open, '
     'mixed with dry and plastic waste (spoiling recyclables), or burned — releasing toxins in '
     'one of India\'s most sensitive ecosystems.', after=4, align=WD_ALIGN_PARAGRAPH.LEFT)

hrule()
sp(4)

# Three-column problem stats
stat_box([
    ('42 MT/day', 'Waste generated\nin the Corbett region'),
    ('600+', 'Bird species at risk\nfrom waste pollution'),
    ('~75%', 'Waste unsegregated\nat source'),
], fill=SAGE_HEX)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 4 — OUR MODEL: THE PARYAVARAN SAKHIS
# ══════════════════════════════════════════════════════════════════════════════
heading1('Our Proven Model: The Paryavaran Sakhis')

heading2('The Sakhi Initiative', underline=True)
body('In 2021, Waste Warriors launched the Paryavaran Sakhi model — "friend of the environment" '
     '— a community entrepreneurship initiative that addresses the waste crisis through women\'s '
     'empowerment, gender equality, and dignified livelihoods.')

body('Today, 25 women from diverse backgrounds are professional environmental service providers, '
     'not passive beneficiaries. They collect, segregate, transport, and process waste — and now, '
     'with a wet waste processing facility in place, they are ready to do even more.')

# Two-column: infra stats left, photo right
t=doc.add_table(rows=1,cols=2); no_border(t)
lc,rc=t.rows[0].cells[0],t.rows[0].cells[1]
lc.width=Cm(7.8); rc.width=Cm(8.0)
cm_(lc,top=0,bot=0,l=0,r=200); cm_(rc,top=0,bot=0,l=120,r=0)

# Stat cells stacked in left
sub=doc.add_table(rows=3,cols=1) # will be placed inline differently
# Instead use paragraph approach in the cell
for stat_txt, label in [
    ('480 sq. ft.', 'Wet waste processing shed at\nHimmatpur Dotiyal'),
    ('400 kg/day', 'Processing capacity of the facility'),
    ('+120 days/month', 'Additional livelihood days for Sakhis\nthrough wet waste operations'),
]:
    p=lc.add_paragraph(); p.paragraph_format.space_after=Pt(8)
    R(p, stat_txt, bold=True, sz=18, col=GREEN)
    p2=lc.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(10)
    R(p2, label, sz=9.5, col=GREY)
# clear first empty para
lc.paragraphs[0]._p.getparent().remove(lc.paragraphs[0]._p)

sc(rc, PHOTO_HEX); cm_(rc,top=300,bot=300,l=80,r=80); vc(rc)
pp=rc.paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; pp.paragraph_format.space_after=Pt(0)
R(pp,'[ Photo — Paryavaran Sakhis\nat the wet waste facility ]', sz=8.5, col=GREY, italic=True)
sp(8)

callout(
    '"The Sakhis were primarily engaged in dry waste. With the wet waste facility, they can now '
    'serve hotels and resorts — creating more consistent, year-round income."',
    fill=LAV_HEX, italic=True, sz=10.5)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 5 — THE PROJECT
# ══════════════════════════════════════════════════════════════════════════════
heading1('The Project: Mitti Se Mitti')
heading2('कचरे से खाद, खाद से समृद्धि   ·   From Soil, To Soil', underline=False)
sp(2)

body('The infrastructure is ready. The Sakhis are trained. What this project funds is the operations, '
     'awareness, and systems needed to run the wet waste machine at full scale — and to ensure the '
     'communities and businesses around Corbett are part of it.')

sp(4)

# 3-month overview strip
heading2('What we will do — in 3 months', underline=True)
sp(2)

# Activity cards in 2x3 grid
activities = [
    ('🌱  Process 6,000 kg/month', 'Convert wet waste to quality compost at the Himmatpur Dotiyal facility every month.'),
    ('👩  Train 6 Paryavaran Sakhis', 'End-to-end training: awareness, collection, transportation, and on-site composting.'),
    ('🏘️  12 Community Sessions', 'Chai-pe-charcha & focus groups on source segregation and home composting.'),
    ('🏨  6 Stakeholder Meetings', 'Engage hotels & resorts to segregate at source and route wet waste to the facility.'),
    ('🪧  Signboards & IEC', 'Direction boards, knowledge-hub signage, home-composting recognition kits for 25 households.'),
    ('🌿  Brand the Compost', 'Design a name and identity for the Sakhis\' ready-to-use compost — turning waste into a product.'),
]
for row_i in range(0, len(activities), 2):
    t=doc.add_table(rows=1,cols=2); no_border(t)
    for col_i in range(2):
        idx=row_i+col_i
        if idx >= len(activities): break
        title, desc = activities[idx]
        c=t.rows[0].cells[col_i]
        sc(c, SAGE_HEX if (row_i+col_i)%2==0 else 'FFFFFF')
        cm_(c,top=100,bot=80,l=140,r=120)
        p=c.paragraphs[0]; p.paragraph_format.space_after=Pt(3)
        R(p, title, bold=True, sz=10, col=GREEN)
        p2=c.add_paragraph(); p2.paragraph_format.space_before=Pt(0); p2.paragraph_format.space_after=Pt(0)
        R(p2, desc, sz=9.5, col=GREY)
    sp(2)

sp(4)

# Vehicle need — woven naturally
callout(
    'Note on field operations:  Reliable waste collection and transportation across the villages '
    'and hotels on Corbett\'s periphery requires a dedicated 4-wheeler vehicle. This is a critical '
    'operational need — for both the day-to-day collection runs and to transport processed compost '
    'to buyers — and support for the same would go a long way in scaling this initiative.',
    fill=SAND_HEX, sz=10.5, color=DARK)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 6 — OUTCOMES + PHOTO GALLERY
# ══════════════════════════════════════════════════════════════════════════════
heading1('What Success Looks Like')

stat_box([
    ('18,000 kg', 'Wet waste processed\ninto compost'),
    ('6', 'Paryavaran Sakhis\ntrained end-to-end'),
    ('12', 'Community awareness\nsessions'),
    ('6', 'Hotels / resorts\nonboarded'),
], fill=SAGE_HEX)

# Outcomes table (clean, no heavy color)
outcomes = [
    ('🌍  Environmental',
     '18,000 kg of wet waste diverted from open dumping — preventing methane, leachate, and '
     'wildlife attraction on Corbett\'s periphery.'),
    ('👩‍💼  Gender & Livelihoods',
     '6 Sakhis with expanded skills and +120 additional livelihood days per month — '
     'growing towards year-round, independent income.'),
    ('🏘️  Behaviour Change',
     'Measurable improvement in household source segregation across target villages, '
     'reinforced by awareness sessions and a home-composting movement.'),
    ('🏨  Industry Adoption',
     'Hotels and resorts formally onboarded to segregate and route wet waste — '
     'a replicable template for the wider hospitality sector.'),
    ('💰  A New Revenue Stream',
     'Branded, ready-to-use compost sold to nurseries, farmers, and hotels — '
     'turning waste into income that outlasts this grant.'),
]

t=doc.add_table(rows=len(outcomes),cols=2); no_border(t)
for i,(icon_title, desc) in enumerate(outcomes):
    lc,rc=t.rows[i].cells[0],t.rows[i].cells[1]
    lc.width=Cm(4.2); rc.width=Cm(11.8)
    fill=SAGE_HEX if i%2==0 else 'F7F7F7'
    sc(lc,fill); sc(rc,fill)
    cm_(lc,top=80,bot=80,l=120,r=120); cm_(rc,top=80,bot=80,l=120,r=120)
    vc(lc); vc(rc)
    p=lc.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    R(p, icon_title, bold=True, sz=9.5, col=GREEN)
    p=rc.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    R(p, desc, sz=9.5, col=DARK)
sp(8)

# Photo gallery row
photo_placeholder(
    caption='', height_cm=3.8, cols=3,
    captions=[
        '[ Photo — Wet waste being processed ]',
        '[ Photo — Sakhis at work ]',
        '[ Photo — Corbett landscape / forest ]',
    ])

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 7 — TIMELINE + BUDGET
# ══════════════════════════════════════════════════════════════════════════════
heading1('Budget and Timeline')

heading2('Timeline: 3 months, FY 2026–27', underline=True)
sp(2)

# Timeline strip
tl_data = [
    ('Month 1', 'Setup & Mobilise',
     'Sakhi training · Procure consumables & safety gear · '
     'Design compost brand · First awareness sessions · Install signboards · Begin processing'),
    ('Month 2', 'Scale & Engage',
     'Full-scale processing (6,000 kg) · Stakeholder meetings with hotels · '
     'Home-composting kits distributed · IEC collateral installed at site'),
    ('Month 3', 'Consolidate & Sustain',
     'Final awareness rounds · Compost market linkage · '
     'Recognise composting households · Documentation & final report'),
]
for m,phase,desc in tl_data:
    t=doc.add_table(rows=1,cols=3); no_border(t)
    c0,c1,c2=t.rows[0].cells[0],t.rows[0].cells[1],t.rows[0].cells[2]
    c0.width=Cm(2.0); c1.width=Cm(3.8); c2.width=Cm(10.2)
    sc(c0,GREEN_HEX); sc(c1,SAGE_HEX); sc(c2,'FAFAFA')
    cm_(c0,top=100,bot=100,l=80,r=80); cm_(c1,top=100,bot=100,l=120,r=120)
    cm_(c2,top=100,bot=100,l=140,r=140)
    vc(c0); vc(c1); vc(c2)
    p=c0.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(0)
    R(p, m, bold=True, sz=10, col=WHITE)
    p=c1.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    R(p, phase, bold=True, sz=10, col=GREEN)
    p=c2.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    R(p, desc, sz=9.5, col=DARK)
    sp(2)

sp(8)
heading2('Budget Summary — Total: ₹2,75,420', underline=True)
p=P(after=4)
R(p, 'The estimated budget required for the Mitti Se Mitti project, FY 2026–27, is ', sz=10.5, col=DARK)
R(p, '₹2,75,420/-', bold=True, sz=10.5, col=GREEN)
R(p, ' (Two Lakh Seventy-Five Thousand Four Hundred and Twenty Only).', sz=10.5, col=DARK)

# Budget table — clean black-bordered (same style as Trek & Trails)
bud_rows = [
    ('A', 'Awareness & IEC', '', '1,47,920'),
    ('A a)', 'Awareness Sessions (chai-pe-charcha, FGDs, exposure visits)', '12 × ₹660', '7,920'),
    ('A b)', "Stakeholders' Meet (hotels & resorts)", '6 × ₹5,000', '30,000'),
    ('A c)', 'IEC Collaterals (on-site knowledge hub)', '–', '50,000'),
    ('A d)', 'Signboards (direction + composting recognition)', '4 × ₹15,000', '60,000'),
    ('B', 'Operations', '', '1,07,500'),
    ('B a)', 'Waste Management Consumables (gloves, drums, trolley, earthen pots, guides, certificates, compost brand)', '–', '92,500'),
    ('B b)', 'Training & Capacity Building of 6 Paryavaran Sakhis', '3 × ₹5,000', '15,000'),
    ('C', 'Office Expenses', '', '20,000'),
    ('C a)', 'Printing & Stationery (stakeholder report & office)', '–', '10,000'),
    ('C b)', 'Local Travel', '–', '10,000'),
    ('TOT', 'TOTAL', '', '2,75,420'),
]

bt=doc.add_table(rows=1+len(bud_rows), cols=4)
grid_border(bt, color='AAAAAA', sz='4')
headers=['Code','Head / Activity','Target / Frequency','Amount (₹)']
for j,h in enumerate(headers):
    cell=bt.rows[0].cells[j]
    sc(cell,'333333'); cm_(cell,top=80,bot=80,l=100,r=100)
    p=cell.paragraphs[0]
    if j>=2: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after=Pt(0)
    R(p,h,bold=True,sz=9,col=WHITE)

for i,row in enumerate(bud_rows):
    code,name,freq,amt=row
    is_section=len(code)==1 or code=='TOT'
    tr=bt.rows[i+1]
    for j in range(4):
        c=tr.cells[j]
        if is_section:
            sc(c,SAGE_HEX if code!='TOT' else GREEN_HEX)
        elif i%2==0:
            sc(c,'F9F9F9')
        cm_(c,top=70,bot=70,l=100,r=100)
        p=c.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
        if j>=2: p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
        vals=[code,name,freq,amt]
        is_tot=code=='TOT'
        R(p, vals[j],
          bold=is_section,
          sz=9 if not is_section else 9.5,
          col=WHITE if is_tot else (GREEN if is_section else DARK))
sp(8)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 8 — DUE DILIGENCE + CONTACT
# ══════════════════════════════════════════════════════════════════════════════
heading1('WWS Due Diligence')

dd=[
    ('1','NGO Registration','Yes','17-11-2012','17-11-2027'),
    ('2','Permanent Account Number','Yes','17-11-2012','Perpetual'),
    ('3','Certificate under 12A','Yes','02-03-2026','31-03-2036'),
    ('4','Certificate under 80G','Yes','02-03-2026','31-03-2031'),
    ('5','CSR 1 Registration','Yes','19-04-2021','Perpetual'),
    ('6','FCRA (Renewal Applied)','Yes','15-11-2017','01-04-2028'),
    ('7','Employee Provident Fund (EPF)','Yes','27-11-2015','Perpetual'),
    ('8','Audited Balance Sheets (last 3 yrs)','Yes','–','–'),
]
ddt=doc.add_table(rows=1+len(dd),cols=5)
grid_border(ddt,'AAAAAA','4')
for j,h in enumerate(['S.No.','Particulars','Applicable?','Registration Date','Valid Till']):
    c=ddt.rows[0].cells[j]; sc(c,'333333'); cm_(c,top=70,bot=70,l=100,r=100)
    p=c.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    R(p,h,bold=True,sz=9,col=WHITE)
for i,row in enumerate(dd):
    for j,val in enumerate(row):
        c=ddt.rows[i+1].cells[j]
        sc(c,'F4F4F4' if i%2==0 else 'FFFFFF')
        cm_(c,top=60,bot=60,l=100,r=100)
        p=c.paragraphs[0]; p.paragraph_format.space_after=Pt(0)
        p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        R(p,val,sz=9,col=DARK)

sp(6)
heading2('Other Donors / Partners of Waste Warriors Society', underline=False)
body('Parle Biscuits Pvt. Ltd. · Airbnb · MacArthur Foundation · HDFC Bank Ltd · HT Parekh Foundation '
     '· Lal Family Foundation · Give India Foundation · Make My Trip Foundation · Rain Matter Foundation '
     '· EdelGive Foundation · Rohini Nilekani Philanthropies · Godrej Consumer Products Ltd · SBI Foundation',
     after=4, align=WD_ALIGN_PARAGRAPH.LEFT)

hrule()
sp(6)

# Contact block
p=P(align=WD_ALIGN_PARAGRAPH.CENTER,after=4)
R(p,'OUR SINCERE GRATITUDE TO YOU FOR CONSIDERING OUR PROPOSAL',
  bold=True, sz=11, col=DARK)

ct=doc.add_table(rows=1,cols=2); no_border(ct)
left,right=ct.rows[0].cells[0],ct.rows[0].cells[1]
sc(left,SAGE_HEX); sc(right,LAV_HEX)
cm_(left,top=140,bot=140,l=160,r=160); cm_(right,top=140,bot=140,l=160,r=160)
for cell, lines in [
    (left,[('Mayank Sharma','CEO, Waste Warriors Society','mayank.sharma@wastewarriors.org')]),
    (right,[('Lehar Arora','Manager — Partnerships','lehar.arora@wastewarriors.org')]),
]:
    for i,(name,role,email) in enumerate(lines):
        p=cell.paragraphs[0] if i==0 else cell.add_paragraph()
        p.paragraph_format.space_after=Pt(2)
        R(p,name,bold=True,sz=10.5,col=DARK)
        p2=cell.add_paragraph(); p2.paragraph_format.space_after=Pt(1); p2.paragraph_format.space_before=Pt(0)
        R(p2,role,sz=9.5,col=GREY)
        p3=cell.add_paragraph(); p3.paragraph_format.space_after=Pt(0); p3.paragraph_format.space_before=Pt(0)
        R(p3,email,sz=9.5,col=GREEN)
sp(8)

# Footer
p=P(align=WD_ALIGN_PARAGRAPH.CENTER,after=1)
R(p,'Waste Warriors Society (WWS)',bold=True,sz=10.5,col=DARK)
p=P(align=WD_ALIGN_PARAGRAPH.CENTER,after=1)
R(p,'Registered under the Societies Registration Act 1860 · Number 243/2012-2013',sz=9,col=GREY)
p=P(align=WD_ALIGN_PARAGRAPH.CENTER,after=1)
R(p,'136/2/2, Shivam Vihar, Jakhan, Rajpur Road, Dehradun, Uttarakhand – 248001',sz=9,col=GREY)
p=P(align=WD_ALIGN_PARAGRAPH.CENTER,after=0)
R(p,'Phone: +91 7505763049  |  Email: partnerships@wastewarriors.org  |  Website: www.wastewarriors.org',
  sz=9,col=GREY)

# ── Save ──────────────────────────────────────────────────────────────────────
out='/home/user/jcode/Mitti_Se_Mitti_v2_EGF.docx'
doc.save(out)
print('Saved:', out)
