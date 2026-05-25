from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3)
    section.right_margin  = Cm(2.5)

# ── Colour palette ─────────────────────────────────────────────────────────────
DARK_GREEN  = RGBColor(0x1A, 0x5C, 0x3A)  # deep forest green
MID_GREEN   = RGBColor(0x2E, 0x86, 0x48)  # heading accent
LIGHT_GREEN = RGBColor(0xD4, 0xED, 0xDA)  # table-header fill (hex: D4EDDA)
ACCENT_BLUE = RGBColor(0x15, 0x63, 0xAA)  # hyperlink / sponsor
DARK_GREY   = RGBColor(0x33, 0x33, 0x33)
BLACK       = RGBColor(0x00, 0x00, 0x00)

# ── Helper: set paragraph shading ─────────────────────────────────────────────
def shade_paragraph(para, hex_fill: str):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_fill)
    pPr.append(shd)

def shade_cell(cell, hex_fill: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_fill)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    """kwargs: top/bottom/left/right = {'sz':4,'val':'single','color':'000000'}"""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        if edge in kwargs:
            tag = OxmlElement(f'w:{edge}')
            for k, v in kwargs[edge].items():
                tag.set(qn(f'w:{k}'), str(v))
            tcBorders.append(tag)
    tcPr.append(tcBorders)

# ── Helper: add a run with explicit formatting ─────────────────────────────────
def add_run(para, text, bold=False, italic=False, size=11,
            color=None, underline=False):
    run = para.add_run(text)
    run.bold      = bold
    run.italic    = italic
    run.underline = underline
    run.font.size = Pt(size)
    run.font.color.rgb = color or BLACK
    return run

# ── Helper: section heading (numbered, dark-green bar) ────────────────────────
def section_heading(doc, number, title):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(14)
    para.paragraph_format.space_after  = Pt(4)
    shade_paragraph(para, '1A5C3A')          # dark green bg
    run = para.add_run(f"  {number}.  {title.upper()}")
    run.bold           = True
    run.font.size      = Pt(12)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)   # white text
    return para

# ── Helper: sub-heading ────────────────────────────────────────────────────────
def sub_heading(doc, text):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(8)
    para.paragraph_format.space_after  = Pt(2)
    run = para.add_run(text)
    run.bold           = True
    run.font.size      = Pt(11)
    run.font.color.rgb = MID_GREEN
    return para

# ── Helper: body paragraph ─────────────────────────────────────────────────────
def body(doc, text, space_after=4):
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(space_after)
    run = para.add_run(text)
    run.font.size      = Pt(10.5)
    run.font.color.rgb = DARK_GREY
    return para

# ── Helper: bullet ────────────────────────────────────────────────────────────
def bullet(doc, text, level=0):
    para = doc.add_paragraph(style='List Bullet')
    para.paragraph_format.left_indent  = Inches(0.25 + level * 0.25)
    para.paragraph_format.space_after  = Pt(2)
    run = para.add_run(text)
    run.font.size      = Pt(10.5)
    run.font.color.rgb = DARK_GREY
    return para

# ── Helper: horizontal rule ───────────────────────────────────────────────────
def hrule(doc):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(4)
    para.paragraph_format.space_after  = Pt(4)
    pPr  = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot  = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    '6')
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), '1A5C3A')
    pBdr.append(bot)
    pPr.append(pBdr)

# ══════════════════════════════════════════════════════════════════════════════
#  COVER / HEADER BLOCK
# ══════════════════════════════════════════════════════════════════════════════
# Annexure label
ann = doc.add_paragraph()
ann.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = ann.add_run("Annexure – 1")
run.italic    = True
run.font.size = Pt(10)
run.font.color.rgb = DARK_GREY

doc.add_paragraph()   # spacer

# Main title
title_para = doc.add_paragraph()
title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_para.paragraph_format.space_before = Pt(6)
title_para.paragraph_format.space_after  = Pt(2)
tr = title_para.add_run("FORMAT OF PROJECT PROPOSAL FOR GRANT UNDER CSR")
tr.bold           = True
tr.font.size      = Pt(14)
tr.underline      = True
tr.font.color.rgb = DARK_GREEN

doc.add_paragraph()   # spacer

# ── Sponsoring-Agency info table ──────────────────────────────────────────────
info_tbl = doc.add_table(rows=2, cols=2)
info_tbl.style = 'Table Grid'
info_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

def info_row(tbl, row_idx, label, value_lines):
    label_cell = tbl.rows[row_idx].cells[0]
    value_cell = tbl.rows[row_idx].cells[1]
    label_cell.width = Cm(6)
    value_cell.width = Cm(11)
    shade_cell(label_cell, 'D4EDDA')
    lp = label_cell.paragraphs[0]
    lr = lp.add_run(label)
    lr.bold = True; lr.font.size = Pt(10); lr.font.color.rgb = DARK_GREEN
    for i, vl in enumerate(value_lines):
        if i == 0:
            vp = value_cell.paragraphs[0]
        else:
            vp = value_cell.add_paragraph()
        vr = vp.add_run(vl)
        vr.font.size = Pt(10); vr.font.color.rgb = DARK_GREY

info_row(info_tbl, 0, "Name of Sponsoring Agency :",
         ["National Backward Classes Finance & Development Corporation (NBCFDC)",
          "A Govt. of India Undertaking, M/o Social Justice & Empowerment, Govt. of India"])
info_row(info_tbl, 1, "Contact Details :\n(Name, Phone, Fax, E-mail)",
         ["Sh. Ajit Kumar Samal, Senior General Manager (Finance/CS)",
          "Phone: 011-45854400 / 9717699911",
          "E-mail: cs@nbcfdc.gov.in"])

doc.add_paragraph()   # spacer

# ── Part I heading ────────────────────────────────────────────────────────────
p1_para = doc.add_paragraph()
p1_para.paragraph_format.space_before = Pt(6)
p1_para.paragraph_format.space_after  = Pt(2)
p1_run = p1_para.add_run("Part – I:    FORMAT FOR SUBMISSION OF PROJECT PROPOSAL")
p1_run.bold           = True
p1_run.underline      = True
p1_run.font.size      = Pt(12)
p1_run.font.color.rgb = DARK_GREEN

intro = doc.add_paragraph()
intro.paragraph_format.space_after = Pt(6)
ir = intro.add_run(
    "All the proposals, unless otherwise desired, shall be submitted by the Implementing "
    "Agency/District Authorities as per the following format. The project proposal should be "
    "as per format in Part I mentioned herein and should be accompanied by the documents as per Part II.")
ir.font.size = Pt(10.5); ir.font.color.rgb = DARK_GREY

hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  1. TITLE
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 1, "Title of the Project")

t = doc.add_paragraph()
t.paragraph_format.space_after = Pt(4)
tr1 = t.add_run("The Good Trek Project: ")
tr1.bold = True; tr1.font.size = Pt(11); tr1.font.color.rgb = MID_GREEN
tr2 = t.add_run("Preserving Himalayan Trails Through Responsible Waste Management")
tr2.font.size = Pt(11); tr2.font.color.rgb = DARK_GREY

hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  2. SOCIO-ECONOMIC STATUS
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 2, "Current Socio-Economic Status of the Area and Target Group")

body(doc,
     "The project is situated in Sankri village, Uttarkashi district, Uttarakhand — the primary "
     "gateway to the Kedarkantha Trek, located within the ecologically sensitive Govind Wildlife "
     "Sanctuary and National Park. The Indian Himalayan Region (IHR) is home to over 50 million "
     "people, of whom approximately 75% reside in rural areas and depend heavily on agriculture, "
     "pastoralism, and an increasingly critical tourism economy.")

sub_heading(doc, "Skill Level of Beneficiaries")
rows_skill = [
    ("Skilled",      "Trek guides with certified training, experienced mule handlers"),
    ("Semi-Skilled", "Porters, waste-collection workers, community mobilisers"),
    ("Unskilled",    "Daily-wage labourers in waste handling, local sorters at drop-off points"),
]
tbl_skill = doc.add_table(rows=1+len(rows_skill), cols=2)
tbl_skill.style = 'Table Grid'
for j, hdr in enumerate(["Skill Category", "Description / Role"]):
    cell = tbl_skill.rows[0].cells[j]
    shade_cell(cell, '1A5C3A')
    r = cell.paragraphs[0].add_run(hdr)
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
for i, (cat, desc) in enumerate(rows_skill):
    shade_cell(tbl_skill.rows[i+1].cells[0], 'D4EDDA') if i%2==0 else None
    for j, txt in enumerate([cat, desc]):
        r = tbl_skill.rows[i+1].cells[j].paragraphs[0].add_run(txt)
        r.font.size = Pt(10); r.font.color.rgb = DARK_GREY

doc.add_paragraph()

sub_heading(doc, "Current Income Level & Gender Profile")
body(doc,
     "Local communities around the Kedarkantha route are predominantly dependent on trekking-season "
     "revenues (October–May). Average daily earnings for porters and guides range from ₹400–₹800 "
     "per day during peak season, with negligible income during the monsoon closure (June–September). "
     "Women remain underrepresented in formal trekking roles; however, Waste Warriors actively "
     "promotes dignified livelihood opportunities for women through waste-handling and community "
     "engagement roles, consistent with its organisational mandate to ignite urgency among women and "
     "youth to earn identity through dignified livelihoods.")

body(doc,
     "Tourism in the region generates over ₹200 crore annually but the burden of unmanaged waste "
     "falls disproportionately on local communities — degrading grazing lands, water sources, and "
     "village aesthetics — while trekkers return to urban centres unaffected.")

hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  3. OBJECTIVES
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 3, "Objectives of the Project")

objectives = [
    ("Establish a Structured Waste Management System",
     "Create a formal, agency-led system for waste collection, segregation, and transportation "
     "along the Kedarkantha Trek route where none previously existed."),
    ("Divert Trek Waste from Fragile Ecosystems",
     "Collect and divert a minimum of 6 metric tonnes (MT) of trek-generated solid waste from "
     "the high-altitude Himalayan ecosystem over the project period."),
    ("Ensure Responsible Processing",
     "Ensure at least 80% of collected waste is properly segregated at source and transported to "
     "Waste Warriors' Material Recovery Facility (MRF) in Dehradun for sorting, recovery, and "
     "verified recycling."),
    ("Engage Trekking Agencies",
     "Formalise collaboration with 45+ trekking agencies operating on the Kedarkantha route, "
     "embedding waste responsibility into their operations and guest experience."),
    ("Promote Behaviour Change Among Trekkers",
     "Reach a minimum of 5,000 trekkers with targeted behaviour-change communication and "
     "responsible-trekking awareness initiatives."),
    ("Build a Scalable and Replicable Model",
     "Develop an institutionalised trek-waste management model that can be adopted across other "
     "trekking corridors in the Indian Himalayan Region."),
    ("Strengthen Local Livelihoods",
     "Create dignified employment opportunities for local community members — particularly women "
     "and youth — in waste collection, handling, and awareness activities."),
]

for i, (obj_title, obj_desc) in enumerate(objectives):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_after = Pt(4)
    p.add_run(f"{i+1}.  ").font.size = Pt(10.5)
    rb = p.add_run(f"{obj_title}: "); rb.bold = True
    rb.font.size = Pt(10.5); rb.font.color.rgb = MID_GREEN
    rd = p.add_run(obj_desc); rd.font.size = Pt(10.5); rd.font.color.rgb = DARK_GREY

hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  4. EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 4, "Executive Summary of Proposal")

body(doc,
     "The Indian Himalayan Region (IHR) is one of the world's most ecologically fragile landscapes. "
     "Nearly one million trekkers explore mountain trails annually through organised expeditions, "
     "generating approximately 288 grams of solid waste per visitor per day. Popular routes such as "
     "Kedarkantha, Har Ki Dun, and Valley of Flowers bear the brunt of this growing waste crisis — "
     "waste that is burned, buried, or left on slopes in the absence of any formal management system.")

body(doc,
     "Waste Warriors Society (WWS) — a non-profit organisation with 240+ field staff across "
     "Uttarakhand and Himachal Pradesh — implemented a successful pilot trek-waste management "
     "initiative in Sankri village (December 2025 – February 2026), collecting approximately 2 "
     "metric tonnes of trek waste for the first time and transporting it ~200 km to WWS's Material "
     "Recovery Facility (MRF) in Dehradun for processing.")

body(doc,
     "Building on this proven pilot, The Good Trek Project seeks CSR support of ₹9,78,650/- "
     "(Nine Lakhs Seventy Eight Thousand Six Hundred and Fifty Only) for FY 2026-27 to scale this "
     "model into a structured, agency-led waste management system across the Kedarkantha Trek. "
     "The project will partner with 45+ trekking agencies, train guides and porters, establish "
     "designated waste drop-off points, and engage 5,000+ trekkers in responsible trekking "
     "practices — collectively diverting 6+ MT of waste from the Himalayan ecosystem and "
     "demonstrating a replicable blueprint for responsible mountain tourism across India.")

hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  5. TARGETED GROUP
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 5, "Targeted Group of the Project Proposal")

groups = [
    ("Trekking Agencies (45+)",
     "Registered trekking operators on the Kedarkantha route who will integrate waste "
     "management into their operational packages."),
    ("Trek Guides & Porters",
     "Semi-skilled individuals who lead trekking groups and carry supplies — trained to "
     "oversee on-trail waste segregation and collection."),
    ("Local Community Members",
     "Residents of Sankri and surrounding villages, particularly women and youth, engaged "
     "in waste collection, drop-off point management, and awareness activities."),
    ("Trekkers (5,000+)",
     "Adventure tourists visiting the Kedarkantha route, reached through behaviour-change "
     "communication for responsible trekking practices."),
    ("Local Governing Bodies & Forest Department",
     "Stakeholders from Uttarakhand Forest Department, Gram Panchayats, and tourism "
     "boards engaged to institutionalise responsible waste practices."),
]

tbl_grp = doc.add_table(rows=1+len(groups), cols=2)
tbl_grp.style = 'Table Grid'
for j, hdr in enumerate(["Target Group", "Role / Engagement"]):
    cell = tbl_grp.rows[0].cells[j]
    shade_cell(cell, '1A5C3A')
    r = cell.paragraphs[0].add_run(hdr)
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
for i, (grp, role) in enumerate(groups):
    row = tbl_grp.rows[i+1]
    if i % 2 == 0:
        shade_cell(row.cells[0], 'D4EDDA')
        shade_cell(row.cells[1], 'D4EDDA')
    for j, txt in enumerate([grp, role]):
        r = row.cells[j].paragraphs[0].add_run(txt)
        r.font.size = Pt(10); r.font.color.rgb = DARK_GREY
        if j == 0: r.bold = True

doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  6. GEOGRAPHICAL AREA
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 6, "Geographical Area of Execution and Reason for Selection")

body(doc,
     "Project Area:  Kedarkantha Trek Route — Sankri Village, Uttarkashi District, Uttarakhand. "
     "The project corridor spans the trail from the Sankri base camp (approximately 1,920 m) to "
     "the Kedarkantha summit (3,810 m), traversing the Govind Wildlife Sanctuary and National Park.")

sub_heading(doc, "Rationale for Selection")
reasons = [
    "High footfall and acute waste impact: Approximately 20,000 trekkers annually (300–500/day at peak), "
     "generating significant solid waste with no prior formal management system.",
    "Existing WWS presence and pilot success: The Sankri pilot (Dec 2025–Feb 2026) already diverted 2 MT "
     "of waste, proving feasibility and establishing community trust.",
    "Ecological sensitivity: The route lies within the snow leopard conservation landscape of Govind "
     "Wildlife Sanctuary — making responsible waste management a conservation imperative.",
    "Connected to WWS's Dehradun MRF: The established logistics corridor (~200 km) to the Material "
     "Recovery Facility in Harrawala, Dehradun, enables responsible downstream processing.",
    "Replication potential: Kedarkantha is a flagship template; success here creates a blueprint for "
     "other trekking corridors (Har Ki Dun, Valley of Flowers, Gaumukh Tapovan, etc.).",
]
for r in reasons:
    bullet(doc, r)

doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  7. PROJECT COST
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 7, "Project Cost")

sub_heading(doc, "(A)  Budget Details — Detailed Cost Break-Up (Head-wise and Activity-wise)")

# Budget table
budget_headers = ["Budget\nCode", "Budget Head", "Unit Cost\n(₹)", "No. of\nUnits", "Annual Budget\n(₹)"]
budget_rows = [
    ("A",  "OPERATIONAL EXPENSES",                                    "",      "",       "2,75,000"),
    ("A1", "Waste Handling, Transportation & Processing to MRF Dehradun", "20,000", "5 trips",  "1,00,000"),
    ("A2", "Labour/Mule charges for waste collection",                "1,000", "12",      "12,000"),
    ("A3", "Community Engagement Activities",                         "25,000","1",       "25,000"),
    ("A4", "IEC and Project Collaterals",                             "10,000","5",       "50,000"),
    ("A5", "Transportation / Local Conveyance",                       "1,000", "20",      "20,000"),
    ("A6", "Signboards",                                              "17,000","4",       "68,000"),
    ("B",  "CAPITAL EXPENSES",                                        "",      "",       "2,16,000"),
    ("B1", "Waste Consumables (bags, gloves, litter pickers, safety equipment, weighing machine)", "60", "1,600 units", "96,000"),
    ("B2", "Waste Drop-Off Points (infrastructure)",                  "1,20,000","1",   "1,20,000"),
    ("C",  "PROJECT TEAM SALARIES",                                   "",      "",       "3,60,000"),
    ("C1", "Project Associate",                                       "30,000","8 months","2,40,000"),
    ("C2", "Project Executive",                                       "15,000","8 months","1,20,000"),
    ("E",  "ADMINISTRATIVE EXPENSES (10%)",                           "",      "",       "1,27,650"),
    ("",   "TOTAL PROJECT COST",                                      "",      "",       "9,78,650"),
]

tbl_bud = doc.add_table(rows=1+len(budget_rows), cols=5)
tbl_bud.style = 'Table Grid'
col_widths = [Cm(1.8), Cm(7.5), Cm(2.4), Cm(2.2), Cm(2.8)]

for j, hdr in enumerate(budget_headers):
    cell = tbl_bud.rows[0].cells[j]
    shade_cell(cell, '1A5C3A')
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(hdr)
    r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)

section_codes = {"A","B","C","E"}
for i, row_data in enumerate(budget_rows):
    row = tbl_bud.rows[i+1]
    code = row_data[0]
    is_section = code in section_codes or row_data[1] == "TOTAL PROJECT COST"
    is_total   = row_data[1] == "TOTAL PROJECT COST"
    for j, txt in enumerate(row_data):
        cell = row.cells[j]
        if is_total:
            shade_cell(cell, '1A5C3A')
        elif is_section:
            shade_cell(cell, '2E8648')
        elif i % 2 == 0:
            shade_cell(cell, 'EAF7EE')
        p = cell.paragraphs[0]
        if j in (2,3,4):
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(txt)
        r.font.size = Pt(9)
        if is_total:
            r.bold = True; r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
        elif is_section:
            r.bold = True; r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
        else:
            r.font.color.rgb = DARK_GREY

doc.add_paragraph()

sub_heading(doc, "(B)  Means of Finance of the Project Cost")
body(doc,
     "The total project cost of ₹9,78,650/- (Nine Lakhs Seventy Eight Thousand Six Hundred and "
     "Fifty Only) is proposed to be funded entirely through the CSR grant from the sponsoring "
     "agency under the NBCFDC CSR framework. Waste Warriors Society shall contribute in-kind "
     "through institutional support, field staff oversight, MRF infrastructure access, and "
     "established recycling-partner networks at no additional charge to the project.")

sub_heading(doc, "(C)  Funding Requirements — Phased Disbursement")

phase_data = [
    ("Start of Project\n(Month 1–2)", "40%  =  ₹3,91,460/-",
     "Initial capital expenses (drop-off points, consumables), hiring of project team, "
     "agency onboarding, and IEC material development."),
    ("Mid-Project\n(Month 3–6)",     "40%  =  ₹3,91,460/-",
     "Operational expenses (transportation trips, mule charges, community engagement), "
     "team salaries, and ongoing signage installation."),
    ("End of Project\n(Month 7–8)",  "20%  =  ₹1,95,730/-",
     "Final transportation runs, documentation, monitoring and evaluation, "
     "recognition event for exemplary agencies, and final report submission."),
]
tbl_phase = doc.add_table(rows=1+len(phase_data), cols=3)
tbl_phase.style = 'Table Grid'
for j, hdr in enumerate(["Phase", "Amount", "Purpose"]):
    cell = tbl_phase.rows[0].cells[j]
    shade_cell(cell, '1A5C3A')
    r = cell.paragraphs[0].add_run(hdr)
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
for i, (ph, amt, purp) in enumerate(phase_data):
    row = tbl_phase.rows[i+1]
    if i % 2 == 0:
        for c in row.cells: shade_cell(c, 'D4EDDA')
    for j, txt in enumerate([ph, amt, purp]):
        r = row.cells[j].paragraphs[0].add_run(txt)
        r.font.size = Pt(10); r.font.color.rgb = DARK_GREY
        if j <= 1: r.bold = True

doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  7 (Contd) — ACTIVITY TIMELINE TABLE (as per format template)
# ══════════════════════════════════════════════════════════════════════════════
sub_heading(doc, "Activity-wise Timeline")

tl_headers = ["Activities", "Budget (₹)", "Start Date", "End Date", "No. of Days"]
tl_rows = [
    ("1. Agency engagement & onboarding; MoU signing with 45+ trekking agencies",
     "25,000",  "Jul 2026", "Aug 2026", "45"),
    ("2. Waste drop-off point construction & installation of signboards",
     "1,88,000","Jul 2026", "Sep 2026", "60"),
    ("3. Procurement of waste consumables (kits, bags, equipment)",
     "96,000",  "Aug 2026", "Sep 2026", "30"),
    ("4. Training of guides, porters & local staff; IEC campaigns for trekkers",
     "1,00,000","Sep 2026", "Feb 2027", "150"),
    ("5. Trek waste collection, segregation & transportation to MRF (5 trips)",
     "1,00,000","Oct 2026", "Feb 2027", "120"),
    ("6. Labour / mule charges for on-trail waste collection",
     "12,000",  "Oct 2026", "Feb 2027", "120"),
    ("7. Community engagement activities & recognition events",
     "25,000",  "Nov 2026", "Feb 2027", "90"),
    ("8. Project team salaries (Project Associate + Project Executive)",
     "3,60,000","Jul 2026", "Feb 2027", "245"),
    ("9. Administrative expenses (10%)",
     "1,27,650","Jul 2026", "Feb 2027", "245"),
    ("TOTAL", "9,78,650", "", "", ""),
]

tbl_tl = doc.add_table(rows=1+len(tl_rows), cols=5)
tbl_tl.style = 'Table Grid'
for j, hdr in enumerate(tl_headers):
    cell = tbl_tl.rows[0].cells[j]
    shade_cell(cell, '1A5C3A')
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(hdr)
    r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
for i, row_data in enumerate(tl_rows):
    row = tbl_tl.rows[i+1]
    is_total = row_data[0] == "TOTAL"
    if is_total:
        for c in row.cells: shade_cell(c, '1A5C3A')
    elif i % 2 == 0:
        for c in row.cells: shade_cell(c, 'EAF7EE')
    for j, txt in enumerate(row_data):
        p = row.cells[j].paragraphs[0]
        if j >= 1: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        r.font.size = Pt(9)
        if is_total:
            r.bold = True; r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
        else:
            r.font.color.rgb = DARK_GREY

doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  8. EXECUTION / IMPLEMENTATION SCHEDULE
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 8, "Execution / Implementation Schedule")

phases = [
    ("Phase 1 — Mobilisation & Setup  (July – September 2026)",
     ["Sign MoUs and onboard 45+ trekking agencies operating on the Kedarkantha route.",
      "Procure and distribute waste collection kits (bags, gloves, litter pickers, weighing machines).",
      "Construct and install waste drop-off points at strategic locations along the trail.",
      "Install informational signboards at key points along the route.",
      "Recruit and induct the project team (Project Associate and Project Executive).",
      "Conduct baseline assessment of waste generation and littering hotspots."]),
    ("Phase 2 — Operations & Capacity Building  (October 2026 – January 2027)",
     ["Conduct structured training workshops for trek guides, porters, and local waste handlers on segregation, safe handling, and responsible disposal.",
      "Deploy IEC (Information, Education, Communication) campaigns targeting trekkers at base camps and trail entry points.",
      "Activate waste collection system across the trail: guides oversee segregation; waste consolidated at drop-off points.",
      "Execute 5 transportation runs to MRF Dehradun (~200 km), ensuring chain-of-custody documentation.",
      "Engage local governing bodies — Gram Panchayat, Forest Department, and Uttarakhand Tourism — for institutional support.",
      "Conduct mule/labour-based collection along sections inaccessible to vehicles."]),
    ("Phase 3 — Consolidation, Recognition & Reporting  (February 2027)",
     ["Complete final transportation run of collected waste to MRF for processing.",
      "Host recognition event for trekking agencies demonstrating exemplary waste management.",
      "Document learnings, scale-up recommendations, and model replication guide.",
      "Submit final project completion report with photographic evidence, weight records, and impact metrics to NBCFDC.",
      "Initiate stakeholder consultations for replication on Har Ki Dun and Valley of Flowers routes."]),
]

for phase_title, phase_items in phases:
    sub_heading(doc, phase_title)
    for item in phase_items:
        bullet(doc, item)
    doc.add_paragraph()

hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  9. HUMAN RESOURCES
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 9, "Human Resources Involved in Project Implementation")

body(doc,
     "The project will be implemented under the overall stewardship of Waste Warriors Society. "
     "A dedicated project team will be deployed on the ground, supported by WWS's existing field "
     "infrastructure across Uttarakhand.")

hr_data = [
    ("Mayank Sharma", "Chief Executive Officer — WWS (Project Oversight & Strategic Direction)",
     "+91-7505763049", "mayank.sharma@wastewarriors.org"),
    ("Lehar Arora", "Manager — Partnerships (Donor Liaison & Agency Coordination)",
     "N/A", "lehar.arora@wastewarriors.org"),
    ("Project Associate\n(To be recruited)", "On-ground operations management, agency coordination, data collection",
     "To be assigned", "WWS Field Office, Sankri/Dehradun"),
    ("Project Executive\n(To be recruited)", "Trekker engagement, IEC delivery, waste collection supervision",
     "To be assigned", "WWS Field Office, Sankri"),
    ("Local Trek Guides\n& Porters (45+)", "On-trail waste segregation oversight and responsible collection",
     "Via trekking agencies", "Kedarkantha Trail, Sankri"),
]

tbl_hr = doc.add_table(rows=1+len(hr_data), cols=4)
tbl_hr.style = 'Table Grid'
for j, hdr in enumerate(["Name / Role", "Designation & Responsibility", "Contact", "Email / Location"]):
    cell = tbl_hr.rows[0].cells[j]
    shade_cell(cell, '1A5C3A')
    r = cell.paragraphs[0].add_run(hdr)
    r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
for i, (name, desig, phone, email) in enumerate(hr_data):
    row = tbl_hr.rows[i+1]
    if i % 2 == 0:
        for c in row.cells: shade_cell(c, 'D4EDDA')
    for j, txt in enumerate([name, desig, phone, email]):
        r = row.cells[j].paragraphs[0].add_run(txt)
        r.font.size = Pt(9); r.font.color.rgb = DARK_GREY
        if j == 0: r.bold = True

doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  10. EXPECTED OUTCOMES / BENEFITS
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 10, "Expected Outcome / Benefits of the Project")

outcomes = [
    ("Environmental Impact",
     "Divert a minimum of 6 MT of trek-generated solid waste from the fragile Kedarkantha "
     "ecosystem, preventing open burning, dumping, and littering that degrade alpine habitats, "
     "contaminate water sources, and harm wildlife."),
    ("Waste Processing",
     "Ensure 80%+ of collected waste is properly segregated at source and responsibly processed "
     "at the Dehradun MRF — closing the loop from mountain to material recovery."),
    ("Industry Adoption",
     "Establish formal, documented waste management protocols with 45+ trekking agencies — "
     "embedding environmental stewardship as a standard operating practice, not an afterthought."),
    ("Trekker Awareness",
     "Reach 5,000+ trekkers with responsible-trekking behaviour change messaging — creating "
     "a generation of mindful mountain travellers."),
    ("Livelihood Creation",
     "Generate dignified employment for local community members (especially women and youth) "
     "through waste collection, handling, and awareness roles."),
    ("Institutional Precedent",
     "Establish the first formal trek-waste management system on the Kedarkantha route — "
     "demonstrating to government bodies (Forest Department, Uttarakhand Tourism) and the "
     "trekking industry that a viable, scalable model exists."),
    ("Scalable Blueprint",
     "Produce a documented replication guide for extending the model to Har Ki Dun, Valley "
     "of Flowers, Bali Pass, Kuari Pass, and Gaumukh Tapovan — benefiting the broader "
     "Himalayan trekking ecosystem."),
]

for num, (title, desc) in enumerate(outcomes, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_after = Pt(4)
    rb = p.add_run(f"{num}.  {title}: "); rb.bold = True
    rb.font.size = Pt(10.5); rb.font.color.rgb = MID_GREEN
    rd = p.add_run(desc); rd.font.size = Pt(10.5); rd.font.color.rgb = DARK_GREY

doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  11. MONITORING PLAN
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 11, "Monitoring Plan of the Implementing Agency")

body(doc,
     "Waste Warriors Society will implement a robust monitoring framework to ensure accountability, "
     "track progress against targets, and enable course correction throughout the project lifecycle.")

mon_data = [
    ("Monthly",
     "Internal review of waste collected (weight records), transportation trips completed, agency "
     "compliance, and trekker engagement figures. Monthly Progress Reports (MPR) submitted to NBCFDC "
     "as per the prescribed appendix format."),
    ("Quarterly",
     "Field visits to the Kedarkantha trail and Sankri base camp by WWS senior staff to audit "
     "drop-off point utilisation, segregation quality, and guide/porter adherence. Review of "
     "photographic and documentary evidence."),
    ("Mid-Project\n(Month 4)",
     "Formal mid-project review: assessment of targets vs. actuals (waste diverted, agencies "
     "onboarded, trekkers reached). Identification of gaps and adjustment of operational plan."),
    ("End of Project\n(Month 8)",
     "Comprehensive project completion assessment: third-party verification of total waste diverted, "
     "MRF processing records, agency participation registers, trekker engagement data, and "
     "submission of final report to NBCFDC with full financial utilisation statement."),
    ("Post-Project\n(6 months after)",
     "Follow-up monitoring to assess sustainability of the system: whether trekking agencies continue "
     "waste practices independently, whether drop-off infrastructure is maintained, and early evidence "
     "of replication on other routes."),
]

tbl_mon = doc.add_table(rows=1+len(mon_data), cols=2)
tbl_mon.style = 'Table Grid'
for j, hdr in enumerate(["Monitoring Frequency", "Activities & Deliverables"]):
    cell = tbl_mon.rows[0].cells[j]
    shade_cell(cell, '1A5C3A')
    r = cell.paragraphs[0].add_run(hdr)
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
for i, (freq, act) in enumerate(mon_data):
    row = tbl_mon.rows[i+1]
    if i % 2 == 0:
        for c in row.cells: shade_cell(c, 'D4EDDA')
    for j, txt in enumerate([freq, act]):
        r = row.cells[j].paragraphs[0].add_run(txt)
        r.font.size = Pt(10); r.font.color.rgb = DARK_GREY
        if j == 0: r.bold = True

doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  12. SUSTAINABILITY COMMITTEE
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 12, "Sustainability Committee — Post-Handover")

body(doc,
     "To ensure the long-term sustainability of the waste management system beyond the CSR project "
     "period, a Trek Waste Management Committee will be constituted, composed primarily of community "
     "representatives and institutional stakeholders who will eventually take ownership of the system.")

committee = [
    ("Village-Level Representative\n(Gram Panchayat, Sankri)",
     "Elected local leader who will liaise with trekking agencies, oversee drop-off point maintenance, "
     "and coordinate with district waste management authorities."),
    ("Forest Department Representative\n(Govind Wildlife Sanctuary)",
     "Ensures ongoing compliance with environmental norms and facilitates access to sanctuary areas "
     "for waste collection activities."),
    ("Trekking Agency Consortium Lead\n(rotating, from 45+ partners)",
     "Represents the trekking industry; drives voluntary compliance, peer accountability, and "
     "collective adoption of responsible waste standards."),
    ("WWS Field Coordinator\n(transition support for 6 months)",
     "Provides technical handholding to the committee during transition, ensuring knowledge transfer "
     "and system continuity."),
    ("Local Youth / Women's Group Representative",
     "Representatives of community groups employed in waste collection roles, ensuring livelihood "
     "continuity and local ownership of the programme."),
]

for comp_name, comp_role in committee:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_after = Pt(4)
    rb = p.add_run(f"◆  {comp_name}: "); rb.bold = True
    rb.font.size = Pt(10.5); rb.font.color.rgb = MID_GREEN
    rd = p.add_run(comp_role); rd.font.size = Pt(10.5); rd.font.color.rgb = DARK_GREY

doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  13. PROJECT COORDINATOR
# ══════════════════════════════════════════════════════════════════════════════
section_heading(doc, 13, "Coordinator for This Project")

body(doc, "The following officer of Waste Warriors Society is the designated Project Coordinator, "
     "authorised to liaise with NBCFDC on all matters pertaining to this project:")

coord_tbl = doc.add_table(rows=6, cols=2)
coord_tbl.style = 'Table Grid'
coord_rows = [
    ("Name",           "Mayank Sharma"),
    ("Designation",    "Chief Executive Officer"),
    ("Organisation",   "Waste Warriors Society"),
    ("Office Address", "136/2/2, Shivam Vihar, Jakhan, Rajpur Road, Dehradun, Uttarakhand – 248001"),
    ("Phone / Mobile", "+91-7505763049"),
    ("Email ID",       "mayank.sharma@wastewarriors.org"),
]
for i, (label, value) in enumerate(coord_rows):
    row = coord_tbl.rows[i]
    shade_cell(row.cells[0], 'D4EDDA')
    rl = row.cells[0].paragraphs[0].add_run(label)
    rl.bold = True; rl.font.size = Pt(10); rl.font.color.rgb = DARK_GREEN
    rv = row.cells[1].paragraphs[0].add_run(value)
    rv.font.size = Pt(10); rv.font.color.rgb = DARK_GREY

doc.add_paragraph()

# Secondary contact
body(doc, "Secondary Contact (Partnerships):")
coord_tbl2 = doc.add_table(rows=3, cols=2)
coord_tbl2.style = 'Table Grid'
coord_rows2 = [
    ("Name",        "Lehar Arora"),
    ("Designation", "Manager — Partnerships"),
    ("Email ID",    "lehar.arora@wastewarriors.org"),
]
for i, (label, value) in enumerate(coord_rows2):
    row = coord_tbl2.rows[i]
    shade_cell(row.cells[0], 'D4EDDA')
    rl = row.cells[0].paragraphs[0].add_run(label)
    rl.bold = True; rl.font.size = Pt(10); rl.font.color.rgb = DARK_GREEN
    rv = row.cells[1].paragraphs[0].add_run(value)
    rv.font.size = Pt(10); rv.font.color.rgb = DARK_GREY

doc.add_paragraph()
doc.add_paragraph()
hrule(doc)

# ══════════════════════════════════════════════════════════════════════════════
#  PART II  — ANNEXURES
# ══════════════════════════════════════════════════════════════════════════════
p2_para = doc.add_paragraph()
p2_para.paragraph_format.space_before = Pt(10)
p2_para.paragraph_format.space_after  = Pt(4)
p2_run = p2_para.add_run("Part – II:    LIST OF ANNEXURES TO PROJECT PROPOSAL")
p2_run.bold           = True
p2_run.underline      = True
p2_run.font.size      = Pt(12)
p2_run.font.color.rgb = DARK_GREEN

ann_items = [
    ("Undertaking by the Implementing Agency",
     "Waste Warriors Society will execute an undertaking on its official letterhead stating that "
     "it will abide by all terms and conditions of the CSR grant; will not divert funds or entrust "
     "execution to other organisations (unless agreed by NBCFDC in writing); and accepts full "
     "liability for refund of entire/unutilised grant amounts in the event of any breach."),
    ("Copy of Registration Certificate / Trust Deed / MOA",
     "WWS is registered under the Societies Registration Act 1860 — Registration Number 243/2012-2013. "
     "Copy enclosed."),
    ("Annual Reports with Audited Accounts (Last 3 Years)",
     "Audited financial statements for the past three years with auditor's reports — to be enclosed."),
    ("PAN Card and TAN Number",
     "Copy of PAN Card and TAN Number — to be enclosed."),
    ("List of Board of Directors / Executive Committee Members",
     "List with addresses and contact numbers — to be enclosed."),
    ("List of Major Donors / Partners",
     "Parle Biscuits Pvt. Ltd., Airbnb, MacArthur Foundation, HDFC Bank Ltd, HT Parekh Foundation, "
     "Give India Foundation, Make My Trip Foundation, Rain Matter Foundation, Edelgive Foundation, "
     "Rohini Nilekani Philanthropies, Godrej Consumer Products Ltd, SBI Foundation, and others."),
    ("Due Diligence Summary",
     "Refer to the WWS Due Diligence table reproduced below."),
    ("Any Other Document",
     "As may be required by NBCFDC."),
]

for i, (ann_title, ann_desc) in enumerate(ann_items, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_after = Pt(4)
    rb = p.add_run(f"{i}.  {ann_title}\n"); rb.bold = True
    rb.font.size = Pt(10.5); rb.font.color.rgb = MID_GREEN
    rd = p.add_run(f"     {ann_desc}"); rd.font.size = Pt(10); rd.font.color.rgb = DARK_GREY

doc.add_paragraph()

# Due Diligence table
sub_heading(doc, "WWS Due Diligence Summary")

dd_data = [
    ("1", "NGO Registration",                      "Yes", "17-11-2012", "17-11-2027"),
    ("2", "Permanent Account Number (PAN)",         "Yes", "17-11-2012", "Perpetual"),
    ("3", "Certificate under 12A",                 "Yes", "02-03-2026", "31-03-2036"),
    ("4", "Certificate under 80G",                 "Yes", "02-03-2026", "31-03-2031"),
    ("5", "CSR-1 Registration",                    "Yes", "19-04-2021", "Perpetual"),
    ("6", "FCRA Renewal Application",              "Yes", "15-11-2017", "01-04-2028"),
    ("7", "Employee Provident Fund (EPF)",          "Yes", "27-11-2015", "Perpetual"),
    ("8", "Audited Balance Sheets (last 3 years)",  "Yes", "—",          "—"),
]

tbl_dd = doc.add_table(rows=1+len(dd_data), cols=5)
tbl_dd.style = 'Table Grid'
for j, hdr in enumerate(["S.No.", "Particulars", "Applicable", "Registration Date", "Valid Till"]):
    cell = tbl_dd.rows[0].cells[j]
    shade_cell(cell, '1A5C3A')
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(hdr)
    r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
for i, row_data in enumerate(dd_data):
    row = tbl_dd.rows[i+1]
    if i % 2 == 0:
        for c in row.cells: shade_cell(c, 'EAF7EE')
    for j, txt in enumerate(row_data):
        p = row.cells[j].paragraphs[0]
        if j in (0, 2): p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        r.font.size = Pt(9); r.font.color.rgb = DARK_GREY

doc.add_paragraph()
doc.add_paragraph()
hrule(doc)

# ── Signature Block ────────────────────────────────────────────────────────────
sig = doc.add_paragraph()
sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
sig.paragraph_format.space_before = Pt(16)
for line in [
    "Name & Signature of Authorised Signatory",
    "Mayank Sharma — Chief Executive Officer",
    "Waste Warriors Society",
    "136/2/2, Shivam Vihar, Jakhan, Rajpur Road,",
    "Dehradun, Uttarakhand – 248001",
    "Phone: +91-7505763049",
    "Email: partnerships@wastewarriors.org",
    "Website: www.wastewarriors.org",
]:
    r = sig.add_run(line + "\n")
    r.font.size = Pt(10)
    if "Mayank" in line or "Waste Warriors" in line:
        r.bold = True; r.font.color.rgb = DARK_GREEN
    else:
        r.font.color.rgb = DARK_GREY

# ── Footer note ───────────────────────────────────────────────────────────────
doc.add_paragraph()
note = doc.add_paragraph()
note.alignment = WD_ALIGN_PARAGRAPH.CENTER
nr = note.add_run(
    "All documents (each page) shall be duly signed and stamped by the applicant and the "
    "Head of the Grantee Organisation. The Institution shall produce original documents on "
    "demand by NBCFDC. The Institution will submit Monthly Progress Reports (MPR) as per "
    "the Appendix format enclosed with this proposal.")
nr.italic = True; nr.font.size = Pt(9); nr.font.color.rgb = DARK_GREY

# ══════════════════════════════════════════════════════════════════════════════
#  SAVE
# ══════════════════════════════════════════════════════════════════════════════
out_path = "/home/user/jcode/The_Good_Trek_Project_CSR_Proposal.docx"
doc.save(out_path)
print(f"Saved: {out_path}")
