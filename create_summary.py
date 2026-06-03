from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

doc = Document()

# ─── Page margins ───────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ─── Helper functions ────────────────────────────────────────────────────────
def set_font(run, name="Calibri", size=11, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    if level == 1:
        set_font(run, size=16, bold=True, color=(0, 51, 102))
    elif level == 2:
        set_font(run, size=13, bold=True, color=(0, 84, 166))
    elif level == 3:
        set_font(run, size=11, bold=True, color=(40, 40, 40))
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(4)
    return p

def add_body(doc, text, indent=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, size=11)
    p.paragraph_format.space_after = Pt(4)
    if indent:
        p.paragraph_format.left_indent = Cm(0.8)
    return p

def add_bullet(doc, text, sub=False):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    set_font(run, size=11)
    if sub:
        p.paragraph_format.left_indent = Cm(1.5)
    return p

def add_key_value(doc, key, value):
    p = doc.add_paragraph()
    r1 = p.add_run(f"{key}: ")
    set_font(r1, size=11, bold=True, color=(0, 51, 102))
    r2 = p.add_run(value)
    set_font(r2, size=11)
    p.paragraph_format.space_after = Pt(3)
    return p

def add_divider(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '4472C4')
    pBdr.append(bottom)
    pPr.append(pBdr)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(8)

def add_callout(doc, text, bg_label="NOTE"):
    p = doc.add_paragraph()
    r1 = p.add_run(f"  {bg_label}:  ")
    set_font(r1, size=10, bold=True, color=(255, 255, 255))
    r1.font.highlight_color = None
    r2 = p.add_run(f" {text}")
    set_font(r2, size=10, color=(60, 60, 60))
    p.paragraph_format.left_indent  = Cm(0.5)
    p.paragraph_format.space_after  = Pt(6)
    shading = OxmlElement('w:shd')
    shading.set(qn('w:val'), 'clear')
    shading.set(qn('w:color'), 'auto')
    shading.set(qn('w:fill'), 'D6E4F0')
    p._p.get_or_add_pPr().append(shading)
    return p

# ════════════════════════════════════════════════════════════════════════════
# COVER / TITLE
# ════════════════════════════════════════════════════════════════════════════
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
tr = title_p.add_run("GOVERNMENT GAZETTE NOTIFICATIONS")
set_font(tr, size=20, bold=True, color=(0, 51, 102))

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sr = sub_p.add_run("Plain-Language Summary — Three Official Notifications (May 2026)")
set_font(sr, size=12, color=(80, 80, 80))

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
dr = date_p.add_run(f"Prepared: {datetime.date.today().strftime('%d %B %Y')}")
set_font(dr, size=10, color=(120, 120, 120))

doc.add_paragraph()
add_divider(doc)

# Quick-reference table of contents
add_heading(doc, "DOCUMENTS COVERED", 2)
toc = [
    ("Document 1", "EPF Wage Ceiling Notification",
     "S.O. 2702(E) | Ministry of Labour & Employment | 29 May 2026"),
    ("Document 2", "CSR – ZCZP Schedule VII Amendment",
     "G.S.R. 416(E) | Ministry of Corporate Affairs | 27 May 2026"),
    ("Document 3", "CSR – ZCZP Policy Amendment Rules",
     "G.S.R. 415(E) | Ministry of Corporate Affairs | 27 May 2026"),
]
for num, title, ref in toc:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    r1 = p.add_run(f"{num}  |  ")
    set_font(r1, size=11, bold=True, color=(0, 84, 166))
    r2 = p.add_run(f"{title}  ")
    set_font(r2, size=11, bold=True)
    r3 = p.add_run(f"({ref})")
    set_font(r3, size=10, color=(100, 100, 100))
    p.paragraph_format.space_after = Pt(4)

add_divider(doc)

# ════════════════════════════════════════════════════════════════════════════
# DOCUMENT 1 — EPF WAGE CEILING
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "DOCUMENT 1 — EPF Wage Ceiling Notification", 1)
add_heading(doc, "At a Glance", 2)
add_key_value(doc, "Notification No.", "S.O. 2702(E)")
add_key_value(doc, "Issuing Authority", "Ministry of Labour and Employment, Government of India")
add_key_value(doc, "Date", "29 May 2026")
add_key_value(doc, "Legal Basis", "Code on Social Security, 2020 — Section 2, Clause (89)")
add_key_value(doc, "Signed by", "Tejaswi S. Naik, Joint Secretary")
doc.add_paragraph()

add_heading(doc, "What Does It Say — In Plain English?", 2)
add_body(doc,
    "The Central Government has officially fixed ₹15,000 per month as the wage ceiling "
    "for the purpose of Employees' Provident Fund (EPF) contributions under Chapter III of "
    "the Code on Social Security, 2020."
)

add_heading(doc, "What Is a 'Wage Ceiling'?", 3)
add_body(doc,
    "A wage ceiling is the maximum salary limit used to calculate mandatory EPF contributions. "
    "If an employee earns ₹15,000 or less per month, their EPF contribution is calculated on "
    "their full salary. If they earn more than ₹15,000, the employer's mandatory EPF contribution "
    "is still calculated based on ₹15,000 — even if the employee voluntarily contributes on a "
    "higher amount."
)

add_heading(doc, "Why Does This Matter?", 3)
add_bullet(doc, "Every employer who has employees earning up to ₹15,000/month MUST enroll them in EPF.")
add_bullet(doc, "Minimum EPF contribution = 12% of ₹15,000 = ₹1,800/month from the employee.")
add_bullet(doc, "The employer also contributes 12% of ₹15,000, split between EPF and EPS (Employee Pension Scheme).")
add_bullet(doc, "This is the threshold that determines who is covered under the mandatory EPF scheme.")

add_heading(doc, "Who Is Affected?", 3)
add_bullet(doc, "All employers with 20 or more employees.")
add_bullet(doc, "Employees drawing a basic wage up to ₹15,000/month are automatically covered.")
add_bullet(doc, "Employees earning above ₹15,000 can still join voluntarily, but the employer's statutory liability is capped at the ₹15,000 ceiling.")

add_callout(doc,
    "This notification retains the existing ₹15,000 wage ceiling that has been in place "
    "since September 2014, now formally renotified under the new Code on Social Security, 2020 "
    "framework — replacing the older EPF & MP Act, 1952 reference.",
    "KEY POINT"
)

add_heading(doc, "Critical Assessment", 3)
add_body(doc,
    "The wage ceiling of ₹15,000/month has not been revised upward despite significant inflation "
    "since 2014. Given that minimum wages in many states now exceed ₹15,000/month, this ceiling "
    "has become less effective at expanding EPF coverage. Critics argue it should be revised to "
    "₹21,000–₹25,000 to bring more workers under social security protection. This notification, "
    "however, simply migrates the existing ceiling into the new legal framework — it is a "
    "housekeeping measure, not a policy change."
)
add_divider(doc)

# ════════════════════════════════════════════════════════════════════════════
# DOCUMENT 2 — CSR ZCZP SCHEDULE VII AMENDMENT
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "DOCUMENT 2 — CSR: Addition of ZCZP to Schedule VII (Companies Act)", 1)
add_heading(doc, "At a Glance", 2)
add_key_value(doc, "Notification No.", "G.S.R. 416(E)")
add_key_value(doc, "Issuing Authority", "Ministry of Corporate Affairs, Government of India")
add_key_value(doc, "Date", "27 May 2026")
add_key_value(doc, "Legal Basis", "Companies Act, 2013 — Section 467(1)")
add_key_value(doc, "Signed by", "Rahul Jain, Joint Secretary")
doc.add_paragraph()

add_heading(doc, "What Does It Say — In Plain English?", 2)
add_body(doc,
    "Schedule VII of the Companies Act lists the activities that qualify as valid Corporate Social "
    "Responsibility (CSR) spending. This notification adds a brand-new item (xiii) to that list:"
)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.left_indent  = Cm(1)
p.paragraph_format.right_indent = Cm(1)
pr = p.add_run(
    '"Subscription to zero coupon zero principal instruments on Social Stock Exchange."'
)
set_font(pr, size=11, bold=True, color=(0, 84, 166))
shading = OxmlElement('w:shd')
shading.set(qn('w:val'), 'clear')
shading.set(qn('w:color'), 'auto')
shading.set(qn('w:fill'), 'EBF5FB')
p._p.get_or_add_pPr().append(shading)

add_heading(doc, "Background: What is Schedule VII?", 3)
add_body(doc,
    "Under Indian law, certain companies (profitable ones meeting size thresholds) must spend at "
    "least 2% of their average net profit on CSR activities. Schedule VII lists the approved "
    "categories — health, education, environment, poverty reduction, etc. "
    "Only spending on these listed activities counts as valid CSR."
)

add_heading(doc, "What is a Zero Coupon Zero Principal (ZCZP) Instrument?", 3)
add_body(doc,
    "A ZCZP instrument is a unique financial security issued by a Non-Profit Organisation (NPO) "
    "listed on India's Social Stock Exchange (SSE). Here's what makes it unusual:"
)
add_bullet(doc, "Zero Coupon: The investor receives NO interest / returns.")
add_bullet(doc, "Zero Principal: The investor gets NO money back at maturity.")
add_bullet(doc, "It is essentially a donation packaged as a financial instrument, giving it SEBI-regulated structure and transparency.")
add_bullet(doc, "The money raised is used entirely for social impact projects.")

add_heading(doc, "What Does This Change Mean Practically?", 3)
add_body(doc,
    "Previously, companies investing in ZCZP instruments were in a grey area regarding CSR credit. "
    "Now, buying ZCZP instruments is officially recognised as a valid CSR activity. "
    "Companies can count this spending toward their mandatory 2% CSR obligation."
)

add_callout(doc,
    "This amendment effectively opens a formal funding channel between corporate India's CSR "
    "budgets and social-sector non-profits operating on the Social Stock Exchange.",
    "SIGNIFICANCE"
)

add_heading(doc, "Critical Assessment", 3)
add_body(doc,
    "This is a positive, enabling amendment. It creates a regulated, transparent pathway for CSR "
    "funds to flow into social enterprises. However, the Social Stock Exchange is still nascent "
    "in India — very few NPOs are listed, and market liquidity is minimal. The real-world impact "
    "depends heavily on how quickly the SSE ecosystem matures. There is also a risk of 'impact "
    "washing' if companies treat ZCZP investments as an easy CSR checkbox without verifying actual "
    "social outcomes."
)
add_divider(doc)

# ════════════════════════════════════════════════════════════════════════════
# DOCUMENT 3 — CSR ZCZP POLICY AMENDMENT RULES
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "DOCUMENT 3 — CSR Policy Amendment Rules: ZCZP Implementation Framework", 1)
add_heading(doc, "At a Glance", 2)
add_key_value(doc, "Notification No.", "G.S.R. 415(E)")
add_key_value(doc, "Issuing Authority", "Ministry of Corporate Affairs, Government of India")
add_key_value(doc, "Date", "27 May 2026")
add_key_value(doc, "Legal Basis", "Companies Act, 2013 — Section 135 and Section 469(1)(2)")
add_key_value(doc, "Full Name", "Companies (Corporate Social Responsibility Policy) Amendment Rules, 2026")
add_key_value(doc, "Signed by", "Rahul Jain, Joint Secretary")
doc.add_paragraph()

add_heading(doc, "What Does It Say — In Plain English?", 2)
add_body(doc,
    "While Document 2 added ZCZP to the list of valid CSR activities, this document provides "
    "the detailed rules — the 'how-to guide' — for companies and NPOs wanting to use ZCZP "
    "instruments for CSR. It amends the CSR Policy Rules, 2014."
)

add_heading(doc, "New Definitions Added", 3)
add_body(doc, "Two key terms are now formally defined in law:")
add_bullet(doc,
    "Not for Profit Organization (NPO): An entity with the same meaning as defined in "
    "SEBI's ICDR Regulations 2018, Regulation 292A(e). In simple terms: a registered "
    "organisation that reinvests all surplus into its social mission."
)
add_bullet(doc,
    "Zero Coupon Zero Principal (ZCZP) Instrument: A SEBI-declared security issued by an NPO "
    "listed on the Social Stock Exchange segment of a recognised stock exchange — offering "
    "no coupon (interest) and no return of principal."
)

add_heading(doc, "New Rule 4A — The ZCZP Implementation Framework", 3)
add_body(doc, "This is the heart of the notification. Rule 4A introduces a complete framework:")

add_heading(doc, "For Companies (Investors in ZCZP):", 3)
add_bullet(doc, "A company CAN use CSR funds to subscribe to (buy) ZCZP instruments.")
add_bullet(doc,
    "CAP: ZCZP spending cannot exceed 10% of the company's total CSR expenditure for that "
    "financial year. Example: If a company's total CSR budget is ₹1 crore, a maximum of "
    "₹10 lakh can go into ZCZP instruments."
)
add_bullet(doc,
    "EXEMPTION from Impact Assessment: Companies investing in ZCZP do NOT have to conduct "
    "an independent impact assessment of ZCZP-funded projects. (Normally required for projects "
    "above ₹1 crore.) This reduces compliance burden."
)

add_heading(doc, "For Non-Profit Organisations (Issuers of ZCZP):", 3)
add_bullet(doc,
    "Project Duration Limit: The NPO must complete the funded project within 3 financial years "
    "from the date of issuance of the ZCZP instrument. No open-ended commitments allowed."
)
add_bullet(doc,
    "Unspent Funds: If the ZCZP instrument is delisted (terminated) before the project is complete, "
    "any unspent funds MUST be transferred to an approved fund listed in Schedule VII of the "
    "Companies Act — the money cannot be retained or used for other purposes."
)
add_bullet(doc,
    "SEBI Compliance Report: The NPO must submit a compliance report to SEBI when the listing ends."
)

add_heading(doc, "Applicability of Existing CSR Rules:", 3)
add_body(doc,
    "Most of the existing Rule 4 (general CSR implementation rules) applies to ZCZP-based CSR — "
    "EXCEPT sub-rules (5) and (6), which deal with unspent CSR account mechanics that are "
    "replaced by the specific ZCZP unspent-fund transfer mechanism above."
)

add_heading(doc, "Quick Reference: Key Numbers", 3)

table = doc.add_table(rows=4, cols=3)
table.style = "Table Grid"
hdr = table.rows[0].cells
hdr[0].text = "Rule"
hdr[1].text = "What It Says"
hdr[2].text = "Limit / Requirement"
for cell in hdr:
    for para in cell.paragraphs:
        for run in para.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '003366')
        cell._tc.get_or_add_tcPr().append(shd)

rows_data = [
    ("ZCZP Spending Cap", "Max % of CSR budget for ZCZP instruments", "10% per financial year"),
    ("Project Duration", "Max project length for NPO using ZCZP funds", "3 financial years from issuance"),
    ("Unspent Funds", "What NPO must do with leftover money on delisting", "Transfer to Schedule VII fund + SEBI report"),
]
for i, (r1, r2, r3) in enumerate(rows_data, start=1):
    cells = table.rows[i].cells
    cells[0].text = r1
    cells[1].text = r2
    cells[2].text = r3

doc.add_paragraph()

add_callout(doc,
    "Documents 2 and 3 work together as a package: Document 2 gives legal recognition, "
    "Document 3 provides the operational rules. Both were notified within two days of each other "
    "and take effect from the date of publication in the Gazette.",
    "NOTE"
)

add_heading(doc, "Critical Assessment", 3)
add_body(doc,
    "The 10% cap on ZCZP spending is a prudent safeguard — it prevents companies from routing "
    "excessive CSR funds through instruments where accountability is lower. The 3-year project "
    "duration limit prevents NPOs from holding funds indefinitely. The SEBI compliance report "
    "requirement adds regulatory oversight. However, the exemption from impact assessment is "
    "a double-edged sword: it reduces company burden but also reduces accountability for "
    "outcomes. The framework is well-structured for a first iteration but will need monitoring "
    "to ensure genuine social impact is delivered."
)
add_divider(doc)

# ════════════════════════════════════════════════════════════════════════════
# OVERALL CONCLUSION
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "OVERALL CONCLUSION", 1)
add_body(doc,
    "These three notifications represent two distinct policy actions from two ministries, "
    "both published in the Gazette of India in late May 2026:"
)

add_bullet(doc,
    "EPF Wage Ceiling (Labour Ministry): A housekeeping measure that reaffirms the existing "
    "₹15,000/month threshold under the new Social Security Code. No practical change for "
    "employers or employees — the ceiling stays the same."
)
add_bullet(doc,
    "CSR + ZCZP Framework (Corporate Affairs Ministry): A significant policy innovation that "
    "allows companies to fund social-sector NPOs through SEBI-regulated instruments on the Social "
    "Stock Exchange — with clear caps, timelines, and oversight rules. This is a meaningful step "
    "toward institutionalising social-impact investing within India's mandatory CSR framework."
)

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
fr = p.add_run("— End of Summary Document —")
set_font(fr, size=10, color=(120, 120, 120))

# ─── Save ────────────────────────────────────────────────────────────────────
out_path = "/home/user/jcode/Gazette_Notifications_Summary_May2026.docx"
doc.save(out_path)
print(f"Saved: {out_path}")
