import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = openpyxl.Workbook()

# ── colour palette ──────────────────────────────────────────────
DARK_NAVY   = "1B2A4A"
MID_BLUE    = "2E5CA8"
LIGHT_BLUE  = "D6E4F7"
ACCENT_GOLD = "F4A523"
WHITE       = "FFFFFF"
LIGHT_GREY  = "F2F2F2"
MED_GREY    = "CCCCCC"
DARK_TEXT   = "1A1A2E"

def hfill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def hfont(hex_color=DARK_TEXT, bold=False, size=10, name="Calibri"):
    return Font(color=hex_color, bold=bold, size=size, name=name)

thin = Side(style="thin", color=MED_GREY)
thick = Side(style="medium", color=MID_BLUE)
border_all   = Border(left=thin, right=thin, top=thin, bottom=thin)
border_thick = Border(left=thick, right=thick, top=thick, bottom=thick)

def wrap_align(h="left", v="top"):
    return Alignment(horizontal=h, vertical=v, wrap_text=True)

# ════════════════════════════════════════════════════════════════
# SHEET 1 – MASTER OVERVIEW TABLE
# ════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Master Overview"
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

# Row 1 – main banner
ws.merge_cells("A1:M1")
ws["A1"] = "CORPORATE EMPLOYEE GIVING PROGRAMS — COMPREHENSIVE RESEARCH REPORT"
ws["A1"].font      = Font(name="Calibri", bold=True, size=14, color=WHITE)
ws["A1"].fill      = hfill(DARK_NAVY)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[1].height = 30

# Row 2 – column headers
headers = [
    "Company", "Program Name", "Match Ratio",
    "Annual Cap (USD)", "Min Donation", "Volunteer Grant",
    "Platform / Portal", "Eligible Orgs", "Submission Process",
    "Special Campaigns / Notes", "Eligibility (Employees)", "Key Stat / Impact",
    "Primary Source"
]
for col, h in enumerate(headers, 1):
    cell = ws.cell(row=2, column=col, value=h)
    cell.font      = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill      = hfill(MID_BLUE)
    cell.alignment = wrap_align("center", "center")
    cell.border    = border_all
ws.row_dimensions[2].height = 30

# ── DATA ────────────────────────────────────────────────────────
data = [
    # (Company, Program Name, Ratio, Cap, Min, VolGrant, Platform,
    #  EligibleOrgs, Process, SpecialCampaigns, Eligibility, KeyStat, Source)
    (
        "Airbnb",
        "Airbnb Community Fund + Employee Matching",
        "1:1",
        "$1,000 / employee / year",
        "$25",
        "Not publicly confirmed",
        "Internal portal (Benevity referenced in nonprofit databases)",
        "Registered 501(c)(3) nonprofits worldwide",
        "Employees donate via internal giving portal; match requested same portal; "
        "Community Fund ERGs guide grant nominations",
        "Community Fund: $8.5M+ distributed to 160+ nonprofits in 30+ countries "
        "(2024-2025). Grants awarded July 2024–early 2025.",
        "Full-time employees",
        "$25M+ total Community Fund grants to orgs in 60 countries",
        "https://news.airbnb.com/2025-community-fund/ | "
        "https://doublethedonation.com/matching-gifts/airbnb"
    ),
    (
        "Amazon",
        "Amazon Employee Giving Program",
        "1:1",
        "$5,000 / employee / year",
        "~$25",
        "Volunteer Time Off available; grant details internal",
        "Internal Amazon portal (YourCause / Benevity referenced externally)",
        "501(c)(3) nonprofits; special focus on homelessness orgs",
        "Employees submit requests via internal HR giving portal; "
        "payroll deduction and credit/debit card options available",
        "Amazon Global Month of Volunteering (June); "
        "Dedicated $5M match for 20 vetted homelessness nonprofits; "
        "AmazonSmile discontinued Jan 2023",
        "Full-time and part-time employees",
        "One of largest employer matching programs by headcount; "
        "Community Impact supports 50+ states & 100+ countries",
        "https://doublethedonation.com/matching-gifts/amazon | "
        "https://www.glassdoor.com/Benefits/Amazon-Charitable-Gift-Matching-US-BNFT21_E6036_N1.htm | "
        "https://www.amazon.jobs/content/en/teams/ccr/aci"
    ),
    (
        "NVIDIA",
        "NVIDIA Foundation — Inspire 365",
        "1:1",
        "$10,000 / employee / year",
        "No minimum specified",
        "Included in $10,000 combined cap (time + money)",
        "Benevity (Inspire 365 platform; support@benevity.com)",
        "Qualified 501(c)(3) charities globally",
        "1) Employee donates externally (cash, card, check, stock, DAF). "
        "2) Logs gift in Inspire 365 Benevity portal — no enrollment required. "
        "3) NVIDIA Foundation issues match. "
        "Funds do NOT roll over to next fiscal year.",
        "Year-end Giving Campaign; FY2024: employees raised $5.5M; "
        "Inspire 365 covers giving + volunteering + community action",
        "All NVIDIA employees globally; no enrollment required",
        "FY2024: 40%+ employee participation; $16M+ donated via Inspire 365",
        "https://www.nvidia.com/en-us/foundation/ | "
        "https://blogs.nvidia.com/blog/nvidia-life-inspire-365-2024/ | "
        "https://files.doublethedonation.com/forms/nvidia-guidelines.pdf"
    ),
    (
        "Google",
        "Google Employee Giving (go/give) — powered by Benevity",
        "1:1",
        "$12,000 / employee / year",
        "$50 minimum",
        "$10 / volunteer hour",
        "Benevity — accessed via internal shortlink go/give "
        "(google.benevity.org for nonprofits)",
        "200,000+ eligible organizations worldwide",
        "1) Employee donates or logs volunteer hours. "
        "2) Submits via go/give Benevity portal. "
        "Deadline: Dec 31 each year (11:59 PM PT). "
        "Payroll deduction and card payment options.",
        "15+ years of giving; year-end giving credit issued to all employees; "
        "Special disaster-relief matching campaigns activated as needed; "
        "Google.org employee fellowship program (pro bono)",
        "Full-time and part-time employees (retired employees NOT eligible)",
        "Among largest corporate giving programs globally by match cap",
        "https://google.benevity.org/terms-and-conditions | "
        "https://doublethedonation.com/forms/google-guidelines.pdf | "
        "https://doublethedonation.com/matching-gifts/google-inc"
    ),
    (
        "HP",
        "HP Foundation — Employee Cash Match + Product Giving Program",
        "1:1 cash; 3:1 product (HP covers 75%, employee 25%)",
        "Cash: $5,000 / employee / year (while funds available); "
        "Product: up to retail value limits",
        "$25",
        "$50 per 10 volunteer hours per quarter ($5/hour equivalent)",
        "HP Foundation portal (hp.com/hpinfo/socialinnovation/us/)",
        "501(c)(3) charities; schools; educational institutions",
        "Cash Match: Employee donates → submits via HP Foundation portal "
        "for 1:1 match while annual fund pool lasts. "
        "Product Match: Employee contributes 25% of tech product cost; "
        "HP Foundation covers 75% via Employee Product Giving Program.",
        "HP LIFE digital skills program; "
        "HP Sustainable Impact partnerships; "
        "Product Giving Program (unique in industry)",
        "Active employees (full-time and part-time); "
        "some benefits extend to retirees",
        "HP Foundation has enabled product donations worth millions in tech assets to nonprofits",
        "https://www.hp.com/us-en/hp-information/hp-foundation.html | "
        "https://doublethedonation.com/matching-gifts/hp-inc | "
        "http://www.hp.com/hpinfo/socialinnovation/us/product_matching.html"
    ),
    (
        "IBM",
        "IBM Matching Grants Program",
        "1:1",
        "$10,000 / employee / year",
        "$25 minimum",
        "IBM Volunteer Grants (separate program; hours → grant to org)",
        "Benevity — ibm.benevity.org",
        "501(c)(3) nonprofits; accredited educational institutions; "
        "hospitals; social welfare organizations",
        "1) Employee donates to eligible org. "
        "2) Submits match request via ibm.benevity.org "
        "within 1 YEAR of donation date. "
        "3) IBM Foundation issues match to nonprofit.",
        "IBM Centennial of Service (100M+ volunteer hours legacy); "
        "IBM Skills Academy for nonprofits; "
        "On Demand Community volunteering portal",
        "Active US full-time and part-time employees; "
        "retirees also eligible",
        "IBM Foundation has contributed hundreds of millions since founding; "
        "volunteers log 3M+ hours annually",
        "https://www.ibm.com/responsibility/programs/volunteerism-giving | "
        "https://doublethedonation.com/matching-gifts/ibm-international-business-machines | "
        "http://www.crhsraiderband.org/Content/3_31/Files/IBM.pdf"
    ),
    (
        "American Express",
        "Give2Gether (powered by GlobalGiving)",
        "1:1",
        "$10,000 / employee / year (via American Express Foundation)",
        "Not publicly specified",
        "Volunteer grant funding available (amount not publicly confirmed)",
        "GlobalGiving platform — globalgiving.org/amex-give2gether/",
        "Vetted nonprofits in 170+ countries via GlobalGiving; "
        "all cause areas eligible",
        "1) Employee accesses Give2Gether portal at globalgiving.org/amex-give2gether/. "
        "2) Selects from vetted global nonprofits. "
        "3) Donates via card or payroll deduction. "
        "4) American Express Foundation matches automatically. "
        "Available in US, Canada, India; 30+ global locations via GlobalGiving.",
        "Global Match Program; "
        "Annual Volunteer Day / AmexCares volunteer events; "
        "Disaster relief campaigns; "
        "2024: 6,000+ colleagues volunteered",
        "Eligible colleagues in US, Canada, India (core); "
        "30+ country expansion via GlobalGiving",
        "Supports 170+ country nonprofits; $10,000 cap is among highest in financial sector",
        "https://www.globalgiving.org/amex-give2gether/ | "
        "https://about.americanexpress.com/we-engage-our-colleagues | "
        "https://www.businesswire.com/news/home/20201013005631/en/American-Express-Partners-With-GlobalGiving"
    ),
    (
        "Apple",
        "Apple Matching Gifts Program",
        "1:1 (standard); up to 2:1 in special campaigns",
        "$10,000 / employee / year",
        "Not publicly specified",
        "$25 / volunteer hour",
        "Benevity (administered by Benevity on behalf of Apple)",
        "501(c)(3) nonprofits; international equivalents eligible",
        "1) Employee donates to eligible org. "
        "2) Submits match request via Benevity portal "
        "within 1 MONTH of donation. "
        "3) Apple matches and Benevity disburses funds. "
        "Cash, check, and payroll deduction accepted.",
        "Disaster relief enhanced matching campaigns; "
        "Apple has run elevated 2:1 matches for specific causes; "
        "Racial equity and justice-focused giving periods",
        "Full-time and part-time employees",
        "One of most generous volunteer grant rates in tech ($25/hour)",
        "https://doublethedonation.com/matching-gifts/apple-inc | "
        "https://forms.matchinggifts.com/AppleGuide.pdf | "
        "https://assets.globalgiving.org/docs/apple/apple-faqs.pdf"
    ),
    (
        "Texas Instruments",
        "TI Foundation Matching Gift Programs "
        "(Educational Matching Gift Program + Community Giving)",
        "1:1",
        "Community Giving: up to $30,000 / employee / year; "
        "Educational Program: up to $10,000 / employee / year",
        "$25 minimum",
        "$1,000 / employee / year in volunteer hour matching",
        "TI Foundation portal (forms.matchinggifts.com/TIEDUGuidelines.pdf)",
        "Educational: public & private accredited educational institutions. "
        "Community: 501(c)(3) nonprofits; public agencies; "
        "government instrumentalities.",
        "1) Employee donates to eligible org. "
        "2) Submits match form via TI Foundation portal. "
        "3) Foundation verifies and issues match. "
        "Volunteer match: log hours → grant to org (up to $1,000/year).",
        "Educational Matching Gift Program established in 1970 — "
        "one of oldest corporate matching programs in the US; "
        "STEM education focus; TI STEM initiatives in Dallas-Ft Worth",
        "Active employees and retirees; directors also eligible",
        "TI Foundation has supported STEM education since 1970; "
        "$30,000 cap is one of highest in the industry",
        "https://www.ti.com/about-ti/citizenship-community/giving.html | "
        "https://www.ti.com/about-ti/citizenship-community/giving/eligible-organizations.html | "
        "https://tialumni.org/volunteer-2/ti-foundation-and-matching-gifts/"
    ),
    (
        "Salesforce",
        "Salesforce 1-1-1 Model / Philanthropy Cloud (Pledge 1%)",
        "1:1",
        "$10,000 / employee / year",
        "Not publicly specified",
        "56 hours Volunteer Time Off (VTO) per employee per year; "
        "Volunteer grants also offered",
        "Salesforce Philanthropy Cloud (proprietary internal platform); "
        "Salesforce.org tools",
        "501(c)(3) nonprofits globally; "
        "wide range of causes including education, equity, environment",
        "1) Employee selects cause / nonprofit in Philanthropy Cloud. "
        "2) Donates or logs VTO hours. "
        "3) Salesforce matches donations and amplifies volunteer impact. "
        "Integrated with Salesforce CRM tools for nonprofits.",
        "Pledge 1% founding member (1% equity, 1% product, 1% time); "
        "Giving Tuesday campaigns; "
        "19,000+ Pledge 1% companies globally inspired by Salesforce model; "
        "Equality groups for targeted giving",
        "All Salesforce employees globally; "
        "56 VTO hours is among highest in tech industry",
        "FY2025: 50% of employees volunteered (vs 15% industry average); "
        "$3B+ generated in philanthropy via Pledge 1% movement",
        "https://www.salesforce.com/company/philanthropy/employee-volunteering-giving/ | "
        "https://www.salesforce.com/news/stories/employee-giving-model/ | "
        "https://www.salesforce.com/company/pledge/"
    ),
    (
        "Mastercard",
        "Mastercard Matching Gifts Program + Mastercard Impact Fund",
        "1:1",
        "Employees: $5,000 / year; "
        "Directors / Board: $15,000 / year",
        "Not publicly specified",
        "Up to 5 paid volunteer days/year; "
        "volunteer rewards redeemable up to $10,000 toward nonprofits",
        "Internal Mastercard portal "
        "(forms.matchinggifts.com/MasterCard.pdf for guidelines)",
        "Registered nonprofits and charities globally",
        "1) Employee donates to eligible nonprofit. "
        "2) Submits match request via Mastercard internal portal. "
        "3) Mastercard Foundation issues 1:1 match. "
        "Volunteer rewards: log hours → redeem rewards as nonprofit donations.",
        "Mastercard Impact Fund (strategic corporate philanthropy); "
        "Financial inclusion programs; "
        "Center for Inclusive Growth partnerships; "
        "Global Giving Tuesday campaigns",
        "Active employees (worldwide); "
        "retirees and board directors also eligible",
        "Mastercard Impact Fund deploys hundreds of millions in inclusion programs globally",
        "https://www.mastercard.com/us/en/for-the-world/people/community-impact.html | "
        "https://doublethedonation.com/matching-gifts/mastercard-worldwide | "
        "https://www.mastercardcenter.org/about-the-center/mastercard-impact-fund"
    ),
    (
        "Intel",
        "Intel Foundation Matching Program (Intel Involved)",
        "1:1",
        "$1,000 / employee / year COMBINED "
        "(donations + volunteer hour matches; revised 2025)",
        "Not publicly specified",
        "$10 / volunteer hour "
        "(counts toward $1,000 combined annual cap)",
        "Benevity — reopened Jan 15, 2025; "
        "intel.benevity.org (referenced)",
        "501(c)(3) nonprofits; accredited educational institutions; "
        "public agencies",
        "1) Employee donates or volunteers. "
        "2) Submits within 60-day window via Benevity portal. "
        "3) Intel Foundation matches (payouts now monthly, not annually). "
        "NOTE: 2025 changes significantly reduced cap from prior $10,000; "
        "reassessment planned for 2027.",
        "Intel Involved volunteer program; "
        "Employee Resource Groups (ERGs) giving campaigns; "
        "Rising Up Grants for education & STEM; "
        "2025: major program restructuring due to company financial pressures",
        "Active employees and retirees eligible",
        "Pre-2025: up to $10,000 cap; "
        "2025 cap reduced to $1,000 combined due to Intel financial restructuring",
        "https://www.intel.com/content/www/us/en/corporate-responsibility/intel-foundation-employee-generosity.html | "
        "https://intelretiree.com/2025/01/17/benevity-donation-volunteer-matching-site-opens/ | "
        "https://www.gomotionapp.com/team/osthsc/page/news/538907/updates-to-intels-donation-and-volunteer-matching-program-2025"
    ),
    (
        "NetApp",
        "NetApp Cares — Matching Gifts Program",
        "1:1 (doubles donations to eligible nonprofits)",
        "Not publicly confirmed (industry est. $2,500–$5,000/year)",
        "Not publicly specified",
        "Volunteer grants tied to individual and team volunteering; "
        "board service grants also available",
        "Benevity (referenced in nonprofit databases); "
        "internal NetApp Cares portal",
        "Eligible nonprofits globally; "
        "focus on underserved communities and education",
        "1) Employee donates to eligible nonprofit. "
        "2) Submits via NetApp Cares portal (Benevity-powered). "
        "3) NetApp matches donation. "
        "Volunteer grants: log hours → grant disbursed.",
        "Global Giving Challenge (December annual campaign); "
        "NetApp Serves volunteer program; "
        "June Volunteer Month events; "
        "40 hours paid Volunteer Time Off/year",
        "All NetApp employees globally; "
        "40 hours VTO is among more generous policies",
        "2024: significant employee volunteer engagement globally across NetApp Serves events",
        "https://www.netapp.com/responsibility/social-impact/netapp-cares/ | "
        "https://www.netapp.com/blog/netapp-serves-employee-volunteer-program/ | "
        "https://doublethedonation.com/matching-gifts/netapp"
    ),
    (
        "Apollo Global Management",
        "Apollo Citizenship Grants + Apollo Opportunity Foundation",
        "Not publicly confirmed "
        "(Citizenship Grants for matching gifts & volunteer rewards)",
        "Not publicly confirmed",
        "Not publicly specified",
        "Volunteer rewards/grants available (amount not confirmed)",
        "Internal Apollo portal; "
        "doublethedonation.com/matching-gifts/apollo-global-management-inc",
        "Nonprofits focused on career education, workforce development, "
        "and economic empowerment (primary focus); "
        "general 501(c)(3) eligible",
        "1) Employee donates or volunteers. "
        "2) Submits Citizenship Grant request via Apollo internal portal. "
        "3) Apollo Foundation matches/grants. "
        "Volunteer events hosted throughout year.",
        "Apollo Opportunity Foundation (launched Feb 2022): "
        "deploys capital + employee engagement for career education & "
        "economic empowerment; "
        "Hosts volunteer events; "
        "ERG-guided giving initiatives",
        "Active Apollo employees globally",
        "Apollo Opportunity Foundation launched with multi-year capital commitment; "
        "community investment in employee home markets",
        "https://doublethedonation.com/matching-gifts/apollo-global-management-inc | "
        "https://www.sec.gov/Archives/edgar/data/0001858681/000119312523215946/d543681ddef14a.htm"
    ),
    (
        "Stryker",
        "Stryker Impact Platform (Driven to Give)",
        "1:1 year-round; elevated rates during special campaigns",
        "Not publicly confirmed (estimated competitive range)",
        "Not publicly specified",
        "Paid volunteer day (1 day/year in US); "
        "volunteer grants via Impact platform; "
        "skills-based and virtual volunteer grants",
        "Stryker Impact Platform (proprietary, launched 2021); "
        "stryker.com/giving-back",
        "501(c)(3) nonprofits globally (5,700+ orgs supported in 2024)",
        "1) Employee donates via Stryker Impact platform. "
        "2) Stryker automatically matches at 1:1 (elevated during campaigns). "
        "3) Volunteer hours logged → grants disbursed. "
        "Personalized volunteer options: virtual, skills-based, in-person.",
        "Annual giving campaigns with elevated match rates; "
        "Giving Tuesday campaigns; "
        "Disaster relief giving; "
        "Stryker Johnston Foundation grants; "
        "Global giving and volunteering events",
        "Active employees globally; "
        "US employees: 1 paid volunteer day/year; "
        "virtual & skills-based options available worldwide",
        "2024: 11,000+ employee donors & volunteers; "
        "5,700+ nonprofits supported globally",
        "https://www.stryker.com/us/en/about/corporate-responsibility/giving-back.html | "
        "https://www.strykercareersblog.com/post/empowering-communities-stryker-s-commitment-to-giving-and-volunteering | "
        "https://doublethedonation.com/matching-gifts/stryker-corporation"
    ),
    (
        "Microsoft",
        "Microsoft Give Match",
        "1:1",
        "$15,000 / employee / year",
        "Not publicly specified",
        "$25 / volunteer hour",
        "Benevity — microsoft.benevity.org",
        "36,500+ nonprofits and schools in 110 countries",
        "1) Employee donates or logs volunteer hours. "
        "2) Submits via microsoft.benevity.org. "
        "3) Microsoft matches at 1:1 up to $15,000; "
        "volunteer hours matched at $25/hour. "
        "Payroll deduction, card, and stock donation options.",
        "Annual Give Campaign (October); "
        "Microsoft Change Agent Program (digital transformation for nonprofits); "
        "AI Nonprofit Advisors Program; "
        "Pro Bono legal and tech services; "
        "Disaster relief enhanced matching",
        "All full-time Microsoft employees globally; "
        "$15,000 cap is highest confirmed among all 16 companies",
        "CY2024: employees donated $255.6M + 1.2M volunteer hours "
        "to 36,500 nonprofits in 110 countries",
        "https://www.microsoft.com/en-us/corporate-responsibility/philanthropies/employee-engagement | "
        "https://cdn-dynmedia-1.microsoft.com/is/content/microsoftcorp/microsoft/msc/documents/presentations/CSR/FY25-Max-Employee-Engagement-March-2025.pdf | "
        "https://www.consiliowealth.com/insights/microsoft-charitable-giving-program-donor-advised-fund"
    ),
]

# fill data rows
row_colors = [LIGHT_GREY, WHITE]
for row_idx, row_data in enumerate(data, start=3):
    alt_color = row_colors[(row_idx - 3) % 2]
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill      = hfill(alt_color)
        cell.font      = hfont(DARK_TEXT, bold=(col_idx == 1), size=9)
        cell.alignment = wrap_align("left", "top")
        cell.border    = border_all
    ws.row_dimensions[row_idx].height = 90

# column widths
col_widths = [14, 28, 20, 30, 14, 28, 35, 35, 55, 55, 35, 45, 80]
for i, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ════════════════════════════════════════════════════════════════
# SHEET 2 – PROCESS DEEP-DIVE
# ════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Process Deep-Dive")
ws2.sheet_view.showGridLines = False
ws2.freeze_panes = "A3"

ws2.merge_cells("A1:G1")
ws2["A1"] = "EMPLOYEE GIVING — STEP-BY-STEP PROCESS & PLATFORM DETAILS"
ws2["A1"].font      = Font(name="Calibri", bold=True, size=14, color=WHITE)
ws2["A1"].fill      = hfill(DARK_NAVY)
ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 30

h2 = ["Company", "Platform / Tool", "How to Access",
      "Step-by-Step Process", "Submission Deadline",
      "Volunteer Grant Process", "Important Notes / Caveats"]
for col, h in enumerate(h2, 1):
    cell = ws2.cell(row=2, column=col, value=h)
    cell.font      = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill      = hfill(MID_BLUE)
    cell.alignment = wrap_align("center", "center")
    cell.border    = border_all
ws2.row_dimensions[2].height = 30

process_data = [
    ("Airbnb",
     "Benevity (internal portal)",
     "Internal HR/giving portal; Benevity referenced in nonprofit databases",
     "1. Log in to Airbnb internal giving portal.\n"
     "2. Search for eligible nonprofit.\n"
     "3. Donate ($25–$1,000).\n"
     "4. Submit match request within portal.\n"
     "5. Airbnb Foundation issues 1:1 match.",
     "Typically within the calendar year",
     "Limited public info; Community Fund ERG nominations for grants",
     "Community Fund (not same as employee matching) distributes $8.5M+/year "
     "separately via ERG-guided nonprofit nominations. "
     "Employee matching cap is $1,000 — lower than most tech peers."),
    ("Amazon",
     "Internal Amazon HR Portal (YourCause/Benevity referenced externally)",
     "Amazon internal employee portal (not publicly accessible)",
     "1. Access internal Amazon giving portal.\n"
     "2. Search and select eligible nonprofit.\n"
     "3. Donate via payroll deduction or card.\n"
     "4. Submit match request.\n"
     "5. Amazon matches up to $5,000/year at 1:1.",
     "Within the calendar year",
     "Volunteer Time Off available; specific volunteer grant $ amount not confirmed publicly",
     "AmazonSmile (customer program) was discontinued Jan 18, 2023. "
     "Employee matching is separate. "
     "Amazon has also run targeted $5M match for homelessness orgs. "
     "Specific 'ACE' program details are internal."),
    ("NVIDIA",
     "Benevity — Inspire 365 Platform",
     "Internal Inspire 365 portal (support: support@benevity.com)",
     "1. Employee makes external donation (cash, card, check, stock, or DAF) "
     "to qualified charity.\n"
     "2. Logs donation in Inspire 365 Benevity portal.\n"
     "3. No enrollment required — any NVIDIA employee can participate.\n"
     "4. NVIDIA Foundation reviews and issues match.\n"
     "5. Funds expire — do NOT roll over to next fiscal year.",
     "Before fiscal year end (NVIDIA fiscal year ends Jan 26)",
     "Volunteer hours also count toward $10,000 combined cap; "
     "log hours in Inspire 365 portal",
     "FY2024 participation: 40%+ of employees. "
     "Donations via Donor Advised Funds (DAFs) ARE eligible — "
     "notable exception vs many other programs. "
     "Stock donations also accepted."),
    ("Google",
     "Benevity — Internal branding: go/give",
     "go/give (internal shortlink); google.benevity.org (for nonprofits)",
     "1. Access go/give Benevity portal.\n"
     "2. Search 200,000+ eligible orgs.\n"
     "3. Donate (minimum $50) via payroll or card.\n"
     "4. Or log volunteer hours (earn $10/hour in grant).\n"
     "5. Submit by Dec 31 at 11:59 PM PT.\n"
     "6. Google matches up to $12,000/year.",
     "December 31 each year (hard deadline)",
     "Log volunteer hours in go/give portal → $10/hour grant to nonprofit of choice",
     "For 15+ years Google has given employees a year-end giving credit. "
     "Retired employees are NOT eligible. "
     "Google.org also operates separate grant and fellowship programs "
     "independent of employee matching."),
    ("HP",
     "HP Foundation Portal + Employee Product Giving Program",
     "hp.com/hpinfo/socialinnovation/us/",
     "Cash Match:\n"
     "1. Employee donates to eligible nonprofit.\n"
     "2. Submits via HP Foundation portal.\n"
     "3. HP Foundation matches 1:1 up to $5,000 while annual fund lasts.\n\n"
     "Product Match:\n"
     "1. Employee identifies technology donation need at nonprofit.\n"
     "2. Employee contributes 25% of retail value.\n"
     "3. HP Foundation contributes remaining 75%.\n"
     "4. HP tech product donated to nonprofit.",
     "While annual funds available (cash match); varies (product match)",
     "Volunteer grant: $50 per 10 hours volunteered per quarter "
     "(submit quarterly through HP Foundation portal)",
     "'While funds available' caveat means cash match pool can be exhausted "
     "before year-end — employees should submit early. "
     "Product Giving Program is unique in industry. "
     "HP Inc. and Hewlett Packard Enterprise (HPE) have separate programs "
     "post-split; confirm which entity you work for."),
    ("IBM",
     "Benevity — ibm.benevity.org",
     "ibm.benevity.org (login with IBM credentials)",
     "1. Employee donates to eligible org (min $25).\n"
     "2. Within 1 YEAR of donation, submits match request at ibm.benevity.org.\n"
     "3. IBM Foundation verifies organization eligibility.\n"
     "4. IBM Foundation issues 1:1 match (up to $10,000/year).\n"
     "Payroll deduction and credit card accepted.",
     "Within 1 year of the donation date",
     "IBM On Demand Community: log volunteer hours → IBM issues volunteer grant "
     "to eligible org; submit via ibm.benevity.org",
     "IBM has one of the longest-running corporate giving programs. "
     "Retirees are also eligible. "
     "IBM also operates the IBM Centennial of Service program "
     "and Skills Academy to provide pro bono tech training to nonprofits."),
    ("American Express",
     "GlobalGiving — Give2Gether platform",
     "globalgiving.org/amex-give2gether/ (use American Express employee credentials)",
     "1. Employee logs in to Give2Gether portal at globalgiving.org/amex-give2gether/.\n"
     "2. Browses verified nonprofits across 170+ countries.\n"
     "3. Donates via card or payroll deduction.\n"
     "4. American Express Foundation automatically matches up to $10,000/year.\n"
     "5. GlobalGiving disburses funds to nonprofit.",
     "Within the calendar year",
     "Volunteer grant program exists; specific $/hour not confirmed publicly; "
     "submit via AmexCares / Give2Gether portal",
     "Give2Gether is available in US, Canada, and India (full program); "
     "30+ countries participate via GlobalGiving expansion. "
     "Unique: uses GlobalGiving (third-party nonprofit marketplace) "
     "rather than Benevity/YourCause. "
     "2024: 6,000+ AmEx colleagues volunteered."),
    ("Apple",
     "Benevity (administered by Benevity on behalf of Apple)",
     "Apple internal giving portal powered by Benevity",
     "1. Employee donates to eligible 501(c)(3) (cash, check, or payroll deduction).\n"
     "2. Within 1 MONTH of donation, submits match request via Benevity portal.\n"
     "3. Apple matches at 1:1 (up to $10,000/year).\n"
     "4. Benevity disburses funds to nonprofit.\n"
     "Volunteer hours: submit in same portal → $25/hour grant.",
     "Within 1 month of donation (strict — shorter than most companies)",
     "$25/hour volunteer grant — among highest in industry; "
     "submit volunteer hours via Apple Benevity portal",
     "The 1-month submission deadline is significantly stricter than "
     "industry standard (most allow 1 year). "
     "Apple has previously run elevated 2:1 matches for racial equity and "
     "disaster relief campaigns. "
     "In 2021, 6 former Apple employees were charged with charity scam "
     "exploiting matching program — Apple has since tightened verification."),
    ("Texas Instruments",
     "TI Foundation Portal",
     "TI internal foundation portal; forms.matchinggifts.com/TIEDUGuidelines.pdf (guidelines)",
     "Educational Program:\n"
     "1. Employee donates to accredited educational institution.\n"
     "2. Submits TI Educational Matching Gift Form.\n"
     "3. TI Foundation matches 1:1 up to $10,000/year.\n\n"
     "Community Program:\n"
     "1. Employee donates to eligible 501(c)(3).\n"
     "2. Submits via TI Foundation portal.\n"
     "3. TI Foundation matches 1:1 up to $30,000/year.\n\n"
     "Volunteer Match:\n"
     "1. Employee logs volunteer hours.\n"
     "2. TI Foundation issues up to $1,000/year in volunteer grants.",
     "Within the calendar year",
     "Volunteer hours matched at up to $1,000/year; "
     "submit via TI Foundation portal",
     "TI Educational Matching Gift Program founded in 1970 — "
     "one of the oldest corporate matching programs in the US. "
     "The $30,000 community cap is among the highest of any company on this list. "
     "Employees AND retirees AND directors are eligible. "
     "Strong STEM education focus reflecting TI's semiconductor heritage."),
    ("Salesforce",
     "Salesforce Philanthropy Cloud (proprietary platform)",
     "Internal Salesforce giving portal; philanthropy.salesforce.com ecosystem",
     "1. Employee accesses Philanthropy Cloud portal.\n"
     "2. Searches and selects nonprofit.\n"
     "3. Donates via payroll deduction or card.\n"
     "4. OR uses 56 hours of annual VTO for volunteering.\n"
     "5. Salesforce matches donations 1:1 up to $10,000/year.\n"
     "6. Volunteer hours tracked → grants disbursed.",
     "Within the calendar year; VTO hours tracked annually",
     "56 hours VTO per employee per year; volunteer grants disbursed "
     "via Philanthropy Cloud; skills-based volunteering via Salesforce.org",
     "Salesforce is the FOUNDING company of the Pledge 1% movement "
     "(founded 1999 by Marc Benioff, before IPO). "
     "The 1-1-1 model (1% equity, 1% product, 1% time) has inspired "
     "19,000+ companies globally to adopt similar models. "
     "FY2025: 50% employee volunteer participation vs 15% industry average. "
     "Uses its own Philanthropy Cloud product — 'eats its own dog food.'"),
    ("Mastercard",
     "Internal Mastercard Portal + Mastercard Impact Fund",
     "Internal Mastercard HR portal; forms.matchinggifts.com/MasterCard.pdf (guidelines)",
     "1. Employee donates to eligible nonprofit.\n"
     "2. Submits match request via Mastercard internal portal.\n"
     "3. Mastercard Foundation issues 1:1 match up to $5,000/year (employees) "
     "or $15,000/year (board directors).\n\n"
     "Volunteer Rewards:\n"
     "1. Employee volunteers and logs hours.\n"
     "2. Earns volunteer rewards redeemable as nonprofit donations (up to $10,000).\n"
     "3. Up to 5 paid volunteer days per year.",
     "Within the calendar year",
     "5 paid volunteer days/year; volunteer rewards redeemable up to $10,000; "
     "log via Mastercard internal portal",
     "Board directors receive $15,000 match vs $5,000 for employees — "
     "higher-tier giving for leadership. "
     "Mastercard Impact Fund is a separate strategic philanthropy vehicle "
     "focused on financial inclusion and digital equity, "
     "independent of the employee matching program. "
     "Center for Inclusive Growth funds research and NGOs globally."),
    ("Intel",
     "Benevity — reopened Jan 15, 2025",
     "Intel Benevity portal (intel.benevity.org referenced); "
     "intel.com/corporate-responsibility/intel-matching-grants",
     "1. Employee donates to eligible org.\n"
     "2. Submits match request via Benevity portal "
     "within 60-DAY WINDOW of donation.\n"
     "3. Intel Foundation matches (payouts now monthly).\n"
     "4. COMBINED annual cap of $1,000 for donations + volunteer hour matches.\n"
     "Volunteer match: $10/hour toward combined cap.",
     "Within 60 days of donation (strict 2025 rule); Dec 15 for volunteer hours",
     "Volunteer: log hours in Benevity → $10/hour grant, "
     "counts toward $1,000 combined annual cap",
     "MAJOR 2025 CHANGES: Annual cap REDUCED from $10,000 to $1,000 combined "
     "(donations + volunteer hours), effective Jan 15, 2025. "
     "Change driven by Intel financial restructuring / layoffs. "
     "Payouts changed from annual to monthly. "
     "60-day submission window is shorter than most peers (was 1 year). "
     "Program reassessment planned for 2027. "
     "Pre-2025 participants should note previous volunteer hour deadline "
     "was Dec 15, 2024 — no retroactive submissions accepted."),
    ("NetApp",
     "Benevity (referenced in databases) — NetApp Cares Portal",
     "Internal NetApp Cares portal; netapp.com/responsibility/social-impact/netapp-cares/",
     "1. Employee donates to eligible nonprofit.\n"
     "2. Submits match request via NetApp Cares / Benevity portal.\n"
     "3. NetApp doubles donation (1:1 match).\n"
     "4. Volunteer hours: log in portal → volunteer grant disbursed.\n"
     "5. Board service grants also available.",
     "Within the calendar year",
     "Individual and team volunteer grants; board service grants; "
     "40 hours paid VTO/year; log hours via NetApp Cares portal",
     "Annual cap not publicly confirmed — estimated $2,500–$5,000 range "
     "based on industry comparables. "
     "Global Giving Challenge each December. "
     "NetApp Serves volunteer program is a flagship initiative. "
     "40 hours VTO is above industry average. "
     "Employees can earn grants for both individual and TEAM volunteering."),
    ("Apollo Global Management",
     "Apollo Internal Portal (Citizenship Grants) + Apollo Opportunity Foundation",
     "Internal Apollo HR portal; apollooportunity.org (foundation)",
     "1. Employee participates in volunteer event or donates.\n"
     "2. Submits Citizenship Grant request via Apollo internal portal.\n"
     "3. Apollo Foundation issues matching/grant.\n"
     "4. Apollo Opportunity Foundation: separate capital deployment "
     "for career education and economic empowerment programs.",
     "Within the calendar year (specific deadlines internal)",
     "Volunteer events hosted; volunteer rewards/grants available; "
     "specific $/hour not publicly confirmed",
     "Apollo Opportunity Foundation launched Feb 2022 with focus on "
     "career education, workforce development, economic empowerment. "
     "Less publicly documented than tech peers — private equity firms "
     "traditionally less transparent on employee giving. "
     "Community giving focuses on markets where employees live/work. "
     "Specific match ratios and caps are not publicly disclosed."),
    ("Stryker",
     "Stryker Impact Platform (launched 2021)",
     "Internal Impact platform; stryker.com/us/en/about/corporate-responsibility/giving-back.html",
     "1. Employee accesses Stryker Impact platform.\n"
     "2. Selects nonprofit (5,700+ orgs supported globally).\n"
     "3. Donates via platform.\n"
     "4. Stryker matches 1:1 year-round (elevated rates during campaigns).\n"
     "5. Volunteer hours logged → grants disbursed.\n"
     "6. Skills-based, virtual, and in-person volunteering options.",
     "Year-round; elevated campaigns have specific windows",
     "1 paid volunteer day/year (US employees); "
     "volunteer grants via Impact platform; "
     "virtual and skills-based volunteer options included",
     "Stryker Impact launched in 2021 as a dedicated global giving platform. "
     "Unique: elevated match rates during specific campaigns incentivize "
     "targeted giving. "
     "2024: 11,000+ donors/volunteers; 5,700+ nonprofits supported — "
     "broad reach for a 50,000-employee medical device company. "
     "Stryker Johnston Foundation also runs separate community grants. "
     "Specific annual cap not publicly confirmed."),
    ("Microsoft",
     "Benevity — microsoft.benevity.org",
     "microsoft.benevity.org (login with Microsoft employee credentials)",
     "1. Employee donates to eligible nonprofit via microsoft.benevity.org.\n"
     "2. OR logs volunteer hours (earns $25/hour grant).\n"
     "3. Microsoft matches donations 1:1 up to $15,000/year.\n"
     "4. Volunteer hours: $25/hour (no separate annual cap stated).\n"
     "Payment: payroll deduction, credit/debit card, stock donation.\n"
     "5. Microsoft Give Campaign (October) for annual pledge.",
     "Within the calendar year; Annual Give Campaign in October",
     "$25/hour volunteer grant; log via microsoft.benevity.org; "
     "no minimum hours stated",
     "Microsoft has the HIGHEST confirmed annual match cap ($15,000) "
     "of all 16 companies on this list. "
     "CY2024: $255.6M donated by employees (including match) — "
     "largest employee giving program by dollar volume in tech. "
     "Microsoft Change Agent Program helps nonprofits adopt Microsoft tech. "
     "AI Nonprofit Advisors: free 1-hour AI strategy sessions for nonprofits. "
     "Pro Bono legal + tech volunteer program. "
     "Stock donations accepted via Benevity portal."),
]

for row_idx, row_data in enumerate(process_data, start=3):
    alt_color = row_colors[(row_idx - 3) % 2]
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill      = hfill(alt_color)
        cell.font      = hfont(DARK_TEXT, bold=(col_idx == 1), size=9)
        cell.alignment = wrap_align("left", "top")
        cell.border    = border_all
    ws2.row_dimensions[row_idx].height = 150

col_widths2 = [14, 30, 35, 65, 30, 50, 65]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# ════════════════════════════════════════════════════════════════
# SHEET 3 – COMPARISON QUICK-REFERENCE
# ════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Quick Comparison")
ws3.sheet_view.showGridLines = False
ws3.freeze_panes = "A3"

ws3.merge_cells("A1:H1")
ws3["A1"] = "QUICK COMPARISON — KEY METRICS AT A GLANCE (RANKED BY ANNUAL CAP)"
ws3["A1"].font      = Font(name="Calibri", bold=True, size=14, color=WHITE)
ws3["A1"].fill      = hfill(DARK_NAVY)
ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[1].height = 30

h3 = ["Rank", "Company", "Program Name", "Match\nRatio",
      "Annual Cap\n(USD)", "Volunteer\nGrant", "Platform", "Key Source"]
for col, h in enumerate(h3, 1):
    cell = ws3.cell(row=2, column=col, value=h)
    cell.font      = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill      = hfill(MID_BLUE)
    cell.alignment = wrap_align("center", "center")
    cell.border    = border_all
ws3.row_dimensions[2].height = 35

quick_data = [
    (1,  "Microsoft",       "Microsoft Give Match",                   "1:1", "$15,000",        "$25/hr",              "Benevity",             "microsoft.com/en-us/corporate-responsibility/philanthropies/employee-engagement"),
    (2,  "Texas Instruments","TI Foundation Matching",                "1:1", "$30,000 (comm.) / $10,000 (edu.)", "$1,000/yr", "TI Foundation",   "ti.com/about-ti/citizenship-community/giving.html"),
    (3,  "Google",           "go/give (Benevity)",                    "1:1", "$12,000",         "$10/hr",              "Benevity",             "doublethedonation.com/forms/google-guidelines.pdf"),
    (4,  "NVIDIA",           "Inspire 365",                           "1:1", "$10,000",         "Incl. in $10k cap",   "Benevity",             "nvidia.com/en-us/foundation/"),
    (5,  "Apple",            "Apple Matching Gifts Program",          "1:1", "$10,000",         "$25/hr",              "Benevity",             "doublethedonation.com/matching-gifts/apple-inc"),
    (6,  "IBM",              "IBM Matching Grants",                   "1:1", "$10,000",         "Volunteer Grants",    "Benevity",             "ibm.com/responsibility/programs/volunteerism-giving"),
    (7,  "American Express", "Give2Gether (GlobalGiving)",            "1:1", "$10,000",         "Yes (amt. TBC)",      "GlobalGiving",         "globalgiving.org/amex-give2gether/"),
    (8,  "Salesforce",       "1-1-1 Model / Philanthropy Cloud",      "1:1", "$10,000",         "56 hrs VTO + grants", "Philanthropy Cloud",   "salesforce.com/company/philanthropy/employee-volunteering-giving/"),
    (9,  "Amazon",           "Amazon Employee Giving",                "1:1", "$5,000",          "VTO (amt. TBC)",      "Internal Portal",      "doublethedonation.com/matching-gifts/amazon"),
    (10, "HP",               "HP Foundation Cash Match",              "1:1", "$5,000",          "$5/hr (~$50/10hrs)",  "HP Foundation Portal", "hp.com/us-en/hp-information/hp-foundation.html"),
    (11, "Mastercard",       "Mastercard Matching Gifts Program",     "1:1", "$5,000 (emp.) / $15,000 (directors)", "Up to $10,000 vol. rewards", "Internal Portal", "mastercard.com/us/en/for-the-world/people/community-impact.html"),
    (12, "NetApp",           "NetApp Cares Matching Gifts",           "1:1", "~$2,500–$5,000 (est.)", "Vol. + board grants", "Benevity",        "netapp.com/responsibility/social-impact/netapp-cares/"),
    (13, "Stryker",          "Stryker Impact Platform",               "1:1 + elevated", "Not confirmed", "1 paid vol. day + grants", "Impact Platform", "stryker.com/us/en/about/corporate-responsibility/giving-back.html"),
    (14, "Airbnb",           "Community Fund + Employee Matching",    "1:1", "$1,000",          "Not confirmed",       "Benevity (ref.)",      "news.airbnb.com/2025-community-fund/"),
    (15, "Intel",            "Intel Foundation (Intel Involved)",     "1:1", "$1,000 combined\n(2025 reduction)", "$10/hr (in cap)", "Benevity", "intel.com/content/www/us/en/corporate-responsibility/intel-foundation-employee-generosity.html"),
    (16, "Apollo Global",    "Citizenship Grants / Opportunity Fdn.", "Not confirmed", "Not confirmed", "Yes (amt. TBC)", "Internal Portal", "doublethedonation.com/matching-gifts/apollo-global-management-inc"),
]

# color coding for tiers
tier_colors = {
    1: "C6EFCE",   # green  (>=$10k)
    2: "FFEB9C",   # yellow ($5k-$9,999)
    3: "FFC7CE",   # red    (<$5k or unconfirmed)
}

for row_idx, row_data in enumerate(quick_data, start=3):
    rank = row_data[0]
    if rank <= 8:
        tier_color = tier_colors[1]
    elif rank <= 12:
        tier_color = tier_colors[2]
    else:
        tier_color = tier_colors[3]

    for col_idx, value in enumerate(row_data, start=1):
        cell = ws3.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill      = hfill(tier_color)
        cell.font      = hfont(DARK_TEXT, bold=(col_idx in [1, 2]), size=9)
        cell.alignment = wrap_align("center" if col_idx <= 2 else "left", "center")
        cell.border    = border_all
    ws3.row_dimensions[row_idx].height = 40

# legend
legend_row = len(quick_data) + 4
ws3.merge_cells(f"A{legend_row}:H{legend_row}")
ws3[f"A{legend_row}"] = "COLOUR LEGEND:"
ws3[f"A{legend_row}"].font = Font(name="Calibri", bold=True, size=10, color=DARK_TEXT)

for label, color, start_col in [
    ("$10,000+ Annual Cap (Top Tier)", "C6EFCE", 1),
    ("$5,000–$9,999 Annual Cap (Mid Tier)", "FFEB9C", 4),
    ("Below $5,000 / Unconfirmed (Lower Tier)", "FFC7CE", 6),
]:
    cell = ws3.cell(row=legend_row + 1, column=start_col, value=label)
    cell.fill = hfill(color)
    cell.font = Font(name="Calibri", bold=True, size=9)
    cell.border = border_all
    cell.alignment = wrap_align("center", "center")

col_widths3 = [6, 18, 30, 12, 22, 22, 20, 55]
for i, w in enumerate(col_widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# ════════════════════════════════════════════════════════════════
# SHEET 4 – SOURCES & METHODOLOGY
# ════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Sources & Methodology")
ws4.sheet_view.showGridLines = False

ws4.merge_cells("A1:C1")
ws4["A1"] = "SOURCES, METHODOLOGY & DATA CONFIDENCE"
ws4["A1"].font      = Font(name="Calibri", bold=True, size=14, color=WHITE)
ws4["A1"].fill      = hfill(DARK_NAVY)
ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws4.row_dimensions[1].height = 30

for col, h in enumerate(["Company / Source Category", "Source URL / Reference", "Data Confidence"], 1):
    cell = ws4.cell(row=2, column=col, value=h)
    cell.font      = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill      = hfill(MID_BLUE)
    cell.alignment = wrap_align("center", "center")
    cell.border    = border_all
ws4.row_dimensions[2].height = 25

sources = [
    ("METHODOLOGY NOTE",
     "Primary sources: Official company CSR/philanthropy pages, "
     "official program guidelines (PDF), Benevity platform references, "
     "GlobalGiving partnership announcements, Glassdoor Benefits pages, "
     "Intel Retiree Organization (for 2025 Intel changes), "
     "Double the Donation database, NVIDIA official blog, "
     "Microsoft official CSR disclosure, Salesforce official philanthropy pages.",
     "N/A"),
    ("Airbnb — Community Fund",          "https://news.airbnb.com/2025-community-fund/",                                                              "HIGH — Official Airbnb newsroom"),
    ("Airbnb — Matching Gifts",          "https://doublethedonation.com/matching-gifts/airbnb",                                                       "MEDIUM — Third-party database"),
    ("Amazon — Matching Gifts",          "https://doublethedonation.com/matching-gifts/amazon",                                                       "MEDIUM — Third-party database"),
    ("Amazon — Community Impact",        "https://www.amazon.jobs/content/en/teams/ccr/aci",                                                          "HIGH — Official Amazon page"),
    ("Amazon — Glassdoor Benefits",      "https://www.glassdoor.com/Benefits/Amazon-Charitable-Gift-Matching-US-BNFT21_E6036_N1.htm",                 "MEDIUM — Employee-reported"),
    ("NVIDIA — Inspire 365 Foundation",  "https://www.nvidia.com/en-us/foundation/",                                                                  "HIGH — Official NVIDIA Foundation page"),
    ("NVIDIA — FY2024 Giving Blog",      "https://blogs.nvidia.com/blog/nvidia-life-inspire-365-2024/",                                               "HIGH — Official NVIDIA blog"),
    ("NVIDIA — Matching Gifts PDF",      "https://files.doublethedonation.com/forms/nvidia-guidelines.pdf",                                           "HIGH — Official program guidelines PDF"),
    ("Google — Benevity Portal T&C",     "https://google.benevity.org/terms-and-conditions",                                                          "HIGH — Official Google/Benevity portal"),
    ("Google — Program Guidelines",      "https://doublethedonation.com/forms/google-guidelines.pdf",                                                 "HIGH — Official Google program PDF"),
    ("HP — HP Foundation",               "https://www.hp.com/us-en/hp-information/hp-foundation.html",                                                "HIGH — Official HP page"),
    ("HP — Product Matching",            "http://www.hp.com/hpinfo/socialinnovation/us/product_matching.html",                                        "HIGH — Official HP page"),
    ("IBM — Responsibility Programs",    "https://www.ibm.com/responsibility/programs/volunteerism-giving",                                            "HIGH — Official IBM page"),
    ("IBM — Benevity Portal",            "https://ibm.benevity.org/user/login",                                                                       "HIGH — Official IBM Benevity portal"),
    ("American Express — Give2Gether",   "https://www.globalgiving.org/amex-give2gether/",                                                            "HIGH — Official GlobalGiving/AmEx page"),
    ("American Express — Colleagues",    "https://about.americanexpress.com/we-engage-our-colleagues",                                                 "HIGH — Official AmEx newsroom"),
    ("American Express — Press Release", "https://www.businesswire.com/news/home/20201013005631/en/American-Express-Partners-With-GlobalGiving",       "HIGH — BusinessWire press release"),
    ("Apple — Matching Gifts DB",        "https://doublethedonation.com/matching-gifts/apple-inc",                                                     "MEDIUM — Third-party database"),
    ("Apple — Program Guide PDF",        "https://forms.matchinggifts.com/AppleGuide.pdf",                                                            "HIGH — Official Apple program PDF"),
    ("Apple — Program FAQs PDF",         "https://assets.globalgiving.org/docs/apple/apple-faqs.pdf",                                                 "HIGH — Official Apple FAQ PDF"),
    ("Texas Instruments — Giving",       "https://www.ti.com/about-ti/citizenship-community/giving.html",                                             "HIGH — Official TI page"),
    ("Texas Instruments — Eligibility",  "https://www.ti.com/about-ti/citizenship-community/giving/eligible-organizations.html",                      "HIGH — Official TI page"),
    ("Texas Instruments — Alumni",       "https://tialumni.org/volunteer-2/ti-foundation-and-matching-gifts/",                                        "MEDIUM — TI Alumni Association"),
    ("Salesforce — Philanthropy",        "https://www.salesforce.com/company/philanthropy/employee-volunteering-giving/",                              "HIGH — Official Salesforce page"),
    ("Salesforce — Employee Giving",     "https://www.salesforce.com/news/stories/employee-giving-model/",                                            "HIGH — Official Salesforce news"),
    ("Salesforce — Pledge 1%",           "https://www.salesforce.com/company/pledge/",                                                                "HIGH — Official Salesforce page"),
    ("Mastercard — Community Impact",    "https://www.mastercard.com/us/en/for-the-world/people/community-impact.html",                               "HIGH — Official Mastercard page"),
    ("Mastercard — Impact Fund",         "https://www.mastercardcenter.org/about-the-center/mastercard-impact-fund",                                  "HIGH — Official Mastercard Center page"),
    ("Mastercard — Guidelines PDF",      "https://forms.matchinggifts.com/MasterCard.pdf",                                                            "HIGH — Official program guidelines PDF"),
    ("Intel — Foundation Generosity",    "https://www.intel.com/content/www/us/en/corporate-responsibility/intel-foundation-employee-generosity.html", "HIGH — Official Intel page"),
    ("Intel — 2025 Benevity Update",     "https://intelretiree.com/2025/01/17/benevity-donation-volunteer-matching-site-opens/",                      "HIGH — Intel Retiree Organization (official communication)"),
    ("Intel — 2025 Program Changes",     "https://www.gomotionapp.com/team/osthsc/page/news/538907/updates-to-intels-donation-and-volunteer-matching-program-2025", "MEDIUM — Third-party relay of Intel communication"),
    ("NetApp — NetApp Cares",            "https://www.netapp.com/responsibility/social-impact/netapp-cares/",                                         "HIGH — Official NetApp page"),
    ("NetApp — Serves Blog",             "https://www.netapp.com/blog/netapp-serves-employee-volunteer-program/",                                     "HIGH — Official NetApp blog"),
    ("NetApp — Matching Gifts DB",       "https://doublethedonation.com/matching-gifts/netapp",                                                       "MEDIUM — Third-party database"),
    ("Apollo Global — Matching Gifts",   "https://doublethedonation.com/matching-gifts/apollo-global-management-inc",                                 "MEDIUM — Third-party database"),
    ("Apollo Global — DEF 14A (SEC)",    "https://www.sec.gov/Archives/edgar/data/0001858681/000119312523215946/d543681ddef14a.htm",                   "HIGH — SEC official filing"),
    ("Stryker — Giving Back",            "https://www.stryker.com/us/en/about/corporate-responsibility/giving-back.html",                             "HIGH — Official Stryker page"),
    ("Stryker — Careers Blog",           "https://www.strykercareersblog.com/post/empowering-communities-stryker-s-commitment-to-giving-and-volunteering", "HIGH — Official Stryker careers blog"),
    ("Stryker — Matching Gifts DB",      "https://doublethedonation.com/matching-gifts/stryker-corporation",                                          "MEDIUM — Third-party database"),
    ("Microsoft — CSR Employee Giving",  "https://www.microsoft.com/en-us/corporate-responsibility/philanthropies/employee-engagement",               "HIGH — Official Microsoft CSR page"),
    ("Microsoft — FY25 Engagement PDF",  "https://cdn-dynmedia-1.microsoft.com/is/content/microsoftcorp/microsoft/msc/documents/presentations/CSR/FY25-Max-Employee-Engagement-March-2025.pdf", "HIGH — Official Microsoft CSR PDF"),
    ("Microsoft — Benevity Guide PDF",   "https://forms.matchinggifts.com/The%20Microsoft%20Employee%20Giving%20Program%20and%20Benevity.pdf",        "HIGH — Official Microsoft/Benevity guide"),
    ("DATA CAVEAT",
     "Program details can change annually. Intel 2025 changes confirmed. "
     "Always verify current terms via each company's official HR/benefits portal. "
     "Annual caps, match ratios, and volunteer grant rates reflect best available "
     "public data as of research date (June 2026). "
     "NetApp, Apollo, Stryker, and Airbnb specific caps are estimated or unconfirmed — "
     "marked accordingly in Master Overview sheet.",
     "READ BEFORE USE"),
]

for row_idx, row_data in enumerate(sources, start=3):
    alt_color = row_colors[(row_idx - 3) % 2]
    is_note = row_data[0] in ("METHODOLOGY NOTE", "DATA CAVEAT")
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws4.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill      = hfill("FFF2CC" if is_note else alt_color)
        cell.font      = hfont(DARK_TEXT, bold=(col_idx == 1 or is_note), size=9)
        cell.alignment = wrap_align("left", "top")
        cell.border    = border_all
    ws4.row_dimensions[row_idx].height = 45 if is_note else 30

col_widths4 = [40, 100, 25]
for i, w in enumerate(col_widths4, 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

# ════════════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════════════
output_path = "/home/user/jcode/Employee_Giving_Programs_Comprehensive.xlsx"
wb.save(output_path)
print(f"Excel file saved: {output_path}")
