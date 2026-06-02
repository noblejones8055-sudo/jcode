import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── Palette ────────────────────────────────────────────────────
DARK_NAVY   = "1B2A4A"
MID_BLUE    = "2E5CA8"
FOREST_GRN  = "1A5C38"
MID_GREEN   = "217A4F"
LIGHT_GREEN = "D5F5E3"
AMBER       = "D4A017"
LIGHT_AMBER = "FFF3CD"
RED_DARK    = "C0392B"
LIGHT_RED   = "FADBD8"
LIGHT_BLUE  = "D6E4F7"
WHITE       = "FFFFFF"
LIGHT_GREY  = "F2F2F2"
MED_GREY    = "CCCCCC"
DARK_TEXT   = "1A1A2E"
YELLOW_HL   = "FFFACD"

def hfill(c): return PatternFill("solid", fgColor=c)
def hfont(c=DARK_TEXT, bold=False, sz=9): return Font(color=c, bold=bold, size=sz, name="Calibri")
thin  = Side(style="thin",   color=MED_GREY)
brd   = Border(left=thin, right=thin, top=thin, bottom=thin)
def walign(h="left", v="top"):
    return Alignment(horizontal=h, vertical=v, wrap_text=True)

def make_title(ws, text, cols, color=DARK_NAVY):
    ws.merge_cells(f"A1:{get_column_letter(cols)}1")
    ws["A1"] = text
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=WHITE)
    ws["A1"].fill = hfill(color)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

def make_header_row(ws, headers, row=2, color=MID_BLUE):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
        cell.fill = hfill(color)
        cell.alignment = walign("center", "center")
        cell.border = brd
    ws.row_dimensions[row].height = 30

ALT = [LIGHT_GREY, WHITE]

# ══════════════════════════════════════════════════════════════════
# SHEET 1 — ACTIVE STATUS CHECK (ALL 32 COMPANIES)
# ══════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Active Status — All 32 Companies"
ws1.sheet_view.showGridLines = False
ws1.freeze_panes = "A3"

make_title(ws1, "ACTIVE STATUS CHECK — ALL 32 COMPANIES (VERIFIED JUNE 2026)", 9, DARK_NAVY)
make_header_row(ws1, [
    "Company", "Program Name", "Status (2025/26)",
    "Last Confirmed Evidence", "Key Change / Note",
    "Platform", "Waste Warriors Eligible?",
    "CAF America Needed?", "Action for Waste Warriors"
])

status_data = [
    # ─── BATCH 1: Original 16 ───
    ("Microsoft",             "Microsoft Give Match",                    "✅ ACTIVE",
     "FY2025: $263.7M + 1.2M volunteer hours, 36,500 nonprofits (official MS page)",
     "No changes; among largest corporate giving programs globally",
     "Benevity (microsoft.benevity.org)",
     "YES — if registered on Benevity",
     "YES (for India NGO receiving US match)",
     "Register on Benevity → employees auto-discover you"),
    ("Google",                "go/give (Spark!) — Benevity",            "✅ ACTIVE",
     "Dec 31 2025 deadline confirmed on google.benevity.org; matching cap $12,000",
     "No changes; 200,000+ eligible orgs",
     "Benevity (google.benevity.org)",
     "YES — if registered on Benevity",
     "YES (for India NGO)",
     "Register on Benevity → available to all Google employees"),
    ("Apple",                 "Apple Matching Gifts Program",            "✅ ACTIVE",
     "Benevity administers program; $25/hour volunteer grant confirmed; $10,000 cap",
     "No changes reported; using Benevity since 2022+",
     "Benevity (causes.benevity.com)",
     "YES — if registered on Benevity",
     "YES",
     "Register on Benevity → auto-eligible for Apple employee matches"),
    ("Salesforce",            "Salesforce 1-1-1 / Philanthropy Cloud",  "✅ ACTIVE",
     "2025: 50% employee participation; $50 given to each of 76,000 employees for charity",
     "No changes; 56 hours VTO; expanded to $50 gift per employee in 2025",
     "Salesforce Philanthropy Cloud (own platform)",
     "YES — apply via Philanthropy Cloud",
     "YES",
     "Apply directly via Salesforce Philanthropy Cloud portal"),
    ("IBM",                   "IBM Matching Grants (Benevity)",          "✅ ACTIVE",
     "ibm.benevity.org confirmed active; $10,000 cap; IBM Sustainability Accelerator running",
     "No changes; IBM added STEM and tech volunteering focus",
     "Benevity (ibm.benevity.org)",
     "YES — if registered on Benevity",
     "YES",
     "Register on Benevity → auto-discoverable by IBM employees"),
    ("NVIDIA",                "NVIDIA Foundation Inspire 365",           "✅ ACTIVE",
     "FY2024: 40%+ employees participated, $16M+ donated; Benevity portal confirmed",
     "No changes; annual year-end campaign each Nov–Dec",
     "Benevity",
     "YES — if registered on Benevity",
     "YES",
     "Register on Benevity; pitch NVIDIA's environmental focus"),
    ("Amazon",                "Amazon Employee Giving",                  "✅ ACTIVE",
     "Matching gift program confirmed via Amazon jobs page and Glassdoor 2025",
     "AmazonSmile shut Jan 2023 — separate from employee giving; matching still active",
     "Internal Amazon portal",
     "YES — apply via Amazon CSR team directly",
     "YES",
     "Submit NGO registration directly to Amazon Community Impact team"),
    ("HP",                    "HP Foundation Employee Cash Match",       "✅ ACTIVE",
     "HP Foundation page confirmed; cash match $5,000 + product giving program",
     "No changes; both cash and product programs active",
     "HP Foundation portal",
     "YES — register with HP Foundation",
     "YES",
     "Register with HP Foundation at hp.com/foundation"),
    ("Texas Instruments",     "TI Foundation Matching Gifts",            "✅ ACTIVE",
     "TI Foundation active; $30,000 cap confirmed; Educational Matching Gift Program since 1970",
     "No changes; educational focus maintained",
     "TI Foundation portal",
     "YES — apply as 501(c)(3) equivalent",
     "YES",
     "Contact TI Foundation; position as environmental education org"),
    ("American Express",      "Give2Gether (GlobalGiving)",              "✅ ACTIVE",
     "GlobalGiving partnership confirmed; $10,000 cap; available in 30+ countries",
     "No changes; still via GlobalGiving platform",
     "GlobalGiving (globalgiving.org/amex-give2gether/)",
     "YES — register as GlobalGiving partner",
     "YES (via GlobalGiving)",
     "Register on GlobalGiving → auto-eligible for Amex employee matches"),
    ("Mastercard",            "Mastercard Matching Gifts Program",       "✅ ACTIVE",
     "Mastercard community impact page confirmed active; $5,000 employee cap",
     "No changes; Mastercard Impact Fund running separately",
     "Mastercard internal portal",
     "YES — register directly with Mastercard",
     "YES",
     "Contact Mastercard CSR / Corporate Responsibility team"),
    ("Intel",                 "Intel Foundation Matching Program",       "⚠️ ACTIVE — REDUCED",
     "Reopened Jan 15 2025 on Benevity; cap CUT to $1,000 combined (was higher)",
     "MAJOR CHANGE: cap reduced to $1,000 total (donations + volunteer hrs combined); "
     "monthly payouts now (was annual); 60-day submission window",
     "Benevity (reopened Jan 2025)",
     "YES — if registered on Benevity",
     "YES",
     "Register on Benevity; note reduced cap; still worth pursuing for visibility"),
    ("NetApp",                "NetApp Cares Matching Gifts",             "✅ ACTIVE",
     "NetApp Cares page active; Global Giving Challenge Dec 2025 confirmed",
     "No changes; June volunteering events and December giving challenge annual",
     "Benevity (referenced)",
     "YES — if registered on Benevity",
     "YES",
     "Register on Benevity; pitch environmental/waste angle to NetApp Cares team"),
    ("Apollo Global Mgmt",    "Apollo Citizenship Grants / Opp. Fdn.",  "✅ ACTIVE",
     "Apollo Opportunity Foundation active; Citizenship Grants for matching confirmed",
     "Focus: career education, workforce development, economic empowerment",
     "Apollo internal portal",
     "PARTIAL — align programs with their focus areas",
     "YES",
     "Contact Apollo Opportunity Foundation with livelihoods angle"),
    ("Stryker",               "Stryker Impact Platform",                 "✅ ACTIVE",
     "2024: 11,000+ employee donors/volunteers; 5,700+ nonprofits globally confirmed",
     "No changes; Impact platform launched 2021, fully operational",
     "Stryker Impact (own platform)",
     "YES — register via Impact platform",
     "YES",
     "Register on Stryker Impact platform; medical waste angle is strong fit"),
    ("Salesforce",            "NOTE: Already listed above",              "—",
     "—", "—", "—", "—", "—", "—"),

    # ─── BATCH 2: New 16 Companies ───
    ("JPMorgan Chase",        "Employee Matching Gift Program",          "✅ ACTIVE",
     "JPMorgan volunteerism page active 2025; Global Month of Service confirmed",
     "CyberGrants portal; salary cap applies (≤$150K base)",
     "CyberGrants (cybergrants.com/jpmc/giving)",
     "YES — register via CyberGrants",
     "YES",
     "Register on CyberGrants platform; align with JPMorgan's community focus"),
    ("Goldman Sachs",         "Goldman Sachs Matching Gift Program",     "✅ ACTIVE",
     "$20,000 cap; EasyMatch portal confirmed; GS Gives DAF active",
     "No changes; quarterly processing; 180-day window",
     "EasyMatch (easymatch.com/gsmg2/)",
     "YES — register via EasyMatch/Goldman portal",
     "YES",
     "Reach out to Goldman Sachs Gives team; high cap makes it high priority"),
    ("Bank of America",       "Employee-Directed Matching Gifts",        "✅ ACTIVE",
     "2024: $69M+ in combined giving to 48,000+ nonprofits confirmed officially",
     "No changes; Benevity-powered; global program",
     "Benevity",
     "YES — if registered on Benevity",
     "YES",
     "Register on Benevity; BofA has environment/climate focus — strong fit"),
    ("Wells Fargo",           "Wells Fargo Matching Gifts Program",      "✅ ACTIVE",
     "Philanthropy FAQs page active; 2025 community giving page confirmed",
     "No changes; 6-month submission window (shorter than peers)",
     "YourCause",
     "YES — register on YourCause/NPOconnect",
     "YES",
     "Register on YourCause (NPOconnect); low cap but large employee base"),
    ("Johnson & Johnson",     "J&J Matching Gifts Program",              "✅ ACTIVE",
     "jnj.com/caring-and-giving active; 2:1 match and program guidelines confirmed",
     "No changes; 2:1 match for current employees still in place",
     "J&J internal portal",
     "YES — register via J&J portal",
     "YES",
     "Contact J&J Foundation; 2:1 match is excellent ROI for your effort"),
    ("Pfizer",                "Pfizer Foundation Matching Gift Program", "✅ ACTIVE",
     "pfizerplus.com active; Give Forward page confirmed 2025; retiree program noted",
     "No changes; March 1 deadline is generous",
     "Pfizer Foundation (pfizerplus.com)",
     "YES — register via Pfizer Foundation",
     "YES",
     "Register at pfizerplus.com; environment/health overlap is strong angle"),
    ("Merck",                 "Partnership for Giving (CyberGrants)",    "✅ ACTIVE",
     "merck.com philanthropy page active; CyberGrants portal confirmed; $30,000 cap",
     "No changes; highest cap in pharma; retirees NOT eligible",
     "CyberGrants",
     "YES — register via CyberGrants",
     "YES",
     "Register on CyberGrants; position waste/health angle with Merck Foundation"),
    ("Deloitte",              "Deloitte Foundation Matching Gifts",      "✅ ACTIVE",
     "Foundation guidelines PDF updated Apr 2024; $32,500 cap confirmed",
     "No changes; part-time and retirees excluded",
     "Deloitte Foundation portal",
     "YES — register via Deloitte Foundation",
     "YES",
     "Apply to Deloitte Foundation; $32,500 cap = highest of all 32 companies"),
    ("Coca-Cola",             "Coca-Cola Matching Gifts Program",        "✅ ACTIVE",
     "Program confirmed; 2:1 ratio; Feb 28 deadline; retirees eligible",
     "No changes; Feb 28 generous deadline",
     "Coca-Cola Foundation portal",
     "YES — register via Coca-Cola Foundation",
     "YES",
     "Contact Coca-Cola Foundation; environment/water/packaging focus = fit"),
    ("PepsiCo",               "PepsiCo Foundation Matching Gifts",       "✅ ACTIVE",
     "pepsico.yourcause.com active; YourCause platform confirmed; pep+ sustainability",
     "No changes; 2:1 for 50+ volunteer hours is active incentive",
     "YourCause (pepsico.yourcause.com)",
     "YES — register on YourCause/NPOconnect",
     "YES",
     "Register on YourCause; PepsiCo's pep+ sustainability focus = great fit"),
    ("Nike",                  "Give Your Best (GYB) — Benevity",        "✅ ACTIVE",
     "Nike community engagement page active; NCIF 2025 cycle confirmed at oregoncf.org",
     "No changes; $25,000 cap; 2:1 on Giving Tuesday",
     "Benevity (NikeNet internal)",
     "YES — register on Benevity",
     "YES",
     "Register on Benevity; ALSO apply for NCIF grant via Oregon Community Foundation"),
    ("Cisco",                 "Cisco Matching Gifts (Benevity)",         "✅ ACTIVE",
     "Benevity case study FY24/FY25 confirmed 80-85% participation; $90M+ since FY23",
     "Previous suspension REINSTATED; now via Benevity; 80h VTO active",
     "Benevity",
     "YES — if registered on Benevity",
     "YES",
     "Register on Benevity; Cisco's 85% participation = large engaged donor base"),
    ("Dell Technologies",     "Dell Employee Giving (YourCause)",        "✅ ACTIVE",
     "dell.yourcause.com active; Global Month of Service June 2025 confirmed",
     "No changes; quarterly volunteer grant structure active",
     "YourCause (dell.yourcause.com)",
     "YES — register on YourCause/NPOconnect",
     "YES",
     "Register on YourCause; pitch to Dell's Global Month of Service June campaign"),
    ("Adobe",                 "Create Change (Benevity)",                "✅ ACTIVE",
     "Record giving season confirmed via Benevity case study; ECF grants active 2025",
     "No changes; team volunteer grant ($1,000/team) highly active",
     "Benevity (causes.benevity.com)",
     "YES — register on Benevity",
     "YES",
     "Register on Benevity; ALSO apply for Adobe ECF (Employee Community Fund) grant"),
    ("SAP",                   "SAP Together (SAP internal portal)",      "✅ ACTIVE",
     "SAP Social Sabbatical program active; $20/hour volunteer grant confirmed",
     "No changes; $750 cap is low but $20/hr volunteer grant is best $/hr of all 32",
     "SAP internal portal",
     "YES — register via SAP portal",
     "YES",
     "Register with SAP; request Social Sabbatical placement (4-week volunteer leave)"),
    ("Qualcomm",              "Qualcomm Employee Giving (Benevity)",     "✅ ACTIVE",
     "Benevity portal referenced; Oct 1–Sep 30 fiscal year; tiered cap confirmed",
     "No changes; Oct–Sep fiscal year means apply by September",
     "Benevity",
     "YES — if registered on Benevity",
     "YES",
     "Register on Benevity; note Oct–Sep fiscal year for campaign timing"),
]

# Remove duplicate Salesforce row
status_data = [r for r in status_data if r[0] != "Salesforce" or r[1] != "NOTE: Already listed above"]

row_colors_status = {
    "✅ ACTIVE":          LIGHT_GREEN,
    "⚠️ ACTIVE — REDUCED": LIGHT_AMBER,
    "❌ SUSPENDED":       LIGHT_RED,
    "PARTIAL":            LIGHT_BLUE,
}

for row_idx, row_data in enumerate(status_data, start=3):
    status = row_data[2]
    if "✅" in status:     fill = LIGHT_GREEN
    elif "⚠️" in status:  fill = LIGHT_AMBER
    elif "❌" in status:   fill = LIGHT_RED
    else:                  fill = ALT[(row_idx-3) % 2]
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws1.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill = hfill(fill)
        cell.font = hfont(DARK_TEXT, bold=(col_idx in [1,3]), sz=9)
        cell.alignment = walign("left", "top")
        cell.border = brd
    ws1.row_dimensions[row_idx].height = 75

col_w1 = [18, 26, 22, 48, 48, 30, 20, 18, 45]
for i, w in enumerate(col_w1, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# Legend
leg_row = len(status_data) + 4
ws1.merge_cells(f"A{leg_row}:I{leg_row}")
ws1[f"A{leg_row}"] = "STATUS LEGEND:"
ws1[f"A{leg_row}"].font = Font(name="Calibri", bold=True, size=10)
for label, color, col in [
    ("✅ ACTIVE — Program confirmed running",          LIGHT_GREEN,  1),
    ("⚠️ ACTIVE — REDUCED / Changed terms",           LIGHT_AMBER,  4),
    ("❌ SUSPENDED / Ended",                           LIGHT_RED,    7),
]:
    c = ws1.cell(row=leg_row + 1, column=col, value=label)
    c.fill = hfill(color)
    c.font = Font(name="Calibri", bold=True, size=9)
    c.border = brd
    c.alignment = walign("center", "center")

# ══════════════════════════════════════════════════════════════════
# SHEET 2 — WASTE WARRIORS ENTRY POINT STRATEGY
# ══════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Waste Warriors — Entry Strategy")
ws2.sheet_view.showGridLines = False
ws2.freeze_panes = "A3"

make_title(ws2,
    "WASTE WARRIORS NGO — HOW TO APPROACH EACH COMPANY (ENTRY POINT STRATEGY)",
    10, FOREST_GRN)
make_header_row(ws2, [
    "Company", "Entry Platform / Route", "Priority",
    "Step-by-Step Approach", "Who to Contact / URL",
    "Pitch Angle (Why WW fits)", "Timeline",
    "Documents Needed", "CAF America Route?", "Expected Match Per Employee"
], color=MID_GREEN)

# ── Waste Warriors background box ─────────────────────────────
ws2.merge_cells("A2:J2")  # overwrite header row 2 — add a sub-banner
# Actually keep headers at row 2, add org info at row 3 as a merged banner
# Insert an info row after headers

# Org profile row
org_row = 3
ws2.merge_cells(f"A{org_row}:J{org_row}")
ws2[f"A{org_row}"] = (
    "WASTE WARRIORS PROFILE:  Registered Society | 12A ✓ | 80G ✓ | FCRA ✓ | "
    "Founded 2012 | Focus: Solid waste management, Himalayas (Himachal Pradesh & Uttarakhand) | "
    "Keeling Curve Prize 2023 (Top 10 global climate solutions) | "
    "Programs: Green Gurukul, Swachhata Ki Pathshala, Young Warriors Club, Community Activation | "
    "Contact: wastewarriors.org | CSR page: wastewarriors.org/csr-projects-in-india/"
)
ws2[f"A{org_row}"].font = Font(name="Calibri", bold=True, size=10, color=WHITE)
ws2[f"A{org_row}"].fill = hfill(MID_GREEN)
ws2[f"A{org_row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws2[f"A{org_row}"].border = brd
ws2.row_dimensions[org_row].height = 50

entry_data = [
    # ── PLATFORM: BENEVITY (covers 12+ companies at once) ──
    ("🌟 BENEVITY PLATFORM\n(One registration covers:\nMicrosoft, Google, Apple,\nIBM, NVIDIA, Cisco,\nBank of America, Nike,\nAdobe, Qualcomm, NetApp,\nIntel + more)",
     "Benevity Nonprofit Portal\ncauseshelp.benevity.org",
     "🔴 HIGHEST PRIORITY",
     "1. Go to causeshelp.benevity.org\n"
     "2. Click 'Register Your Nonprofit'\n"
     "3. Select India as country\n"
     "4. Search for 'Waste Warriors Society'\n"
     "5. Fill organization profile: mission, programs, impact metrics\n"
     "6. Upload: 80G certificate, 12A certificate, FCRA certificate, "
     "PAN card, bank account details (FCRA-designated account), "
     "society registration certificate\n"
     "7. Benevity reviews (2–4 weeks)\n"
     "8. Once approved → auto-visible to ALL Benevity client employees globally\n"
     "9. Employees of 1,000+ companies can then search, donate, and request matches",
     "causeshelp.benevity.org\nSupport: nonprofits@benevity.com\nPhone: 1-855-236-3848",
     "Environment & Climate: Keeling Curve Prize 2023 winner. "
     "Himalayan ecosystem protection — directly relevant to tech companies' "
     "sustainability/climate pledges (Google, Microsoft, Apple all have net-zero goals). "
     "Circular economy / waste reduction aligns with Nike, Cisco, Adobe ESG reports.",
     "Registration: 2–4 weeks\n"
     "First matches: 1–3 months post-registration",
     "• FCRA certificate + FCRA bank account details\n"
     "• 80G registration certificate\n"
     "• 12A registration certificate\n"
     "• Society registration certificate\n"
     "• PAN card\n"
     "• Audited financials (last 2 years)\n"
     "• Annual report / impact report\n"
     "• Board member list\n"
     "• Organization description (250 words)",
     "YES — US company donations to Indian NGOs via Benevity route through "
     "CAF America or Benevity's own global giving mechanism",
     "Varies by company: $10,000–$25,000 per active employee/year"),

    # ── PLATFORM: YOURCAUSE / NPOconnect ──
    ("🌟 YOURCAUSE / NPOconnect\n(Covers:\nPepsiCo, Dell, Wells Fargo,\nAmerican Express,\nSalesforce + more)",
     "YourCause NPOconnect\nyourcause.com/nonprofits\nNPOconnect portal",
     "🔴 HIGH PRIORITY",
     "1. Go to yourcause.com → Nonprofits section\n"
     "2. Register on NPOconnect\n"
     "3. Complete nonprofit profile with India registration details\n"
     "4. Upload FCRA, 80G, 12A certificates\n"
     "5. YourCause verifies via Blackbaud Giving Fund / CAF America\n"
     "6. Once listed → visible to all YourCause client employees\n"
     "NOTE: YourCause routes India donations via CAF America for FCRA compliance\n"
     "7. For American Express specifically: "
     "ALSO register directly on GlobalGiving (globalgiving.org) — "
     "AmEx uses GlobalGiving as their employee giving platform",
     "yourcause.com/nonprofits\nGlobalGiving: globalgiving.org/become-a-partner\n"
     "CAF America: cafamerica.org",
     "PepsiCo's pep+ sustainability strategy includes waste reduction — "
     "direct alignment. Dell's Planet program focuses on circular economy. "
     "Wells Fargo's environmental grants. AmEx's community impact "
     "in 30+ countries includes India.",
     "Registration: 2–3 weeks\n"
     "First donations: 1–2 months",
     "• FCRA certificate\n"
     "• 80G certificate\n"
     "• 12A certificate\n"
     "• CAF America eligibility confirmation\n"
     "• PAN card\n"
     "• Audited financials",
     "YES — YourCause uses CAF America for India donations",
     "$2,000–$10,000 per employee/year"),

    # ── CAF AMERICA — UNLOCK ALL US COMPANIES ──
    ("🌟 CAF AMERICA\n(Critical gateway for ALL\nUS company donations\nto Indian NGOs)",
     "CAF America\ncafamerica.org\nEquivalency Determination (ED)",
     "🔴 CRITICAL FIRST STEP",
     "1. Apply for CAF America 'Equivalency Determination' (ED)\n"
     "   This confirms Waste Warriors is equivalent to a US 501(c)(3) nonprofit\n"
     "2. Go to cafamerica.org → Services → Equivalency Determination\n"
     "3. Submit: 12A, 80G, FCRA, financials, governance docs, bylaws, "
     "program reports, tax ID (PAN)\n"
     "4. CAF America reviews (4–8 weeks), issues ED letter\n"
     "5. With ED letter: Waste Warriors becomes eligible to receive "
     "tax-deductible donations from ALL US corporations\n"
     "6. This unlocks: Microsoft, Google, Apple, Amazon, IBM, "
     "JPMorgan Chase, Goldman Sachs, Merck, Deloitte — ALL 32 companies researched\n"
     "7. Benevity and YourCause both use CAF America for India-based orgs",
     "cafamerica.org/services/equivalency-determination/\n"
     "Email: info@cafamerica.org\nPhone: +1 (202) 296-8641",
     "FCRA-registered Indian NGO with proven 10+ years of work, "
     "international recognition (Keeling Curve Prize), and clear impact metrics "
     "= ideal candidate for CAF America ED.",
     "Application: 4–8 weeks\n"
     "Cost: ~$1,500–$2,500 USD for ED",
     "• Articles of Association / Bylaws\n"
     "• Society registration certificate\n"
     "• FCRA registration (critical)\n"
     "• 12A & 80G certificates\n"
     "• PAN card\n"
     "• 3 years audited financials\n"
     "• Annual reports\n"
     "• Board member list with bios\n"
     "• Program descriptions\n"
     "• Recent IRS-equivalent filings",
     "THIS IS THE CAF AMERICA STEP",
     "Unlocks all US companies = up to $30,000/employee/year (Merck/Deloitte)"),

    # ── GLOBALGIVING ──
    ("🌟 GLOBALGIVING\n(Direct entry for:\nAmerican Express,\nand 600+ corporate partners)",
     "GlobalGiving Partner\nglobalgiving.org/become-a-partner",
     "🟡 HIGH VALUE",
     "1. Go to globalgiving.org/become-a-partner\n"
     "2. Create organization account\n"
     "3. Submit project proposal (project title, description, "
     "fundraising goal, timeline, impact metrics)\n"
     "4. Pass GlobalGiving vetting process (4–6 weeks)\n"
     "5. Run an 'Open Challenge' fundraising campaign to earn full partner status\n"
     "   (raise $5,000 from 40+ donors in the challenge period)\n"
     "6. Once full partner: auto-eligible for American Express Give2Gether "
     "employee matching ($10,000/employee/year)\n"
     "7. Also becomes visible to 600+ corporate partners on GlobalGiving",
     "globalgiving.org/become-a-partner\n"
     "American Express Give2Gether: globalgiving.org/amex-give2gether/",
     "GlobalGiving's India program is extensive. "
     "American Express specifically partners with GlobalGiving for employee giving "
     "in 30+ countries. Waste Warriors' mountain ecosystem work has "
     "strong international donor appeal.",
     "Vetting + challenge: 6–10 weeks\n"
     "Challenge fundraising: 3–4 weeks",
     "• FCRA certificate\n"
     "• 80G certificate\n"
     "• Organization registration\n"
     "• Project description\n"
     "• Impact metrics\n"
     "• Beneficiary photos/stories",
     "Indirect — GlobalGiving handles compliance",
     "Up to $10,000/employee/year (American Express)"),

    # ── DIRECT: COMPANIES WITH OWN PORTALS ──
    ("DIRECT APPROACH:\nJPMorgan Chase\nGoldman Sachs\nJ&J / Pfizer / Merck\nDeloitte",
     "CyberGrants (JPMorgan, Merck)\nEasyMatch (Goldman)\nDirect CSR portals (J&J, Pfizer)\nDeloitte Foundation",
     "🟡 MEDIUM–HIGH PRIORITY",
     "For each company:\n"
     "1. JPMorgan Chase: Register on CyberGrants at cybergrants.com → "
     "submit India NGO registration details\n"
     "2. Goldman Sachs: Contact gs@easymatch.com or easymatch.com → "
     "register as eligible organization\n"
     "3. Johnson & Johnson: Contact J&J Foundation directly → "
     "jnj.com/about-jnj/global-impact → 'For NGO Partners' section\n"
     "4. Pfizer: Register at pfizerplus.com and contact "
     "Pfizer Foundation India team\n"
     "5. Merck: Submit via CyberGrants (cybergrants.com/merck/merckgivesback); "
     "contact Partnership for Giving team\n"
     "6. Deloitte: Contact Deloitte Foundation India desk → "
     "Position as education/skills partner (fits Deloitte's STEM/WorldClass focus)",
     "JPMorgan: cybergrants.com/jpmc/giving\n"
     "Goldman: gs@easymatch.com\n"
     "J&J: jnj.com/about-jnj/global-impact\n"
     "Pfizer: pfizerplus.com/gi/give_back.aspx\n"
     "Merck: cybergrants.com/merck/merckgivesback\n"
     "Deloitte: deloitte.com/global/en/about/social-impact",
     "J&J / Pfizer / Merck: Waste in healthcare supply chains, "
     "plastic waste in rural India — direct public health angle.\n"
     "JPMorgan / Goldman / Deloitte: Employee engagement in India offices "
     "(they all have large India teams who can volunteer with Waste Warriors).",
     "Registration: 3–6 weeks per company\n"
     "First matches: 1–3 months",
     "• CAF America ED letter (critical for US donations)\n"
     "• FCRA certificate\n"
     "• 80G, 12A certificates\n"
     "• Annual report\n"
     "• Impact metrics / beneficiary data",
     "YES — required for US corporate donations to Indian NGOs",
     "Up to $30,000/employee (Merck/Deloitte)"),

    # ── INDIA-SPECIFIC ──
    ("INDIA DIRECT:\nHCL Tech\nWipro\nMicrosoft India\nMahindra\nHDFC Bank",
     "HCL Foundation Grant portal\nWipro Cares portal\nMicrosoft India NGO form\nMahindra Foundation\nGive India (HDFC route)",
     "🟢 IMMEDIATE OPPORTUNITY",
     "1. HCL Tech Grant: Apply at hclfoundation.org → HCL Tech Grant → "
     "Waste/Environment category NOW open (FY26: added Biodiversity/Water category)\n"
     "2. Wipro Cares: Contact Wiprofoundation.org/wipro-cares/ → "
     "Wipro employees contribute ₹1/2/5/day + Wipro matches 1:1; "
     "NGOs apply to be listed as partner organizations\n"
     "3. Microsoft India: Contact Microsoft India CSR team → "
     "microsoft.com/en-in/about/employee-giving → "
     "Apply to be listed as eligible NGO for Oct giving campaign\n"
     "4. Mahindra Foundation: Apply via Mahindra Esops NGO partnership program; "
     "contact Mahindra Foundation at mahindra.com/sustainability\n"
     "5. HDFC Bank: Register on GiveIndia platform (giveindia.org) → "
     "HDFC Bank routes employee payroll giving through GiveIndia → "
     "auto-eligible for bank's 1:1 match",
     "HCL Foundation: hclfoundation.org/hcltech-grant\n"
     "Wipro Cares: wiprofoundation.org/wipro-cares\n"
     "Microsoft India: microsoft.com/en-in/about/employee-giving\n"
     "Mahindra: mahindrafoundation.org\n"
     "HDFC + GiveIndia: giveindia.org/ngo-onboarding\n"
     "Give India Payroll: indiadonates.org",
     "HCL Tech Grant (FY26): Biodiversity/Water = EXACT FIT for Himalayan ecosystem work.\n"
     "Wipro: Ecology focus = waste management aligns perfectly.\n"
     "HDFC: Environment category available for NGO selection.\n"
     "NO CAF AMERICA NEEDED for India-based giving — all under Companies Act 2013 CSR",
     "HCL Grant: Apply by deadline (check hclfoundation.org)\n"
     "GiveIndia: 2–4 weeks registration\n"
     "Microsoft India: October campaign (apply by August)",
     "India programs:\n"
     "• 12A certificate\n"
     "• 80G certificate\n"
     "• CSR-1 registration (Ministry of Corporate Affairs)\n"
     "• PAN card\n"
     "• Society registration\n"
     "• Annual report\n"
     "• Project proposal (for HCL Grant)\n"
     "• Impact data",
     "NO — India donations don't require FCRA/CAF America",
     "HCL Grant: up to ₹50–75 lakhs per NGO\n"
     "Wipro Cares: ₹1–5 lakhs/year from payroll\n"
     "HDFC: 1:1 match on employee donations"),

    # ── SALESFORCE SPECIFIC ──
    ("Salesforce\n(Philanthropy Cloud +\nPledge 1% movement)",
     "Salesforce Philanthropy Cloud\nPledge 1% movement",
     "🟡 MEDIUM PRIORITY",
     "1. Register on Salesforce Philanthropy Cloud: "
     "salesforce.com/company/philanthropy/ → 'For Nonprofits'\n"
     "2. Apply to be listed as a Pledge 1% partner nonprofit\n"
     "3. Attend Salesforce.org's nonprofit onboarding webinars\n"
     "4. Once listed: 76,000 Salesforce employees can donate ($50 credit each)\n"
     "   + additional personal donations with $10,000 match\n"
     "5. Consider applying for Salesforce.org tech grant "
     "(free CRM software for nonprofits = operational benefit)",
     "salesforce.com/company/philanthropy/\n"
     "salesforce.org (tech grants for nonprofits)\n"
     "pledge1percent.org",
     "Salesforce's employee-led giving model means employees in Salesforce India "
     "(Hyderabad, Bangalore) can donate locally. "
     "Sustainability/environment is a key pillar of Salesforce's Stakeholder 360.",
     "Philanthropy Cloud registration: 3–5 weeks\n"
     "Pledge 1% partner: 2–4 weeks",
     "• 80G, 12A, FCRA\n"
     "• Salesforce.org may require US 501(c)(3) equivalent (CAF America ED)\n"
     "• Annual report\n"
     "• Impact metrics",
     "YES for US employee match portion",
     "Up to $10,000/employee + $50 default donation credit"),

    # ── STRYKER ──
    ("Stryker\n(Impact platform +\nMedical waste angle)",
     "Stryker Impact Platform\n(direct registration)",
     "🟡 MEDIUM PRIORITY",
     "1. Go to stryker.com/us/en/about/corporate-responsibility/giving-back.html\n"
     "2. Contact Stryker's Corporate Responsibility team\n"
     "3. Pitch: Medical device/surgical waste management — "
     "Stryker is a medical device company; medical waste in Himalayas "
     "(near hospitals, healthcare camps) is a very specific and resonant angle\n"
     "4. Request listing on Stryker Impact platform\n"
     "5. Employees can then donate and trigger Stryker's 1:1 match",
     "stryker.com/us/en/about/corporate-responsibility/giving-back.html\n"
     "CSR contact: via stryker.com/contact",
     "Medical waste angle: Stryker makes surgical/medical devices. "
     "Waste Warriors handles medical waste in mountain communities — "
     "unique and very relevant pitch. "
     "2024: 11,000+ Stryker employee donors to 5,700 nonprofits globally.",
     "Outreach + listing: 4–8 weeks",
     "• FCRA, 80G, 12A\n"
     "• CAF America ED letter\n"
     "• Medical waste program documentation",
     "YES",
     "1:1 match, elevated rates on special campaigns"),
]

for row_idx, row_data in enumerate(entry_data, start=4):
    pri = row_data[2]
    if "HIGHEST" in pri or "CRITICAL" in pri: fill = "FFE0E0"   # light red/urgent
    elif "HIGH" in pri and "🔴" in pri:        fill = LIGHT_AMBER
    elif "IMMEDIATE" in pri:                   fill = LIGHT_GREEN
    else:                                      fill = ALT[(row_idx-4) % 2]
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill = hfill(fill)
        cell.font = hfont(DARK_TEXT, bold=(col_idx in [1,3]), sz=9)
        cell.alignment = walign("left", "top")
        cell.border = brd
    ws2.row_dimensions[row_idx].height = 180

col_w2 = [22, 26, 18, 65, 38, 50, 22, 40, 22, 30]
for i, w in enumerate(col_w2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════════
# SHEET 3 — INDIA COMPANIES
# ══════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("India Companies")
ws3.sheet_view.showGridLines = False
ws3.freeze_panes = "A3"

make_title(ws3,
    "INDIA CORPORATE EMPLOYEE GIVING PROGRAMS — 8 COMPANIES (WITH SOURCES)",
    10, "0D3B2E")
make_header_row(ws3, [
    "Company", "Sector", "Program Name", "Type of Giving",
    "Match / Contribution Details", "Platform / Portal",
    "Eligible NGO Categories", "How Waste Warriors Can Apply",
    "Regulatory Framework", "Sources"
], color=FOREST_GRN)

india_data = [
    ("Microsoft India",
     "Technology",
     "Microsoft India Employee Giving Program",
     "Payroll giving + Microsoft match",
     "Match cap:\n• Engineering team: INR 2,50,000/employee/year\n"
     "• Other employees: INR 50,000/employee/year\n"
     "Equal 1:1 match by Microsoft India\n"
     "3 days paid Volunteer Time Off/year\n"
     "Annual giving campaign: October\n"
     "Total contributed: ₹19.32 crores to 144 NGOs since inception",
     "Microsoft India internal portal\n"
     "(microsoft.com/en-in/about/employee-giving)",
     "Registered NGOs with 80G status; "
     "focus on education, community development, environment",
     "1. Contact Microsoft India CSR team at microsoft.com/en-in/about/employee-giving\n"
     "2. Apply to be listed as eligible NGO for October campaign\n"
     "3. Submit: 80G, 12A, registration, annual report\n"
     "4. Once listed: Microsoft India employees can donate + get matched\n"
     "Best time to apply: July–August (before October campaign)",
     "Companies Act 2013, Section 135 (CSR); "
     "Income Tax Act Section 80G (donor deduction); "
     "No FCRA needed for India domestic giving",
     "https://www.microsoft.com/en-in/about/employee-giving"),

    ("HCL Technologies",
     "Technology / IT Services",
     "Power of One + HCL Tech Grant",
     "Payroll giving (₹1/2/5/day) + Corporate CSR grant",
     "Power of One: employees contribute ₹1, ₹2, or ₹5 per day via payroll\n"
     "36,000+ employees enrolled\n"
     "Funds go to scholarships and community programs\n\n"
     "HCL Tech Grant (flagship): \n"
     "• FY26 Annual outlay: ₹24 CRORE to NGOs\n"
     "• Categories: Education, Healthcare, Biodiversity (NEW FY26), Water (NEW FY26)\n"
     "• Grants of ₹50–75 lakhs per selected NGO per year\n"
     "• Multi-year grants available",
     "HCL Foundation portal\nhclfoundation.org/hcltech-grant\n"
     "Annual application cycle (usually Dec–Feb)",
     "Rural development, education, healthcare, biodiversity, "
     "water conservation, environment",
     "IMMEDIATE OPPORTUNITY:\n"
     "1. Apply for HCL Tech Grant FY26 at hclfoundation.org/hcltech-grant\n"
     "2. Apply under: Biodiversity AND Water categories (both new in FY26)\n"
     "3. Himalayan ecosystem / waste management in rivers = perfect fit\n"
     "4. Grants: ₹50–75 lakhs/year; multi-year possible\n"
     "5. Deadline: check hclfoundation.org for current cycle",
     "Companies Act 2013, Section 135; "
     "80G eligible; no FCRA needed",
     "https://www.hclfoundation.org/hcltech-grant | "
     "https://www.hcltech.com/corporate-social-responsibility"),

    ("Wipro Limited",
     "Technology / IT Services",
     "Wipro Cares (Employee-Matched Giving)",
     "Employee payroll contribution + 1:1 Wipro match",
     "1:1 MATCH by Wipro on employee contributions\n"
     "45,000+ employees enrolled in Wipro Cares\n"
     "3 funding tracks:\n"
     "1. Employee contributions (matched 1:1 by Wipro)\n"
     "2. Direct Wipro CSR budget allocation\n"
     "3. Donations from Wipro group companies\n\n"
     "Focus areas: Education, Health, Ecology, Disaster Response",
     "Wipro Foundation portal\nwiprofoundation.org/wipro-cares/\n"
     "FAQs: wiprofoundation.org/faqs/",
     "Education, health, ecology, environment, disaster response; "
     "NGOs must be registered 80G-certified in India",
     "1. Apply as Wipro Cares partner NGO at wiprofoundation.org/wipro-cares/\n"
     "2. Focus on Ecology/Environment track\n"
     "3. Submit: 80G, 12A, registration, 2-year financials, program details\n"
     "4. Once listed: 45,000+ employees see you and can direct contributions\n"
     "5. Wipro matches 1:1 — potential ₹5–15 lakhs/year if well-promoted",
     "Companies Act 2013, Section 135; "
     "Wipro Foundation is the CSR implementation arm",
     "https://www.wiprofoundation.org/wipro-cares/ | "
     "https://www.wipro.com/investors/corporate-governance/corporate-social-responsibility/"),

    ("Tata Consultancy Services (TCS)",
     "Technology / IT Services",
     "HOPE (Hours of Purpose by Employees) + TECH4HOPE",
     "Employee volunteering (no cash match confirmed); "
     "TECH4HOPE: in-kind tech pro bono for NGOs",
     "HOPE Program:\n"
     "FY2025: 8.9 MILLION volunteer hours (30% growth vs FY24)\n"
     "86.5% of TCS volunteering is through HOPE\n"
     "146,000+ TCSers volunteered\n"
     "Focus: education, healthcare, environment, digital literacy\n\n"
     "TECH4HOPE Pro Bono:\n"
     "TCS employees deliver in-kind tech solutions to NGOs\n"
     "In North America: $1.4M in-kind services donated\n"
     "No cash matching confirmed publicly",
     "TCS CSR portal: tcs.com/who-we-are/corporate-social-responsibility\n"
     "HOPE NGO partnerships: apply via TCS Foundation\n"
     "TCS Foundation: tcsfoundation.in",
     "Education, environment, health, digital inclusion; "
     "NGOs applying as HOPE project partners",
     "1. Apply as HOPE program NGO partner via TCS Foundation (tcsfoundation.in)\n"
     "2. Apply for TECH4HOPE: request TCS volunteer teams to help with "
     "tech needs (database, website, waste tracking systems)\n"
     "3. TCS India employees can volunteer at Waste Warriors sites\n"
     "4. Annual CSR budget: ₹1,038 crores; apply for direct project grants\n"
     "5. Contact TCS Foundation via tcsfoundation.in",
     "Companies Act 2013, Section 135; "
     "TCS CSR spend: ₹1,038 crores FY25",
     "https://www.tcs.com/who-we-are/corporate-social-responsibility | "
     "https://tcsfoundation.in | "
     "https://www.tcs.com/who-we-are/corporate-social-responsibility/article/building-hope-tcs-employees-contributed-millions-hours-of-purpose"),

    ("Infosys / Infosys Foundation",
     "Technology / IT Services",
     "Infosys Foundation Grants + Employee Volunteering",
     "Corporate CSR grants via Infosys Foundation; "
     "employee volunteering (no formal cash match confirmed)",
     "Infosys Foundation:\n"
     "• One of India's largest corporate foundations\n"
     "• Focus: education, rural development, healthcare, arts, destitute care\n"
     "• CSR spend FY2024-25: ₹18+ crores (one quarter alone)\n"
     "• Employee volunteering: 150,000+ hours, 48,000+ volunteers across 21 DCs\n\n"
     "Application process: Infosys Foundation invites project proposals from NGOs\n"
     "Focus on programs with measurable community impact",
     "Infosys Foundation: infosys.org\n"
     "CSR reporting: infosys.com/investors/corporate-governance/social-responsibility.html\n"
     "Foundation report FY2024-25: infosys.org/infosys-foundation/about/reports/",
     "Education, rural development, healthcare, ecology; "
     "Projects in Infosys operating locations (Bangalore, Mysore, Pune, etc.)",
     "1. Contact Infosys Foundation at infosys.org → 'Contact Us'\n"
     "2. Submit project proposal focusing on: "
     "waste management education for rural/tribal communities\n"
     "3. Align with Infosys' environmental goals (Infosys is carbon neutral)\n"
     "4. Approach Infosys India offices near Waste Warriors' geography "
     "(Delhi, Chandigarh offices closest to Himachal/Uttarakhand)\n"
     "5. Request employee volunteer engagement in addition to grant",
     "Companies Act 2013, Section 135; "
     "Infosys Foundation is a public charitable trust",
     "https://www.infosys.org | "
     "https://www.infosys.org/infosys-foundation/about/reports/documents/infosys-foundation-report-2024-25.pdf"),

    ("Mahindra Group",
     "Diversified Conglomerate",
     "Employee Social Options (Esops) + Mahindra Foundation",
     "Employee volunteering + community project grants; "
     "CSR budget from group companies",
     "Esops Program:\n"
     "Running since 2005; employees volunteer in community projects\n"
     "Activities: NGO partnerships, village development, tree plantation (Hariyali)\n"
     "Hariyali: 1.5 million+ trees planted/year\n\n"
     "Mahindra Foundation:\n"
     "• Manages CSR across Mahindra Group companies\n"
     "• Focus: education, health, environment, livelihoods\n"
     "• Partners with NGOs for multi-year programs",
     "Mahindra Foundation: mahindrafoundation.org\n"
     "CSR: mahindra.com/en/sustainability\n"
     "Esops: via Mahindra HR portal",
     "Education, environment, rural livelihoods, health, "
     "women empowerment; NGOs working in Mahindra operating locations",
     "1. Apply to Mahindra Foundation as implementation partner NGO\n"
     "2. Pitch: Waste Warriors' eco-tourism/trekking route waste management "
     "aligns with Mahindra Holidays (Mahindra Group company)\n"
     "3. Request Esops employee volunteer program engagement\n"
     "4. Mahindra Foundation grants: submit letter of intent to "
     "mahindrafoundation.org → apply for community project funding",
     "Companies Act 2013, Section 135; "
     "Mahindra Group CSR commitment is large across all group companies",
     "https://mahindrafoundation.org | "
     "https://ngofundings.org/csr/mahindra-mahindra-ltd-csr-activities/"),

    ("HDFC Bank",
     "Banking / Financial Services",
     "HDFC Bank Payroll Giving + Parivartan CSR",
     "Payroll giving via GiveIndia platform + 1:1 bank match",
     "Payroll Giving:\n"
     "• Platform: GiveIndia (giveindia.org)\n"
     "• Employees donate monthly via payroll to chosen NGO\n"
     "• HDFC Bank matches employee donations (1:1 confirmed)\n"
     "• 1,500+ employees have donated; 190+ NGOs listed\n"
     "• Running for 7 years\n\n"
     "Parivartan (direct CSR):\n"
     "• Rural development, education, skill development, livelihood\n"
     "• Healthcare & hygiene\n"
     "• Financial literacy\n"
     "• Environment category available",
     "GiveIndia: giveindia.org/ngo-onboarding\n"
     "HDFC CSR: v.hdfc.bank.in/csr\n"
     "Parivartan: hdfc.bank.in/personal/about-us/csr",
     "Environment, education, healthcare, rural development; "
     "NGOs registered on GiveIndia platform",
     "1. Register on GiveIndia platform at giveindia.org/ngo-onboarding\n"
     "2. Once listed: HDFC Bank employees can choose Waste Warriors for payroll giving\n"
     "3. HDFC Bank matches 1:1 → double impact\n"
     "4. Also apply for Parivartan CSR grants directly via HDFC Bank CSR team\n"
     "5. Submit: 80G, 12A, registration, 2-year financials, impact report\n"
     "Note: Environment category is strong fit for waste management",
     "Companies Act 2013, Section 135; "
     "Income Tax Act 80G; "
     "No FCRA needed (domestic giving)",
     "https://v.hdfc.bank.in/content/dam/hdfc-aem-microsites/csr/pdfs/impact-reports/P0040-Payroll-Project-FDP-Impact-Assessment-Report-FY22-23.pdf | "
     "https://giveindia.org"),

    ("Accenture India",
     "Consulting / Technology Services",
     "Accenture Skills to Succeed + Employee Giving (Benevity)",
     "Employee matching gifts + corporate skills-based volunteering",
     "Matching Gift Program:\n"
     "• Uses Benevity globally (India employees included)\n"
     "• Match ratio and cap: internal (not publicly confirmed)\n"
     "• Estimated: 1:1, up to $2,500–$5,000\n\n"
     "Skills to Succeed:\n"
     "• Equipping 3M+ people/year with job skills globally\n"
     "• India: 2M+ young learners skilled since 2021 (with UNICEF Gen Unlimited)\n"
     "• Accenture India employees volunteer professional skills to NGOs\n\n"
     "FY2024: ~5 million people reached globally with partners",
     "Benevity (global platform)\n"
     "Accenture Corporate Citizenship: accenture.com/us-en/about/corporate-citizenship\n"
     "India CSR contact: accenture.com/in-en",
     "Skills development, education, environment, community development; "
     "NGOs receiving both cash and pro bono skills support",
     "1. Register on Benevity (causeshelp.benevity.org) → "
     "Accenture India employees can then donate and request matches\n"
     "2. Approach Accenture India CSR team for skills-based volunteering:\n"
     "   - Request waste management technology volunteers\n"
     "   - Request data analytics support for impact measurement\n"
     "3. Contact Accenture India corporate citizenship team via LinkedIn\n"
     "4. Annual Accenture India CSR budget is substantial (large India presence)",
     "Companies Act 2013, Section 135 (India CSR); "
     "Also subject to Accenture global giving policies via Benevity for international matching",
     "https://www.accenture.com/us-en/about/corporate-citizenship | "
     "https://doublethedonation.com/matching-gifts/accenture"),
]

for row_idx, row_data in enumerate(india_data, start=3):
    fill = ALT[(row_idx - 3) % 2]
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws3.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill = hfill(fill)
        cell.font = hfont(DARK_TEXT, bold=(col_idx == 1), sz=9)
        cell.alignment = walign("left", "top")
        cell.border = brd
    ws3.row_dimensions[row_idx].height = 160

col_w3 = [18, 18, 28, 22, 42, 28, 30, 55, 28, 55]
for i, w in enumerate(col_w3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════════
# SHEET 4 — MASTER ACTION PLAN (PRIORITIZED ROADMAP)
# ══════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Action Plan — Waste Warriors")
ws4.sheet_view.showGridLines = False
ws4.freeze_panes = "A3"

make_title(ws4,
    "WASTE WARRIORS — PRIORITIZED 90-DAY ACTION PLAN TO ACCESS CORPORATE GIVING",
    8, DARK_NAVY)
make_header_row(ws4, [
    "Phase", "Action", "Platform / Contact",
    "Companies Unlocked", "Timeline",
    "Cost / Effort", "Expected Annual Potential",
    "Key Documents"
])

action_plan = [
    ("PHASE 1\nWEEK 1–2\n🔴 DO FIRST",
     "Apply to CAF America for Equivalency Determination (ED).\n"
     "This is the MASTER KEY that unlocks ALL 28 US companies.\n"
     "Without this, US companies legally cannot route donations to Waste Warriors.",
     "CAF America\ncafamerica.org/services/equivalency-determination/\n"
     "Email: info@cafamerica.org\nPhone: +1 (202) 296-8641",
     "ALL 28 US-based companies researched:\nMicrosoft, Google, Apple, Amazon,\n"
     "IBM, NVIDIA, Salesforce, HP, TI,\nAmerican Express, Mastercard, Intel,\n"
     "NetApp, Stryker, Apollo, + all 16 new cos.",
     "Apply in Week 1\nED issued in 4–8 weeks",
     "Cost: ~$1,500–$2,500 USD\n"
     "Effort: High (document gathering)\n"
     "ONE-TIME cost; ED valid for years",
     "UNLOCKS $10,000–$30,000/employee/year potential across all US companies",
     "• Society registration certificate\n"
     "• FCRA certificate + FCRA bank account\n"
     "• 12A & 80G certificates\n"
     "• PAN card\n"
     "• Bylaws/Articles of Association\n"
     "• 3 years audited financials\n"
     "• Annual reports (last 2 years)\n"
     "• Board member list with bios\n"
     "• Program descriptions with impact data"),

    ("PHASE 1\nWEEK 1–2\n🔴 SIMULTANEOUS",
     "Register on Benevity Nonprofit Portal.\n"
     "ONE registration → visible to 1,000+ corporate clients globally.\n"
     "12 of the 32 companies researched use Benevity.",
     "Benevity Nonprofit Portal\ncauseshelp.benevity.org\n"
     "Email: nonprofits@benevity.com\nPhone: 1-855-236-3848",
     "Microsoft, Google, Apple, IBM,\nNVIDIA, Cisco, Bank of America,\n"
     "Nike, Adobe, Qualcomm, NetApp, Intel\n"
     "(12 companies, 1 registration)",
     "Register Week 1\nApproval: 2–4 weeks\n"
     "First matches: Month 2",
     "Cost: FREE\nEffort: Medium (profile setup + document upload)\n"
     "HIGHEST ROI of any single action",
     "12 companies × avg $12,000 cap =\npotential $144,000/year if well-promoted",
     "Same as CAF America docs above\n"
     "(Benevity uses these for India org verification)"),

    ("PHASE 1\nWEEK 2–3\n🔴 SIMULTANEOUS",
     "Register on YourCause NPOconnect.\n"
     "Covers PepsiCo, Dell, Wells Fargo, American Express.\n"
     "Also register on GlobalGiving to unlock American Express Give2Gether.",
     "YourCause: yourcause.com/nonprofits\n"
     "GlobalGiving: globalgiving.org/become-a-partner\n"
     "CAF America routes India donations on YourCause",
     "PepsiCo, Dell Technologies,\nWells Fargo, American Express\n"
     "(4 additional companies)",
     "YourCause: 2–3 weeks\n"
     "GlobalGiving: 6–10 weeks (includes fundraising challenge)",
     "Cost: FREE (YourCause)\n"
     "GlobalGiving: FREE to list; challenge requires raising $5,000 from 40+ donors\n"
     "Effort: Medium",
     "PepsiCo: $10,000/employee\n"
     "AmEx: $10,000/employee\n"
     "Dell: $10,000/employee",
     "FCRA certificate, 80G, 12A,\nannual report, project photos/stories"),

    ("PHASE 1\nWEEK 2–3\n🟢 INDIA — IMMEDIATE",
     "Apply for HCL Tech Grant FY26.\n"
     "NEW categories FY26: Biodiversity + Water = PERFECT FIT.\n"
     "Grant: ₹50–75 lakhs. Multi-year. No FCRA needed.",
     "HCL Foundation\nhclfoundation.org/hcltech-grant\n"
     "Contact: hclfoundation.org/contact",
     "HCL Technologies\n(India domestic — no CAF America needed)",
     "Check deadline NOW at hclfoundation.org\n"
     "Typically applies Dec–Feb cycle",
     "Cost: FREE to apply\nEffort: Medium (detailed project proposal required)\n"
     "This is a GRANT, not a matching program",
     "₹50–75 lakhs per year (grant)\n"
     "Multi-year grants possible",
     "• 12A, 80G certificates\n"
     "• CSR-1 registration (MCA portal)\n"
     "• Society registration\n"
     "• Detailed project proposal\n"
     "• Budget breakdown\n"
     "• Impact metrics\n"
     "• 2-year financials"),

    ("PHASE 2\nWEEK 3–5\n🟡 INDIA PLATFORMS",
     "Register on GiveIndia platform.\n"
     "HDFC Bank routes employee payroll giving (with 1:1 match) through GiveIndia.\n"
     "Also apply as Wipro Cares partner NGO (1:1 match, 45,000+ employees).",
     "GiveIndia: giveindia.org/ngo-onboarding\n"
     "Wipro Cares: wiprofoundation.org/wipro-cares/\n"
     "Concern India: concernindiafoundation.org/payrollgiving",
     "HDFC Bank (payroll + 1:1 match)\n"
     "Wipro (1:1 match, 45,000 employees)\n"
     "+ any other GiveIndia corporate partner",
     "GiveIndia: 2–4 weeks\n"
     "Wipro Cares: 4–8 weeks",
     "Cost: FREE\nEffort: Low–Medium",
     "HDFC + Wipro combined:\n"
     "₹5–15 lakhs/year in matched employee giving",
     "12A, 80G, CSR-1 (MCA), PAN,\nregistration, financials, impact report"),

    ("PHASE 2\nWEEK 4–6\n🟡 DIRECT OUTREACH",
     "Direct email/LinkedIn outreach to CSR/Foundation teams at:\n"
     "Goldman Sachs, J&J, Merck, Deloitte, Stryker.\n"
     "These companies have OWN portals (not Benevity/YourCause).\n"
     "High caps ($20,000–$32,500/employee) make them priority.",
     "Goldman: gs@easymatch.com\n"
     "J&J: jnj.com/about-jnj/global-impact\n"
     "Merck: CyberGrants + Partnership for Giving team\n"
     "Deloitte: Deloitte Foundation India desk\n"
     "Stryker: stryker.com CSR team",
     "Goldman Sachs ($20K cap)\n"
     "Johnson & Johnson ($20K match, 2:1)\n"
     "Merck ($30K cap)\n"
     "Deloitte ($32.5K cap)\n"
     "Stryker (medical waste angle)",
     "Outreach: Week 4–6\n"
     "Listing approval: 4–12 weeks per company",
     "Cost: FREE (time investment)\n"
     "Effort: High (personalized pitches per company)\n"
     "CAF America ED letter required",
     "5 companies × avg $20K cap =\n"
     "potential $100,000+/year from these alone",
     "CAF America ED letter (essential)\n"
     "Tailored 1-page pitch deck per company\n"
     "Impact report with metrics"),

    ("PHASE 3\nWEEK 6–12\n🟢 AWARENESS CAMPAIGN",
     "Launch internal awareness campaign:\n"
     "Email/social posts to Waste Warriors donors asking them to:\n"
     "1. Check if employer matches (use doublethedonation.com)\n"
     "2. Submit matching gift requests for past/current donations\n"
     "3. Share with colleagues at matching companies\n\n"
     "Add matching gift widget to wastewarriors.org donation page.",
     "Double the Donation: doublethedonation.com (add widget to website)\n"
     "Email templates: vsuw.org/wp-content/uploads/Campaign-Email-Templates-1.pdf\n"
     "Matching gift reminder at 30/60/90 days post-donation",
     "ANY company whose employees already donate to Waste Warriors\n"
     "Industry data: 84% of donors would submit matches IF REMINDED",
     "Week 6 onwards\n"
     "December = peak: all year-end deadlines",
     "Cost: FREE (Double the Donation widget: ~$499/year)\n"
     "Effort: Low–Medium\n"
     "84% match submission rate when reminded",
     "Recovery of $4–7B in unclaimed matches industry-wide annually;\n"
     "nonprofits reminding donors recover 40–60% of eligible matches",
     "No additional docs needed\n"
     "Just need to be registered on Benevity/YourCause first"),

    ("PHASE 3\nWEEK 8–12\n🟡 INDIA TECH FIRMS",
     "Apply to TCS Foundation (HOPE program partnership),\n"
     "Infosys Foundation (project proposal),\n"
     "Microsoft India (October campaign),\n"
     "Mahindra Foundation (Esops NGO partner),\n"
     "Accenture India (skills volunteering + Benevity).",
     "TCS Foundation: tcsfoundation.in\n"
     "Infosys Foundation: infosys.org\n"
     "Microsoft India: microsoft.com/en-in/about/employee-giving\n"
     "Mahindra: mahindrafoundation.org\n"
     "Accenture India: accenture.com/in-en",
     "TCS (8.9M volunteer hours potential)\n"
     "Infosys Foundation (grant + volunteers)\n"
     "Microsoft India (Oct campaign)\n"
     "Mahindra (Hariyali/environment)\n"
     "Accenture India (skills + matching via Benevity)",
     "Apply: Month 2–3\n"
     "Microsoft India: Apply Aug for Oct campaign\n"
     "TCS/Infosys: Year-round",
     "Cost: FREE\nEffort: Medium (multiple applications)",
     "Combined India potential:\n"
     "₹25–50 lakhs/year in grants + employee giving",
     "12A, 80G, CSR-1,\nproject proposals per company,\nimpact metrics"),
]

phase_colors = {
    "PHASE 1": "FFD6D6",   # light red — urgent
    "PHASE 2": LIGHT_AMBER,
    "PHASE 3": LIGHT_GREEN,
}

for row_idx, row_data in enumerate(action_plan, start=3):
    phase_text = row_data[0]
    if "PHASE 1" in phase_text:   fill = "FFD6D6"
    elif "PHASE 2" in phase_text: fill = LIGHT_AMBER
    else:                          fill = LIGHT_GREEN

    for col_idx, value in enumerate(row_data, start=1):
        cell = ws4.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill = hfill(fill)
        cell.font = hfont(DARK_TEXT, bold=(col_idx in [1,2]), sz=9)
        cell.alignment = walign("left", "top")
        cell.border = brd
    ws4.row_dimensions[row_idx].height = 160

col_w4 = [18, 55, 38, 32, 18, 25, 32, 40]
for i, w in enumerate(col_w4, 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

# ── Legend + Summary box ──
summary_row = len(action_plan) + 4
ws4.merge_cells(f"A{summary_row}:H{summary_row}")
ws4[f"A{summary_row}"] = "TOTAL POTENTIAL: If Waste Warriors completes all Phase 1–3 actions, "
ws4[f"A{summary_row}"].value = (
    "TOTAL ANNUAL POTENTIAL:  "
    "India programs: ₹75 lakhs–1.5 crore/year in grants + matched giving  |  "
    "US companies via Benevity/YourCause (10 active employee donors/company × $500 avg): "
    "$50,000–$200,000 USD/year  |  "
    "Deloitte/Merck/Goldman high-cap programs: additional $50,000–$100,000 USD/year  |  "
    "KEY INSIGHT: Registering on Benevity alone (FREE) makes Waste Warriors discoverable "
    "to 1,000+ corporations and their combined 30M+ employees worldwide."
)
ws4[f"A{summary_row}"].font = Font(name="Calibri", bold=True, size=10, color=WHITE)
ws4[f"A{summary_row}"].fill = hfill(DARK_NAVY)
ws4[f"A{summary_row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws4[f"A{summary_row}"].border = brd
ws4.row_dimensions[summary_row].height = 60

# ══════════════════════════════════════════════════════════════════
out = "/home/user/jcode/Waste_Warriors_Corporate_Giving_Strategy.xlsx"
wb.save(out)
print(f"Saved: {out}")
