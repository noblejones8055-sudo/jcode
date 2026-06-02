import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── palette ──────────────────────────────────────────────────────
DARK_NAVY   = "1B2A4A"
MID_BLUE    = "2E5CA8"
LIGHT_BLUE  = "D6E4F7"
DEEP_GREEN  = "1A5C38"
MID_GREEN   = "2E7D52"
WHITE       = "FFFFFF"
LIGHT_GREY  = "F2F2F2"
MED_GREY    = "CCCCCC"
DARK_TEXT   = "1A1A2E"

def hfill(c): return PatternFill("solid", fgColor=c)
def hfont(c=DARK_TEXT, bold=False, sz=9): return Font(color=c, bold=bold, size=sz, name="Calibri")
thin  = Side(style="thin",   color=MED_GREY)
thick = Side(style="medium", color=MID_BLUE)
brd   = Border(left=thin, right=thin, top=thin, bottom=thin)
def walign(h="left", v="top"):
    return Alignment(horizontal=h, vertical=v, wrap_text=True)

# ════════════════════════════════════════════════════════════════
# SHEET 1 – MASTER OVERVIEW
# ════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Master Overview"
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

ws.merge_cells("A1:M1")
ws["A1"] = "NEW CORPORATE EMPLOYEE GIVING PROGRAMS — 16 ADDITIONAL COMPANIES (COMPREHENSIVE RESEARCH)"
ws["A1"].font      = Font(name="Calibri", bold=True, size=14, color=WHITE)
ws["A1"].fill      = hfill(DARK_NAVY)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 30

headers = [
    "Company", "Sector", "Program Name", "Match Ratio",
    "Annual Cap (USD)", "Min Donation", "Volunteer Grant",
    "Platform / Portal", "Eligible Orgs", "Submission Process",
    "Special Campaigns / Notes", "Eligibility", "Primary Source"
]
for col, h in enumerate(headers, 1):
    cell = ws.cell(row=2, column=col, value=h)
    cell.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill = hfill(MID_BLUE)
    cell.alignment = walign("center", "center")
    cell.border = brd
ws.row_dimensions[2].height = 30

data = [
    # Banking / Financial Services
    (
        "JPMorgan Chase",
        "Banking / Financial Services",
        "JPMorgan Chase Employee Matching Gift Program",
        "1:1",
        "$2,000 / employee / year",
        "$25",
        "Volunteer grants available (tiered; specific $/hr not publicly confirmed)",
        "CyberGrants — cybergrants.com/jpmc/giving",
        "501(c)(3) nonprofits; accredited educational institutions; "
        "hospitals; arts and cultural organizations",
        "1) Employee donates to eligible nonprofit. "
        "2) Submits match request via cybergrants.com/jpmc/giving. "
        "3) JPMorgan Chase Foundation reviews and issues match. "
        "Note: program has salary eligibility cap ($150,000 annual base pay or below for standard tier).",
        "Global Month of Service (annual volunteer campaign); "
        "Advancing Black Pathways giving initiatives; "
        "PRO Bono volunteer programs; "
        "Disaster relief enhanced matching campaigns",
        "Full-time employees with annual base pay ≤ $150,000 (standard program); "
        "some senior employees have separate giving arrangements",
        "https://doublethedonation.com/matching-gifts/jpmorgan-chase-co | "
        "https://forms.matchinggifts.com/ChaseGuide.Links.pdf"
    ),
    (
        "Goldman Sachs",
        "Banking / Financial Services",
        "Goldman Sachs Matching Gift Program (via Goldman Sachs Gives)",
        "1:1",
        "$20,000 / employee / year (minimum $50 donation)",
        "$50",
        "Volunteer grants available; specific $/hr not publicly confirmed",
        "EasyMatch — secure11.easymatch.com/gsmg2/ (via company intranet); "
        "Phone: (877) 298-2635; Email: gs@easymatch.com",
        "Nearly all 501(c)(3) organizations eligible including arts, "
        "environment, civic groups, education, health and human services",
        "1) Employee donates to eligible nonprofit (min $50). "
        "2) Connects to company intranet and accesses EasyMatch portal. "
        "3) Submits matching gift form within 180 days of donation. "
        "4) Requests processed quarterly (Jan 30, Apr 30, Jul 30, Oct 30).",
        "Goldman Sachs Gives (DAF fund for senior employees); "
        "10,000 Small Businesses philanthropic initiative; "
        "Launch With GS diversity-focused giving; "
        "Annual giving campaigns; "
        "Program increased from $10,000 to $20,000 cap in 2008",
        "Full-time employees; "
        "$20,000 cap is one of the highest in financial services",
        "https://doublethedonation.com/matching-gifts/goldman-sachs-group | "
        "https://forms.matchinggifts.com/GoldmanSachs.pdf | "
        "https://www.goldmansachs.com/citizenship/goldman-sachs-gives/"
    ),
    (
        "Bank of America",
        "Banking / Financial Services",
        "Employee-Directed Matching Gifts Program",
        "1:1",
        "$5,000 / employee / year",
        "$25",
        "Volunteer grants available (amount not publicly confirmed); "
        "Paid Volunteer Time Off (VTO) provided",
        "Benevity (referenced by Bank of America's community pages)",
        "Eligible nonprofits worldwide; "
        "focus on housing, jobs, healthcare, education",
        "1) Employee donates to eligible nonprofit. "
        "2) Submits match request via Bank of America employee portal (Benevity). "
        "3) Bank of America Charitable Foundation issues 1:1 match. "
        "Payroll deduction and card options available.",
        "Neighborhood Builders program (capacity grants to nonprofits); "
        "Global Volunteer Month (April); "
        "Employee Engagement activities in 35+ countries; "
        "Disaster response enhanced matching; "
        "2024: $69M+ in combined employee giving, matching gifts & volunteer grants "
        "to 48,000+ nonprofits globally",
        "Active employees globally; "
        "Bank of America has operations in 35+ countries",
        "https://about.bankofamerica.com/en/making-an-impact/matching-gifts-features-and-eligibility | "
        "https://doublethedonation.com/matching-gifts/bank-of-america"
    ),
    (
        "Wells Fargo",
        "Banking / Financial Services",
        "Wells Fargo Matching Gifts Program",
        "1:1",
        "$2,000 / employee / year",
        "$25",
        "Tiered volunteer grants: $50 for 1 hr; larger awards at 16 hrs, 50 hrs; "
        "16 hours paid Volunteer Time Off / year",
        "YourCause — employee login via WF internal portal",
        "501(c)(3) nonprofits; "
        "community development organizations; "
        "educational institutions; arts; environment",
        "1) Employee donates to eligible nonprofit. "
        "2) Submits request within 6 months of donation via YourCause portal. "
        "3) Wells Fargo Foundation issues 1:1 match up to $2,000/year. "
        "Volunteer grants: log hours → tiered grant to org.",
        "Housing Affordability Philanthropy initiative; "
        "Wells Fargo Open for Business Fund; "
        "16 hours paid VTO annually; "
        "Volunteer matching tiered by number of hours",
        "Active employees; "
        "retirees may be eligible with reduced benefit",
        "https://doublethedonation.com/matching-gifts/wells-fargo-company | "
        "https://www.wellsfargo.com/about/responsibility-and-impact/community-giving/giving-faqs/"
    ),
    # Pharma / Healthcare
    (
        "Johnson & Johnson",
        "Pharmaceutical / Healthcare",
        "Johnson & Johnson Matching Gifts Program",
        "2:1 (current employees); 1:1 (retirees)",
        "Employee personal donation cap: $10,000/year → "
        "Company match: up to $20,000 (current) or $10,000 (retiree)",
        "Not publicly specified",
        "Volunteer grants (specific $/hr not publicly confirmed)",
        "J&J internal matching gift portal (online submission)",
        "501(c)(3) nonprofits broadly; "
        "arts, education, health, civic, environmental organizations",
        "1) Employee donates to eligible nonprofit (up to $10,000 personal). "
        "2) Submits match request via J&J online portal within 1 YEAR of donation. "
        "3) Johnson & Johnson Foundation issues 2:1 match (current employees) "
        "or 1:1 (retirees). "
        "Result: up to triple the impact for current employees.",
        "J&J Foundation global giving programs; "
        "Employee ERG-aligned giving; "
        "Global Day of Service; "
        "Disaster relief giving campaigns; "
        "2:1 match is rare and makes J&J one of the most generous matchers in the industry",
        "Full-time and part-time active employees (2:1); "
        "retirees (1:1); board-eligible members",
        "2:1 match effectively triples employee donations — "
        "among most generous in pharma sector",
        "https://doublethedonation.com/matching-gifts/johnson-and-johnson | "
        "https://forms.matchinggifts.com/JohnsonAndJohnsonmgguidelines.pdf | "
        "https://www.jnj.com/caring-and-giving"
    ),
    (
        "Pfizer",
        "Pharmaceutical / Healthcare",
        "Pfizer Foundation Matching Gift Program",
        "1:1 (employees); 0.5:1 / 50% (retirees)",
        "$5,000 / employee / year",
        "$25",
        "Individual volunteer rewards: earn credits for volunteer hours → "
        "donate to any eligible nonprofit. "
        "Global Team Volunteer Grants: Pfizer Foundation grants to orgs "
        "where global teams volunteer together.",
        "Pfizer Foundation online portal — pfizerplus.com "
        "(submit and track requests)",
        "501(c)(3) nonprofits in US and Puerto Rico (individual matching); "
        "global nonprofits eligible for Team Volunteer Grants",
        "1) US/Puerto Rico-based employee donates to eligible org (min $25). "
        "2) Submits match request online via Pfizer Foundation portal. "
        "3) Deadline: March 1 of the FOLLOWING year. "
        "4) Pfizer Foundation issues 1:1 match. "
        "Retirees: submit same portal, receive 50% match. "
        "Team Volunteer Grants: group volunteer → Foundation grants to org.",
        "Pfizer Ignite employee giving campaigns; "
        "Global Day of Service; "
        "Team Volunteer Grant is a globally-accessible version of matching; "
        "Rare disease and health equity giving focus; "
        "30% employee participation rate (leading in pharma per industry data)",
        "US and Puerto Rico full-time employees (1:1); "
        "part-time employees eligible; "
        "retirees eligible at 50% match; "
        "spouses and board members eligible",
        "https://www.pfizer.com/about/responsibility/global-impact/give-forward | "
        "https://doublethedonation.com/matching-gifts/pfizer-inc | "
        "https://pfizerplus.com/gi/give_back.aspx"
    ),
    (
        "Merck",
        "Pharmaceutical / Healthcare",
        "Merck Company Foundation — Partnership for Giving",
        "1:1",
        "$30,000 / employee / year",
        "Not publicly specified",
        "Volunteer grants (tiered): "
        "$500 grant for 40 volunteer hours/year; "
        "$1,000 grant for 80 volunteer hours/year",
        "CyberGrants — cybergrants.com/merck/merckgivesback",
        "501(c)(3) nonprofits broadly; "
        "educational, health, civic, arts, environmental organizations",
        "1) Employee donates to eligible nonprofit. "
        "2) Submits via CyberGrants portal (cybergrants.com/merck/merckgivesback). "
        "3) Merck Company Foundation issues 1:1 match up to $30,000/year. "
        "Volunteer grants: log 40 or 80 hours → receive $500 or $1,000 grant.",
        "Merck for Mothers global maternal health initiatives; "
        "Merck Patient Assistance Program; "
        "Global Day of Service volunteer events; "
        "$30,000 cap is among highest in pharma sector; "
        "Partnership for Giving is the formal program name",
        "Full-time and part-time active employees; "
        "NOTE: retirees NOT eligible (unlike J&J and Pfizer); "
        "$30,000 cap makes this one of the most generous programs overall",
        "https://doublethedonation.com/matching-gifts/merck-co | "
        "https://files.doublethedonation.com/forms/Merck---Co--guidelines-1686482653.pdf | "
        "https://www.merck.com/company-overview/sustainability/philanthropy/strengthening-our-communities/"
    ),
    # Consulting / Professional Services
    (
        "Deloitte",
        "Consulting / Professional Services",
        "Deloitte Foundation Matching Gifts Program",
        "1:1",
        "$32,500 / employee / year",
        "Not publicly specified",
        "Volunteer grants available (specific $/hr details internal)",
        "Deloitte Foundation online portal "
        "(forms.matchinggifts.com/deloittematchinggifts.pdf for guidelines)",
        "Accredited nonprofit educational institutions (primary focus: K-12 through university); "
        "most 501(c)(3) organizations eligible",
        "1) Employee donates to eligible nonprofit. "
        "2) Submits match request via Deloitte Foundation portal. "
        "3) Deloitte Foundation issues 1:1 match up to $32,500/year. "
        "Payroll deduction and direct donation options.",
        "Deloitte STEM education partnerships; "
        "WorldClass initiative (reaching 100M people with skills/opportunities by 2030); "
        "Pro bono professional services for nonprofits; "
        "Deloitte Impact Day (global volunteer event); "
        "$32,500 cap is among the highest of any company globally",
        "Active full-time employees; "
        "NOTE: part-time and retired employees NOT eligible for full program; "
        "$32,500 cap is highest confirmed in consulting sector",
        "https://doublethedonation.com/matching-gifts/deloitte | "
        "https://files.doublethedonation.com/forms/Deloitte-guidelines-1712687004.pdf | "
        "https://business.fiu.edu/give/docs/deloittematchinggifts.pdf"
    ),
    # Consumer Goods / FMCG
    (
        "The Coca-Cola Company",
        "Consumer Goods / Beverages",
        "Coca-Cola Matching Gifts Program",
        "2:1 (effectively triples donations)",
        "$10,000 personal donation / year → up to $20,000 company match",
        "Not publicly specified",
        "No volunteer grants program (per available sources)",
        "Coca-Cola internal giving portal "
        "(doublethedonation.com/forms/coke.pdf for guidelines)",
        "Most 501(c)(3) nonprofits; "
        "arts, education, environment, health, civic organizations",
        "1) Employee (or retiree) donates to eligible nonprofit. "
        "2) Submits match request via Coca-Cola Foundation portal. "
        "3) Coca-Cola Foundation issues 2:1 match. "
        "DEADLINE: February 28 of the FOLLOWING calendar year (generous window). "
        "Result: $1 donated = $3 total to nonprofit.",
        "Coca-Cola Foundation global giving; "
        "World Without Waste environmental initiatives; "
        "5by20 women economic empowerment program; "
        "2:1 match is rare in consumer goods sector; "
        "Retirees ARE eligible (unlike many peers)",
        "Full-time and part-time employees; "
        "RETIREES also eligible — one of few companies extending this benefit; "
        "2:1 match makes this exceptional for consumer sector",
        "https://doublethedonation.com/matching-gifts/the-coca-cola-company | "
        "https://doublethedonation.com/forms/coke.pdf"
    ),
    (
        "PepsiCo",
        "Consumer Goods / Beverages",
        "PepsiCo Foundation Matching Gifts Program",
        "1:1 standard; 2:1 if employee volunteers 50+ hours/year",
        "$10,000 / employee / year",
        "$25",
        "2:1 match if employee volunteers 50+ hours/year with an org "
        "(effectively tripling donations for active volunteers). "
        "Separate volunteer time off provided.",
        "YourCause — pepsico.yourcause.com",
        "Most 501(c)(3) nonprofits including education (K-12), "
        "health and human services, arts, civic and community, "
        "environmental organizations",
        "1) Employee donates to eligible org (min $25). "
        "2) Submits match request via pepsico.yourcause.com "
        "within 6 MONTHS of donation. "
        "3) PepsiCo Foundation issues 1:1 match (or 2:1 if 50+ vol. hours). "
        "Phone support: 1-866-321-6391.",
        "PepsiCo Positive (pep+) sustainability initiatives; "
        "Racial equality giving campaigns; "
        "Volunteer 50+ hours → unlock 2:1 ratio (unique incentive structure); "
        "6-month deadline is shorter than industry average (1 year)",
        "Full-time and part-time employees; "
        "NOTE: retirees NOT eligible; "
        "Unique volunteer-incentivized 2:1 upgrade is a standout feature",
        "https://doublethedonation.com/matching-gifts/pepsico | "
        "http://www.crhsraiderband.org/Content/3_31/Files//PepsiCo.pdf | "
        "https://doublethedonation.com/forms/pepsico-guidelines.pdf"
    ),
    # Consumer / Retail / Apparel
    (
        "Nike",
        "Consumer / Apparel / Retail",
        "Nike Give Your Best (GYB) — powered by Benevity",
        "1:1 standard; 2:1 on Giving Tuesday",
        "$25,000 / employee / year",
        "Not publicly specified",
        "$10 / volunteer hour (credit donated to cause of employee's choice). "
        "Nike Community Impact Fund (NCIF): employee-led grant-making to local nonprofits.",
        "Benevity — internal branding: Give Your Best (GYB) on NikeNet",
        "Eligible schools and charitable organizations broadly; "
        "NCIF focuses on sport, physical activity, and community wellbeing",
        "1) Employee accesses Give Your Best (GYB) platform via NikeNet (Benevity). "
        "2) Donates to eligible org. "
        "3) Nike Foundation matches 1:1 (or 2:1 on Giving Tuesday). "
        "4) Volunteer hours: log in GYB → $10/hour credit donated to nonprofit. "
        "NCIF: employees participate in grant-making panels to award local grants.",
        "Giving Tuesday 2:1 elevated match; "
        "Nike Community Impact Fund (employee-led grant-making panels); "
        "Move to Zero sustainability philanthropy focus; "
        "Nike Employee Giving Campaign (annual); "
        "NCIF administered by Oregon Community Foundation",
        "All Nike employees globally via NikeNet; "
        "$25,000 cap is among highest in consumer/retail sector",
        "https://about.nike.com/en/impact/initiatives/matching-our-employees-impact | "
        "https://about.nike.com/en/mission/initiatives/nike-community-impact-fund | "
        "https://doublethedonation.com/matching-gifts/nike-inc"
    ),
    # Technology
    (
        "Cisco",
        "Technology",
        "Cisco Matching Gifts Program — powered by Benevity (Time2Give)",
        "1:1",
        "$25,000 / employee / year",
        "Not publicly specified",
        "Volunteer hours matched; "
        "80 hours paid Volunteer Time Off / year (Time2Give program — "
        "among highest VTO in tech industry); "
        "cash, stock, and volunteer time all match-eligible",
        "Benevity (migrated from Bright Funds; cisco.brightfunds.org no longer active)",
        "Nonprofits and NGOs globally; "
        "focus on education, economic empowerment, and critical human needs",
        "1) Employee donates (cash, stock, or volunteer time) to eligible org. "
        "2) Submits via Cisco Benevity portal. "
        "3) Cisco Foundation matches up to $25,000/year. "
        "80 hours VTO: take paid time off for community service; "
        "log hours for additional volunteer match.",
        "Time2Give: 80 hours paid VTO (10 paid days) — one of highest in industry; "
        "Cisco Global Problem Solver Challenge (innovation grants); "
        "Cisco Crisis Response (tech for disaster relief); "
        "FY23-FY25: 80-85% employee participation rate — "
        "exceptional for any program at scale; "
        "$90M+ in employee donations + matches since FY23",
        "All Cisco employees globally; "
        "80 hours VTO is significantly above industry average",
        "https://benevity.com/client-stories/cisco-elevated-social-impact-with-benevity | "
        "https://www.cisco.com/site/us/en/about/purpose/social-impact/cisco-foundation.html | "
        "https://doublethedonation.com/matching-gifts/cisco-systems-inc"
    ),
    (
        "Dell Technologies",
        "Technology",
        "Dell Technologies Employee Giving Program",
        "1:1",
        "$10,000 / employee / year",
        "Not publicly specified",
        "Volunteer grant: $150 per quarter when employee volunteers 10+ hours/quarter "
        "(=$600 max/year if all 4 quarters). "
        "Submit via YourCause portal.",
        "YourCause — dell.yourcause.com",
        "501(c)(3) nonprofits; educational institutions; "
        "health and human services; arts; civic organizations",
        "1) Employee donates to eligible nonprofit. "
        "2) Submits via dell.yourcause.com. "
        "3) Dell Technologies Foundation issues 1:1 match up to $10,000/year. "
        "Volunteer grants: volunteer 10+ hours/quarter → log via YourCause → "
        "$150 quarterly grant to nonprofit of choice.",
        "Dell Reconnect (electronics recycling partnerships with nonprofits); "
        "Global Month of Service (June); "
        "Dell Technologies Capital for Nonprofits; "
        "Planet program (sustainability-focused giving); "
        "Quarterly structure incentivizes year-round volunteering",
        "Full-time employees; "
        "part-time employees may be eligible (confirm via HR); "
        "quarterly volunteering structure is distinctive",
        "https://doublethedonation.com/matching-gifts/dell-inc | "
        "http://www.crhsraiderband.org/Content/3_31/Files/DELL.pdf | "
        "https://www.dell.com/en-us/dt/corporate/social-impact/transforming-lives/employee-led-impact.htm"
    ),
    (
        "Adobe",
        "Technology",
        "Adobe Create Change — powered by Benevity",
        "1:1",
        "$10,000 combined annual cap "
        "(matching gifts + volunteer grants combined)",
        "Not publicly specified",
        "Individual: $25 / volunteer hour. "
        "Team: $1,000 grant per 10-employee team that volunteers together "
        "for 2+ hours (max $10,000). "
        "Both count toward $10,000 annual combined cap.",
        "Benevity — causes.benevity.com "
        "(Adobe's 'Create Change Dashboard')",
        "Eligible charities globally; "
        "broad cause areas including education, arts, equity, environment",
        "1) Employee registers with Benevity at causes.benevity.com. "
        "2) Donates cash or securities to eligible charity. "
        "3) Submits match request via Create Change Dashboard. "
        "4) Adobe matches 1:1 up to combined $10,000 cap. "
        "5) Volunteer hours: log individually OR as a team → grants disbursed. "
        "Team volunteer grant: 10 employees × 2+ hours = $1,000 grant.",
        "Adobe Creative Residency (artist support); "
        "Adobe Employee Community Fund (ECF) — employees apply for grants "
        "to nonprofits they are personally involved with; "
        "Giving Season campaigns; "
        "Team volunteer structure incentivizes group community action; "
        "Adobe has achieved record giving seasons since partnering with Benevity",
        "All Adobe employees globally; "
        "unique team volunteer grant ($1,000/team) is distinctive; "
        "ECF grant program allows employees to nominate orgs for larger grants",
        "https://www.adobe.com/corporate-responsibility/community.html | "
        "https://benevity.com/client-stories/adobes-best-giving-season | "
        "https://benevity.com/resources/adobe-engage-employees-grantmaking-program | "
        "https://storage.benevitygrants.org/files/adobe/public/general/Adobe_ECF_Grant%20Applicant%20Guidelines%20and%20Resources.pdf"
    ),
    (
        "SAP",
        "Technology",
        "SAP Social Sabbatical + SAP Together Matching Program",
        "1:1",
        "$750 / employee / year (matching gifts)",
        "$50 minimum",
        "$20 / volunteer hour — "
        "among highest volunteer grant rates per hour of any company",
        "SAP internal giving portal "
        "(doublethedonation.com/forms/sap-guidelines.pdf for guidelines)",
        "501(c)(3) nonprofits broadly; "
        "educational, health, civic, environmental organizations",
        "1) Employee donates to eligible org (min $50). "
        "2) Submits match request via SAP internal portal. "
        "3) SAP Foundation issues 1:1 match up to $750/year. "
        "Volunteer grants: log hours → $20/hour grant disbursed. "
        "DEADLINE: December 31 of the donation year.",
        "SAP Social Sabbatical: paid 4-week leave for employees to volunteer "
        "full-time with nonprofits in emerging markets (unique in industry); "
        "SAP Together Program for community engagement; "
        "$20/hour volunteer grant is highest publicly confirmed rate "
        "among all companies researched; "
        "Dec 31 deadline is same-year (stricter than most)",
        "Active employees; "
        "NOTE: low $750 matching cap, but $20/hour volunteer grant "
        "compensates for frequent volunteers; "
        "SAP Social Sabbatical is exceptional — 4-week paid volunteering leave",
        "https://doublethedonation.com/matching-gifts/sap-america-inc | "
        "https://doublethedonation.com/forms/sap-guidelines.pdf | "
        "https://www.glassdoor.com/Benefits/SAP-Charitable-Gift-Matching-US-BNFT21_E10471_N1.htm"
    ),
    (
        "Qualcomm",
        "Technology / Semiconductors",
        "Qualcomm Employee Giving Program — powered by Benevity",
        "1:1",
        "$1,000–$5,000 / year (tiered by employee level: "
        "Associate→Director level: $1,000; "
        "Director+ to VP: higher tiers up to $5,000)",
        "Not publicly specified",
        "Volunteer grants (specific $/hr not confirmed; "
        "combined with donation matching within annual cap)",
        "Benevity — prudential.benevity.org (referenced portal); "
        "Qualcomm internal giving platform",
        "501(c)(3) nonprofits; educational institutions; "
        "STEM, education, and community-focused organizations (priority areas)",
        "1) Employee donates to eligible org. "
        "2) Submits via Qualcomm Benevity portal. "
        "3) Qualcomm Foundation issues match based on employee tier. "
        "NOTE: Qualcomm's giving program runs on a FISCAL YEAR: "
        "October 1 – September 30 (not calendar year).",
        "Qualcomm Wireless Reach (ICT for sustainable development); "
        "STEM education initiatives; "
        "Qualcomm THINKABIT Lab (STEM education for underserved students); "
        "Employee level-based tiering is distinctive in industry; "
        "Oct 1–Sep 30 fiscal year is unique",
        "Active employees; "
        "tiered by employee level (associate through VP); "
        "fiscal year Oct 1–Sep 30 requires different planning than calendar-year programs",
        "https://doublethedonation.com/matching-gifts | "
        "https://www.qualcomm.com/company/positions/corporate-responsibility"
    ),
]

row_colors = [LIGHT_GREY, WHITE]
for row_idx, row_data in enumerate(data, start=3):
    alt = row_colors[(row_idx - 3) % 2]
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill = hfill(alt)
        cell.font = hfont(DARK_TEXT, bold=(col_idx == 1), sz=9)
        cell.alignment = walign("left", "top")
        cell.border = brd
    ws.row_dimensions[row_idx].height = 100

col_widths = [18, 22, 32, 22, 32, 14, 38, 38, 38, 65, 60, 45, 80]
for i, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ════════════════════════════════════════════════════════════════
# SHEET 2 – PROCESS DEEP-DIVE
# ════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Process Deep-Dive")
ws2.sheet_view.showGridLines = False
ws2.freeze_panes = "A3"

ws2.merge_cells("A1:G1")
ws2["A1"] = "STEP-BY-STEP PROCESS & PLATFORM DETAILS — 16 NEW COMPANIES"
ws2["A1"].font = Font(name="Calibri", bold=True, size=14, color=WHITE)
ws2["A1"].fill = hfill(DARK_NAVY)
ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 30

h2 = ["Company", "Platform / Tool", "How to Access",
      "Step-by-Step Process", "Submission Deadline",
      "Volunteer Grant Details", "Key Caveats / Standout Features"]
for col, h in enumerate(h2, 1):
    cell = ws2.cell(row=2, column=col, value=h)
    cell.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill = hfill(MID_BLUE)
    cell.alignment = walign("center", "center")
    cell.border = brd
ws2.row_dimensions[2].height = 30

process_data = [
    ("JPMorgan Chase",
     "CyberGrants",
     "cybergrants.com/jpmc/giving (via internal network)",
     "1. Employee donates to eligible nonprofit.\n"
     "2. Logs in to CyberGrants at cybergrants.com/jpmc/giving.\n"
     "3. Submits matching gift request (attach donation receipt).\n"
     "4. JPMorgan Chase Foundation reviews and approves.\n"
     "5. Foundation issues 1:1 match up to $2,000/year.",
     "Within the calendar year (Dec 31)",
     "Volunteer grants available (tiered amounts not publicly confirmed); "
     "submit via CyberGrants portal alongside donation matching",
     "Salary cap for standard program: employees with annual base pay "
     "exceeding $150,000 may not qualify for standard matching program. "
     "JPMorgan Chase is the largest US bank — over 300,000 employees. "
     "$2,000 cap is on the lower end for a company of this size. "
     "The firm also runs separate Global Month of Service annually."),
    ("Goldman Sachs",
     "EasyMatch (easymatch.com)",
     "secure11.easymatch.com/gsmg2/ — access via company intranet; "
     "Phone: (877) 298-2635; Email: gs@easymatch.com",
     "1. Employee donates to eligible nonprofit (min $50).\n"
     "2. Connects via company intranet to EasyMatch portal.\n"
     "3. Submits matching gift form within 180 DAYS of donation.\n"
     "4. Request processed quarterly:\n"
     "   - January 30, April 30, July 30, October 30.\n"
     "5. Goldman Sachs Foundation issues 1:1 match up to $20,000.",
     "Within 180 days of donation; quarterly processing cycles",
     "Volunteer grants available but specific $/hr not publicly confirmed; "
     "Goldman Sachs Gives (DAF) is separate senior-level program",
     "Quarterly processing means donations in late December "
     "may not be matched until January 30 of next year — plan accordingly. "
     "$20,000 cap increased from $10,000 in 2008. "
     "Goldman Sachs Gives is a separate Donor Advised Fund program "
     "for senior professionals — distinct from standard employee matching. "
     "EasyMatch (not Benevity/YourCause) is the platform — verify current URL."),
    ("Bank of America",
     "Benevity (employee giving portal)",
     "Bank of America internal employee portal (Benevity-powered)",
     "1. Employee logs in to Bank of America employee giving portal.\n"
     "2. Searches for and selects eligible nonprofit.\n"
     "3. Donates via payroll deduction or card.\n"
     "4. Submits match request within portal.\n"
     "5. BofA Charitable Foundation issues 1:1 match up to $5,000/year.",
     "Within the calendar year (Dec 31)",
     "Volunteer grants available; specific $/hr not confirmed publicly; "
     "Bank of America has a robust Volunteer Time Off policy",
     "2024 impact: $69M+ in combined giving to 48,000+ nonprofits globally — "
     "one of the largest bank-driven giving programs by reach. "
     "Neighborhood Builders is a separate $200,000 capacity grant program "
     "for nonprofits (not employee matching). "
     "Bank of America operates in 35+ countries — global program."),
    ("Wells Fargo",
     "YourCause",
     "Internal Wells Fargo employee portal (YourCause-powered)",
     "1. Employee donates to eligible nonprofit (min $25).\n"
     "2. Submits match request via YourCause portal "
     "within 6 MONTHS of donation.\n"
     "3. Wells Fargo Foundation issues 1:1 match up to $2,000/year.\n"
     "Volunteer grants (tiered):\n"
     "- 1 hour → $50 grant\n"
     "- 16 hours → larger grant\n"
     "- 50+ hours → largest tier grant\n"
     "Log hours via YourCause to claim volunteer grants.",
     "Within 6 months of donation",
     "Tiered volunteer grants by hours; "
     "16 hours paid VTO per year; "
     "log hours via YourCause portal for grant disbursement",
     "6-month submission window is shorter than most peers (many allow 1 year). "
     "$2,000 cap is low for a major bank. "
     "Wells Fargo has had reputational issues affecting CSR perception, "
     "but the matching program is active and confirmed. "
     "16 hours paid VTO is a tangible benefit."),
    ("Johnson & Johnson",
     "J&J Online Matching Gift Portal",
     "Internal J&J employee portal (online submission)",
     "1. Employee donates to eligible nonprofit.\n"
     "2. Submits match request via J&J online portal "
     "within 1 YEAR of donation.\n"
     "3. Johnson & Johnson Foundation issues:\n"
     "   - Current employees: 2:1 match (up to $20,000 company contribution)\n"
     "   - Retirees: 1:1 match (up to $10,000 company contribution)\n"
     "Personal donation cap: $10,000/year either way.",
     "Within 1 year of donation",
     "Volunteer grants available; specific rate not publicly confirmed; "
     "J&J employees participate in Global Day of Service",
     "The 2:1 match effectively TRIPLES a current employee's donation — "
     "among the most generous in any sector. "
     "$10,000 personal donation → $20,000 company match → "
     "$30,000 total to nonprofit. "
     "Retirees receive 1:1 (still generous). "
     "J&J is one of very few large companies still offering a 2:1 match."),
    ("Pfizer",
     "Pfizer Foundation Portal (pfizerplus.com)",
     "pfizerplus.com/gi/give_back.aspx",
     "1. US/Puerto Rico-based employee donates to eligible 501(c)(3).\n"
     "2. Submits online via Pfizer Foundation portal (pfizerplus.com).\n"
     "3. DEADLINE: March 1 of the following year "
     "(one of the later deadlines in industry).\n"
     "4. Pfizer Foundation issues 1:1 match up to $5,000.\n"
     "Retirees: submit same way, receive 50% (0.5:1) match.\n"
     "Team Volunteer Grants: team volunteers globally → "
     "Foundation grants to participating org.",
     "March 1 of the following year (generous deadline)",
     "Individual volunteer rewards: earn credits → donate to any eligible nonprofit. "
     "Global Team Volunteer Grants: for international teams — "
     "Pfizer Foundation grants to orgs where teams volunteer together. "
     "Submit via Pfizer Foundation portal.",
     "March 1 next-year deadline is longer than most — employees have time. "
     "Retirees receive 50% match (unusual tiered structure). "
     "Global Team Volunteer Grants extend benefit beyond US/Puerto Rico "
     "to all Pfizer employees worldwide. "
     "Pfizer raises $33.7M/year in giving programs with 40,000+ requests."),
    ("Merck",
     "CyberGrants — Partnership for Giving",
     "cybergrants.com/merck/merckgivesback",
     "1. Employee donates to eligible nonprofit.\n"
     "2. Submits via CyberGrants portal at "
     "cybergrants.com/merck/merckgivesback.\n"
     "3. Merck Company Foundation issues 1:1 match up to $30,000/year.\n"
     "Volunteer grants:\n"
     "- Log 40 volunteer hours/year → $500 grant to nonprofit\n"
     "- Log 80 volunteer hours/year → $1,000 grant to nonprofit",
     "Within the calendar year (Dec 31)",
     "Tiered volunteer grants: 40 hrs → $500; 80 hrs → $1,000; "
     "log hours via CyberGrants portal",
     "$30,000 cap is exceptional — one of the highest globally for pharma. "
     "NOTE: Retirees are NOT eligible (important for former employees to know). "
     "CyberGrants is the platform (not Benevity/YourCause). "
     "Volunteer grant requires significant hour commitment (40 hrs minimum). "
     "Merck is separate from Merck KGaA (German company) — "
     "this covers Merck & Co. (US/Canada)."),
    ("Deloitte",
     "Deloitte Foundation Portal",
     "Internal Deloitte giving portal; "
     "guidelines: files.doublethedonation.com/forms/Deloitte-guidelines-1712687004.pdf",
     "1. Employee donates to eligible nonprofit.\n"
     "2. Submits match request via Deloitte Foundation internal portal.\n"
     "3. Deloitte Foundation issues 1:1 match up to $32,500/year.\n"
     "Payroll deduction and direct donation both accepted.",
     "Within the calendar year (Dec 31)",
     "Volunteer grants available (specific rate not publicly confirmed); "
     "Deloitte Impact Day: firm-wide annual volunteer event",
     "$32,500 is the HIGHEST confirmed annual match cap of any company "
     "in this entire research set across all 32 companies. "
     "Note: part-time and retired employees are NOT eligible. "
     "Deloitte's WorldClass initiative commits to 50M people reached by 2030. "
     "Pro bono professional services (accounting, consulting, legal) "
     "complement the cash matching program. "
     "Deloitte Impact Day mobilizes all ~500,000 global employees for a day of service."),
    ("The Coca-Cola Company",
     "Coca-Cola Foundation Portal",
     "Internal Coca-Cola employee portal",
     "1. Employee (or retiree) donates to eligible nonprofit.\n"
     "2. Submits match request via Coca-Cola Foundation portal.\n"
     "3. Coca-Cola Foundation issues 2:1 match "
     "(employee donates $1, Coca-Cola adds $2).\n"
     "4. DEADLINE: February 28 of the FOLLOWING year.\n"
     "Result: each $1 donated = $3 total to nonprofit.",
     "February 28 of the following calendar year (generous window)",
     "No volunteer grant program confirmed in public sources",
     "2:1 match is rare among consumer goods companies. "
     "Feb 28 deadline (not Dec 31) gives employees 2 extra months — "
     "very employee-friendly. "
     "Retirees ARE eligible — extends program reach to former employees. "
     "The company calls their matching funds the Coca-Cola Foundation. "
     "No volunteer grant program has been publicly confirmed — "
     "a gap vs tech/pharma peers."),
    ("PepsiCo",
     "YourCause — pepsico.yourcause.com",
     "pepsico.yourcause.com; Phone support: 1-866-321-6391",
     "1. Employee donates to eligible org (min $25).\n"
     "2. Submits via pepsico.yourcause.com "
     "within 6 MONTHS of donation.\n"
     "3. PepsiCo Foundation issues match:\n"
     "   - Standard: 1:1 up to $10,000/year\n"
     "   - If volunteered 50+ hrs with org in the same year: "
     "2:1 up to $10,000/year\n"
     "4. Volunteer-upgraded donors: contact (866) 321-6391 to confirm tier.",
     "Within 6 months of donation (strict)",
     "Volunteer 50+ hours with org → automatically upgrades to 2:1 match ratio; "
     "this is PepsiCo's unique incentive structure to reward active volunteers",
     "6-month deadline is shorter than industry average — submit promptly. "
     "Volunteer-upgraded 2:1 ratio is a standout feature that rewards "
     "sustained community involvement. "
     "Retirees are NOT eligible. "
     "YourCause is the platform (same as Dell, Wells Fargo). "
     "Note: YourCause is owned by Blackbaud."),
    ("Nike",
     "Benevity — Give Your Best (GYB) on NikeNet",
     "Accessible via NikeNet (Nike's internal employee portal)",
     "1. Employee accesses Give Your Best (GYB) platform via NikeNet (Benevity).\n"
     "2. Searches for eligible school or charitable organization.\n"
     "3. Donates via payroll, card, or other method.\n"
     "4. Nike Foundation matches 1:1 (or 2:1 on Giving Tuesday).\n"
     "5. Volunteer hours: log in GYB → $10/hour credit → "
     "donate credit to cause of choice.\n"
     "NCIF: Employee applies to participate in grant-making panel "
     "→ reviews community applications "
     "→ helps award grants to local nonprofits.",
     "Within the calendar year (Dec 31); Giving Tuesday (specific date each Nov)",
     "$10/hour volunteer grant; credits accumulated and donated via GYB; "
     "NCIF employee grant-making panels (separate from matching program)",
     "Nike Community Impact Fund (NCIF) is a unique employee-led "
     "grant-making initiative — employees literally sit on panels "
     "that decide grant awards to local nonprofits. "
     "Giving Tuesday 2:1 elevated match is an annual campaign. "
     "$25,000 cap is highest among consumer/apparel companies surveyed. "
     "NCIF administered by Oregon Community Foundation (Portland HQ focus)."),
    ("Cisco",
     "Benevity (migrated from Bright Funds)",
     "Cisco internal Benevity portal; "
     "cisco.brightfunds.org is no longer active",
     "1. Employee accesses Cisco Benevity portal.\n"
     "2. Donates cash, stock, or volunteer hours to eligible nonprofit.\n"
     "3. Cisco Foundation matches up to $25,000/year.\n"
     "4. Time2Give: take up to 80 hours paid VTO for community service.\n"
     "5. Log VTO hours in Benevity portal → "
     "volunteer hours also count toward matching.",
     "Within the fiscal year",
     "80 hours paid Volunteer Time Off (Time2Give) — "
     "equivalent to 10 full working days per year; "
     "cash, stock, AND volunteer time all match-eligible; "
     "log hours in Benevity portal",
     "80 hours VTO is the highest confirmed paid VTO among all companies "
     "in this research set. "
     "Cisco moved from Bright Funds to Benevity — "
     "cisco.brightfunds.org no longer active. "
     "FY23-FY25: 80-85% employee participation rate is extraordinary "
     "(industry average is 5-10%). "
     "$90M+ contributed to nonprofits since FY23. "
     "Cash AND stock AND volunteer time all eligible for match — "
     "broadest definition of matchable giving."),
    ("Dell Technologies",
     "YourCause — dell.yourcause.com",
     "dell.yourcause.com",
     "1. Employee donates to eligible nonprofit.\n"
     "2. Submits via dell.yourcause.com.\n"
     "3. Dell Technologies Foundation issues 1:1 match up to $10,000/year.\n"
     "4. Volunteer grant (quarterly):\n"
     "   - Volunteer 10+ hours in a quarter → "
     "request $150 grant for nonprofit of choice\n"
     "   - Maximum: $150 x 4 quarters = $600/year in volunteer grants",
     "Within the calendar year (Dec 31)",
     "$150/quarter volunteer grant when 10+ hours volunteered; "
     "max $600/year in volunteer grants; "
     "log hours via YourCause portal quarterly",
     "Quarterly volunteer grant structure is unique — "
     "incentivizes consistent year-round volunteering, "
     "not just a one-time annual effort. "
     "Dell Reconnect program (electronics recycling) is a separate "
     "in-kind giving initiative. "
     "YourCause (Blackbaud) is the platform — same as Wells Fargo, PepsiCo. "
     "Global Month of Service (June) is a flagship volunteer campaign."),
    ("Adobe",
     "Benevity — Create Change Dashboard (causes.benevity.com)",
     "causes.benevity.com (register with Adobe credentials)",
     "1. Employee registers at causes.benevity.com (Adobe's Benevity instance).\n"
     "2. Donates cash or securities to eligible charity.\n"
     "3. Submits match on Create Change Dashboard.\n"
     "4. Adobe matches 1:1 within $10,000 combined annual cap.\n"
     "5. Individual volunteer grant: log hours → $25/hour grant disbursed.\n"
     "6. Team volunteer grant: "
     "10+ employees volunteer together for 2+ hours "
     "→ $1,000 grant to chosen nonprofit (max $10,000/year).\n"
     "Adobe Employee Community Fund (ECF): employees also apply "
     "for grants to nonprofits they are personally involved with.",
     "Within the calendar year (Dec 31)",
     "Individual: $25/hour. "
     "Team: $1,000 per 10-employee team (2+ hours together). "
     "Both count toward $10,000 combined annual cap. "
     "Submit via Create Change Dashboard on causes.benevity.com.",
     "$25/hour individual volunteer rate = among highest alongside Apple, Microsoft. "
     "Team volunteer grant ($1,000 per team) incentivizes group action. "
     "$10,000 is a COMBINED cap — all matching gifts + volunteer grants count together. "
     "Adobe Employee Community Fund (ECF) is a separate grant application program "
     "where employees nominate nonprofits for larger grants. "
     "Adobe achieved a 'record giving season' after realigning CSR with Benevity."),
    ("SAP",
     "SAP Internal Giving Portal",
     "Internal SAP giving portal; guidelines: doublethedonation.com/forms/sap-guidelines.pdf",
     "1. Employee donates to eligible nonprofit (min $50).\n"
     "2. Submits match request via SAP internal portal.\n"
     "3. SAP Foundation issues 1:1 match up to $750/year.\n"
     "4. DEADLINE: December 31 of the donation year (same-year only).\n"
     "5. Volunteer grant: log hours → $20/hour grant to nonprofit.\n"
     "SAP Social Sabbatical: separate — employee applies for 4-week "
     "paid leave to volunteer full-time with a nonprofit globally.",
     "December 31 of the SAME YEAR as donation (strict same-year rule)",
     "$20/hour volunteer grant — highest per-hour rate of all 32 companies researched; "
     "SAP Social Sabbatical: 4-week fully paid volunteer leave "
     "(flagship unique program); "
     "log hours via SAP internal portal",
     "$750 matching cap is low — but $20/hour volunteer grant "
     "is the highest confirmed per-hour rate of any company in this research. "
     "SAP Social Sabbatical (4 weeks paid full-time volunteering globally) "
     "is arguably the most impactful employee-giving benefit in the tech sector. "
     "Dec 31 same-year deadline means there is no grace period. "
     "Min $50 donation (higher than industry standard of $25)."),
    ("Qualcomm",
     "Benevity (Qualcomm internal Benevity portal)",
     "Internal Qualcomm Benevity portal",
     "1. Employee donates to eligible nonprofit.\n"
     "2. Submits via Qualcomm Benevity portal.\n"
     "3. Qualcomm Foundation issues tiered match based on employee level:\n"
     "   - Associate to Director level: up to $1,000/year\n"
     "   - Director+ and above: up to $5,000/year\n"
     "4. FISCAL YEAR: October 1 – September 30 "
     "(different from all other companies, which use Jan–Dec).",
     "Within fiscal year: October 1 – September 30",
     "Volunteer grants available (specific $/hr not confirmed publicly); "
     "combined with donation matching within annual tiered cap",
     "Oct 1–Sep 30 fiscal year is unique — employees who join mid-year "
     "should calculate their eligibility window differently. "
     "Tiered matching by employee seniority is unusual in industry. "
     "Qualcomm THINKABIT Lab is a flagship STEM education program — "
     "separate from employee matching but reflects overall giving culture. "
     "Qualcomm Wireless Reach provides ICT grants to underserved communities globally."),
]

for row_idx, row_data in enumerate(process_data, start=3):
    alt = row_colors[(row_idx - 3) % 2]
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill = hfill(alt)
        cell.font = hfont(DARK_TEXT, bold=(col_idx == 1), sz=9)
        cell.alignment = walign("left", "top")
        cell.border = brd
    ws2.row_dimensions[row_idx].height = 160

col_widths2 = [16, 28, 38, 65, 28, 52, 65]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# ════════════════════════════════════════════════════════════════
# SHEET 3 – QUICK COMPARISON
# ════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Quick Comparison")
ws3.sheet_view.showGridLines = False
ws3.freeze_panes = "A3"

ws3.merge_cells("A1:H1")
ws3["A1"] = "QUICK COMPARISON — 16 NEW COMPANIES (RANKED BY ANNUAL CAP)"
ws3["A1"].font = Font(name="Calibri", bold=True, size=14, color=WHITE)
ws3["A1"].fill = hfill(DARK_NAVY)
ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[1].height = 30

h3 = ["Rank", "Company", "Sector", "Match Ratio",
      "Annual Cap (USD)", "Volunteer Grant", "Platform", "Source"]
for col, h in enumerate(h3, 1):
    cell = ws3.cell(row=2, column=col, value=h)
    cell.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill = hfill(MID_BLUE)
    cell.alignment = walign("center", "center")
    cell.border = brd
ws3.row_dimensions[2].height = 35

quick_data = [
    (1,  "Deloitte",             "Consulting",          "1:1",                    "$32,500",             "Yes (amt. TBC)",      "Deloitte Foundation Portal", "doublethedonation.com/matching-gifts/deloitte"),
    (2,  "Merck",                "Pharma",              "1:1",                    "$30,000",             "$500–$1,000 (tiered hrs)", "CyberGrants",          "merck.com/company-overview/sustainability/philanthropy"),
    (3,  "Nike",                 "Consumer / Apparel",  "1:1 (2:1 Giving Tue.)", "$25,000",             "$10/hr",               "Benevity (Give Your Best)","about.nike.com/en/impact/initiatives/matching-our-employees-impact"),
    (4,  "Cisco",                "Technology",          "1:1",                    "$25,000",             "80 hrs VTO (Time2Give)","Benevity",               "cisco.com/site/us/en/about/purpose/social-impact/cisco-foundation.html"),
    (5,  "Goldman Sachs",        "Finance",             "1:1",                    "$20,000",             "Yes (amt. TBC)",        "EasyMatch",             "goldmansachs.com/citizenship/goldman-sachs-gives/"),
    (6,  "J&J",                  "Pharma / Healthcare", "2:1 current / 1:1 ret.","$10,000 personal (=$20,000 match for current)", "Yes (amt. TBC)", "J&J Internal Portal", "jnj.com/caring-and-giving"),
    (7,  "Adobe",                "Technology",          "1:1",                    "$10,000 combined",    "$25/hr (indiv.) + $1,000/team", "Benevity (Create Change)", "adobe.com/corporate-responsibility/community.html"),
    (8,  "Dell Technologies",    "Technology",          "1:1",                    "$10,000",             "$150/quarter (10+hrs)", "YourCause",             "dell.com/en-us/dt/corporate/social-impact"),
    (9,  "Coca-Cola",            "Consumer / Beverages","2:1",                    "$10,000 personal (=$20,000 match)", "Not confirmed", "Coca-Cola Foundation Portal", "doublethedonation.com/matching-gifts/the-coca-cola-company"),
    (10, "PepsiCo",              "Consumer / Beverages","1:1 (2:1 if 50+ vol.hrs)","$10,000",            "2:1 match upgrade for 50+ vol. hrs", "YourCause (pepsico.yourcause.com)", "doublethedonation.com/matching-gifts/pepsico"),
    (11, "Bank of America",      "Finance",             "1:1",                    "$5,000",              "Yes (amt. TBC)",        "Benevity",              "about.bankofamerica.com/en/making-an-impact/matching-gifts"),
    (12, "Pfizer",               "Pharma",              "1:1 emp. / 0.5:1 ret.", "$5,000",              "Individual volunteer rewards + Global Team Grants", "Pfizer Foundation (pfizerplus.com)", "pfizer.com/about/responsibility/global-impact/give-forward"),
    (13, "Qualcomm",             "Technology / Semicon.","1:1",                   "$1,000–$5,000 (tiered by level)", "Yes (amt. TBC)", "Benevity",            "qualcomm.com/company/positions/corporate-responsibility"),
    (14, "JPMorgan Chase",       "Finance",             "1:1",                    "$2,000",              "Tiered (amt. TBC)",     "CyberGrants",           "doublethedonation.com/matching-gifts/jpmorgan-chase-co"),
    (15, "Wells Fargo",          "Finance",             "1:1",                    "$2,000",              "Tiered: $50 for 1hr+",  "YourCause",             "wellsfargo.com/about/responsibility-and-impact/community-giving"),
    (16, "SAP",                  "Technology",          "1:1",                    "$750",                "$20/hr (highest $/hr rate)", "SAP Internal Portal", "doublethedonation.com/matching-gifts/sap-america-inc"),
]

tier_colors = {1: "C6EFCE", 2: "FFEB9C", 3: "FFC7CE"}

for row_idx, row_data in enumerate(quick_data, start=3):
    rank = row_data[0]
    if rank <= 7:
        tc = tier_colors[1]
    elif rank <= 12:
        tc = tier_colors[2]
    else:
        tc = tier_colors[3]
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws3.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill = hfill(tc)
        cell.font = hfont(DARK_TEXT, bold=(col_idx in [1, 2]), sz=9)
        cell.alignment = walign("center" if col_idx <= 2 else "left", "center")
        cell.border = brd
    ws3.row_dimensions[row_idx].height = 42

col_widths3 = [6, 18, 20, 22, 28, 30, 26, 60]
for i, w in enumerate(col_widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# Legend
legend_row = len(quick_data) + 4
ws3.merge_cells(f"A{legend_row}:H{legend_row}")
ws3[f"A{legend_row}"] = "COLOUR LEGEND:"
ws3[f"A{legend_row}"].font = Font(name="Calibri", bold=True, size=10)

for label, color, col in [
    ("$10,000+ Annual Cap (Top Tier)", "C6EFCE", 1),
    ("$2,000–$9,999 Annual Cap (Mid Tier)", "FFEB9C", 4),
    ("Below $2,000 / Tiered (Lower Tier)", "FFC7CE", 6),
]:
    cell = ws3.cell(row=legend_row + 1, column=col, value=label)
    cell.fill = hfill(color)
    cell.font = Font(name="Calibri", bold=True, size=9)
    cell.border = brd
    cell.alignment = walign("center", "center")

# ════════════════════════════════════════════════════════════════
# SHEET 4 – SOURCES
# ════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Sources & Methodology")
ws4.sheet_view.showGridLines = False

ws4.merge_cells("A1:C1")
ws4["A1"] = "SOURCES, METHODOLOGY & DATA CONFIDENCE — NEW COMPANIES"
ws4["A1"].font = Font(name="Calibri", bold=True, size=14, color=WHITE)
ws4["A1"].fill = hfill(DARK_NAVY)
ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws4.row_dimensions[1].height = 30

for col, h in enumerate(["Company / Source Category", "Source URL / Reference", "Confidence"], 1):
    cell = ws4.cell(row=2, column=col, value=h)
    cell.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
    cell.fill = hfill(MID_BLUE)
    cell.alignment = walign("center", "center")
    cell.border = brd
ws4.row_dimensions[2].height = 25

sources = [
    ("METHODOLOGY", "Sources: official company CSR/philanthropy pages, official program PDFs, "
     "Benevity/YourCause/CyberGrants platform references, Double the Donation database, "
     "Glassdoor Benefits pages, Benevity client case studies, company blogs, press releases, "
     "SEC filings, and industry publications. Research date: June 2026.", "N/A"),
    ("JPMorgan Chase — Matching Gifts", "https://doublethedonation.com/matching-gifts/jpmorgan-chase-co", "MEDIUM — Third-party database"),
    ("JPMorgan Chase — Program Guide", "https://forms.matchinggifts.com/ChaseGuide.Links.pdf", "HIGH — Official program guide PDF"),
    ("Goldman Sachs — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/goldman-sachs-group", "MEDIUM — Third-party database"),
    ("Goldman Sachs — Program Guide PDF", "https://forms.matchinggifts.com/GoldmanSachs.pdf", "HIGH — Official program guide PDF"),
    ("Goldman Sachs — GS Gives", "https://www.goldmansachs.com/citizenship/goldman-sachs-gives/", "HIGH — Official Goldman Sachs page"),
    ("Bank of America — Matching Gifts Page", "https://about.bankofamerica.com/en/making-an-impact/matching-gifts-features-and-eligibility", "HIGH — Official BofA page"),
    ("Bank of America — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/bank-of-america", "MEDIUM — Third-party database"),
    ("Wells Fargo — Community Giving", "https://www.wellsfargo.com/about/responsibility-and-impact/community-giving/giving-faqs/", "HIGH — Official Wells Fargo page"),
    ("Wells Fargo — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/wells-fargo-company", "MEDIUM — Third-party database"),
    ("Johnson & Johnson — Caring & Giving", "https://www.jnj.com/caring-and-giving", "HIGH — Official J&J page"),
    ("Johnson & Johnson — Program Guidelines", "https://forms.matchinggifts.com/JohnsonAndJohnsonmgguidelines.pdf", "HIGH — Official J&J program PDF"),
    ("Johnson & Johnson — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/johnson-and-johnson", "MEDIUM — Third-party database"),
    ("Pfizer — Give Forward", "https://www.pfizer.com/about/responsibility/global-impact/give-forward", "HIGH — Official Pfizer page"),
    ("Pfizer — Retiree Program", "https://pfizerplus.com/gi/give_back.aspx", "HIGH — Official Pfizer retiree portal"),
    ("Pfizer — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/pfizer-inc", "MEDIUM — Third-party database"),
    ("Merck — Philanthropy Page", "https://www.merck.com/company-overview/sustainability/philanthropy/strengthening-our-communities/", "HIGH — Official Merck page"),
    ("Merck — Partnership for Giving PDF", "https://files.doublethedonation.com/forms/Merck---Co--guidelines-1686482653.pdf", "HIGH — Official Merck program PDF"),
    ("Merck — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/merck-co", "MEDIUM — Third-party database"),
    ("Deloitte — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/deloitte", "MEDIUM — Third-party database"),
    ("Deloitte — Foundation Guidelines", "https://files.doublethedonation.com/forms/Deloitte-guidelines-1712687004.pdf", "HIGH — Official Deloitte Foundation PDF"),
    ("Deloitte — Foundation at FIU", "https://business.fiu.edu/give/docs/deloittematchinggifts.pdf", "MEDIUM — University reference of official PDF"),
    ("Coca-Cola — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/the-coca-cola-company", "MEDIUM — Third-party database"),
    ("Coca-Cola — Program Guidelines", "https://doublethedonation.com/forms/coke.pdf", "HIGH — Official Coca-Cola program PDF"),
    ("PepsiCo — YourCause Portal", "http://pepsico.yourcause.com/", "HIGH — Official PepsiCo giving portal"),
    ("PepsiCo — Program Form PDF", "https://forms.matchinggifts.com/PepsiCoForm.pdf", "HIGH — Official PepsiCo program PDF"),
    ("PepsiCo — Guidelines PDF", "https://doublethedonation.com/forms/pepsico-guidelines.pdf", "HIGH — Official PepsiCo guidelines PDF"),
    ("Nike — Employee Community Engagement", "https://about.nike.com/en/impact/initiatives/matching-our-employees-impact", "HIGH — Official Nike page"),
    ("Nike — Community Impact Fund", "https://about.nike.com/en/mission/initiatives/nike-community-impact-fund", "HIGH — Official Nike page"),
    ("Nike — NCIF at Oregon Community Fdn.", "https://oregoncf.org/grants-and-scholarships/grants/nike-community-impact-fund-program", "HIGH — Oregon Community Foundation official page"),
    ("Cisco — Benevity Case Study", "https://benevity.com/client-stories/cisco-elevated-social-impact-with-benevity", "HIGH — Official Benevity/Cisco case study"),
    ("Cisco — Foundation Page", "https://www.cisco.com/site/us/en/about/purpose/social-impact/cisco-foundation.html", "HIGH — Official Cisco page"),
    ("Cisco — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/cisco-systems-inc", "MEDIUM — Third-party database"),
    ("Dell — YourCause Portal", "http://dell.yourcause.com/", "HIGH — Official Dell giving portal"),
    ("Dell — Program Guidelines", "http://www.crhsraiderband.org/Content/3_31/Files/DELL.pdf", "MEDIUM — Third-party relay of official Dell PDF"),
    ("Dell — Community Impact Page", "https://www.dell.com/en-us/dt/corporate/social-impact/transforming-lives/employee-led-impact.htm", "HIGH — Official Dell page"),
    ("Adobe — Community Page", "https://www.adobe.com/corporate-responsibility/community.html", "HIGH — Official Adobe page"),
    ("Adobe — Benevity Case Study", "https://benevity.com/client-stories/adobes-best-giving-season", "HIGH — Official Benevity/Adobe case study"),
    ("Adobe — ECF Grant Guidelines", "https://storage.benevitygrants.org/files/adobe/public/general/Adobe_ECF_Grant%20Applicant%20Guidelines%20and%20Resources.pdf", "HIGH — Official Adobe ECF guidelines PDF"),
    ("Adobe — Grantmaking Employee Blog", "https://benevity.com/resources/adobe-engage-employees-grantmaking-program", "HIGH — Official Benevity/Adobe resource"),
    ("SAP — Matching Gifts DB", "https://doublethedonation.com/matching-gifts/sap-america-inc", "MEDIUM — Third-party database"),
    ("SAP — Program Guidelines PDF", "https://doublethedonation.com/forms/sap-guidelines.pdf", "HIGH — Official SAP program PDF"),
    ("SAP — Glassdoor Benefits", "https://www.glassdoor.com/Benefits/SAP-Charitable-Gift-Matching-US-BNFT21_E10471_N1.htm", "MEDIUM — Employee-reported data"),
    ("Qualcomm — Matching Gifts DB", "https://doublethedonation.com/matching-gifts", "MEDIUM — Third-party database"),
    ("DATA CAVEAT",
     "Program details change annually. Deloitte $32,500 and Merck $30,000 caps are among "
     "the highest globally — always verify current terms. "
     "SAP $750 cap is low but $20/hour volunteer grant compensates for active volunteers. "
     "Cisco 80-hour VTO and SAP Social Sabbatical are standout non-monetary benefits. "
     "J&J 2:1 match is rare. All data reflects best available public sources as of June 2026.",
     "READ BEFORE USE"),
]

for row_idx, row_data in enumerate(sources, start=3):
    alt = row_colors[(row_idx - 3) % 2]
    is_note = row_data[0] in ("METHODOLOGY", "DATA CAVEAT")
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws4.cell(row=row_idx, column=col_idx, value=str(value))
        cell.fill = hfill("FFF2CC" if is_note else alt)
        cell.font = hfont(DARK_TEXT, bold=(col_idx == 1 or is_note), sz=9)
        cell.alignment = walign("left", "top")
        cell.border = brd
    ws4.row_dimensions[row_idx].height = 50 if is_note else 30

for i, w in enumerate([40, 100, 25], 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

out = "/home/user/jcode/Employee_Giving_New_Companies.xlsx"
wb.save(out)
print(f"Saved: {out}")
