import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── palette ──────────────────────────────────────────────────────
DARK_NAVY="1B2A4A"; MID_BLUE="2E5CA8"; FOREST="1A5C38"; MID_GREEN="217A4F"
AMBER="B8860B"; DEEP_RED="8B0000"; TEAL="0F4C5C"; PURPLE="4B0082"; WHITE="FFFFFF"
LT_GREEN="D5F5E3"; LT_AMBER="FFF3CD"; LT_RED="FADBD8"; LT_BLUE="D6E4F7"
LT_GREY="F2F2F2"; MED_GREY="CCCCCC"; LT_PURPLE="EDE7F6"; DARK_TEXT="1A1A2E"
HI_GREEN="ABEBC6"; HOT="FFD6D6"

def hf(c): return PatternFill("solid", fgColor=c)
def fn(c=DARK_TEXT, b=False, s=9): return Font(color=c, bold=b, size=s, name="Calibri")
th = Side(style="thin", color=MED_GREY)
brd = Border(left=th, right=th, top=th, bottom=th)
def wa(h="left", v="top"): return Alignment(horizontal=h, vertical=v, wrap_text=True)

def title_row(ws, txt, cols, bg=DARK_NAVY, r=1, h=32, sz=13):
    ws.merge_cells(f"A{r}:{get_column_letter(cols)}{r}")
    c = ws[f"A{r}"]
    c.value, c.font = txt, Font(name="Calibri", bold=True, size=sz, color=WHITE)
    c.fill, c.alignment = hf(bg), Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = h

def subtitle(ws, txt, cols, bg=TEAL, r=2, h=26):
    ws.merge_cells(f"A{r}:{get_column_letter(cols)}{r}")
    c = ws[f"A{r}"]
    c.value, c.font = txt, Font(name="Calibri", bold=True, size=10, color=WHITE, italic=True)
    c.fill, c.alignment = hf(bg), Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = h

def hdr(ws, cols, row, bg=MID_BLUE):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
        c.fill, c.alignment, c.border = hf(bg), wa("center","center"), brd
    ws.row_dimensions[row].height = 30

ALT = [LT_GREY, WHITE]
FIT_FILL = {"HIGH":HI_GREEN, "MED-HIGH":LT_GREEN, "MEDIUM":LT_BLUE,
            "MED-LOW":LT_AMBER, "LOW":LT_RED}

# ══════════════════════════════════════════════════════════════════
# MASTER DATA  (32 original + 12 new = 44 companies)
# fields: name, sector, india, emp_prog, emp_src, csr_arm, fit, tier, prio
# ══════════════════════════════════════════════════════════════════
# fit: HIGH / MED-HIGH / MEDIUM / MED-LOW / LOW
M = [
# ── TIER A — HIGH FIT (real India env/waste/water/conservation funding) ──
("Coca-Cola","Beverages / FMCG","Coca-Cola India (Gurugram); Anandana Foundation",
 "Matching Gifts 2:1, up to $20,000/yr","doublethedonation.com/matching-gifts/the-coca-cola-company",
 "Anandana – Coca-Cola India Foundation","HIGH","A",1),
("PepsiCo","Beverages / FMCG","PepsiCo India (Gurugram); PepsiCo Foundation",
 "Foundation Matching 1:1 (2:1 if 50+ vol hrs), $10,000/yr","pepsico.yourcause.com",
 "PepsiCo Foundation / PepsiCo India CSR","HIGH","A",1),
("Amazon","E-commerce / Tech","Amazon India (Bengaluru/Hyderabad)",
 "Employee Giving 1:1, $5,000/yr","amazon.jobs/content/en/teams/ccr/aci",
 "Amazon Right Now Climate Fund (RNCF) – India","HIGH","A",2),
("Apple","Technology","Apple India (Bengaluru/Hyderabad)",
 "Matching Gifts 1:1 (2:1 special), $10,000/yr","forms.matchinggifts.com/AppleGuide.pdf",
 "Apple Environment / India partners (WWF-India, Acumen)","HIGH","A",2),
("NVIDIA","Technology / Semiconductors","NVIDIA Graphics Pvt Ltd (Bengaluru/Pune/Hyderabad)",
 "Inspire 365 Foundation 1:1, $10,000/yr","nvidia.com/en-us/foundation/",
 "NVIDIA India CSR (biodiversity/water/mangroves)","HIGH","A",1),
# ── TIER B — MED-HIGH / MEDIUM (env secondary; open or semi-open route) ──
("Cisco","Technology","Cisco India (Bengaluru) – statutory 2% CSR",
 "Matching 1:1 (Time2Give), $25,000/yr","benevity.com/client-stories/cisco-elevated-social-impact-with-benevity",
 "Cisco India Cash Grants (OPEN LOI portal)","MEDIUM","B",1),
("Salesforce","Technology / SaaS","Salesforce India (Hyderabad/Bengaluru/Mumbai)",
 "1-1-1 / Philanthropy Cloud 1:1, $10,000/yr","salesforce.com/company/philanthropy/",
 "Salesforce.org India (lake restoration, afforestation)","MED-HIGH","B",2),
("American Express","Financial Services","Amex India (Gurugram) – big GBT/tech hub",
 "Give2Gether 1:1, $10,000/yr","globalgiving.org/amex-give2gether/",
 "Amex Foundation – Environment pillar (Program Vasundhara)","MED-HIGH","B",2),
("Google","Technology","Google India (Bengaluru/Hyderabad/Gurugram)",
 "Employee Giving (go/give) 1:1, $12,000/yr","google.benevity.org/terms-and-conditions",
 "Google.org APAC Sustainability Seed Fund","MED-HIGH","B",2),
("Microsoft","Technology","Microsoft India (Hyderabad/Bengaluru/Noida)",
 "Give Match 1:1, $15,000/yr","microsoft.com/.../philanthropies/employee-engagement",
 "Microsoft India CSR + Climate Innovation Fund / AI for Good","MEDIUM","B",2),
("HP","Technology / Hardware","HP India (Bengaluru/Gurugram)",
 "HP Foundation Cash Match 1:1; Product 3:1, $5,000/yr","hp.com/us-en/hp-information/hp-foundation.html",
 "HP India CSR (forest restoration, recycled plastics)","MEDIUM","B",2),
("IBM","Technology","IBM India (Bengaluru + nationwide)",
 "Matching Grants 1:1, $10,000/yr","ibm.com/responsibility/programs/volunteerism-giving",
 "IBM Sustainability Accelerator (OPEN RFP portal)","MEDIUM","B",1),
# ── TIER C — MED-LOW (India CSR skews edu/health; env incidental) ──
("SAP","Technology / SaaS","SAP Labs India (Bengaluru – 2nd largest SAP site)",
 "SAP Together Matching 1:1, $750/yr ($20/vol hr)","doublethedonation.com/matching-gifts/sap-america-inc",
 "SAP Labs India CSR (env = minor pillar)","MED-LOW","C",3),
("Intel","Semiconductors","Intel India (Bengaluru)",
 "Intel Involved 1:1, $1,000/yr combined (2025)","intel.com/.../intel-foundation-employee-generosity.html",
 "Intel Foundation (email org info) + Intel India local office","MED-LOW","C",3),
("Goldman Sachs","Financial Services","GS India (Bengaluru/Hyderabad/Mumbai)",
 "Matching 1:1, $20,000/yr (GS Gives)","goldmansachs.com/citizenship/goldman-sachs-gives/",
 "GS India CSR + GS Gives (env = 'One for Each Tree')","MED-LOW","C",3),
("JPMorgan Chase","Financial Services","JPMC India (Mumbai/Bengaluru/Hyderabad)",
 "Matching 1:1, $2,000/yr","forms.matchinggifts.com/ChaseGuide.Links.pdf",
 "JPMC India CSR + Global Philanthropy (Mumbai PO)","MED-LOW","C",3),
("Bank of America","Financial Services","BofA India (Mumbai/Gurugram/Hyderabad/Chennai)",
 "Matching 1:1, $5,000/yr","about.bankofamerica.com/.../matching-gifts-features-and-eligibility",
 "BofA India CSR (env pillar; intl funding by invitation)","MED-LOW","C",3),
("Mastercard","Financial Services","Mastercard India (Pune/Gurugram/Vadodara)",
 "Matching 1:1, $5,000/yr + Impact Fund","mastercard.com/.../community-impact.html",
 "Mastercard Center for Inclusive Growth (e-waste precedent)","MED-LOW","C",3),
("Deloitte","Professional Services","Deloitte India (nationwide)",
 "Foundation Matching 1:1, $32,500/yr","files.doublethedonation.com/forms/Deloitte-guidelines-1712687004.pdf",
 "Deloitte India CSR (WorldClimate; Rainmatter partner)","MED-LOW","C",3),
("NetApp","Technology / Storage","NetApp India (Bengaluru – large R&D)",
 "NetApp Cares 1:1, ~$2,500–5,000/yr","netapp.com/responsibility/social-impact/netapp-cares/",
 "NetApp India CSR (water at Akshaya Patra; mostly edu)","MED-LOW","C",3),
("Adobe","Technology / SaaS","Adobe India (Noida/Bengaluru)",
 "Create Change 1:1, $10,000/yr combined","adobe.com/corporate-responsibility/community.html",
 "Adobe Foundation Employee Community Fund (employee-nominated)","LOW","C",3),
# ── TIER D — LOW FIT (no meaningful India env route) ──
("Wells Fargo","Financial Services","WF India Solutions/EGS (Bengaluru/Hyderabad/Chennai)",
 "Matching 1:1, $2,000/yr (16h VTO)","wellsfargo.com/.../community-giving/giving-faqs/",
 "WF India CSR (overwhelmingly EDUCATION)","LOW","D",3),
("Texas Instruments","Semiconductors","TI India (Bengaluru)",
 "TI Foundation Matching 1:1, up to $30,000/yr","ti.com/about-ti/citizenship-community/giving.html",
 "TI India CSR (Akshaya Patra meals + WiSH)","LOW","D",3),
("Qualcomm","Semiconductors","Qualcomm India (Hyderabad/Bengaluru/Chennai/Noida)",
 "Employee Giving 1:1, $1,000–5,000/yr","doublethedonation.com/matching-gifts",
 "Qualcomm India CSR / Wireless Reach (tech/edu)","LOW","D",3),
("Nike","Apparel / Retail","Nike India (limited corporate footprint)",
 "Give Your Best 1:1 (2:1 GivingTuesday), $25,000/yr","about.nike.com/.../matching-our-employees-impact",
 "Nike Community Impact Fund (sport-for-kids; not env grants)","LOW","D",3),
("Johnson & Johnson","Pharma / Healthcare","J&J India (Mumbai)",
 "Matching 2:1, up to $20,000/yr","jnj.com/caring-and-giving",
 "J&J India CSR (HEALTH only; no unsolicited proposals)","LOW","D",3),
("Pfizer","Pharma","Pfizer India (Mumbai/Chennai)",
 "Foundation Matching 1:1, $5,000/yr","pfizer.com/about/responsibility/global-impact/give-forward",
 "Pfizer India CSR (health + school sanitation)","LOW","D",3),
("Merck (MSD)","Pharma","MSD India (Mumbai/Pune/Hyderabad)",
 "Partnership for Giving 1:1, $30,000/yr","merck.com/.../philanthropy/strengthening-our-communities/",
 "MSD India (health; water via Safe Water Network) – portal only","MED-LOW","D",3),
("Apollo Global Mgmt","Private Equity / Finance","Apollo India (Mumbai)",
 "Citizenship Grants (matching + volunteer rewards)","sec.gov DEF 14A (Apollo)",
 "Apollo Opportunity Foundation (career/workforce only)","LOW","D",3),
("Stryker","Medical Devices","Stryker India (Gurugram/Bengaluru)",
 "Stryker Impact Platform 1:1","stryker.com/us/en/about/corporate-responsibility/giving-back.html",
 "Stryker Johnston Foundation (US/health-focused)","LOW","D",3),
("Airbnb","Travel / Tech","Airbnb India (Gurugram)",
 "Community Fund + Matching 1:1, $1,000/yr","news.airbnb.com/2025-community-fund/",
 "Airbnb Community Fund (host-NOMINATED; funds env NGOs in India)","LOW","D",3),
("Dell Technologies","Technology","Dell India (Bengaluru/Hyderabad)",
 "Employee Giving 1:1, $10,000/yr","dell.yourcause.com",
 "Dell India CSR (digital/edu; env incidental)","LOW","D",3),
]

# ══════════════════════════════════════════════════════════════════
# SHEET 1 — MASTER SHEET (all companies collated)
# ══════════════════════════════════════════════════════════════════
ws1 = wb.active; ws1.title = "MASTER SHEET"
ws1.sheet_view.showGridLines = False; ws1.freeze_panes = "A4"
COLS1 = ["#","Company","Sector","India Presence","Employee Giving Program (PROOF it runs)",
         "Proof Source","Direct-Giving Arm / Foundation (the real entry)","Env Fit","Tier","Priority"]
title_row(ws1, "WASTE WARRIORS — MASTER SHEET: 32 COMPANIES, EMPLOYEE-GIVING PROOF + DIRECT-GIVING ENTRY", len(COLS1))
subtitle(ws1, "Goal = DIRECT corporate/foundation/India-CSR giving (NOT Benevity platform). Fit rated for an Indian environmental / waste / conservation NGO.", len(COLS1))
hdr(ws1, COLS1, 3)
for i, row in enumerate(M, 1):
    name, sector, india, emp, src, csr, fit, tier, prio = row
    r = i + 3
    vals = [i, name, sector, india, emp, src, csr, fit, tier, f"P{prio}"]
    base = ALT[i % 2]
    for ci, v in enumerate(vals, 1):
        c = ws1.cell(r, ci, v)
        c.alignment = wa("center" if ci in (1,8,9,10) else "left"); c.border = brd
        c.font = fn(DARK_TEXT, b=(ci==2), s=9)
        if ci == 8: c.fill = hf(FIT_FILL[fit]); c.font = fn(DARK_TEXT, b=True, s=9)
        elif ci == 10:
            c.fill = hf(HOT if prio==1 else LT_AMBER if prio==2 else LT_GREY)
            c.font = fn(DEEP_RED if prio==1 else AMBER if prio==2 else DARK_TEXT, b=True, s=9)
        else: c.fill = hf(base)
    ws1.row_dimensions[r].height = 46
cw1 = [4,18,20,30,34,34,38,10,7,9]
for i,w in enumerate(cw1,1): ws1.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════════
# SHEET 2 — DIRECT ENTRY STRATEGY (per company, the heart)
# fields: company, route, team, entry_point, decision_maker, angle, legal, contact
# ══════════════════════════════════════════════════════════════════
S = [
("Coca-Cola",
 "Anandana (Coca-Cola India Foundation) – DIRECT grants to NGO partners",
 "Anandana Foundation programme team (water + waste)",
 "Apply / pitch via anandana.org contact form; reference their existing NGO model (SM Sehgal Fdn, SARA, Harikritika)",
 "Programme Director / CSR Lead – Anandana / Coca-Cola India",
 "Position WW as a waste-management + water-body restoration partner in the Himalayas; Anandana's core is water + waste = direct fit",
 "India CSR (INR, Companies Act 2%) — no FCRA needed",
 "anandana.org  •  CSRBox profile: csrbox.org (Anandana Foundation)"),
("PepsiCo",
 "PepsiCo Foundation / PepsiCo India CSR – solicits proposals for large grants (>$100K)",
 "PepsiCo India Sustainability / Foundation team",
 "Pitch as a plastic-waste & waste-worker (Safai Saathi) partner; mirror Tidy Trails / Recity / Smile Foundation 'Waste no More'",
 "Director – Sustainability / PepsiCo Foundation India",
 "Frame WW around plastic-waste collection + waste-worker livelihoods (their proven model); add Himalayan hill-town angle",
 "India CSR (INR) + PepsiCo Foundation grant",
 "pepsicoindia.co.in/our-impact/community/pepsico-foundation"),
("Amazon",
 "Amazon Right Now Climate Fund (RNCF) – $3.28M already deployed in India",
 "Amazon India Sustainability / Worldwide Sustainability team",
 "WARM INTRO needed (no open portal). Target via existing RNCF partners (Centre for Wildlife Studies, ICLEI). Pitch ecosystem restoration + plastic-waste removal",
 "Head of Sustainability – Amazon India",
 "Himalayan forest/biodiversity restoration + plastic-waste removal (mirrors their Mumbai mangrove + Western Ghats grants)",
 "India CSR + corporate climate fund",
 "aboutamazon.in/news/sustainability  •  theclimatepledge.com (RNCF)"),
("Apple",
 "Apple Environment – India partners via WWF-India + Acumen catalytic grants",
 "Apple Environment team (VP Sarah Chandler) + India partners",
 "WARM INTRO via WWF-India or Acumen (they granted 6 green/waste/circular enterprises). No open NGO portal",
 "Environment Programs Lead – Apple / WWF-India contact",
 "Circular economy + waste-management enterprise angle (their May-2026 India launch funds exactly this)",
 "Routed via partner NGOs (WWF-India / Acumen)",
 "apple.com/in/newsroom (2026 India environment initiatives)"),
("NVIDIA",
 "NVIDIA India CSR (NVIDIA Graphics Pvt Ltd) – statutory 2%",
 "NVIDIA India CSR committee",
 "Pitch to India CSR team; their stated pillars = biodiversity, habitat conservation, water-body rejuvenation, mangrove restoration",
 "India CSR / Community Lead – NVIDIA India",
 "Himalayan biodiversity + water-body rejuvenation directly matches their published India CSR pillars",
 "India CSR (INR, 2%) — no FCRA needed",
 "nvidia.com/en-in/csr-india/ (India CSR Policy PDF lists contact)"),
("Cisco",
 "Cisco India Cash Grants — OPEN Letter-of-Inquiry portal (best mechanics on the list)",
 "Cisco India CSR / Social Impact team",
 "APPLY DIRECTLY: eligibility quiz → submit LOI / concept note. Frame env via climate-resilience + waste-worker livelihoods",
 "Cisco India CSR Manager / Social Impact",
 "Climate-resilient livelihoods + tech-enabled waste management (fit their education/economic-empowerment/resilience themes)",
 "India CSR (INR, statutory 2%) — open application",
 "cisco.com/c/en_in/about/csr/india-cash-grants.html (+ India Cash Grant Policy 2025 PDF)"),
("Salesforce",
 "Salesforce.org India – direct grants to env partners",
 "Salesforce.org India / Philanthropy team",
 "Pitch via Salesforce.org partnerships; they fund lake restoration (Malligavad) + 400k+ sapling afforestation in India",
 "Philanthropy / Salesforce.org Lead – India",
 "Water-body & forest restoration in the Himalayas (mirrors their Bengaluru/Hyderabad lake-restoration grants)",
 "India CSR + Salesforce.org grant",
 "salesforce.com/solutions/philanthropy/corporate-social-responsibility/"),
("American Express",
 "Amex Foundation (Environment is a stated pillar) + Amex India CSR",
 "Amex India CSR + Foundation Environment programme",
 "Reference live env program 'Program Vasundhara' (Safe Water Network) – water, biodiversity, climate resilience",
 "CSR / Corporate Sustainability Lead – Amex India",
 "Water + biodiversity + climate-resilience (their Vasundhara model) extended to the Himalayas",
 "India CSR (INR) + Amex Foundation",
 "about.americanexpress.com/corporate-responsibility/ (India CSR policy PDF)"),
("Google",
 "Google.org APAC Sustainability Seed Fund ($5M) + Impact Challenges",
 "Google.org India team",
 "Watch impactchallenge.withgoogle.com for open calls; pitch water/climate-resilience (grantees: INREM, WELL Labs, CRDF, MHT)",
 "Google.org Lead – India",
 "Climate-resilience + water (tech/data angle helps); pure waste-mgmt is a stretch — lead with water/climate data",
 "Google.org grant + India CSR",
 "blog.google/intl/en-in (Google.org India changemakers)"),
("Microsoft",
 "Microsoft India CSR + Climate Innovation Fund + AI for Good/Earth grants",
 "Microsoft India CSR + Sustainability team",
 "India CSR for direct grants; Climate Innovation Fund (form) skews tech/AI. Lead with data/AI-for-environment angle",
 "CSR Lead – Microsoft India",
 "AI/data-for-environment (waste mapping, forest monitoring) to fit their tech-for-good slant",
 "India CSR (INR) + global climate fund (investment)",
 "microsoft.com/en-in/corporate-social-responsibility"),
("HP",
 "HP India CSR (Sustainable Impact / Living Progress)",
 "HP India CSR / Sustainability team",
 "Partner-led (no open portal); reference their forest restoration (WWF, Arbor Day) + recycled-plastics work",
 "Sustainability / CSR Lead – HP India",
 "Forest restoration + circular plastics in the Himalayas (mirrors their WWF forest partnerships)",
 "India CSR (INR)",
 "hp.com/in-en/sustainable-impact/climate-action.html"),
("IBM",
 "IBM Sustainability Accelerator — OPEN public RFP/Proposal Submission Portal (annual call ~Apr)",
 "IBM Corporate Social Responsibility / Accelerator team",
 "APPLY DIRECTLY in the annual call. Note: it's pro-bono tech (AI/cloud) + some cash, for environment-vulnerable communities",
 "IBM CSR / Sustainability Accelerator Manager",
 "Tech-enabled climate adaptation for vulnerable Himalayan communities (their grantee model, e.g. TNC India)",
 "Programme grant + pro-bono tech (global + India)",
 "ibm.com/responsibility/programs/ibm-impact-accelerator (apply via IBM Proposal Portal)"),
("SAP",
 "SAP Labs India CSR (env is a minor pillar)",
 "SAP Labs India CSR team (CSR head APJ/India: Gunjan Patel)",
 "LinkedIn route to named CSR head; frame as social-enterprise / tech-for-good (their core)",
 "CSR Head APJ/India – SAP Labs",
 "Tech-enabled waste social enterprise (their Northeast Growth Labs model) — env framing weak, use enterprise angle",
 "India CSR (INR)",
 "sap.com/india/about/csr.html"),
("Intel",
 "Intel Foundation (email org info) routed to Intel India local office",
 "Intel India site/community team",
 "No unsolicited proposals — email foundation, then push via Intel India Bengaluru office relationships",
 "Community Affairs – Intel India",
 "Env not a headline India theme — only pursue with a Bengaluru-area project or warm intro",
 "India CSR (INR) via local office",
 "intel.foundation@intel.com  •  +1 503-696-8080"),
("Goldman Sachs",
 "GS India CSR + GS Gives (DAF) + Office of Corporate Engagement",
 "GS India Office of Corporate Engagement",
 "Partners 40+ NGOs/yr. Env incidental ('One for Each Tree' afforestation). Pitch ONLY with jobs/livelihoods angle",
 "VP – Corporate Engagement, GS India",
 "Waste-worker livelihoods / women's economic empowerment (their core) + afforestation add-on",
 "India CSR (INR)",
 "goldmansachs.com/worldwide/india/citizenship/"),
("JPMorgan Chase",
 "JPMC India CSR + Global Philanthropy (Mumbai-based Program Officer)",
 "JPMC Global Philanthropy – Mumbai team",
 "They co-create proposals with NGOs. Core = jobs/financial health; climate possible. Reach Mumbai PO",
 "Program Officer, Global Philanthropy – India",
 "Green jobs / circular-economy livelihoods (fits their jobs + climate grantmaking)",
 "India CSR (INR)",
 "jpmorgan.com (India CSR Policy PDF)"),
("Bank of America",
 "BofA India CSR + Charitable Foundation (env pillar, intl funding BY INVITATION)",
 "BofA India CSR team",
 "Cultivate relationship → invitation (no open intl applications). Env is a real pillar",
 "CSR Lead – BofA India",
 "Environmental sustainability (their stated pillar) — needs warm intro before any ask",
 "India CSR (INR)",
 "business.bofa.com/in (India CSR policy PDF)"),
("Mastercard",
 "Mastercard Center for Inclusive Growth / Impact Fund (Spark Grants)",
 "Mastercard Center / Impact Fund team",
 "Reference India e-waste precedent (Karo Sambhav circular economy). Core = financial inclusion",
 "Programme Lead – Center for Inclusive Growth",
 "Circular economy / e-waste + waste-worker financial inclusion (their precedent)",
 "Impact Fund grant + India CSR",
 "mastercardcenter.org (contact form)"),
("Deloitte",
 "Deloitte India CSR (WorldClimate + WorldClass)",
 "Deloitte India CSR / WorldClimate team",
 "Grantmaking is partnership-led + pro-bono; partnered Rainmatter Foundation for climate action",
 "CSR / WorldClimate Lead – Deloitte India",
 "Community-led climate action (their Rainmatter model) + pro-bono advisory for WW",
 "India CSR (INR) + pro-bono",
 "deloitte.com/in/en/about/people/social-responsibility.html"),
("NetApp",
 "NetApp India CSR (Global Giving team)",
 "NetApp India CSR (MD India: Ravi Chhabria)",
 "Env is a listed focus (water conservation) but spend skews hunger/edu (Akshaya Patra). LinkedIn route",
 "CSR Lead – NetApp India",
 "Water conservation tie-in; otherwise low priority",
 "India CSR (INR)",
 "netapp.com (India CSR Policy PDF)"),
("Adobe",
 "Adobe Foundation Employee Community Fund (employee-NOMINATED, Noida/Bengaluru)",
 "Adobe India action teams (employee-led)",
 "Borders on employee-giving (you're avoiding this). Needs an Adobe-employee champion to nominate WW",
 "Adobe India CSR / action-team lead",
 "Only viable with an internal employee champion — deprioritize for direct strategy",
 "Employee-nominated grant",
 "socialimpact@adobe.com  •  adobe.com/corporate-responsibility/community/community-fund.html"),
("Wells Fargo",
 "WF India CSR (overwhelmingly EDUCATION)",
 "WF India Solutions/EGS CSR team",
 "India spend is education-dominated (Mantra, Agastya, Room to Read). Parent env funding intl = by invitation",
 "CSR Lead – WF India Solutions",
 "Weak env fit — only via an education+environment blended pitch",
 "India CSR (INR) — education-skewed",
 "wellsfargojobs.com (India CSR projects 2024-25 PDF)"),
("Texas Instruments",
 "TI India CSR (Akshaya Patra meals + WiSH women-in-semiconductors)",
 "TI India CSR / TI Foundation",
 "Env = own-ops emissions only; no community env grants. Very low fit",
 "CSR Lead – TI India",
 "No realistic env route — deprioritize",
 "India CSR (INR) — non-env",
 "ti.com/about-ti/citizenship-community/giving.html"),
("Qualcomm",
 "Qualcomm India CSR / Wireless Reach (tech-for-development)",
 "Qualcomm India CSR team",
 "Tech/education focus; SootSwap clean cookstoves is the only env-adjacent program",
 "CSR Lead – Qualcomm India",
 "Only via a tech-for-environment (clean cookstove / sensor) framing",
 "India CSR (INR)",
 "qualcomm.com/company/locations/india/CSRPolicy"),
("Nike",
 "Nike Community Impact Fund (sport-for-kids; NOT env grants)",
 "Nike NCIF (locally employee-reviewed)",
 "'Move to Zero' is internal supply-chain sustainability, not grants. Low fit",
 "Community Impact – Nike",
 "No env grantmaking route — deprioritize",
 "Community fund (sport)",
 "about.nike.com/en/mission/initiatives/nike-community-impact-fund"),
("Johnson & Johnson",
 "J&J India CSR (HEALTH only) — Foundation does NOT accept unsolicited proposals",
 "J&J India CSR (health mission)",
 "Health-only mission; nursing/education/health. No env route",
 "CSR Lead – J&J India",
 "No env fit — deprioritize",
 "India CSR (INR) — health only",
 "jnj.in/csr-projects"),
("Pfizer",
 "Pfizer India CSR (health + school sanitation)",
 "Pfizer India CSR team",
 "Sanitation is the only adjacency; no environment/waste/conservation",
 "CSR Lead – Pfizer India",
 "Only via a sanitation/WASH framing — weak",
 "India CSR (INR) — health/WASH",
 "pfizer.co.in/our-community-efforts/csr-initiatives"),
("Merck (MSD)",
 "MSD India – health + Safe Water Network water-access ($1.5M Andhra); global portal only (msd.versaic.com)",
 "MSD India CSR / Foundation",
 "No unsolicited (online portal only). The ONLY angle = water-access",
 "CSR Lead – MSD India",
 "Water-access framing (their Safe Water Network model) — narrow but possible",
 "India CSR (INR) — via water angle",
 "msd.versaic.com (proposal portal)"),
("Apollo Global Mgmt",
 "Apollo Opportunity Foundation (career/workforce/economic empowerment only)",
 "Apollo India (Mumbai)",
 "No environmental grantmaking (clean-energy is investment, not philanthropy)",
 "Foundation Lead – Apollo",
 "No env route — deprioritize",
 "Foundation grant — non-env",
 "apollo.com/impact"),
("Stryker",
 "Stryker Johnston Foundation (US/health-focused) + product donations",
 "Stryker Johnston Foundation",
 "US/health focus, not India environment. Low fit",
 "Foundation grants team",
 "No env fit — deprioritize",
 "Foundation grant — non-env",
 "grants@strykerjohnstonfoundation.org"),
("Airbnb",
 "Airbnb Community Fund — does NOT accept applications; recipients are HOST-NOMINATED",
 "Local Airbnb Host Club / Host Advisory Board",
 "Cultivate Himachal/Uttarakhand/Corbett-area Airbnb HOSTS to nominate WW (2026 cohort funded Keystone, Stree Mukti waste-pickers, Goonj)",
 "Host Club organizers (Dharamshala/Mussoorie/Corbett)",
 "Sustainability + waste-pickers (their funded themes); access ONLY via host nomination",
 "Community Fund grant (host-nominated)",
 "news.airbnb.com/airbnb-community-fund-grants..."),
("Dell Technologies",
 "Dell India CSR (digital/education; env incidental)",
 "Dell India CSR team",
 "Digital-inclusion focus; env not a headline theme",
 "CSR Lead – Dell India",
 "Only via a digital/tech-for-environment framing — weak",
 "India CSR (INR)",
 "dell.com/.../social-impact/transforming-lives/employee-led-impact.htm"),
]

ws2 = wb.create_sheet("Direct Entry Strategy")
ws2.sheet_view.showGridLines = False; ws2.freeze_panes = "B4"
COLS2 = ["Company","The DIRECT Route (not platform)","Team / Dept to Target",
         "Entry Point — how to get in the door","Decision-Maker to Find (LinkedIn)",
         "Tailored Proposal Angle for Waste Warriors","Legal Mechanism","Real Contact / URL"]
title_row(ws2, "WASTE WARRIORS — DIRECT ENTRY STRATEGY: HOW TO GET FIRSTHAND CONTACT WITH EACH COMPANY", len(COLS2), FOREST)
subtitle(ws2, "For every company: skip the employee portal — go straight to the India CSR team / foundation / ESG fund and pitch a proposal.", len(COLS2), MID_GREEN)
hdr(ws2, COLS2, 3, MID_GREEN)
# fit lookup for coloring
fitmap = {r[0]: r[6] for r in M}
for i, row in enumerate(S, 1):
    r = i + 3
    fit = fitmap.get(row[0], "MEDIUM")
    rowfill = FIT_FILL[fit]
    for ci, v in enumerate(row, 1):
        c = ws2.cell(r, ci, v)
        c.alignment = wa(); c.border = brd
        c.font = fn(DARK_TEXT, b=(ci==1), s=9)
        c.fill = hf(rowfill if ci==1 else ALT[i%2])
    ws2.row_dimensions[r].height = 78
cw2 = [16,34,26,40,26,40,24,34]
for i,w in enumerate(cw2,1): ws2.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════════
# SHEET 3 — 12 NEW COMPANIES (added) — best India env + employee giving
# fields: name, why, india_csr_env, emp_giving, entry, fit, open_portal
# ══════════════════════════════════════════════════════════════════
NEW = [
("Hindustan Unilever (Unilever)","FMCG","HIGH","YES (open windows via HUF)",
 "Hindustan Unilever Foundation (HUF) — 'Water for Public Good' + Project Circular Bharat (plastic) + Safai Saathi waste-picker inclusion in 10+ cities",
 "Unilever runs employee matching/volunteering",
 "Best thematic match: water + plastic waste. Pitch HUF as a water + waste-picker partner",
 "csrbox.org (HUF profile) • hul.co.in/sustainability/plastics"),
("HSBC","Banking","HIGH","Partner via MoU after due diligence",
 "HSBC India CSR + global HSBC Water Programme (WWF, WaterAid, Earthwatch) + Climate Solutions Partnership in 8 Indian states",
 "Employee citizen-science engagement",
 "Water + climate via the WWF/WaterAid model; cultivate via due-diligence MoU",
 "about.hsbc.co.in/.../corporate-social-responsibility-policy"),
("3M","Industrials / Tech","HIGH","YES — OPEN 3M India grant portal",
 "3M India Foundation — explicit Environmental Programs pillar (open grant guidelines)",
 "YES — employee/retiree Matching Gift Program since 1971",
 "CLEANEST combo: open direct-apply env portal + matching. APPLY DIRECTLY",
 "3mindia.in/3M/en_IN/gives-in/3m-foundation/grant-guidelines/"),
("Cummins","Industrials","HIGH","YES — open CSR grant calls",
 "Cummins India Foundation (CIF) — Energy & Environment pillar (tree planting, water-body cleaning, waste mgmt); grants up to ₹20 lakh",
 "Every Employee Every Community (EEEC) + matching",
 "Direct env grants in India; apply to CIF open calls",
 "cummins.com/en/in/.../cummins-india-foundation"),
("Bosch","Industrials / Tech","HIGH (geo: South India)","Partner (NGO eligibility stated)",
 "Bosch India Foundation (BIF) — Environment & Water Conservation pillar (lake rejuvenation, 100k+ trees)",
 "'Act of Kindness' volunteering",
 "Strong env, but geo-limited to Bengaluru/Coimbatore — fit only if WW has a South-India project",
 "bosch.in/.../environment-water-conservation/"),
("Schneider Electric","Industrials / Energy","HIGH","YES — OPEN 'Submit Your Project' portal",
 "Schneider Electric India Foundation (SEIF) — Environment + energy access (solar microgrids/pumps)",
 "Dollars for Doers ($25/hr)",
 "Open global project portal + India env pillar; APPLY DIRECTLY",
 "seifoundation.se.com • submityourproject.foundation.schneider-electric.com"),
("Nestlé India","FMCG","HIGH (geo MATCH!)","Pitch as expansion partner",
 "Project Hilldaari (with PLAN Foundation + Recity) — solid-waste mgmt in MUSSOORIE, DALHOUSIE, PALAMPUR, DARJEELING, MUNNAR (~47k MT diverted)",
 "Employee volunteering cleanups",
 "★ HIGHEST geo+theme fit: hill-town waste = WW's exact footprint. Pitch as a Hilldaari expansion partner",
 "thecsruniverse.com (Nestlé Hilldaari) • csrbox.org"),
("EY","Professional Services","MED-HIGH","Accelerator (grant + skills)",
 "EY Foundation + EY Ripples + TRANSFORM accelerator (scaled TrashCon plastic-waste 100x)",
 "Higher-education matching gifts + Ripples volunteering",
 "Accelerator model — good for tech-enabling WW's waste operations",
 "ey.com/en_in/about-us/corporate-responsibility"),
("Diageo (United Spirits)","Beverages","MED-HIGH","India CSR partner",
 "United Spirits 'Society 2030' — water replenishment (479k cu.m/yr) + 100% EPR plastic-waste collection (24k T)",
 "Employee programs (parent Diageo)",
 "Water + plastic-EPR partner; pitch via United Spirits CSR",
 "diageoindia.com/en/investors/corporate-social-responsibility"),
("P&G","FMCG","MED-HIGH","India CSR partner",
 "P&G India — plastic-packaging recycling (collects > it uses) + water restoration + Children's Safe Drinking Water",
 "Employee giving programs",
 "Plastic recycling + water restoration; pitch via P&G India CSR",
 "in.pg.com/environmental-sustainability • in.pg.com/community-impact"),
("HCLTech","IT Services (Indian)","HIGH","YES — OPEN competitive grant",
 "HCLTech Grant — open ₹5 Crore / 4-yr grants in BIODIVERSITY, WATER, Health, Education + 'Harit' env program",
 "'Power of One' payroll-giving + volunteering",
 "★ APPLY DIRECTLY to the Biodiversity/Water track — exactly WW's profile",
 "hclfoundation.org/hcltech-grant"),
("Wipro","IT Services (Indian)","HIGH","YES — OPEN ecology grants",
 "Wipro Foundation Ecology — open Community Ecology Program (afforestation, biodiversity, restoration, climate adaptation) + Urban Ecology Small Grants",
 "Employee giving exists",
 "★ APPLY DIRECTLY — eligibility 80G/12A + 3yrs ecology experience = WW qualifies",
 "wiprofoundation.org/our-initiatives/ecology/"),
]
ws3 = wb.create_sheet("12 NEW Companies")
ws3.sheet_view.showGridLines = False; ws3.freeze_panes = "B4"
COLS3 = ["Company","Sector","Env Fit","Open Direct Route?","India CSR / Env Program (PROOF)",
         "Employee Giving (2nd box)","Why Added / How to Approach","Entry Point / Source"]
title_row(ws3, "WASTE WARRIORS — 12 NEW COMPANIES: STRONG INDIA ENVIRONMENTAL CSR + EMPLOYEE GIVING", len(COLS3), PURPLE)
subtitle(ws3, "Not in the original 32. Prioritized for REAL environmental/waste/water funding in India. ★ = open application you can submit directly NOW.", len(COLS3), "5E35B1")
hdr(ws3, COLS3, 3, "5E35B1")
for i, row in enumerate(NEW, 1):
    r = i + 3
    name, sector, fit, openr, prog, emp, why, entry = row
    fitkey = fit.split()[0]
    fitkey = {"MED-HIGH":"MED-HIGH","HIGH":"HIGH","MEDIUM":"MEDIUM","MED-LOW":"MED-LOW","LOW":"LOW"}.get(fitkey,"MED-HIGH")
    vals = [name, sector, fit, openr, prog, emp, why, entry]
    for ci, v in enumerate(vals, 1):
        c = ws3.cell(r, ci, v)
        c.alignment = wa("center" if ci in (3,) else "left"); c.border = brd
        c.font = fn(DARK_TEXT, b=(ci==1), s=9)
        if ci==3: c.fill = hf(FIT_FILL.get(fitkey, LT_GREEN)); c.font = fn(DARK_TEXT,b=True,s=9)
        elif ci==4 and "YES" in openr: c.fill = hf(HI_GREEN); c.font=fn(FOREST,b=True,s=9)
        else: c.fill = hf(ALT[i%2])
    ws3.row_dimensions[r].height = 74
cw3 = [22,16,12,22,42,22,40,34]
for i,w in enumerate(cw3,1): ws3.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════════
# SHEET 4 — PROOF & SOURCES (employee giving + direct giving)
# ══════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Proof & Sources")
ws4.sheet_view.showGridLines = False; ws4.freeze_panes = "A4"
COLS4 = ["Company","PROOF — Employee Giving runs (specific)","PROOF — Direct/Foundation/India-CSR giving","Source(s)","Confidence"]
title_row(ws4, "WASTE WARRIORS — PROOF & SOURCES: EVIDENCE EACH COMPANY DOES EMPLOYEE GIVING + DIRECT GIVING", len(COLS4), TEAL)
subtitle(ws4, "Employee-giving proof from official program pages/PDFs; direct-giving proof from India CSR pages, foundation pages & CSRBox.", len(COLS4), "0A6E80")
hdr(ws4, COLS4, 3, "0A6E80")
PROOF = [
("Coca-Cola","2:1 Matching Gifts (official program guideline PDF), up to $20k/yr",
 "Anandana (Coca-Cola India Foundation) gives monetary grants to NGO partners; 500+ water projects; National Water Award",
 "doublethedonation.com/.../coca-cola; thecsruniverse.com; anandana.org","HIGH"),
("PepsiCo","Foundation Matching 1:1 (2:1 for 50+ vol hrs) via YourCause, $10k/yr",
 "PepsiCo Foundation funds plastic-waste NGOs (Tidy Trails, Recity, Smile 'Waste no More', Pyxera Assam)",
 "pepsico.yourcause.com; csrbox.org (Tidy Trails)","HIGH"),
("Amazon","Employee Giving 1:1, $5k/yr (Amazon ACI page)",
 "Right Now Climate Fund: $3.28M deployed in India (Western Ghats, Mumbai mangroves+plastic, ICLEI)",
 "amazon.jobs ACI; aboutamazon.in; theclimatepledge.com","HIGH"),
("Apple","Matching Gifts 1:1 (2:1 special), $10k/yr (official Apple guide PDF)",
 "India 2026: WWF-India waste partnership + Acumen catalytic grants to 6 green/circular enterprises",
 "forms.matchinggifts.com/AppleGuide.pdf; apple.com/in/newsroom","HIGH"),
("NVIDIA","Inspire 365 Foundation 1:1, $10k/yr (official guidelines PDF)",
 "NVIDIA India CSR pillars: biodiversity, water-body rejuvenation, mangrove restoration (Sundarbans)",
 "nvidia.com/en-us/foundation; nvidia.com/en-in/csr-india","HIGH"),
("Cisco","Matching 1:1 (Time2Give), $25k/yr (Benevity case study)",
 "Cisco India Cash Grants — OPEN LOI portal (statutory 2% CSR via Cisco Inc.)",
 "benevity.com/.../cisco; cisco.com/c/en_in/about/csr/india-cash-grants.html","HIGH"),
("Salesforce","1-1-1 model 1:1, $10k/yr (official philanthropy page)",
 "Salesforce.org India funds lake restoration (Malligavad) + 400k+ saplings; >$10M all-time India",
 "salesforce.com/company/philanthropy; Salesforce.org India announcements","HIGH"),
("American Express","Give2Gether 1:1, $10k/yr (GlobalGiving official page)",
 "Amex Foundation Environment pillar; Program Vasundhara (Safe Water Network) water+biodiversity",
 "globalgiving.org/amex-give2gether; csrbox.org (Amex India env)","HIGH"),
("Google","Employee Giving 1:1, $12k/yr (google.benevity.org T&C)",
 "Google.org APAC Sustainability Seed Fund $5M (India water/climate grantees)",
 "google.benevity.org; blog.google/intl/en-in","HIGH"),
("Microsoft","Give Match 1:1, $15k/yr (official MS CSR page + PDF)",
 "Microsoft India CSR + $1B Climate Innovation Fund + AI for Good/Earth grants",
 "microsoft.com/.../employee-engagement; microsoft.com/en-in/corporate-social-responsibility","HIGH"),
("HP","HP Foundation Cash Match 1:1 / Product 3:1, $5k/yr (official HP page)",
 "HP India CSR: forest restoration (WWF, Arbor Day), recycled-plastics, climate",
 "hp.com/.../hp-foundation; hp.com/in-en/sustainable-impact/climate-action","HIGH"),
("IBM","Matching Grants 1:1, $10k/yr (official IBM responsibility page)",
 "IBM Sustainability Accelerator — OPEN RFP portal; India grantee TNC India (crop-burning)",
 "ibm.com/responsibility/...; ibm.com/responsibility/programs/ibm-impact-accelerator","HIGH"),
("SAP","Together Matching 1:1, $750/yr, $20/vol hr (official SAP PDF)",
 "SAP Labs India CSR — env minor pillar (plantation/rejuvenation); core edu/social-enterprise",
 "doublethedonation.com/.../sap; sap.com/india/about/csr.html","MEDIUM"),
("Intel","Intel Involved 1:1, $1k/yr combined (2025) (official Intel page)",
 "Intel Foundation — no unsolicited; email org info, grants at local office level",
 "intel.com/.../intel-foundation; intel.foundation@intel.com","MEDIUM"),
("Goldman Sachs","Matching 1:1, $20k/yr (official program PDF + GS Gives)",
 "GS India CSR; env incidental ('One for Each Tree' Bengaluru afforestation)",
 "forms.matchinggifts.com/GoldmanSachs.pdf; goldmansachs.com/.../india/citizenship","MEDIUM"),
("JPMorgan Chase","Matching 1:1, $2k/yr (official program guide)",
 "JPMC India CSR + Global Philanthropy (Mumbai PO); $13M+ climate globally",
 "forms.matchinggifts.com/ChaseGuide; jpmorgan.com (India CSR Policy PDF)","MEDIUM"),
("Bank of America","Matching 1:1, $5k/yr (official BofA page)",
 "BofA India CSR + Charitable Foundation env pillar (intl funding by invitation)",
 "about.bankofamerica.com/.../matching-gifts; business.bofa.com/in","MEDIUM"),
("Mastercard","Matching 1:1, $5k/yr + Impact Fund (official MC page + PDF)",
 "Center for Inclusive Growth; India e-waste precedent (Karo Sambhav, Spark Grants)",
 "mastercard.com/.../community-impact; mastercardcenter.org","MEDIUM"),
("Deloitte","Foundation Matching 1:1, $32,500/yr (official Foundation PDF)",
 "Deloitte India CSR (WorldClimate); partnered Rainmatter Foundation",
 "files.doublethedonation.com/.../Deloitte; deloitte.com/in","MEDIUM"),
("NetApp","NetApp Cares 1:1 (official NetApp page)",
 "NetApp India CSR — env listed (water) but spend skews hunger/edu (Akshaya Patra)",
 "netapp.com/.../netapp-cares; netapp.com India CSR Policy PDF","MEDIUM"),
("Adobe","Create Change 1:1, $10k combined (Benevity case study)",
 "Adobe Foundation Employee Community Fund — employee-nominated (Noida/Bengaluru)",
 "adobe.com/corporate-responsibility/community; socialimpact@adobe.com","MEDIUM"),
("Wells Fargo","Matching 1:1, $2k/yr, 16h VTO (official WF page)",
 "WF India CSR — education-dominated (Mantra, Agastya, Room to Read)",
 "wellsfargo.com/.../giving-faqs; wellsfargojobs.com India CSR PDF","HIGH"),
("Texas Instruments","TI Foundation Matching 1:1, up to $30k/yr (official TI page)",
 "TI India CSR — Akshaya Patra meals + WiSH; env = own-ops only",
 "ti.com/.../giving.html; csrbox.org (TI India)","HIGH"),
("Qualcomm","Employee Giving 1:1, $1k–5k/yr (Double the Donation DB)",
 "Qualcomm India CSR / Wireless Reach — tech/edu; SootSwap cookstoves only env-adjacent",
 "doublethedonation.com/matching-gifts; qualcomm.com/.../india/CSRPolicy","MEDIUM"),
("Nike","Give Your Best 1:1 (2:1 GivingTuesday), $25k/yr (official Nike page)",
 "Nike Community Impact Fund = sport-for-kids; 'Move to Zero' internal (not grants)",
 "about.nike.com/.../matching-our-employees-impact","HIGH"),
("Johnson & Johnson","Matching 2:1, up to $20k/yr (official J&J PDF)",
 "J&J India CSR = health only; Foundation no unsolicited proposals",
 "jnj.com/caring-and-giving; jnj.in/csr-projects","HIGH"),
("Pfizer","Foundation Matching 1:1, $5k/yr (official Pfizer page)",
 "Pfizer India CSR = health + school sanitation; no env/waste",
 "pfizer.com/.../give-forward; pfizer.co.in/.../csr-initiatives","HIGH"),
("Merck (MSD)","Partnership for Giving 1:1, $30k/yr (official Merck page+PDF)",
 "MSD India = health + Safe Water Network water-access; global portal only (versaic)",
 "merck.com/.../philanthropy; msd.versaic.com","MEDIUM"),
("Apollo Global Mgmt","Citizenship Grants matching+vol rewards (SEC DEF 14A)",
 "Apollo Opportunity Foundation = career/workforce only; no env grantmaking",
 "sec.gov DEF 14A (Apollo); apollo.com/impact","MEDIUM"),
("Stryker","Impact Platform 1:1 (official Stryker page)",
 "Stryker Johnston Foundation = US/health; product donations only",
 "stryker.com/.../giving-back; grants@strykerjohnstonfoundation.org","MEDIUM"),
("Airbnb","Community Fund + Matching 1:1, $1k/yr (Airbnb newsroom)",
 "Airbnb Community Fund (host-nominated) funds India env NGOs (Keystone, Stree Mukti, Goonj)",
 "news.airbnb.com/2025-community-fund; indiacsr.in (2026 cohort)","HIGH"),
("Dell Technologies","Employee Giving 1:1, $10k/yr (YourCause portal)",
 "Dell India CSR — digital/edu focus; env incidental",
 "dell.yourcause.com; dell.com/.../employee-led-impact","HIGH"),
]
for i, row in enumerate(PROOF, 1):
    r = i + 3
    conf = row[4]
    for ci, v in enumerate(row, 1):
        c = ws4.cell(r, ci, v)
        c.alignment = wa("center" if ci==5 else "left"); c.border = brd
        c.font = fn(DARK_TEXT, b=(ci==1), s=9)
        if ci==5:
            c.fill = hf(HI_GREEN if conf=="HIGH" else LT_AMBER); c.font=fn(DARK_TEXT,b=True,s=9)
        else: c.fill = hf(ALT[i%2])
    ws4.row_dimensions[r].height = 52
cw4 = [18,42,46,40,12]
for i,w in enumerate(cw4,1): ws4.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════════
# SHEET 5 — ACTION PLAN & PROPOSAL PLAYBOOK
# ══════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Action Plan & Playbook")
ws5.sheet_view.showGridLines = False
title_row(ws5, "WASTE WARRIORS — DIRECT-GIVING ACTION PLAN & PROPOSAL PLAYBOOK", 4, AMBER)

def block(ws, r, heading, lines, bg, span=4):
    ws.merge_cells(f"A{r}:{get_column_letter(span)}{r}")
    c = ws[f"A{r}"]; c.value=heading
    c.font=Font(name="Calibri",bold=True,size=11,color=WHITE)
    c.fill=hf(bg); c.alignment=wa("left","center"); ws.row_dimensions[r].height=24
    r+=1
    for txt in lines:
        ws.merge_cells(f"A{r}:{get_column_letter(span)}{r}")
        cc=ws[f"A{r}"]; cc.value=txt
        cc.font=fn(DARK_TEXT, b=txt.startswith("►") or txt.startswith("★"), s=10)
        cc.alignment=wa("left","top"); cc.fill=hf(WHITE)
        cc.border=Border(bottom=Side(style="hair",color=MED_GREY))
        ws.row_dimensions[r].height = 16 + 14*max(0,(len(txt)//95))
        r+=1
    return r+1

r=3
r=block(ws5,r,"WHY THIS PLAN (the pivot away from Benevity)",[
 "On Benevity you wait passively to be discovered by one of 200,000+ orgs. This plan does the opposite: YOU pitch the company directly and they grant corporate money to Waste Warriors.",
 "The real money for a US/EU company giving in India is its INDIA SUBSIDIARY's statutory 2% CSR (Companies Act 2013) — paid in INR to local NGOs. No FCRA needed; you only need 12A + 80G + CSR-1 (you have these).",
 "Three direct doors per company: (1) India CSR team  (2) Corporate Foundation  (3) ESG/Sustainability fund. Always look the company up on csrbox.org first to find its India CSR entity, current NGO partners, and contact.",
], DARK_NAVY)

r=block(ws5,r,"★ TIER 1 — APPLY DIRECTLY THIS MONTH (open portals / perfect fit)",[
 "★ HCLTech Grant — open ₹5 Crore / 4-yr competitive grant, BIODIVERSITY & WATER tracks = exactly Waste Warriors. APPLY. (hclfoundation.org/hcltech-grant)",
 "★ Wipro Foundation Ecology — open Community Ecology Program; you meet eligibility (80G/12A + 3yrs ecology). Register & apply. (wiprofoundation.org/our-initiatives/ecology)",
 "★ Nestlé India – Project Hilldaari — runs hill-town waste mgmt in Mussoorie/Dalhousie/Palampur/Darjeeling = your footprint. Pitch as an expansion partner. (thecsruniverse.com)",
 "★ Cisco India Cash Grants — open Letter-of-Inquiry portal. Frame waste-worker livelihoods + climate resilience. (cisco.com/c/en_in/about/csr/india-cash-grants.html)",
 "★ IBM Sustainability Accelerator — open RFP (annual call ~April); tech + cash for climate-vulnerable communities. (ibm.com/responsibility/programs/ibm-impact-accelerator)",
 "★ 3M India — open env grant portal + matching gifts. (3mindia.in/.../grant-guidelines)   ★ Schneider Electric — 'Submit Your Project' portal. (seifoundation.se.com)",
 "★ Coca-Cola Anandana & PepsiCo Foundation — both grant NGOs directly for water + plastic waste. Send proposals. (anandana.org / pepsicoindia.co.in)",
], FOREST)

r=block(ws5,r,"TIER 2 — WARM INTRO NEEDED (high fit, no open portal)",[
 "► Amazon Right Now Climate Fund — intro via existing partners (Centre for Wildlife Studies, ICLEI). Pitch restoration + plastic-waste removal.",
 "► Apple Environment — intro via WWF-India / Acumen (they fund waste/circular enterprises in India).",
 "► NVIDIA India, Salesforce.org India, American Express (Program Vasundhara), Google.org, Microsoft India, HP India, HSBC, Cummins, HUF (Unilever) — approach the named India CSR head on LinkedIn + a 1-page concept note.",
], MID_BLUE)

r=block(ws5,r,"TIER 3 — DEPRIORITIZE (weak/no India env route)",[
 "► Low env fit: Wells Fargo, TI, Qualcomm, Nike, J&J, Pfizer, Apollo, Stryker, Dell, SAP, Adobe, Goldman, JPMorgan, BofA, Mastercard, Deloitte, NetApp, Merck.",
 "► Only pursue if you can reframe (e.g. waste-worker livelihoods, WASH/sanitation, tech-for-environment) OR you land an internal employee champion.",
 "► Airbnb is theme-fit but access is host-NOMINATED — cultivate Dharamshala/Mussoorie/Corbett-area Airbnb hosts to nominate you.",
], DEEP_RED)

r=block(ws5,r,"THE DIRECT-GIVING ENTRY SEQUENCE (repeat per target)",[
 "1. RESEARCH: look up the company on csrbox.org → find its India CSR entity, focus areas, current NGO partners, and any named CSR contact.",
 "2. FIND THE PERSON: on LinkedIn search '<Company> India CSR' / 'Sustainability' / 'Corporate Social Responsibility'. Note the name + title (Sheet 'Direct Entry Strategy').",
 "3. CONCEPT NOTE (1 page): problem (Himalayan waste/biodiversity) → your proven model + numbers → the specific outcome their CSR pillar wants → budget range → why you (FCRA/80G/12A/CSR-1, audited).",
 "4. OUTREACH: email the CSR/foundation contact (or apply on the open portal) + a short LinkedIn message to the named person referencing their existing env program by name.",
 "5. FOLLOW UP after 7–10 days; offer a 20-min call + a site visit invite to a Waste Warriors project.",
 "6. CONVERT: turn interest into a full proposal + MoU; align reporting to their CSR/ESG metrics (tonnes diverted, trees, waste-workers, beneficiaries).",
], TEAL)

r=block(ws5,r,"WHAT GOES IN THE 1-PAGE PROPOSAL / CONCEPT NOTE",[
 "• Hook: one sharp line on the Himalayan waste/biodiversity problem with a number.",
 "• Track record: WW's model, geographies (Dehradun/Dharamshala/Corbett), tonnes diverted, waste-workers supported, communities reached.",
 "• The fit: name THEIR program/pillar (e.g. 'aligns with Anandana's water+waste mandate' / 'HCLTech Grant Water track') — show you did the homework.",
 "• The ask: clear scope + budget range + duration; what their money buys in measurable units.",
 "• Credibility: 12A, 80G, CSR-1, FCRA, audited accounts, board, key partners/awards.",
 "• Reporting: the exact metrics you'll report back (this is what CSR teams need for compliance).",
], PURPLE)

for col,w in zip("ABCD",[55,45,45,45]): ws5.column_dimensions[col].width=w

# ── save ──────────────────────────────────────────────────────────
out = "/home/user/jcode/Waste_Warriors_Direct_Corporate_Giving_Mastersheet.xlsx"
wb.save(out)
print("Saved:", out)
print("Sheets:", wb.sheetnames)
print("Companies in master:", len(M), "| New companies:", len(NEW), "| Strategy rows:", len(S))
