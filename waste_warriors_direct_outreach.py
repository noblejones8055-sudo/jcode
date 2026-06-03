import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

DARK_NAVY  = "1B2A4A"; MID_BLUE   = "2E5CA8"; FOREST     = "1A5C38"
MID_GREEN  = "217A4F"; AMBER      = "B8860B"; DEEP_RED   = "8B0000"
TEAL       = "0F4C5C"; PURPLE     = "4B0082"; WHITE      = "FFFFFF"
LT_GREEN   = "D5F5E3"; LT_AMBER   = "FFF3CD"; LT_RED     = "FADBD8"
LT_BLUE    = "D6E4F7"; LT_GREY    = "F2F2F2"; MED_GREY   = "CCCCCC"
LT_PURPLE  = "EDE7F6"; DARK_TEXT  = "1A1A2E"

def hf(c): return PatternFill("solid", fgColor=c)
def fn(c=DARK_TEXT, b=False, s=9): return Font(color=c, bold=b, size=s, name="Calibri")
th = Side(style="thin", color=MED_GREY)
brd = Border(left=th, right=th, top=th, bottom=th)
def wa(h="left", v="top"): return Alignment(horizontal=h, vertical=v, wrap_text=True)

def title_row(ws, txt, cols, bg=DARK_NAVY, r=1):
    ws.merge_cells(f"A{r}:{get_column_letter(cols)}{r}")
    c = ws[f"A{r}"]
    c.value, c.font = txt, Font(name="Calibri", bold=True, size=13, color=WHITE)
    c.fill, c.alignment = hf(bg), Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 30

def hdr(ws, cols, row=2, bg=MID_BLUE):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
        c.fill, c.alignment, c.border = hf(bg), wa("center","center"), brd
    ws.row_dimensions[row].height = 28

ALT = [LT_GREY, WHITE]

# ══════════════════════════════════════════════════════════════
# SHEET 1 – DIRECT CONTACT DIRECTORY (ALL 32 + INDIA)
# ══════════════════════════════════════════════════════════════
ws1 = wb.active; ws1.title = "Direct Contact Directory"
ws1.sheet_view.showGridLines = False; ws1.freeze_panes = "A3"

title_row(ws1, "WASTE WARRIORS — DIRECT CONTACT DIRECTORY: ALL 32 COMPANIES + INDIA (REAL CONTACTS)", 9)
hdr(ws1, ["Company","Sector","Contact Type","Direct Email / Phone",
          "Portal / URL","Contact Name (if known)","Open for Unsolicited?",
          "Best Entry Point","Priority"])

contacts = [
    # ── BENEVITY COMPANIES ──────────────────────────────────────
    ("Microsoft","Technology","Benevity platform + NGO programme",
     "No direct public email for NGO listing.\nChange Agent Programme submissions via:\nmicrosoft.com/en-us/corporate-responsibility/philanthropies/employee-engagement",
     "microsoft.com/en-us/corporate-responsibility/philanthropies/employee-engagement\nmicrosoft.benevity.org (employees)",
     "Sonia Fiorenza-Lowe — Sr Director, Employee Giving & Volunteer Programs (LinkedIn)\nAlternate: Microsoft Philanthropies team via philanthropies@microsoft.com",
     "YES — register on Benevity; also apply to Change Agent Programme",
     "1. Register Benevity (causes@benevity.org)\n2. Email philanthropies@microsoft.com introducing Waste Warriors\n3. Apply to Change Agent Programme on MS website",
     "🔴 HIGHEST"),

    ("Google","Technology","Benevity platform + Google.org",
     "Google for Nonprofits: support.google.com/nonprofits\nGoogle.org grants: google.org/our-work/\nBenevity: causes@benevity.org",
     "google.benevity.org (employees)\ngoogle.org (grants)\nsupport.google.com/nonprofits",
     "Google.org team contact via google.org/contact\nGoogle for Nonprofits support form",
     "YES — register on Benevity AND apply to Google for Nonprofits",
     "1. Register on Benevity\n2. Apply for Google for Nonprofits (free tools + visibility)\n3. Submit expression of interest to google.org",
     "🔴 HIGHEST"),

    ("Apple","Technology","Benevity + GlobalGiving dual-track",
     "Benevity charity portal: causes@benevity.org\nGlobalGiving (Apple due diligence): apple@globalgiving.org",
     "apple.benevity.org (employees)\nglobalgiving.org (due diligence)",
     "Apple's program is administered via Benevity + GlobalGiving.\nNo named individual publicly available.",
     "YES — register on Benevity AND on GlobalGiving. Apple requires BOTH.",
     "1. Register on Benevity → causes@benevity.org\n2. ALSO register on GlobalGiving + submit to apple@globalgiving.org for Apple-specific due diligence\nBOTH steps required to appear in Apple's program",
     "🔴 HIGHEST"),

    ("IBM","Technology","Benevity platform",
     "IBM giving portal: ibm.benevity.org\nIBM Smarter Cities / Tech grants: ibm.com/responsibility/programs\nGeneral CSR: corporatecitizenship@us.ibm.com",
     "ibm.benevity.org\nwww.ibm.com/responsibility/programs",
     "Contact via: corporatecitizenship@us.ibm.com\nor IBM Foundation (via Candid directory)",
     "YES — register on Benevity + reach out via email",
     "1. Register on Benevity\n2. Email corporatecitizenship@us.ibm.com with 1-page introduction\n3. Apply for IBM Community Grants if eligible",
     "🔴 HIGH"),

    ("NVIDIA","Technology","Benevity (Inspire 365) + NVIDIA Foundation",
     "NVIDIA Foundation via Silicon Valley Community Foundation: svcf.org/nvidia\nSkills-based volunteering partnerships: nvidia.com/foundation\nBenevity support: causes@benevity.org",
     "nvidia.com/en-us/foundation/\nsvcf.org/nvidia",
     "NVIDIA Foundation is managed through Silicon Valley Community Foundation.\nContact SVCF for foundation grants.",
     "YES — register on Benevity; foundation grants via SVCF",
     "1. Register on Benevity (auto-eligible for Inspire 365 matching)\n2. Contact SVCF at svcf.org/nvidia for foundation grant consideration\n3. Apply for Skills for Impact programme (nvidia.com/foundation/skills-based-volunteerism/)",
     "🔴 HIGH"),

    ("Salesforce","Technology","Salesforce Philanthropy Cloud + Power of Us",
     "Salesforce.org nonprofits: salesforce.org/nonprofit/\nPower of Us programme: salesforce.com/company/power-of-us/\nPhilanthropy Cloud: salesforce.com/company/philanthropy/",
     "salesforce.com/company/philanthropy/\nsalesforce.org/nonprofit/",
     "Salesforce.org Nonprofit Success team\nContact form at salesforce.org/nonprofit/",
     "YES — apply for Power of Us (free Salesforce CRM for NGOs) + submit to Philanthropy Cloud",
     "1. Apply for Power of Us programme (free CRM + visibility to Salesforce employees)\n2. Register on Philanthropy Cloud as eligible charity\n3. Email Salesforce India CSR team for local employee engagement",
     "🟡 MEDIUM"),

    ("Cisco","Technology","Benevity (Community Impact Portal)",
     "DIRECT EMAIL: communityimpact@cisco.com ✓\nCisco Foundation grants: cisco.com/site/us/en/about/purpose/social-impact",
     "cisco.com/site/us/en/about/purpose/social-impact/cisco-foundation.html",
     "Contact: communityimpact@cisco.com\nCisco Foundation team",
     "YES — email communityimpact@cisco.com directly",
     "1. Email communityimpact@cisco.com with Waste Warriors intro\n2. Register on Benevity (covers Cisco employees)\n3. Request to be featured in Cisco's monthly giving campaigns",
     "🔴 HIGH"),

    ("Adobe","Technology","Benevity (Create Change Dashboard)",
     "Benevity: causes@benevity.org\nAdobe ECF (Employee Community Fund) grants:\nstorage.benevitygrants.org/adobe\nAdobe CSR: adobe.com/corporate-responsibility",
     "causes.benevity.com\nadobe.com/corporate-responsibility/community.html",
     "Adobe Corporate Responsibility team\nEmail via adobe.com/corporate-responsibility contact form",
     "YES — register on Benevity + apply for Adobe ECF grant",
     "1. Register on Benevity\n2. Apply for Adobe Employee Community Fund (ECF) grant separately\n3. Contact Adobe CSR team to be featured in a giving campaign",
     "🔴 HIGH"),

    ("Amazon","Technology","Internal Amazon portal (invite-only)",
     "Amazon Community Impact: aboutamazon.com/impact/community\nAWS for Nonprofits: aws.amazon.com/government-education/nonprofits/\nHousing Equity Fund (specific): housingequityfund@amazon.com",
     "aboutamazon.com/impact/community\naws.amazon.com/government-education/nonprofits/",
     "Amazon generally does NOT accept unsolicited proposals.\nBest route: apply for AWS for Nonprofits (free cloud credits)\nLinkedIn: contact Amazon Community Impact team members directly",
     "PARTIAL — via AWS Nonprofits + LinkedIn outreach to Community Impact team",
     "1. Apply for AWS for Nonprofits (free cloud credits = builds relationship)\n2. Submit expression of interest via aboutamazon.com/impact/community\n3. LinkedIn: search 'Amazon Community Impact' and connect with managers\n4. Monitor amazon.jobs/content/en/teams/ccr/aci for open application windows",
     "🟡 MEDIUM"),

    ("HP","Technology","HP Foundation portal",
     "HP Philanthropy contact form: hp.com/hpinfo/socialinnovation/us/contactus.html\nHP Employee Giving: hp.com/hpinfo/socialinnovation/us/contactus_employeegiving.html\nHP Foundation: hp.com/us-en/hp-information/hp-foundation.html",
     "hp.com/us-en/hp-information/hp-foundation.html",
     "HP Foundation (via contact form)\nHP Philanthropy team via form on website",
     "PARTIAL — no direct email; use contact form",
     "1. Submit via HP Philanthropy contact form\n2. Apply for HP Future of Work Accelerator (hp.com/accelerator)\n3. Register on relevant matching gift platform",
     "🟡 MEDIUM"),

    ("Texas Instruments","Technology/Semiconductor","TI Foundation portal",
     "TI Foundation: ti.com/about-ti/citizenship-community/giving.html\nTI Alumni matching: tialumni.org/volunteer-2/ti-foundation-and-matching-gifts/\nTI Eligible orgs: ti.com/about-ti/citizenship-community/giving/eligible-organizations.html",
     "ti.com/about-ti/citizenship-community/giving.html",
     "Texas Instruments Foundation team\nContact via ti.com/about-ti/citizenship-community/contact.html",
     "YES — educational and 501(c)(3) equivalent NGOs eligible",
     "1. Verify eligibility at TI eligible organizations page\n2. Contact TI Foundation via website form\n3. Ensure 80G/FCRA + CAF America ED letter available",
     "🟡 MEDIUM"),

    ("Intel","Technology/Semiconductor","Benevity (reopened 2025)",
     "Intel Foundation: intel.com/content/www/us/en/corporate-responsibility/intel-foundation-employee-generosity.html\nIntel Involved volunteering: intel.com/content/www/us/en/corporate-responsibility/intel-matching-grants.html\nBenevity: causes@benevity.org",
     "intel.benevity.org",
     "Intel Foundation team\nContact via intel.com corporate responsibility pages",
     "YES — register on Benevity (Intel's platform since Jan 2025)",
     "1. Register on Benevity (Intel now uses Benevity since Jan 2025)\n2. Note: $1,000 combined cap (reduced in 2025)\n3. Contact Intel Foundation for skills-based volunteering (Skills for Impact equivalent)",
     "🟡 MEDIUM"),

    ("NetApp","Technology","Benevity (NetApp Cares)",
     "NetApp Cares: netapp.com/responsibility/social-impact/netapp-cares/\nBenevity: causes@benevity.org\nNetApp CSR: netapp.com/responsibility/",
     "netapp.com/responsibility/social-impact/netapp-cares/",
     "NetApp Corporate Social Responsibility team\nContact form at netapp.com/company/contact-us/",
     "YES — register on Benevity + direct outreach",
     "1. Register on Benevity\n2. Email NetApp CSR team via contact form with Global Giving Challenge pitch\n3. Aim for December Global Giving Challenge feature",
     "🟡 MEDIUM"),

    ("Mastercard","Financial Services","Mastercard internal portal",
     "Mastercard Community Impact: mastercard.com/us/en/for-the-world/people/community-impact.html\nMastercard Impact Fund: mastercardcenter.org/about-the-center/mastercard-impact-fund\nContact via: mastercard.com/global/en/about/contact.html",
     "mastercard.com/us/en/for-the-world/people/community-impact.html",
     "Mastercard Corporate Responsibility team\nImpact Fund: mastercardcenter.org contact form",
     "PARTIAL — primarily strategic partnerships",
     "1. Submit expression of interest via Mastercard contact form\n2. Apply to Mastercard Impact Fund via mastercardcenter.org\n3. Connect with Mastercard India CSR team (Mastercard has Pune/Vadodara offices)",
     "🟡 MEDIUM"),

    ("American Express","Financial Services","GlobalGiving (Give2Gether)",
     "GlobalGiving (AmEx route): globalgiving.org/amex-give2gether/\nAmEx Give2Gether: globalgiving.org/amex-give2gether/\nGlobalGiving partner: globalgiving.org/become-a-partner",
     "globalgiving.org/amex-give2gether/",
     "GlobalGiving partner team\nEmail: partners@globalgiving.org",
     "YES — register as GlobalGiving partner → auto-eligible for AmEx Give2Gether",
     "1. Email partners@globalgiving.org to start partnership process\n2. Register as GlobalGiving partner\n3. Run Open Challenge to earn full partner status\n4. Once full partner: auto-eligible for AmEx $10,000/employee match",
     "🔴 HIGH"),

    ("Stryker","Healthcare/Medical","Stryker Impact + Johnston Foundation",
     "DIRECT EMAIL: info@strykerjohnstonfoundation.org ✓\nPhone: 269.488.8484\nAddress: 180 E. Water Street, Suite 3000, Kalamazoo, MI 49007\nStryker Impact platform: stryker.com/us/en/about/corporate-responsibility/giving-back.html",
     "strykerjohnstonfoundation.org/apply\nstryker.com/us/en/about/corporate-responsibility/giving-back.html",
     "Stryker Johnston Foundation\nDIRECT CONTACT: info@strykerjohnstonfoundation.org",
     "YES — open grant applications + Impact platform registration",
     "1. Email info@strykerjohnstonfoundation.org introducing Waste Warriors\n2. Apply for Stryker Johnston Foundation grant\n3. Request listing on Stryker Impact employee giving platform\n4. PITCH ANGLE: medical/hospital waste in Himalayan communities",
     "🔴 HIGH"),

    ("Apollo Global Mgmt","Private Equity/Finance","Apollo Opportunity Foundation",
     "Apollo Opportunity Foundation: apollo.com/responsibility\nApollo Global CSR: apollo.com/citizenship\nLinkedIn: search 'Apollo Global Corporate Citizenship'",
     "apollo.com/responsibility",
     "Apollo Opportunity Foundation team\nContact via apollo.com/contact",
     "PARTIAL — focus on career education / workforce development",
     "1. Contact Apollo Opportunity Foundation directly via apollo.com/responsibility\n2. Align pitch with their focus: livelihoods + workforce development in waste sector\n3. LinkedIn: search Apollo Corporate Citizenship team",
     "🟡 MEDIUM"),

    # ── BANKING ──────────────────────────────────────────────────
    ("JPMorgan Chase","Banking","JPMorgan Chase Foundation (invite-only)",
     "Foundation: jpmorganchase.com/impact\nGrant search: jpmorgan.com/private-bank/foundations/online-applications/search\nGiving Point: jpmgiving.org/contact\nLinkedIn: JPMorgan Chase Corporate Responsibility team",
     "jpmorganchase.com/impact/volunteerism\njpmgiving.org",
     "JPMorgan Chase Foundation — generally invite-only.\nBest: LinkedIn outreach to JP Morgan India Corporate Responsibility team\n(Mumbai / Bangalore offices)",
     "PARTIAL — invite-only for grants; CyberGrants for employee matching",
     "1. Register on CyberGrants (cybergrants.com/jpmc/giving) for employee matching\n2. LinkedIn: connect with JPMorgan Chase India CSR managers\n3. Monitor jpmorganchase.com/impact for open application windows",
     "🟡 MEDIUM"),

    ("Goldman Sachs","Banking","Office of Corporate Engagement",
     "Corporate Engagement: goldmansachs.com/community-impact\nCommunity TeamWorks volunteering: goldmansachs.com/community-impact/community-teamworks\nGoldman Sachs Gives: goldmansachs.com/community-impact/goldman-sachs-gives\nMatching: gs@easymatch.com",
     "goldmansachs.com/community-impact",
     "Office of Corporate Engagement team\nContact via goldmansachs.com/community-impact\n$20,000 cap — high priority",
     "YES — Corporate Engagement team actively seeks NGO partners for Community TeamWorks",
     "1. Email via goldmansachs.com/community-impact contact form\n2. Pitch Community TeamWorks volunteering + employee matching\n3. EasyMatch: gs@easymatch.com for technical matching questions\n4. Goldman has large Bengaluru/HYD offices — India angle is strong",
     "🔴 HIGH"),

    ("Bank of America","Banking","Benevity + BofA Charitable Foundation",
     "BofA Charitable Foundation grants: about.bankofamerica.com/en/making-an-impact/charitable-foundation-funding\nGrant cycles: Basic Needs (Feb), Stable Housing (May)\nNeighborhood Builders: about.bankofamerica.com/en/making-an-impact",
     "about.bankofamerica.com/en/making-an-impact/charitable-foundation-funding",
     "Bank of America Charitable Foundation team\nApplication form via NPO Hub (npohub.bankofamerica.com)",
     "YES — open RFP cycles twice a year + Benevity for employee matching",
     "1. Register on Benevity (BofA employees can then donate + get matched)\n2. Apply in BofA Charitable Foundation RFP cycle:\n   - Basic Needs: Feb 2 – Mar 2\n   - Empowering Communities: May 18 – Jun 29\n3. Register on NPO Hub: npohub.bankofamerica.com",
     "🔴 HIGH"),

    ("Wells Fargo","Banking","YourCause + Wells Fargo Foundation",
     "WF Community Giving: wellsfargo.com/about/responsibility-and-impact/community-giving/\nWF Foundation: wellsfargofoundation.org\nPhilanthropy FAQs: wellsfargo.com/about/responsibility-and-impact/community-giving/giving-faqs/",
     "wellsfargo.com/about/responsibility-and-impact/community-giving/",
     "Wells Fargo Foundation team\nContact via wellsfargo.com contact channels",
     "PARTIAL — Wells Fargo Foundation is largely invite-only; YourCause for matching",
     "1. Register on YourCause (NPOconnect) for employee matching\n2. Contact Wells Fargo Foundation via website for grant consideration\n3. Note: 6-month submission deadline on employee matches",
     "🟡 MEDIUM"),

    # ── PHARMA ───────────────────────────────────────────────────
    ("Johnson & Johnson","Pharma/Healthcare","J&J Foundation (direct contact)",
     "DIRECT PHONE: +1 (732) 524-2455\nJ&J Caring website: jnj.com/caring\nFoundation: jnj.com/about-jnj/global-impact\nAddress: One Johnson & Johnson Plaza, New Brunswick, NJ 08933",
     "jnj.com/caring\njnj.com/about-jnj/global-impact",
     "Johnson & Johnson Foundation\nPhone: +1 (732) 524-2455\nContact via jnj.com/caring",
     "YES — J&J accepts NGO partner applications; focus on health + women/children",
     "1. Call +1 (732) 524-2455 or submit via jnj.com/caring\n2. PITCH ANGLE: waste impact on women's health / child health in Himalayas\n3. Register on J&J matching gift portal for employee donations\n4. 2:1 match = exceptional ROI; prioritize this outreach",
     "🔴 HIGHEST"),

    ("Pfizer","Pharma","Pfizer Foundation competitive grants",
     "Competitive Grants: pfizer.com/about/programs-policies/grants/competitive-grants\nGlobal Impact: pfizer.com/about/responsibility/global-impact\nFoundation: pfizer.com/about/responsibility/global-impact/the-pfizer-foundation",
     "pfizer.com/about/programs-policies/grants/competitive-grants",
     "Pfizer Foundation team\nApply via competitive grants portal\nIndia contact: pfizer.com/en_IN (India page)",
     "YES — competitive grants portal is open; employee matching via pfizerplus.com",
     "1. Submit via Pfizer competitive grants portal (pfizer.com/about/programs-policies/grants)\n2. Register at pfizerplus.com to be eligible for employee matching\n3. PITCH: Himalayan community health + waste as public health issue",
     "🟡 MEDIUM"),

    ("Merck","Pharma","CyberGrants (Partnership for Giving)",
     "DIRECT EMAIL (matching): merckp4g@easymatch.com ✓\nDIRECT PHONE: 1-866-205-2857 ✓\nMerck Foundation contact: leslie_hardy@merck.com ✓\nPhilanthropy: merck.com/company-overview/sustainability/philanthropy/",
     "merck.com/company-overview/sustainability/philanthropy/\ncybergrants.com/merck/merckgivesback",
     "DIRECT CONTACTS:\n1. Partnership for Giving: merckp4g@easymatch.com\n2. Phone: 1-866-205-2857\n3. Merck Foundation: leslie_hardy@merck.com",
     "YES — direct contacts available; $30,000 cap is highest in pharma",
     "1. Email merckp4g@easymatch.com to register Waste Warriors as eligible NGO\n2. Contact leslie_hardy@merck.com for Merck Foundation grant consideration\n3. Register on CyberGrants platform\n4. High cap ($30,000/employee) = prioritize this",
     "🔴 HIGHEST"),

    # ── CONSULTING ────────────────────────────────────────────────
    ("Deloitte","Consulting","Deloitte Foundation (invite-only with key contact)",
     "DIRECT CONTACT: kshoztic@deloitte.com ✓\nDIRECT PHONE: (203) 761-3248 ✓\nDeloitte Foundation: deloitte.com/us/en/about/deloitte-foundation.html\nImpact Day: deloitte.com/global/en/about/social-impact",
     "deloitte.com/us/en/about/deloitte-foundation.html",
     "DIRECT CONTACT: Kathy Shoztic\nEmail: kshoztic@deloitte.com\nPhone: (203) 761-3248",
     "PARTIAL — strategic grants are invite-only BUT local Impact grants and matching gifts are accessible",
     "1. Email kshoztic@deloitte.com directly\n2. Subject: 'Waste Warriors — Himalayas Waste Management | Deloitte Foundation Partnership'\n3. Deloitte India offices (Mumbai, Gurgaon, Bengaluru) = employees can donate via matching\n4. Apply for Deloitte Foundation Local Impact Grants\n5. $32,500 cap = absolute highest of all 32 companies",
     "🔴 HIGHEST"),

    # ── CONSUMER ──────────────────────────────────────────────────
    ("Coca-Cola","Consumer/Beverages","Coca-Cola Foundation (online application)",
     "Foundation: coca-colacompany.com/social/coca-cola-foundation\nApplication: coca-colacompany.com (via online application system)\nIndia CSR: coca-colaindia.com/sustainability",
     "coca-colacompany.com/social/coca-cola-foundation",
     "Coca-Cola Foundation team\nApply via online system at coca-colacompany.com/social/coca-cola-foundation\nIndia: coca-colaindia.com contact form",
     "YES — online application system open (eligibility quiz first)",
     "1. Submit online application at coca-colacompany.com/social/coca-cola-foundation\n2. Coca-Cola India office: coca-colaindia.com/sustainability → strong water/environment angle\n3. Employee matching: via Coca-Cola Foundation portal (2:1 match!)\n4. PITCH: plastic waste in Himalayan rivers / water conservation",
     "🔴 HIGH"),

    ("PepsiCo","Consumer/Beverages","YourCause (pepsico.yourcause.com)",
     "PepsiCo YourCause portal: pepsico.yourcause.com\nPepsiCo Foundation: pepsico.com/who-we-are/pepsico-foundation\nContact: pepsico.com/contact-us\nPhone support: 1-866-321-6391",
     "pepsico.yourcause.com\npepsico.com/who-we-are/pepsico-foundation",
     "PepsiCo Foundation team\nMatching support: 1-866-321-6391\nPepsico.com/contact-us for foundation inquiries",
     "YES — register on YourCause + apply to PepsiCo Foundation",
     "1. Register on YourCause (NPOconnect) for employee matching\n2. Apply to PepsiCo Foundation grants programme\n3. PITCH: pep+ (PepsiCo Positive) sustainability aligns perfectly with waste management\n4. Note unique 2:1 match if employees volunteer 50+ hours!",
     "🔴 HIGH"),

    ("Nike","Consumer/Apparel","Benevity (GYB) + NCIF via Oregon Community Foundation",
     "NCIF Contact: grants@oregoncf.org ✓\nNike Employee Community Engagement: about.nike.com/en/impact/initiatives/employee-community-engagement\nNike Social & Community Impact: media.about.nike.com (grantmaking guide)\nCommunity Grants: about.nike.com/en/resources/community-grants",
     "oregoncf.org/grants-and-scholarships/grants/nike-community-impact-fund-program\nabout.nike.com/en/resources/community-grants",
     "Oregon Community Foundation (manages NCIF):\nEmail: grants@oregoncf.org ✓\nNIKE, Inc. Social & Community Impact team",
     "YES — NCIF is open annually (Jan–Feb application window)",
     "1. Email grants@oregoncf.org for NCIF application details\n2. NOTE: NCIF is primarily for Portland/Oregon area\n3. Register on Benevity (Nike uses GYB/Benevity for employee matching globally)\n4. For India: Nike India employee giving → approach Nike India CSR via linkedin",
     "🔴 HIGH"),

    # ── TECH (NEW 16) ──────────────────────────────────────────────
    ("Cisco","Technology","Benevity + Community Impact Portal",
     "DIRECT EMAIL: communityimpact@cisco.com ✓\nCisco Foundation: cisco.com/site/us/en/about/purpose/social-impact/cisco-foundation.html\nGrant Policies: cisco.com/site/us/en/about/purpose/social-impact/investments/grant-giving-policies.html",
     "cisco.com/site/us/en/about/purpose/social-impact/cisco-foundation.html",
     "DIRECT: communityimpact@cisco.com ✓\nCisco Foundation team",
     "YES — communityimpact@cisco.com is the direct contact",
     "1. Email communityimpact@cisco.com with subject:\n'Partnership Request — Waste Warriors | Himalayas Waste Management | Cisco Matching Gift Programme'\n2. Ask to be listed/featured in Cisco's employee giving platform\n3. Register on Benevity in parallel\n4. Cisco's 85% employee participation = enormous donor pool",
     "🔴 HIGHEST"),

    ("Dell Technologies","Technology","YourCause (dell.yourcause.com)",
     "Dell YourCause: dell.yourcause.com\nDell Social Impact: dell.com/en-us/dt/corporate/social-impact\nDell Community Grant: delltechnologies.com/en-us/social-impact\nContact: dell.com/en-us/work/forms/contact-us",
     "dell.yourcause.com\ndell.com/en-us/dt/corporate/social-impact",
     "Dell Technologies Social Impact team\nContact via dell.com contact form",
     "YES — register on YourCause + approach Social Impact team",
     "1. Register on YourCause (NPOconnect)\n2. Contact Dell Social Impact team via website\n3. June = Dell Global Month of Service — perfect time to reach out\n4. PITCH: circular economy / e-waste management aligns with Dell's Planet programme",
     "🔴 HIGH"),

    ("Goldman Sachs","Banking","Office of Corporate Engagement",
     "Already listed above — duplicate removed","","","","",""),

    ("SAP","Technology","SAP internal portal",
     "SAP Social Sabbatical applications: sap.com/corporate-responsibility\nSAP Together programme: sap.com/corporate-responsibility/social-investment\nEmployee giving: via SAP internal portal\nContact: sap.com/corporate-responsibility/contact.html",
     "sap.com/corporate-responsibility\nsap.com/corporate-responsibility/social-investment",
     "SAP Corporate Responsibility team\nSocial Sabbatical applications via sap.com/corporate-responsibility",
     "YES — apply for Social Sabbatical + contact CSR team",
     "1. Apply for SAP Social Sabbatical (4-week paid volunteer placement for SAP employee)\n2. Contact via sap.com/corporate-responsibility/contact.html\n3. SAP has major Bengaluru office — India employees directly relevant\n4. $20/hour volunteer grant = incentivize SAP employees to visit WW sites",
     "🔴 HIGH"),

    ("Qualcomm","Technology/Semiconductor","Benevity",
     "Qualcomm CSR: qualcomm.com/company/positions/corporate-responsibility\nQualcomm Foundation: qualcommfoundation.org\nBenevity: causes@benevity.org",
     "qualcomm.com/company/positions/corporate-responsibility\nqualcommfoundation.org",
     "Qualcomm Foundation\nContact via qualcommfoundation.org or qualcomm.com/contact",
     "YES — register on Benevity + apply to Qualcomm Foundation",
     "1. Register on Benevity (Qualcomm uses Benevity)\n2. Apply to Qualcomm Foundation for grants (STEM/education/environment focus)\n3. Note Oct–Sep fiscal year — reach out by August for Oct start",
     "🟡 MEDIUM"),

    # ── INDIA ────────────────────────────────────────────────────
    ("HCL Foundation (India)","Technology/IT Services","Direct grant application",
     "DIRECT EMAIL: helgrant@hcl.com ✓\nDIRECT PHONE: 1800 572 0608 (Mon–Fri 9am–6pm)\nGrant portal: hclfoundation.org/hcltech-grant\nEdition XI: applications were May 1 – June 25, 2025",
     "hclfoundation.org/hcltech-grant",
     "HCL Foundation\nDIRECT: helgrant@hcl.com ✓\nPhone: 1800 572 0608",
     "YES — open competitive grant; ₹5 crore per winner, ₹50 lakh per runner-up",
     "1. Email helgrant@hcl.com immediately\n2. Apply at hclfoundation.org/hcltech-grant (check Edition XII dates)\n3. Categories: Water + Biodiversity = EXACT FIT\n4. Grant: ₹5 Crore (winner) or ₹50 Lakhs (runner-up) for 2–4 years!",
     "🔴 HIGHEST (India)"),

    ("Wipro / Wipro Cares (India)","Technology/IT Services","Wipro Foundation direct contact",
     "Wipro Cares: wiprofoundation.org/wipro-cares/\nWipro Foundation FAQs: wiprofoundation.org/faqs/\nWipro CSR: wipro.com/investors/corporate-governance/corporate-social-responsibility/\nContact: wiprofoundation.org/contact/",
     "wiprofoundation.org/wipro-cares/",
     "Wipro Foundation team\nContact: wiprofoundation.org/contact/\n45,000+ employees contribute; 1:1 Wipro match",
     "YES — apply as Wipro Cares partner NGO",
     "1. Contact Wipro Foundation via wiprofoundation.org/contact/\n2. Apply as Wipro Cares ecology/environment partner\n3. PITCH: Wipro's ecology focus + Himalayan ecosystem = direct match\n4. 45,000 employees × ₹5/day = significant payroll giving potential",
     "🔴 HIGHEST (India)"),

    ("HDFC Bank (India)","Banking/Finance","GiveIndia platform + Parivartan",
     "GiveIndia: giveindia.org/ngo-onboarding\nHDFC CSR: v.hdfc.bank.in/csr/index.html\nParivartan: hdfc.bank.in/personal/about-us/csr\nGiveIndia partner email: partners@giveindia.org",
     "giveindia.org/ngo-onboarding\nv.hdfc.bank.in/csr",
     "GiveIndia (platform):\npartners@giveindia.org ✓\nHDFC Bank CSR team via v.hdfc.bank.in/csr",
     "YES — register on GiveIndia; HDFC Bank matches 1:1",
     "1. Email partners@giveindia.org to register Waste Warriors on GiveIndia\n2. HDFC Bank routes payroll giving through GiveIndia — 1:1 match\n3. Also apply for Parivartan CSR grants directly via HDFC Bank\n4. Environment category is available",
     "🔴 HIGHEST (India)"),

    ("Microsoft India","Technology","Microsoft India Employee Giving",
     "Microsoft India Employee Giving: microsoft.com/en-in/about/employee-giving\nMicrosoft India CSR: microsoft.com/en-in/about/corporate-responsibility\nContact: microsoft.com/en-in/about/contact",
     "microsoft.com/en-in/about/employee-giving",
     "Microsoft India Corporate Responsibility team\nContact via microsoft.com/en-in/about/contact",
     "YES — apply to be listed as eligible NGO for October campaign",
     "1. Contact Microsoft India CSR team directly\n2. Apply to be listed for October Employee Giving Campaign\n3. INR 2,50,000 match (engineering) + INR 50,000 (others)\n4. 3 paid volunteer days: invite Microsoft employees to WW site visits",
     "🔴 HIGH (India)"),

    ("TCS Foundation (India)","Technology/IT Services","HOPE / TECH4HOPE / direct CSR",
     "TCS Foundation: tcsfoundation.in\nTCS CSR: tcs.com/who-we-are/corporate-social-responsibility\nHOPE programme: tcs.com/who-we-are/corporate-social-responsibility (HOPE section)\nCSR spend: ₹1,038 crores FY25",
     "tcsfoundation.in\ntcs.com/who-we-are/corporate-social-responsibility",
     "TCS Foundation team\nContact via tcsfoundation.in/contact-us\nor email csr@tcs.com",
     "YES — HOPE partner NGOs accepted; TECH4HOPE pro bono for NGOs",
     "1. Apply as HOPE programme partner NGO via tcsfoundation.in\n2. Request TECH4HOPE (free tech help: databases, tracking systems, website)\n3. Email csr@tcs.com with 1-page introduction\n4. 8.9 million volunteer hours = enormous volunteer potential",
     "🔴 HIGH (India)"),
]

# remove duplicate Goldman row
contacts = [r for r in contacts if not (r[0] == "Goldman Sachs" and r[3] == "Already listed above — duplicate removed")]

for ri, row in enumerate(contacts, 3):
    sec = row[1]
    if "HIGHEST" in row[8]:   fill = "FFD6D6"
    elif "🔴 HIGH" in row[8]: fill = LT_AMBER
    elif "India" in row[8]:   fill = LT_GREEN
    else:                      fill = ALT[(ri-3)%2]
    for ci, val in enumerate(row, 1):
        c = ws1.cell(row=ri, column=ci, value=str(val))
        c.fill = hf(fill); c.font = fn(DARK_TEXT, b=(ci==1), s=9)
        c.alignment = wa(); c.border = brd
    ws1.row_dimensions[ri].height = 90

cw1 = [18, 18, 24, 42, 42, 38, 22, 55, 16]
for i, w in enumerate(cw1, 1): ws1.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════
# SHEET 2 – EMAIL TEMPLATES (6 TYPES)
# ══════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Email Templates")
ws2.sheet_view.showGridLines = False; ws2.freeze_panes = "A3"

title_row(ws2, "WASTE WARRIORS — 6 TAILORED EMAIL TEMPLATES FOR DIRECT CORPORATE OUTREACH", 5, FOREST)
hdr(ws2, ["Template Type","Companies to Use With","Subject Line","Email Body (Copy-paste ready)","Follow-up Plan"], bg=MID_GREEN)

templates = [
    ("TYPE A\nBenevity Platform Companies\n(One email unlocks 12+ companies)",
     "Microsoft, Google, Apple, IBM, NVIDIA, Cisco, Adobe, Bank of America, Nike, Qualcomm, NetApp, Intel",
     "Waste Warriors on Benevity — Himalayan Waste Management NGO | Employee Giving Partnership",
     """Dear [Company] Corporate Social Responsibility Team,

My name is [Your Name], and I am writing on behalf of Waste Warriors Society (wastewarriors.org) — a registered environmental NGO working to solve the waste management crisis in the Indian Himalayan Region.

About Waste Warriors:
• Founded 2012 | Registered Society | 12A ✓ | 80G ✓ | FCRA ✓
• Operating in Himachal Pradesh and Uttarakhand (Himalayan eco-zone)
• 2023 Keeling Curve Prize Winner (Top 10 global climate solutions)
• Programs: Community waste management, Green Gurukul (school education), Swachhata Ki Pathshala, Young Warriors Club
• Impact: Tonnes of waste diverted from rivers and forests; thousands of community members activated

We are registered on Benevity (Causes Portal) and are now reaching out directly to companies whose employees can make a difference. Your organization's employee giving programme is a perfect match for our mission — particularly given [Company]'s commitments to [sustainability/net-zero/environmental goals].

What we are asking:
1. Feature Waste Warriors in your next employee giving campaign or newsletter
2. Share our Benevity profile (search "Waste Warriors" on [company].benevity.org) with employees who care about the environment and climate
3. If applicable, help us connect with your [India/CSR/Philanthropy] team for a deeper partnership

We would love to schedule a 20-minute call to share our work and explore how [Company] employees can have a direct, measurable impact on one of India's most fragile ecosystems.

Thank you for your time and for your organization's commitment to making a difference.

Warm regards,
[Your Name]
[Title] | Waste Warriors Society
📧 [email] | 📞 [phone] | 🌐 wastewarriors.org
Registered: 12A | 80G | FCRA | Keeling Curve Prize 2023""",
     "Day 1: Send email\nDay 8: LinkedIn follow-up (connect + brief message)\nDay 21: Second email if no reply (shorter, 3 sentences)\nDay 45: Try a different contact at same company"),

    ("TYPE B\nDirect Email Contacts Found\n(Use exact names/emails)",
     "Cisco (communityimpact@cisco.com)\nDeloitte (kshoztic@deloitte.com)\nMerck (merckp4g@easymatch.com + leslie_hardy@merck.com)\nStryker (info@strykerjohnstonfoundation.org)\nJ&J (+1 732-524-2455)\nApple (apple@globalgiving.org)",
     "[CISCO] Waste Warriors — NGO Listing Request | Benevity Community Impact Portal",
     """Dear [Name / Team],

I am reaching out on behalf of Waste Warriors Society (wastewarriors.org), a Keeling Curve Prize-winning environmental NGO working on solid waste management in India's Himalayan Region (Himachal Pradesh & Uttarakhand).

Organisation Details:
• Registered Society | 12A | 80G | FCRA registered (eligible to receive international funds)
• Founded 2012 | 12+ years of field operations
• Focus: Systemic waste management, river/forest protection, community education
• Award: Keeling Curve Prize 2023 — Top 10 global climate solutions

We are already registered on Benevity and are writing to request:
1. Listing / verification as an eligible NGO in [Company]'s employee giving programme
2. The opportunity to be featured in a monthly giving campaign or newsletter to employees
3. A brief introductory call with your CSR/community impact team

[FOR STRYKER SPECIFICALLY: We also work on medical waste management in remote Himalayan communities near government health camps — which we believe aligns closely with Stryker's mission in healthcare environments.]

[FOR MERCK/J&J/PFIZER: Waste, open burning, and contaminated water directly affect community health outcomes. Our work reduces disease burden in underserved mountain communities.]

Our FCRA registration, 80G status, and CAF America Equivalency Determination (in process) ensure full compliance for international fund transfers.

Would you be available for a 20-minute introduction call this week or next?

Best regards,
[Your Name]
Waste Warriors Society | wastewarriors.org
📧 [email] | 📞 [phone]""",
     "Day 1: Send personalised version to each direct contact\nDay 3: Call the phone numbers provided (Merck: 1-866-205-2857; J&J: +1-732-524-2455; Stryker: 269-488-8484)\nDay 10: Follow up with 3-line email\nDay 21: Try LinkedIn if no response"),

    ("TYPE C\nGoldman Sachs\n(Community TeamWorks + GS Gives)",
     "Goldman Sachs (Office of Corporate Engagement)\ngoldmansachs.com/community-impact",
     "Waste Warriors — NGO Partner Proposal | Community TeamWorks + Employee Matching | Himalayas",
     """Dear Goldman Sachs Office of Corporate Engagement,

I am writing on behalf of Waste Warriors Society, a Keeling Curve Prize 2023 laureate working on transformative waste management in India's Himalayan ecosystem.

Goldman Sachs' $20,000 employee matching programme and Community TeamWorks initiative are well known for their depth of impact. We believe Waste Warriors represents a compelling opportunity for Goldman Sachs employees — particularly those in your Bengaluru and Mumbai offices — to engage meaningfully with India's environmental challenges.

About Waste Warriors:
• 12+ years operating in Himachal Pradesh and Uttarakhand
• FCRA registered + 80G + 12A certified
• Keeling Curve Prize 2023 (Top 10 global climate solutions)
• Impact: tonnes of waste removed from protected Himalayan river catchments; thousands of community members trained in waste practices

What we propose:
1. List Waste Warriors as an eligible nonprofit for Goldman Sachs' employee matching programme
2. Organise a Community TeamWorks volunteer day at our Waste Warriors site (Dharamshala / Dehradun area) for Goldman Sachs India employees
3. Share our work via Goldman Sachs' internal giving channels

Goldman Sachs has disbursed over $2.7 billion through Goldman Sachs Gives and partnered with 10,000+ nonprofits. We would be honoured to be one of them.

I would welcome a brief call with your team at your convenience.

Warm regards,
[Your Name] | Waste Warriors Society
wastewarriors.org | [email] | [phone]""",
     "Day 1: Submit via goldmansachs.com/community-impact contact form + email\nDay 7: LinkedIn message to Goldman Sachs Corporate Engagement team members (Bengaluru/Mumbai)\nDay 21: Second follow-up email"),

    ("TYPE D\nHCL Foundation (India)\n(HIGHEST PRIORITY — ₹5 Crore available)",
     "HCL Foundation\nEmail: helgrant@hcl.com\nPhone: 1800 572 0608",
     "HCLTech Grant Application — Waste Warriors Society | Water + Biodiversity Category",
     """Dear HCL Foundation Team,

I am writing to express Waste Warriors Society's strong interest in applying for the HCLTech Grant (Edition XII).

Organisation: Waste Warriors Society
Registration: Registered Society | 12A | 80G | FCRA
Founded: 2012 | Location: Dharamshala (HP) and Dehradun (Uttarakhand)
Award: Keeling Curve Prize 2023 — Top 10 global climate solutions

Why we fit the HCLTech Grant:

WATER CATEGORY: Waste Warriors addresses plastic and solid waste contamination in Himalayan rivers — directly threatening freshwater ecosystems that supply water to millions downstream. Our interventions include riverside clean-ups, waste collection infrastructure, and community behaviour change.

BIODIVERSITY CATEGORY: The Himalayan ecosystem is one of the world's most biodiverse and fragile regions. Unmanaged waste in forests and river catchments devastates wildlife habitat, migration corridors, and plant diversity. Waste Warriors' systematic waste management approach directly protects this biodiversity.

We would like to:
1. Confirm the current application timeline for Edition XII
2. Understand any pre-qualification steps we should take now
3. Schedule a call with the Foundation team to discuss our proposal

We have been operating for 12+ years with proven, measurable impact. Our team would welcome the opportunity to present our work and demonstrate why Waste Warriors is an ideal partner for HCLFoundation's water and biodiversity mission.

Please feel free to call us at [phone] or email [email].

Warm regards,
[Your Name]
Waste Warriors Society | wastewarriors.org
📞 [phone] | 📧 [email]
Awards: Keeling Curve Prize 2023""",
     "Day 1: Email helgrant@hcl.com\nDay 2: Call 1800 572 0608 to confirm receipt and ask about Edition XII dates\nDay 7: Follow up with any additional documents requested\nSubmit formal application as soon as portal opens"),

    ("TYPE E\nIndia: Wipro Cares + HDFC + TCS",
     "Wipro Foundation (wiprofoundation.org/contact)\nGiveIndia (partners@giveindia.org)\nTCS Foundation (csr@tcs.com / tcsfoundation.in)",
     "Waste Warriors Society — Partner NGO Application | Ecology / Environment Category",
     """Dear [Wipro Foundation / GiveIndia / TCS Foundation] Team,

I am writing on behalf of Waste Warriors Society (wastewarriors.org), a registered NGO solving the waste crisis in India's Himalayan Region since 2012.

Our Credentials:
• Registered Society | 12A | 80G | FCRA
• Operations: Himachal Pradesh (Dharamshala) and Uttarakhand (Dehradun, Corbett region)
• Keeling Curve Prize 2023 — Top 10 global climate solutions
• Programs: Community waste management, Swachhata Ki Pathshala (schools), Green Gurukul, Young Warriors Club

We are writing to apply as a partner NGO under your [Ecology / Environment / Disaster Response] category.

[FOR WIPRO CARES]: Wipro's ecology focus directly aligns with our work. We would like to be listed as an eligible organisation for Wipro Cares, so Wipro employees can direct their daily contributions towards protecting the Himalayan ecosystem — with Wipro's 1:1 match doubling the impact.

[FOR GIVEINDIA / HDFC BANK]: We would like to register on the GiveIndia platform as an eligible NGO. This will allow HDFC Bank employees and other corporate payroll giving participants to choose Waste Warriors as their cause — with HDFC's 1:1 match multiplying every rupee.

[FOR TCS FOUNDATION]: We would like to explore: (1) HOPE programme partnership — TCS employee volunteers at Waste Warriors sites; (2) TECH4HOPE — pro bono technology support for our waste tracking and community management systems.

All required documentation (12A, 80G, registration, 3-year financials, annual report) is available on request.

We would welcome a brief introduction call at your earliest convenience.

Warm regards,
[Your Name]
Waste Warriors Society | wastewarriors.org
📞 [phone] | 📧 [email]""",
     "Wipro: Call wiprofoundation.org/contact/ Day 1; follow up Day 10\nGiveIndia: Email partners@giveindia.org Day 1; they typically respond in 5–7 days\nTCS: Email csr@tcs.com + apply at tcsfoundation.in"),

    ("TYPE F\nLinkedIn Connection Note\n(When emailing doesn't work — 200 chars max)",
     "ALL companies — for CSR/Philanthropy managers found on LinkedIn",
     "N/A — LinkedIn message (300 char limit)",
     """CONNECTION REQUEST NOTE (200 chars max):

Hi [Name], I lead Waste Warriors Society — we won the 2023 Keeling Curve Prize for Himalayan waste management. I'd love to share how [Company] employees can make a real impact in the Himalayas. Would you be open to connecting?

——

FOLLOW-UP MESSAGE (after acceptance, 300 words):

Hi [Name], thank you for connecting!

Waste Warriors Society (wastewarriors.org) is a Keeling Curve Prize 2023-winning NGO solving the solid waste crisis in the Indian Himalayas (Himachal Pradesh and Uttarakhand). We've been operating for 12+ years, are registered with 12A, 80G, and FCRA, and are listed on Benevity.

I'm reaching out because [Company]'s employee giving programme is something we believe could unlock real, measurable impact for the Himalayan ecosystem — one of the world's most biodiverse and fragile regions.

Specifically, I'd love to explore:
1. Listing Waste Warriors as a featured NGO in [Company]'s employee giving campaigns
2. Connecting with your India CSR team for local employee volunteer engagement
3. Any upcoming giving campaigns or company match programmes we should be aware of

Would you have 20 minutes for a call? I'm happy to share our impact report and a short video of our work.

Thank you — and thanks again for the work you do on [Company]'s CSR programmes.

[Your Name] | Waste Warriors Society

HOW TO FIND THE RIGHT PEOPLE ON LINKEDIN:
Search: "[Company name]" + "CSR" OR "Corporate Giving" OR "Social Impact" OR "Philanthropy" OR "Employee Engagement"
Titles to look for: Director/Manager/VP of: Corporate Social Responsibility, Employee Giving, Philanthropy, Social Impact, Community Affairs, Foundation
Filter by: India (for India companies) or USA (for US companies)""",
     "Accept connection: respond within 24 hours with follow-up message\nNo response in 7 days: try another person at same company\nLinkedIn InMail (paid): use for senior contacts who haven't accepted"),
]

for ri, row in enumerate(templates, 3):
    colors = ["FFD6D6", LT_AMBER, LT_BLUE, LT_GREEN, LT_PURPLE, LT_GREY]
    fill = colors[(ri-3) % len(colors)]
    for ci, val in enumerate(row, 1):
        c = ws2.cell(row=ri, column=ci, value=str(val))
        c.fill = hf(fill); c.font = fn(DARK_TEXT, b=(ci==1), s=9)
        c.alignment = wa(); c.border = brd
    ws2.row_dimensions[ri].height = 350

cw2 = [22, 30, 45, 100, 35]
for i, w in enumerate(cw2, 1): ws2.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════
# SHEET 3 – LINKEDIN SEARCH GUIDE (DECISION-MAKER FINDER)
# ══════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("LinkedIn Decision-Maker Guide")
ws3.sheet_view.showGridLines = False; ws3.freeze_panes = "A3"

title_row(ws3, "WASTE WARRIORS — LINKEDIN STRATEGY: HOW TO FIND + CONTACT EVERY COMPANY'S CSR DECISION-MAKER", 7, TEAL)
hdr(ws3, ["Company","LinkedIn Search Keywords","Job Titles to Target",
          "India Office Focus","Key People Strategy","Message Angle","Priority"], bg=TEAL)

linkedin_data = [
    ("Microsoft","'Microsoft' + 'Employee Giving' OR 'Philanthropy'",
     "Sr Director Employee Giving\nCorporate Philanthropy Manager\nSocial Impact Lead",
     "Microsoft India: Hyderabad, Bengaluru, Noida\nSearch: Microsoft India CSR",
     "Find: Sonia Fiorenza-Lowe (Sr Director Employee Giving) or equivalent\nAlso: Microsoft India Corporate Responsibility Director",
     "Already on Benevity — need them to feature WW in Oct campaign newsletter to 50,000+ India employees",
     "🔴 HIGH"),
    ("Google","'Google' OR 'Alphabet' + 'Corporate Giving' OR 'Social Impact'",
     "Manager Corporate Giving\nGoogle.org Programme Officer\nEmployee Engagement Lead",
     "Google India: Hyderabad, Bengaluru, Mumbai\nSearch: Google India Social Impact",
     "Google.org has India-specific teams\nAlso target: Google for Nonprofits India contact",
     "Already on Benevity. Ask to be featured in Google's annual giving campaign (Dec deadline)",
     "🔴 HIGH"),
    ("Apple","'Apple' + 'Community' OR 'Philanthropy' OR 'Corporate Giving'",
     "Corporate Responsibility Manager\nPhilanthropy Lead",
     "Apple India: Hyderabad\nSmall India team — focus on US contacts",
     "Apple's giving is via Benevity + GlobalGiving.\nKey contact route: apple@globalgiving.org (not LinkedIn)",
     "Follow up after GlobalGiving registration. LinkedIn reinforces the email.",
     "🟡 MEDIUM"),
    ("IBM","'IBM' + 'Corporate Citizenship' OR 'CSR' OR 'Foundation'",
     "Corporate Citizenship Manager\nIBM Foundation Programme Officer\nSocial Impact Director",
     "IBM India: Bengaluru, Pune, Chennai, Noida\nIBM India Corporate Citizenship team exists",
     "Search 'IBM India Corporate Citizenship' on LinkedIn\nIBM has dedicated India CSR team",
     "IBM's Sustainability Accelerator: apply as an environment NGO. Large India team = direct volunteer access.",
     "🔴 HIGH"),
    ("Cisco","'Cisco' + 'Community Impact' OR 'CSR' OR 'Foundation'",
     "Director Community Impact\nCommunity Affairs Manager\nCSR Lead",
     "Cisco India: Bengaluru (large office)\nSearch: Cisco India Social Impact",
     "Already have direct email: communityimpact@cisco.com\nLinkedIn: Backup for no email response",
     "Cisco's 85% employee participation rate + 80-hr VTO = most active giving programme. Ask for India volunteer trip.",
     "🔴 HIGHEST"),
    ("NVIDIA","'NVIDIA' + 'Foundation' OR 'CSR' OR 'Corporate Giving'",
     "NVIDIA Foundation Manager\nCorporate Responsibility Director\nInspire 365 Programme Manager",
     "NVIDIA India: Pune, Bengaluru\nSmaller India CSR team",
     "NVIDIA Foundation managed via SVCF.\nSearch for NVIDIA Foundation staff on LinkedIn.",
     "Inspire 365 is active — already on Benevity. LinkedIn to request India employee volunteer day.",
     "🟡 MEDIUM"),
    ("Goldman Sachs","'Goldman Sachs' + 'Corporate Engagement' OR 'Community Affairs'",
     "VP Corporate Engagement\nCommunity Relations Manager\nESG / Social Impact",
     "Goldman Sachs India: Bengaluru, Hyderabad, Mumbai\nGS India Corporate Engagement team",
     "Goldman Sachs India has dedicated Corporate Engagement team.\nSearch 'Goldman Sachs India Corporate Engagement'",
     "$20,000 match cap. India offices are large. Community TeamWorks volunteer events in India = strong pitch.",
     "🔴 HIGHEST"),
    ("Deloitte","'Deloitte' + 'Foundation' OR 'CSR' OR 'WorldClass'",
     "Deloitte Foundation Director\nCSR Manager\nWorldClass Initiative Lead",
     "Deloitte India: Mumbai, Delhi, Bengaluru, Hyderabad\nDeloitte India CSR team",
     "Deloitte India has significant CSR programme.\nSearch 'Deloitte India Corporate Responsibility'",
     "Direct contact: kshoztic@deloitte.com — use LinkedIn as backup only.\n$32,500 cap = absolute priority.",
     "🔴 HIGHEST"),
    ("J&J","'Johnson Johnson' + 'Foundation' OR 'Corporate Responsibility'",
     "J&J Foundation Programme Manager\nCorporate Responsibility Director",
     "J&J India: Mumbai, Bengaluru\nJ&J India CSR is active",
     "J&J India has dedicated Corporate Responsibility team.\nSearch 'Johnson Johnson India Corporate Responsibility'",
     "2:1 match. Phone: +1-732-524-2455. LinkedIn for India team specifically.",
     "🔴 HIGHEST"),
    ("Merck","'Merck' + 'Foundation' OR 'Partnership for Giving' OR 'CSR'",
     "Merck Foundation Director\nCorporate Giving Manager",
     "Merck India: Mumbai (Merck has India presence)\nSearch 'Merck India CSR'",
     "Direct contacts: merckp4g@easymatch.com + leslie_hardy@merck.com\nLinkedIn as backup",
     "$30,000 cap. Email is primary route. LinkedIn reinforces.",
     "🔴 HIGH"),
    ("Stryker","'Stryker' + 'Johnston Foundation' OR 'Corporate Responsibility' OR 'Impact'",
     "Stryker Johnston Foundation Director\nCommunity Affairs Manager\nCSR Lead",
     "Stryker India: Gurgaon, Pune\nMedical device company — India operations",
     "Direct: info@strykerjohnstonfoundation.org\nLinkedIn: Stryker India CSR team",
     "Medical waste angle. Stryker India employees can volunteer at WW mountain health camp waste clean-ups.",
     "🔴 HIGH"),
    ("Coca-Cola","'Coca-Cola' + 'Foundation' OR 'Sustainability' OR 'CSR'",
     "Foundation Programme Manager\nSustainability Lead\nCommunity Affairs Director",
     "Coca-Cola India: Gurgaon, Bengaluru, Mumbai\nCoca-Cola India Sustainability team",
     "Coca-Cola India has strong sustainability focus.\nSearch 'Coca-Cola India Sustainability'",
     "2:1 match. Plastic waste is a key Coca-Cola issue globally — perfect alignment with WW's plastics work.",
     "🔴 HIGH"),
    ("PepsiCo","'PepsiCo' + 'Foundation' OR 'pep+' OR 'CSR' OR 'Sustainability'",
     "PepsiCo Foundation Manager\nSustainability Lead\npep+ Programme Manager",
     "PepsiCo India: Gurgaon, Mumbai\nPepsiCo India Sustainability team",
     "PepsiCo has large India presence and pep+ sustainability programme.\nSearch 'PepsiCo India pep+ OR Sustainability'",
     "pep+ focus = waste and water = direct WW fit. YourCause is the matching platform.",
     "🔴 HIGH"),
    ("HCL Foundation","'HCL Technologies' + 'Foundation' OR 'CSR' OR 'Tech Grant'",
     "HCL Foundation Director\nHCLTech Grant Manager\nCSR Programme Officer",
     "HCL India: Noida (HQ), Bengaluru, Chennai\nHCL Foundation is based in Noida",
     "MOST IMPORTANT India LinkedIn target.\nSearch 'HCL Foundation' directly on LinkedIn",
     "helgrant@hcl.com is primary. LinkedIn to find grant programme managers for follow-up + relationship building.",
     "🔴 HIGHEST (India)"),
    ("Wipro / Wipro Foundation","'Wipro' + 'Foundation' OR 'Wipro Cares' OR 'CSR'",
     "Wipro Foundation Director\nWipro Cares Programme Manager\nEcology Programme Officer",
     "Wipro India: Bengaluru (HQ), Hyderabad, Pune, Delhi\nWipro Foundation based in Bengaluru",
     "Wipro Foundation is Bengaluru-based.\nSearch 'Wipro Foundation Ecology' or 'Wipro Cares'",
     "Wipro Cares 1:1 match + 45,000 employees = powerful. Ecology track is exact fit.",
     "🔴 HIGHEST (India)"),
    ("TCS Foundation","'TCS' OR 'Tata Consultancy' + 'Foundation' OR 'HOPE' OR 'CSR'",
     "TCS Foundation Director\nHOPE Programme Manager\nCSR Lead",
     "TCS India: Mumbai (HQ), Pune, Chennai, Bengaluru, Hyderabad, Noida\nTCS Foundation based in Mumbai",
     "TCS Foundation has dedicated team.\nSearch 'TCS Foundation' or 'TCS HOPE Programme'",
     "8.9 million volunteer hours. Ask for HOPE partnership + TECH4HOPE. India's largest IT employer.",
     "🔴 HIGH (India)"),
    ("HDFC Bank","'HDFC Bank' + 'CSR' OR 'Parivartan' OR 'Corporate Responsibility'",
     "HDFC Bank CSR Director\nParivartan Programme Manager\nSustainability Lead",
     "HDFC Bank India: Mumbai (HQ), pan-India\nHDFC CSR is Mumbai-based",
     "HDFC Bank is India-headquartered.\nSearch 'HDFC Bank Parivartan' or 'HDFC Bank CSR'",
     "Register on GiveIndia first. LinkedIn to find Parivartan programme manager for direct conversation.",
     "🔴 HIGH (India)"),
]

for ri, row in enumerate(linkedin_data, 3):
    fill = ALT[(ri-3)%2]
    for ci, val in enumerate(row, 1):
        c = ws3.cell(row=ri, column=ci, value=str(val))
        c.fill = hf(fill); c.font = fn(DARK_TEXT, b=(ci==1), s=9)
        c.alignment = wa(); c.border = brd
    ws3.row_dimensions[ri].height = 80

cw3 = [16, 32, 28, 28, 40, 38, 14]
for i, w in enumerate(cw3, 1): ws3.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════
# SHEET 4 – MASTER OUTREACH CALENDAR
# ══════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("30-Day Outreach Calendar")
ws4.sheet_view.showGridLines = False; ws4.freeze_panes = "A3"

title_row(ws4, "WASTE WARRIORS — 30-DAY DIRECT OUTREACH CALENDAR (WHO TO CONTACT, WHEN, HOW)", 7, DARK_NAVY)
hdr(ws4, ["Day","Action","Company / Platform","Contact (Email/Phone/URL)","What to Say (in 1 line)","Expected Response Time","Status"])

calendar = [
    ("DAY 1", "Email Template B — Direct contact", "Cisco", "communityimpact@cisco.com", "NGO listing request on Benevity Community Impact Portal", "3–5 business days", "☐ Pending"),
    ("DAY 1", "Email Template B — Direct contact", "Merck (matching)", "merckp4g@easymatch.com", "Register WW as eligible NGO on CyberGrants Partnership for Giving", "5–7 days", "☐ Pending"),
    ("DAY 1", "Email Template B — Direct contact", "Merck (foundation)", "leslie_hardy@merck.com", "Merck Foundation partnership + grant consideration", "1–2 weeks", "☐ Pending"),
    ("DAY 1", "Email Template B — Direct contact", "Deloitte Foundation", "kshoztic@deloitte.com", "WW introduction + request listing for employee matching + Impact Day", "1–2 weeks", "☐ Pending"),
    ("DAY 1", "Email Template B — Direct contact", "Stryker Johnston Foundation", "info@strykerjohnstonfoundation.org", "NGO listing + grant application + medical waste angle pitch", "1–2 weeks", "☐ Pending"),
    ("DAY 1", "Email Template D — HCL Grant", "HCL Foundation", "helgrant@hcl.com", "HCLTech Grant Edition XII inquiry + Water + Biodiversity categories", "3–5 days", "☐ Pending"),
    ("DAY 1", "Email Template E — India", "GiveIndia", "partners@giveindia.org", "Register WW for HDFC Bank payroll giving platform", "5–7 days", "☐ Pending"),
    ("DAY 1", "Email Template E — India", "Wipro Foundation", "wiprofoundation.org/contact/", "Apply as Wipro Cares ecology partner NGO", "1–2 weeks", "☐ Pending"),
    ("DAY 2", "Phone call follow-up", "HCL Foundation", "1800 572 0608 (9am–6pm IST)", "Confirm receipt of Day 1 email; ask Edition XII dates", "Same day", "☐ Pending"),
    ("DAY 2", "Phone call — J&J Foundation", "Johnson & Johnson", "+1 (732) 524-2455", "Introduction + request for NGO listing + 2:1 match programme", "Same call", "☐ Pending"),
    ("DAY 2", "Phone call — Merck P4G", "Merck Partnership for Giving", "1-866-205-2857", "Confirm NGO registration + matching gift eligibility", "Same call", "☐ Pending"),
    ("DAY 3", "Email Template A — Benevity", "Microsoft (India focus)", "philanthropies@microsoft.com", "WW on Benevity + request October campaign feature in India", "1–2 weeks", "☐ Pending"),
    ("DAY 3", "Email Template A — Benevity", "Google", "support.google.com/nonprofits + google.org", "Apply for Google for Nonprofits + Benevity listing confirmation", "1–2 weeks", "☐ Pending"),
    ("DAY 3", "Email Template B — Apple", "Apple (GlobalGiving)", "apple@globalgiving.org", "Start Apple due diligence process via GlobalGiving", "2–3 weeks", "☐ Pending"),
    ("DAY 3", "Email Template E — India", "TCS Foundation", "csr@tcs.com", "HOPE partner NGO + TECH4HOPE request", "1–2 weeks", "☐ Pending"),
    ("DAY 4", "Register on platforms", "Benevity", "causeshelp.benevity.org", "Complete full registration if not yet done", "2–4 week approval", "☐ Pending"),
    ("DAY 4", "Register on platforms", "YourCause NPOconnect", "yourcause.com/nonprofits", "Register for PepsiCo, Dell, Wells Fargo, AmEx coverage", "2–3 week approval", "☐ Pending"),
    ("DAY 4", "Register on platforms", "GlobalGiving", "globalgiving.org/become-a-partner", "Start GlobalGiving partner registration for AmEx Give2Gether", "6–10 weeks", "☐ Pending"),
    ("DAY 5", "Email Template C — Goldman Sachs", "Goldman Sachs", "goldmansachs.com/community-impact form", "Community TeamWorks + $20K matching + India office engagement", "1–2 weeks", "☐ Pending"),
    ("DAY 5", "Email Template A — Benevity", "IBM", "corporatecitizenship@us.ibm.com", "WW on Benevity + IBM India CSR partnership + TECH4GOOD", "1–2 weeks", "☐ Pending"),
    ("DAY 5", "Email Template A — Benevity", "Adobe", "adobe.com/corporate-responsibility contact form", "Benevity listing + Adobe ECF grant application request", "1–2 weeks", "☐ Pending"),
    ("DAY 7", "LinkedIn outreach wave 1", "ALL companies", "LinkedIn (search CSR/Philanthropy managers)", "Send 5–10 connection requests with TYPE F note (200 chars)", "3–7 days to accept", "☐ Pending"),
    ("DAY 8", "Email Template A — Benevity", "Cisco (2nd touch)", "communityimpact@cisco.com", "Follow up on Day 1 email — offer 20-min call", "2–3 days", "☐ Pending"),
    ("DAY 10", "Email — Coca-Cola", "Coca-Cola Foundation", "coca-colacompany.com/social/coca-cola-foundation", "Apply via online system + Coca-Cola India sustainability team", "2–3 weeks", "☐ Pending"),
    ("DAY 10", "Email — PepsiCo", "PepsiCo Foundation", "pepsico.com/who-we-are/pepsico-foundation", "pep+ sustainability alignment + YourCause employee matching", "1–2 weeks", "☐ Pending"),
    ("DAY 10", "Email — Mastercard India", "Mastercard India CSR", "mastercard.com/us/en/for-the-world/people/community-impact.html", "Environmental giving alignment + India office engagement", "2–3 weeks", "☐ Pending"),
    ("DAY 14", "LinkedIn follow-up wave 2", "ALL accepted connections", "LinkedIn messages to accepted contacts", "Send Type F follow-up message to all who accepted connection", "7–14 days", "☐ Pending"),
    ("DAY 14", "Bank of America grant application", "Bank of America", "about.bankofamerica.com/en/making-an-impact/charitable-foundation-funding", "Check RFP cycle dates; register on NPO Hub", "2–3 months for grant decision", "☐ Pending"),
    ("DAY 21", "Follow-up emails (no response)", "All Day 1–5 companies", "Same contacts as initial email", "3-line follow-up: 'Did you receive my email? Happy to send our impact report.'", "Immediate", "☐ Pending"),
    ("DAY 30", "Evaluate + Escalate", "All companies", "Track responses in CRM / spreadsheet", "Move responded companies to pipeline; re-approach non-responders via different route", "Ongoing", "☐ Pending"),
]

cal_colors = {"DAY 1": "FFD6D6", "DAY 2": LT_AMBER, "DAY 3": LT_AMBER,
              "DAY 4": LT_GREEN, "DAY 5": LT_BLUE, "DAY 7": LT_PURPLE,
              "DAY 8": LT_AMBER, "DAY 10": LT_GREEN, "DAY 14": LT_BLUE,
              "DAY 21": "FFECB3", "DAY 30": LT_GREY}

for ri, row in enumerate(calendar, 3):
    fill = cal_colors.get(row[0], WHITE)
    for ci, val in enumerate(row, 1):
        c = ws4.cell(row=ri, column=ci, value=str(val))
        c.fill = hf(fill); c.font = fn(DARK_TEXT, b=(ci in [1,2]), s=9)
        c.alignment = wa(); c.border = brd
    ws4.row_dimensions[ri].height = 45

cw4 = [10, 30, 22, 40, 50, 22, 14]
for i, w in enumerate(cw4, 1): ws4.column_dimensions[get_column_letter(i)].width = w

# ══════════════════════════════════════════════════════════════
out = "/home/user/jcode/Waste_Warriors_Direct_Outreach_Playbook.xlsx"
wb.save(out)
print(f"Saved: {out}")
