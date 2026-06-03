import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "CSR & Employee Giving Contacts"
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

# ── palette ──────────────────────────────────────────────────────────
DARK_NAVY="1B2A4A"; MID_BLUE="2E5CA8"; FOREST="1A5C38"; MID_GREEN="217A4F"
AMBER="B8860B"; TEAL="0F4C5C"; WHITE="FFFFFF"; DARK_TEXT="1A1A2E"
LT_GREEN="D5F5E3"; LT_BLUE="D6E4F7"; LT_GREY="F2F2F2"; MED_GREY="CCCCCC"
HI_GREEN="ABEBC6"; ORANGE="E65100"; LT_ORANGE="FFF3CD"; PURPLE="4B0082"; LT_PURPLE="EDE7F6"
DEEP_RED="8B0000"

def hf(c): return PatternFill("solid", fgColor=c)
def fn(c=DARK_TEXT, b=False, s=10): return Font(color=c, bold=b, size=s, name="Calibri")
th = Side(style="thin", color=MED_GREY)
brd = Border(left=th, right=th, top=th, bottom=th)
def wa(h="left", v="center"): return Alignment(horizontal=h, vertical=v, wrap_text=True)

# ── title ─────────────────────────────────────────────────────────────
ws.merge_cells("A1:E1")
c = ws["A1"]
c.value = "WASTE WARRIORS — 100+ CSR & EMPLOYEE GIVING DECISION-MAKERS (INDIA-FIRST)"
c.font = Font(name="Calibri", bold=True, size=13, color=WHITE)
c.fill, c.alignment = hf(DARK_NAVY), Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 32

# ── header row ────────────────────────────────────────────────────────
headers = ["Name", "Org", "Role", "Location", "LinkedIn Link"]
for i, h in enumerate(headers, 1):
    c = ws.cell(2, i, h)
    c.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
    c.fill, c.alignment, c.border = hf(MID_BLUE), wa("center"), brd
ws.row_dimensions[2].height = 26

# ── section divider helper ────────────────────────────────────────────
def section(ws, r, label, bg, txt_color=WHITE):
    ws.merge_cells(f"A{r}:E{r}")
    c = ws[f"A{r}"]
    c.value, c.font = label, Font(name="Calibri", bold=True, size=11, color=txt_color)
    c.fill, c.alignment = hf(bg), Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 22
    return r + 1

# ════════════════════════════════════════════════════════════════════
# DATA — INDIA (priority)
# ════════════════════════════════════════════════════════════════════

INDIA = [
    # ── Already in user's sheet ──────────────────────────────────────
    ("Ashwin Sesodia", "Deloitte", "Lead: Deloitte Employee Giving Program (DEGP)",
     "Hyderabad, India", "linkedin.com/in/ashwin-sesodia-56343916"),

    # ── Tech / IT ─────────────────────────────────────────────────────
    ("Manju Dhasmana", "Microsoft India",
     "Asia Regional Director, Microsoft Philanthropies; Senior Director CSR",
     "New Delhi, India", "linkedin.com/in/manjudhasmana"),
    ("Meenakshi Jha", "Intel India",
     "Sr. Manager, Corporate Social Responsibility",
     "Bengaluru, India", "linkedin.com/in/meenakshi-jha-she-her-5030b873"),
    ("Mudita Lall", "Adobe India",
     "Country Lead, CSR",
     "Noida, India", "linkedin.com/in/muditalall"),
    ("Kiranmayi B", "Salesforce India",
     "Philanthropy Senior Manager",
     "Bengaluru, India", "linkedin.com/in/kiranmayi-b-0582a131"),
    ("Kheerthan Karunakar", "Salesforce India",
     "Corporate Social Responsibility",
     "Bengaluru, India", "linkedin.com/in/kheerthan-karunakar-290344b1"),
    ("Pratibha Sharma", "Salesforce India",
     "Employee Impact / Volunteering Programs",
     "Hyderabad, India", "linkedin.com/in/pratibha-sharma-317852140"),
    ("Smitha N", "SAP India",
     "Corporate Social Responsibility Lead",
     "Bengaluru, India", "linkedin.com/in/smitha-n-4a767611b"),
    ("Hemang Desai", "SAP India",
     "Corporate Social Responsibility",
     "Bengaluru, India", "linkedin.com/in/hemang-desai-ba21194"),
    ("Shipra Sharma", "IBM India",
     "CSR Leader, India / South Asia",
     "Bengaluru, India", "linkedin.com/in/shipra-sharma"),
    ("Harish Krishnan", "Cisco India & SAARC",
     "Managing Director, Public Affairs & Strategic Engagements",
     "Bengaluru, India", "linkedin.com/in/harish-krishnan-0bbb99"),
    ("Anupam Trehan", "Cisco APJC",
     "Vice President, People & Communities, APJC",
     "Bengaluru, India", "linkedin.com/in/anupamshringi"),
    ("Ravi Viswanathan", "Accenture India",
     "Global Giving Lead / Global Giving and Partnerships",
     "Bengaluru, India", "linkedin.com/in/ravivisw"),
    ("Vinitha Nair", "Amazon India",
     "Corporate Social Responsibility",
     "Bengaluru, India", "linkedin.com/in/vinitha-nair-47b784197"),
    ("Sushil Bhatla", "Hewlett Packard Enterprise India",
     "Head – CSR",
     "Bengaluru, India", "linkedin.com/in/sushil-bhatla-39865a128"),
    ("Vijay Singh", "Infosys",
     "Director – Corporate Social Responsibility",
     "Bengaluru, India", "linkedin.com/in/vijay-kr-singh"),
    ("Prashant Hegde", "Infosys Foundation",
     "Programme Officer / CSR",
     "Bengaluru, India", "linkedin.com/in/prashant-hegde-1b164a266"),

    # ── HCL ───────────────────────────────────────────────────────────
    ("Dr. Nidhi Pundhir", "HCL Foundation",
     "Director, HCL Foundation",
     "Noida, India", "linkedin.com/in/dr-nidhi-pundhir-1215b823"),
    ("Anoop Narayan", "HCL Foundation",
     "Project Director / CSR Lead",
     "Noida, India", "linkedin.com/in/anoop-narayan-a400aa12"),
    ("Prakriti Khar", "HCLTech",
     "CSR Engagements, Communication and ESG",
     "Noida, India", "linkedin.com/in/prakriti-khar-19026a67"),
    ("Simi Suri", "HCL Technologies",
     "CSR",
     "Noida, India", "linkedin.com/in/simi-suri-40a74616"),

    # ── Wipro ─────────────────────────────────────────────────────────
    ("Abhijit Paul Zacharia", "Wipro Foundation",
     "Programme Manager & Lead, Sustainability & Education",
     "Bengaluru, India", "linkedin.com/in/abhijit-paul-zacharia-bb62095"),
    ("Praveen Beechagondahalli", "Wipro Foundation",
     "Lead, Wipro Cares (Employee Volunteering & Community Initiatives)",
     "Bengaluru, India", "linkedin.com/in/praveen-beechagondahalli-6542692"),

    # ── TCS ───────────────────────────────────────────────────────────
    ("Jennifer Ghosh", "Tata Consultancy Services (TCS)",
     "CSR Lead",
     "Kolkata, India", "linkedin.com/in/jennifer-ghosh-3a76b7163"),
    ("Tarun Sharma", "Tata Consultancy Services (TCS)",
     "Corporate Social Responsibility Manager",
     "Indore, India", "linkedin.com/in/tarun-sharma-330a4416b"),
    ("Chandrasekar Natarajan", "Tata Consultancy Services (TCS)",
     "Zonal Head – India South & East, Corporate Social Responsibility",
     "Chennai, India", "linkedin.com/in/chandrasekar-natarajan-3ba7128"),
    ("Debasubhra Banerjee", "Tata Consultancy Services (TCS)",
     "Assistant Manager, Corporate Social Responsibility",
     "Kolkata, India", "linkedin.com/in/debasubhra-banerjee"),

    # ── Tata Group ────────────────────────────────────────────────────
    ("Kaushik Das", "Tata Steel",
     "Head – CSR",
     "Jamshedpur, India", "linkedin.com/in/kaushik-das-936476185"),
    ("Nirmal Mahanta", "Tata Steel",
     "Head, CSR (WBOK Division)",
     "Jamshedpur, India", "linkedin.com/in/nirmal-mahanta-0a25a021"),
    ("Randhir Kumar", "Tata Power",
     "Group Head – CSR",
     "Mumbai, India", "linkedin.com/in/randhir-kumar-7b968baa"),
    ("Mukesh Solanki", "Tata Chemicals",
     "Head – CSR",
     "Mithapur, Gujarat, India", "linkedin.com/in/mukesh-solanki-92589033"),
    ("Benny Antony", "Tata Projects",
     "Assistant General Manager – CSR",
     "Hyderabad, India", "linkedin.com/in/benny-antony-3388144a"),
    ("Dilith Castleton", "Tata Metaliks",
     "Head – CSR",
     "Kharagpur, India", "linkedin.com/in/dilith-castleton-98a168175"),

    # ── Mahindra ──────────────────────────────────────────────────────
    ("Nirmal Parmar", "Mahindra and Mahindra",
     "Corporate Social Responsibility Sr. Manager",
     "Mumbai, India", "linkedin.com/in/nirmal-parmar-4640161b"),
    ("Varun Ram", "Mahindra Group",
     "CSR Professional (15+ years)",
     "Mumbai, India", "linkedin.com/in/varun-ram-74009313"),

    # ── Godrej ────────────────────────────────────────────────────────
    ("Pakzan Dastoor", "Godrej",
     "Head, Sustainability and CSR",
     "Mumbai, India", "linkedin.com/in/pakzan-dastoor-80735415b"),
    ("Kiran Gopale", "Godrej Interio (Godrej & Boyce)",
     "ACM & Head of Sustainability",
     "Mumbai, India", "linkedin.com/in/kiran-gopale-a6b91383"),

    # ── Banking (India) ───────────────────────────────────────────────
    ("Nusrat Pathan", "HDFC Bank",
     "Head – CSR (Parivartan Program)",
     "Mumbai, India", "linkedin.com/in/nusrat-pathan-019b5631"),
    ("Ranjan Sharma", "HDFC Bank",
     "Regional CSR Head",
     "India", "linkedin.com/in/ranjansh197545"),
    ("Barun Kumar", "HDFC Bank",
     "Regional Head CSR (East)",
     "Kolkata, India", "linkedin.com/in/barun-kumar-47632423"),
    ("Ameet Shekhar", "HDFC Bank",
     "Impact Strategist & CSR Leader",
     "Mumbai, India", "linkedin.com/in/ameet-shekhar-7654996"),
    ("Animesh Sharma", "ICICI Bank",
     "Manager – ESG & CSR",
     "Mumbai, India", "linkedin.com/in/animesh-sharma-125701127"),
    ("Archies Divekar", "ICICI Foundation",
     "Corporate Social Responsibility Manager",
     "Mumbai, India", "linkedin.com/in/archies-divekar-a420bb230"),
    ("Sharukh R. Taraporewala", "Axis Bank",
     "CSR, Sustainability & ESG",
     "Mumbai, India", "linkedin.com/in/sharukh-r-taraporewala-621a7428"),
    ("Shubhanjali Roye", "Axis Bank Foundation",
     "CSR",
     "Mumbai, India", "linkedin.com/in/shubhanjali-roye-27a2819b"),
    ("Adil Ghadiali", "HSBC India",
     "Corporate Sustainability Lead",
     "Mumbai, India", "linkedin.com/in/adilghadiali"),
    ("Gopal Krishnan", "HSBC India",
     "CSR",
     "Hyderabad, India", "linkedin.com/in/gopal-krishnan-3b2b7219"),
    ("Aparna Achuthan", "JPMorgan Chase India",
     "Senior Associate, Global Philanthropy",
     "Mumbai, India", "linkedin.com/in/aparnaachuthan"),
    ("Aashna Roy", "Mastercard India",
     "CSR & Stakeholder Relations",
     "Mumbai, India", "linkedin.com/in/aashna-roy-a0692021"),

    # ── Professional Services (India) ─────────────────────────────────
    ("Parul Goel", "Deloitte India",
     "Corporate Responsibility",
     "New Delhi, India", "linkedin.com/in/parulgoel22"),
    ("Suruchi Pawar", "Deloitte India",
     "Social Impact and CSR Professional",
     "Mumbai, India", "linkedin.com/in/suruchi-pawar-402837b9"),
    ("Parul Mukhtyar", "Deloitte Shared Services India",
     "CSR Advisory",
     "Hyderabad, India", "linkedin.com/in/parul-mukhtyar-67a717a8"),
    ("Rumi Mallick Mitra", "EY GDS India",
     "Global Head, Director – CSR and Sustainability",
     "Bengaluru, India", "linkedin.com/in/rumi-mallick-mitra-2ab6835"),
    ("Soma Roy", "Ernst & Young India",
     "Corporate Social Responsibility",
     "Bengaluru, India", "linkedin.com/in/soma-roy-17a6b136"),
    ("Farheen S", "KPMG India",
     "Corporate Social Responsibility",
     "Bengaluru, India", "linkedin.com/in/farheen94"),

    # ── Manufacturing / Industrial (India) ────────────────────────────
    ("Prashanth Shetty", "Schneider Electric India",
     "Deputy General Manager – CSR, India & South Asia",
     "Bengaluru, India", "linkedin.com/in/prashanth-shetty-a94a5878"),
    ("Sikha V S", "Schneider Electric India",
     "CSR Manager",
     "Bengaluru, India", "linkedin.com/in/sikha-v-s-41ab4868"),
    ("Ameet Lele", "Cummins India",
     "CSR",
     "Pune, India", "linkedin.com/in/ameet-lele-2a81668"),
    ("Praveen Parihar", "Bosch India Foundation",
     "CSR",
     "Bengaluru, India", "linkedin.com/in/praveen-parihar-a75301a5"),
    ("Hari Prasad S", "3M India",
     "CSR / Social Responsibility",
     "Bengaluru, India", "linkedin.com/in/hari-prasad-s-0b7a21118"),

    # ── FMCG / Consumer (India) ────────────────────────────────────────
    ("Shveta Arora", "Nestlé India",
     "Head – CSR and Societal Initiatives",
     "Gurugram, India", "linkedin.com/in/shveta-arora-25533610"),
    ("Kanika Pal", "Hindustan Unilever (HUL)",
     "Lead – CSR / Prabhat Community Development",
     "Mumbai, India", "linkedin.com/in/kanika-pal-4130836"),
    ("Mansi Goyal", "Diageo India",
     "Senior Manager, CSR & Sustainability",
     "Bengaluru, India", "linkedin.com/in/mansi-goyal-78762436"),
    ("Rajesh Ayapilla", "Coca-Cola India (INSWA)",
     "Senior Director – CSR and Sustainability",
     "New Delhi, India", "linkedin.com/in/rajesh-ayapilla-84156593"),
    ("Rajeev Srivastava", "Procter & Gamble India",
     "Director – Sustainability",
     "Mumbai, India", "linkedin.com/in/rajeev-srivastava-b457795"),
    ("Srinivas Vemuri", "PepsiCo India",
     "CSR Core Team Lead",
     "Hyderabad, India", "linkedin.com/in/srinivas-vemuri"),

    # ── Tech Mahindra ─────────────────────────────────────────────────
    ("Chetan Kapoor", "Tech Mahindra Foundation",
     "Chief Executive Officer",
     "New Delhi, India", "linkedin.com/company/tech-mahindra-foundation"),
]

# ════════════════════════════════════════════════════════════════════
# DATA — GLOBAL
# ════════════════════════════════════════════════════════════════════

GLOBAL = [
    # ── Already in user's sheet ──────────────────────────────────────
    ("Tony Dixon", "PepsiCo",
     "PepsiCo Foundation Director, Global Employee Giving & Community Engagement",
     "Plano, Texas, USA", "linkedin.com/in/tonydixon-tdf24"),
    ("Sonia Naeemi", "BMO",
     "Senior Ambassadors",
     "Toronto, Ontario, Canada", "linkedin.com/in/sonianaeemi"),

    # ── Microsoft ─────────────────────────────────────────────────────
    ("Steve Lippman", "Microsoft",
     "VP – Corporate Responsibility",
     "Redmond, WA, USA", "linkedin.com/in/stevelippman"),
    ("Jeana Jorgensen", "Microsoft",
     "Employee Giving Program",
     "Redmond, WA, USA", "linkedin.com/in/jeanajorgensen"),
    ("Kim Manis", "Microsoft",
     "Employee Giving",
     "Redmond, WA, USA", "linkedin.com/in/kimmanis"),

    # ── Salesforce ────────────────────────────────────────────────────
    ("Erinn Corbett-Wright", "Salesforce",
     "Senior Manager, Philanthropy",
     "Philadelphia, PA, USA", "linkedin.com/in/erinn-corbett-wright"),
    ("Jeffrey Huang", "Salesforce",
     "Employee Engagement / Philanthropy",
     "San Francisco, CA, USA", "linkedin.com/in/thejeffreyhuang"),
    ("Sunya Norman", "Salesforce",
     "ESG, Sustainability & Social Impact",
     "San Francisco, CA, USA", "linkedin.com/in/sunya"),

    # ── Cisco ─────────────────────────────────────────────────────────
    ("Kyle Thornton", "Cisco Foundation",
     "Global Sector Lead; Social Impact & Community Investment",
     "San Jose, CA, USA", "linkedin.com/in/kylethornton"),
    ("Lynne Elliott", "Cisco Systems Foundation",
     "Grants Administrator – Community Investment, Corporate Philanthropy",
     "San Jose, CA, USA", "linkedin.com/in/lynne-elliott-0b8b505"),
    ("Peter Tavernise", "Cisco",
     "Corporate Affairs / Social Impact",
     "San Jose, CA, USA", "linkedin.com/in/petertavernise"),

    # ── Amazon ────────────────────────────────────────────────────────
    ("Joanna Sylwester", "Amazon",
     "Head of Global Engagement – Social Responsibility",
     "Seattle, WA, USA", "linkedin.com/in/joanna-sylwester-37566837"),

    # ── JPMorgan Chase ────────────────────────────────────────────────
    ("Roma Kaundal", "JPMorgan Chase",
     "Managing Director, Global Philanthropy",
     "New York, NY, USA", "linkedin.com/in/romakaundal"),

    # ── Bank of America ───────────────────────────────────────────────
    ("Kerry Sullivan", "Bank of America Charitable Foundation",
     "President",
     "Charlotte, NC, USA", "linkedin.com/in/kerry-sullivan-766438165"),
    ("Ximena Delgado", "Bank of America",
     "Corporate Social Responsibility",
     "Charlotte, NC, USA", "linkedin.com/in/ximena-delgado-3934363"),
    ("Katy Atwell", "Bank of America",
     "Business Support Manager, Corporate Social Responsibility (VP level)",
     "Charlotte, NC, USA", "linkedin.com/in/katy-atwell-8b8044b"),

    # ── Wells Fargo ───────────────────────────────────────────────────
    ("Kristina Christy", "Wells Fargo",
     "Community Affairs and Corporate Social Responsibility Leader",
     "Charlotte, NC, USA", "linkedin.com/in/kristinachristy"),

    # ── Goldman Sachs ─────────────────────────────────────────────────
    ("Greg Shell", "Goldman Sachs",
     "Social Impact / Impact Investing",
     "New York, NY, USA", "linkedin.com/in/greg-shell-313164166"),

    # ── PepsiCo Global ────────────────────────────────────────────────
    ("Meredith Shull", "PepsiCo",
     "Social Impact & Strategy / Corporate Social Responsibility",
     "Purchase, NY, USA", "linkedin.com/in/meredith-shull-51649735"),

    # ── Unilever ──────────────────────────────────────────────────────
    ("Rebecca Marmot", "Unilever",
     "Chief Advocacy Officer / Sustainability",
     "London, UK", "linkedin.com/in/rebecca-marmot-742a31130"),
    ("Eric Ostern", "Unilever",
     "Global Sustainability and Corporate Affairs",
     "London, UK", "linkedin.com/in/ericostern"),

    # ── Johnson & Johnson ─────────────────────────────────────────────
    ("Jenn Attridge Schacht", "Johnson & Johnson",
     "Community Impact Senior Manager",
     "New Brunswick, NJ, USA", "linkedin.com/in/jennschacht"),

    # ── Nike ──────────────────────────────────────────────────────────
    ("Miki Morimoto", "Nike",
     "Senior Director, Asia Pacific – Social & Community Impact",
     "Portland, OR, USA", "linkedin.com/in/mikimorimoto"),

    # ── Pfizer ────────────────────────────────────────────────────────
    ("Caroline Roan", "Pfizer",
     "VP, Corporate Affairs & Global Health / Social Impact Lead",
     "New York, NY, USA", "linkedin.com/in/carolineroan"),

    # ── Merck ─────────────────────────────────────────────────────────
    ("Brian K. Johnson", "Merck",
     "Executive Director & Head of External Affairs",
     "Rahway, NJ, USA", "linkedin.com/in/brian-k-johnson-20568b4a"),

    # ── Deloitte Global ───────────────────────────────────────────────
    ("Shubha Srinivasan", "Deloitte",
     "Executive Director, Social Impact",
     "New York, NY, USA", "linkedin.com/in/shubha-srinivasan-3462a48a"),
    ("Claire Burton", "Deloitte UK",
     "Leading Social Impact",
     "London, UK", "linkedin.com/in/claire-burton-b6a62154"),

    # ── KPMG ──────────────────────────────────────────────────────────
    ("Gillian Hinde", "KPMG (Global)",
     "Global Corporate Responsibility Leader",
     "London, UK", "linkedin.com/in/gillian-hinde-895a6024"),
    ("Anya T.", "KPMG UK",
     "Strategic Philanthropy & Pro Bono Lead",
     "London, UK", "linkedin.com/in/anya-t-10834239"),
    ("Roisin Murphy", "KPMG Ireland",
     "Corporate Responsibility",
     "Dublin, Ireland", "linkedin.com/in/murphyroisin"),

    # ── SAP Global ────────────────────────────────────────────────────
    ("Alexandra van der Ploeg", "SAP",
     "Head of Corporate Social Responsibility (Global)",
     "Walldorf, Germany", "linkedin.com/in/alexandra-van-der-ploeg-276558"),
    ("Kate Morgan", "SAP Americas",
     "Head of Corporate Social Responsibility, Americas",
     "Newtown Square, PA, USA", "linkedin.com/in/katemmorgan"),

    # ── Accenture Global ──────────────────────────────────────────────
    ("Deborah Swartz", "Accenture Canada",
     "Corporate Social Responsibility Lead",
     "Toronto, Canada", "linkedin.com/in/deborah-swartz-a412557"),

    # ── Texas Instruments ─────────────────────────────────────────────
    ("Lindsey Richmond", "Texas Instruments",
     "Director, Environmental Sustainability",
     "Dallas, TX, USA", "linkedin.com/in/lindsey-richmond-mba"),

    # ── Cummins Global ────────────────────────────────────────────────
    ("Jim Schacht", "Cummins Inc.",
     "Vice President, Corporate Responsibility",
     "Columbus, IN, USA", "linkedin.com/in/jim-schacht"),

    # ── IBM Global ────────────────────────────────────────────────────
    ("Shaun Wilson", "IBM",
     "Corporate Social Responsibility Manager",
     "Armonk, NY, USA", "linkedin.com/in/shaunkwilson1"),
    ("Judy M. Chin", "IBM",
     "Corporate Social Responsibility",
     "Armonk, NY, USA", "linkedin.com/in/judymchin"),
]

# ════════════════════════════════════════════════════════════════════
# WRITE ROWS
# ════════════════════════════════════════════════════════════════════
ROW = 3
ALT_IND = [LT_GREEN, HI_GREEN]   # alternating for India
ALT_GLB = [LT_BLUE, "EAF0FB"]    # alternating for global

# ── INDIA section ─────────────────────────────────────────────────
ROW = section(ws, ROW, f"  ★ INDIA (Priority) — {len(INDIA)} contacts", FOREST)
for i, (name, org, role, loc, link) in enumerate(INDIA, 1):
    fill = ALT_IND[i % 2]
    vals = [name, org, role, loc, link]
    for ci, v in enumerate(vals, 1):
        c = ws.cell(ROW, ci, v)
        c.alignment = wa()
        c.border = brd
        c.font = fn(DARK_TEXT, b=(ci == 1), s=10)
        c.fill = hf(fill)
        if ci == 5:    # LinkedIn — make it blue + underline
            c.font = Font(name="Calibri", color="1155CC", size=10, underline="single")
    ws.row_dimensions[ROW].height = 30
    ROW += 1

# ── GLOBAL section ────────────────────────────────────────────────
ROW = section(ws, ROW, f"  🌐 GLOBAL (Rest of World) — {len(GLOBAL)} contacts", MID_BLUE)
for i, (name, org, role, loc, link) in enumerate(GLOBAL, 1):
    fill = ALT_GLB[i % 2]
    vals = [name, org, role, loc, link]
    for ci, v in enumerate(vals, 1):
        c = ws.cell(ROW, ci, v)
        c.alignment = wa()
        c.border = brd
        c.font = fn(DARK_TEXT, b=(ci == 1), s=10)
        c.fill = hf(fill)
        if ci == 5:
            c.font = Font(name="Calibri", color="1155CC", size=10, underline="single")
    ws.row_dimensions[ROW].height = 30
    ROW += 1

# ── note row ──────────────────────────────────────────────────────
ROW += 1
ws.merge_cells(f"A{ROW}:E{ROW}")
note = ws[f"A{ROW}"]
note.value = (
    "Note: All LinkedIn URLs sourced from public search results. "
    f"Total: {len(INDIA)} India + {len(GLOBAL)} Global = {len(INDIA)+len(GLOBAL)} individuals. "
    "India contacts include 3 originally in your sheet (Ashwin Sesodia, Tony Dixon moved to Global, Sonia Naeemi moved to Global)."
)
note.font = fn("888888", b=False, s=9)
note.fill = hf(LT_GREY)
note.alignment = wa("left", "center")
ws.row_dimensions[ROW].height = 22

# ── column widths ─────────────────────────────────────────────────
cw = [28, 32, 48, 24, 44]
for i, w in enumerate(cw, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ── row 1 height ──────────────────────────────────────────────────
ws.row_dimensions[1].height = 32

out = "/home/user/jcode/Waste_Warriors_100_Individuals.xlsx"
wb.save(out)
total = len(INDIA) + len(GLOBAL)
print(f"Saved: {out}")
print(f"India: {len(INDIA)}  |  Global: {len(GLOBAL)}  |  Total: {total}")
