const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';

const GREEN  = '1B5E20';
const CREAM  = 'F5EDD8';
const PANEL  = 'EDE0C4';
const YELLOW = 'F5C518';
const WHITE  = 'FFFFFF';
const DARK   = '2E2E2E';

const W = 13.33;
const H = 7.5;

function addBackground(slide) {
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:CREAM}, line:{color:CREAM} });
}
function addTitle(slide, title) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x:2.0, y:0.22, w:9.33, h:0.72,
    fill:{color:CREAM}, line:{color:GREEN, width:2}, rectRadius:0.15,
  });
  slide.addText(title, {
    x:2.0, y:0.22, w:9.33, h:0.72,
    fontSize:20, bold:true, color:GREEN, fontFace:'Calibri', align:'center', valign:'middle',
  });
}
function addFooter(slide, text) {
  slide.addShape(pptx.ShapeType.rect, { x:0, y:H-0.72, w:W, h:0.72, fill:{color:GREEN}, line:{color:GREEN} });
  slide.addText(text, {
    x:0.3, y:H-0.72, w:W-0.6, h:0.72,
    fontSize:13, color:WHITE, fontFace:'Calibri', align:'center', valign:'middle',
  });
}
function addBeigePanel(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, fill:{color:PANEL}, line:{color:GREEN, width:1.5}, rectRadius:0.1,
  });
}
function addYellowPanel(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill:{color:YELLOW}, line:{color:YELLOW} });
}
function addGreenPanel(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill:{color:GREEN}, line:{color:GREEN} });
}

const FOOTER = 'Airbnb Community Fund  |  Airbnb × Waste Warriors Society  |  FY 2025–26';

// ── SLIDE 5 — About the Project ───────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'About the Project');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Community-owned SWM systems in the Indian Himalayan Region — powered by the Airbnb Community Fund', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Airbnb Community Fund Partnership\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Waste Warriors operates across 12+ locations in Himachal Pradesh and Uttarakhand — high-footfall, eco-sensitive geographies where waste systems are still being built from the ground up.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Thanks to your support, this year\'s partnership went beyond just field operations. The Airbnb Community Fund enabled us to strengthen the inner architecture of the organisation — the teams, tools, systems, and learning that make everything else possible.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'You enabled this. We are the hands on the ground.', options:{fontSize:12.5, color:DARK, italic:true} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Your Partnership\nEnabled Us To\n\n', options:{bold:true, fontSize:14, color:YELLOW} },
    { text:'• Conduct Waste Management\n   Awareness at scale\n\n', options:{fontSize:12, color:WHITE} },
    { text:'• Co-fund ongoing project\n   field operations\n\n', options:{fontSize:12, color:WHITE} },
    { text:'• Introduce Leadership &\n   Team Trainings\n\n', options:{fontSize:12, color:WHITE} },
    { text:'• Drive Digital Transformation\n   & M&E systems\n\n', options:{fontSize:12, color:WHITE} },
    { text:'• Build organisational resilience\n   through enabling functions', options:{fontSize:12, color:WHITE} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 6 — How Your Support Was Used ───────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'How Your Support Was Used');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Four fund heads — all pointed at the same goal: a stronger, more resilient organisation', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  const heads = [
    {
      label: '01  Project Team Cost',
      desc: 'Supporting salaries of the Digital Transformation & M&E function, and co-funding project team costs at Dharamshala Urban, Dharamshala Rural, and Bir — keeping people on the ground.',
      color: 'beige',
    },
    {
      label: '02  Awareness',
      desc: 'Backing community outreach, IEC activities at Project Flyer\'s Paradise (Bir), and the landmark World Tourism Day event with Airbnb and the Uttarakhand Tourism Development Board.',
      color: 'green',
    },
    {
      label: '03  Data & Storytelling',
      desc: 'Building the Impact Dashboard on Looker — moving from unwieldy spreadsheets to a real-time, standardised system for tracking and communicating organisational impact.',
      color: 'beige',
    },
    {
      label: '04  Learning & Development',
      desc: 'Cross-project visits that create learning exchange between field teams, M&E trainings, and team capacity building — strengthening the people who hold the work together.',
      color: 'green',
    },
  ];

  const stripH = 1.02;
  const stripY0 = 2.03;
  const gap = 0.15;
  heads.forEach((h, i) => {
    const y = stripY0 + i * (stripH + gap);
    if (h.color === 'beige') {
      addBeigePanel(slide, 0.5, y, 12.33, stripH);
      slide.addText([
        { text:h.label + '   ', options:{bold:true, fontSize:14, color:GREEN} },
        { text:h.desc, options:{fontSize:12, color:DARK} },
      ], { x:0.78, y:y, w:11.83, h:stripH, fontFace:'Calibri', valign:'middle' });
    } else {
      addGreenPanel(slide, 0.5, y, 12.33, stripH);
      slide.addText([
        { text:h.label + '   ', options:{bold:true, fontSize:14, color:YELLOW} },
        { text:h.desc, options:{fontSize:12, color:WHITE} },
      ], { x:0.78, y:y, w:11.83, h:stripH, fontFace:'Calibri', valign:'middle' });
    }
  });

  addFooter(slide, FOOTER);
}

// ── SLIDE 7 — Project Team Cost: DT / M&E ─────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Project Team Cost — Digital Transformation & M&E');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('You cannot manage what you cannot measure — your support helped us build the systems to do both', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'Why This Investment Matters\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Waste Warriors operates across dozens of locations, field teams, and data streams. Without dedicated Digital Transformation and Monitoring & Evaluation staff, that complexity turns into noise — and impact becomes invisible.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'With your backing, we funded the DT/M&E function that keeps the organisation accountable, measurable, and legible — both internally and to our partners. This team manages our technology stack, monitors project performance, and ensures that every rupee and every kilogram of waste is tracked and reported.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Your support also co-funded part-salaries for project teams at Dharamshala Urban, Dharamshala Rural, and Bir — keeping people in the field when it mattered.', options:{fontSize:12.5, color:DARK} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Our Technology Stack\n\n', options:{bold:true, fontSize:14, color:YELLOW} },
    { text:'Google Workspace\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'collaboration, docs & team ops\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Wati (WhatsApp)\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'community & stakeholder comms\n\n', options:{fontSize:11, color:CREAM} },
    { text:'ODK\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'field data collection\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Zoho Creator\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'custom apps & workflows\n\n', options:{fontSize:11, color:CREAM} },
    { text:'HRone\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'HR & people management\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Canva Pro\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'communications & design', options:{fontSize:11, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 8 — Data & Storytelling: Impact Dashboard ───────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Data & Storytelling — The Impact Dashboard');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('From cumbersome spreadsheets to a living, real-time view of organisational impact', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Problem We Solved\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'After 11+ years of operations, Waste Warriors had accumulated vast amounts of data — but it lived in siloed, cumbersome spreadsheets that were difficult to collate and nearly impossible to make sense of at an organisational level.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Made possible by your partnership, our M&E team built an ', options:{fontSize:12.5, color:DARK} },
    { text:'Impact Dashboard on Looker', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' — a standardised system that collects project data quarterly, gets it validated by Senior Managers, and updates the dashboard in near real-time.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'This tool serves two purposes: a ', options:{fontSize:12.5, color:DARK} },
    { text:'project-level lens', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' for day-to-day performance tracking, and a ', options:{fontSize:12.5, color:DARK} },
    { text:'bird\'s eye view', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' for leadership to see cumulative impact across all locations — enabling data-driven decisions and stronger reporting to partners.', options:{fontSize:12.5, color:DARK} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'What It Unlocks\n\n', options:{bold:true, fontSize:14, color:YELLOW} },
    { text:'Real-time visibility\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'across all projects & locations\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Standardised recording\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'consistent across field teams\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Senior Manager validation\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'quarterly data accountability loop\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Project + org-level view\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'dual lens for operations & impact\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Partner reporting\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'stronger evidence for funders', options:{fontSize:11, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 9 — Awareness: World Tourism Day & Good Guest Guidebook ──────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Awareness — World Tourism Day & The Good Guest Guidebook');

  // Dark green BG strip with Atithi Devo Bhava
  addGreenPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('"Atithi Devo Bhava" — Reimagining Responsible Tourism in the Himalayan Region', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13.5, bold:true, color:YELLOW, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Event\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'On World Tourism Day in Dehradun, powered by your support, Waste Warriors and the ', options:{fontSize:12.5, color:DARK} },
    { text:'Uttarakhand Tourism Development Board (UTDB)', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' co-hosted a landmark event to reimagine the future of responsible tourism in the Indian Himalayas.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Policymakers, industry leaders, influencers, and community voices came together for vibrant panel discussions, folk performances, and interactive activities — all centred on what it means to be a ', options:{fontSize:12.5, color:DARK} },
    { text:'good guest', options:{bold:true, italic:true, fontSize:12.5, color:GREEN} },
    { text:' in a fragile mountain ecosystem.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'The highlight of the event was the launch of the ', options:{fontSize:12.5, color:DARK} },
    { text:'Good Guest Guidebook', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' — a practical guide to sustainable travel for visitors to Uttarakhand and beyond.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Read the Guidebook: heyzine.com/flip-book/f293cdab8f.html', options:{fontSize:11.5, color:GREEN, italic:true, bold:true} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Event Highlights\n\n', options:{bold:true, fontSize:14, color:YELLOW} },
    { text:'Co-hosted with UTDB\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'Uttarakhand Tourism Development Board\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Panel discussions\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'tourism, sustainability, community\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Folk performances\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'celebrating Uttarakhand culture\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Good Guest Guidebook\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'practical guide to responsible travel\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Citizen-led activities\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'lead the change, not just witness it', options:{fontSize:11, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 10 — Awareness: Project Flyer's Paradise (Bir) ──────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Awareness — Project Flyer\'s Paradise, Bir');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Bir-Billing, the Paragliding Capital of India — where adventure tourism meets the need for systemic waste management', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:12.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Work on the Ground\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'With your backing, our team continued strengthening solid waste management across ', options:{fontSize:12.5, color:DARK} },
    { text:'four gram panchayats — Bir, Chaugan, Keori, and Gunehar', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' in the Baijnath Block of Kangra. Despite statewide plastic bans, peak tourist seasons bring a surge in single-use waste, and our system works year-round to manage it.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Awareness drove behaviour. 16 campaigns and workshops engaged local residents, Mahila Mandals, youth groups, and school children on waste management and menstrual hygiene. The ', options:{fontSize:12.5, color:DARK} },
    { text:'Gau Seva campaign', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' extended our reach into wet waste — diverting edible waste from 10 local establishments to the Gaushala, closing a loop that had previously gone unaddressed.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'The MRF at Keori (Him Ira Swachhta Cafe) reached operational readiness — a major infrastructure milestone.', options:{fontSize:12.5, color:DARK} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Project Highlights\n\n', options:{bold:true, fontSize:14, color:YELLOW} },
    { text:'124.37 MT\n', options:{bold:true, fontSize:13, color:WHITE} },
    { text:'dry waste diverted from landfills\n\n', options:{fontSize:11, color:CREAM} },
    { text:'19 cleanup drives\n', options:{bold:true, fontSize:13, color:WHITE} },
    { text:'5.3 MT collected from hotspots\n\n', options:{fontSize:11, color:CREAM} },
    { text:'1,727 people\n', options:{bold:true, fontSize:13, color:WHITE} },
    { text:'engaged via IEC activities\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Gau Seva Campaign\n', options:{bold:true, fontSize:13, color:WHITE} },
    { text:'1,739 kg wet waste → Gaushala\n\n', options:{fontSize:11, color:CREAM} },
    { text:'MRF construction\n', options:{bold:true, fontSize:13, color:WHITE} },
    { text:'Him Ira Swachhta Cafe, Keori', options:{fontSize:11, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 11 — Learning & Development ─────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Learning & Development');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Building people who build systems — cross-project learning and team capacity at the heart of this fund', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Three horizontal strips
  const strips = [
    {
      title: 'Cross-Project Visits & Learning Exchange',
      body: 'With your support, team members from different project locations visited each other\'s sites — not as supervisors, but as learners. These cross-visits create peer learning, build shared understanding of challenges, and help our teams adopt what works elsewhere rather than reinventing the wheel. A field team in Bir learning from Dehradun; a coordinator in Dharamshala observing Govind Wildlife Sanctuary — this is how institutional knowledge travels.',
      color: 'beige',
    },
    {
      title: 'M&E Team Training',
      body: 'Powered by your partnership, our Monitoring & Evaluation team received structured training on data collection, validation, and storytelling. This directly fed into the Impact Dashboard initiative — equipping the team to not just record data, but to make sense of it, challenge assumptions, and communicate impact in ways that are meaningful to both internal teams and external partners.',
      color: 'green',
    },
    {
      title: 'Organisational Capacity Building',
      body: 'Learning and development at Waste Warriors is not just individual — it is systemic. Team trainings, tool onboarding (ODK, Zoho, HRone), and structured reflection sessions ensure that as the organisation grows, its people grow with it. Thanks to your backing, this layer of investment in people and process remained active and well-resourced through the year.',
      color: 'beige',
    },
  ];

  const stripH = 1.25;
  const y0 = 2.05;
  const gap = 0.14;
  strips.forEach((s, i) => {
    const y = y0 + i * (stripH + gap);
    if (s.color === 'beige') {
      addBeigePanel(slide, 0.5, y, 12.33, stripH);
      slide.addText([
        { text:s.title + '   ', options:{bold:true, fontSize:13, color:GREEN} },
        { text:s.body, options:{fontSize:11.5, color:DARK} },
      ], { x:0.78, y:y, w:11.83, h:stripH, fontFace:'Calibri', valign:'middle' });
    } else {
      addGreenPanel(slide, 0.5, y, 12.33, stripH);
      slide.addText([
        { text:s.title + '   ', options:{bold:true, fontSize:13, color:YELLOW} },
        { text:s.body, options:{fontSize:11.5, color:WHITE} },
      ], { x:0.78, y:y, w:11.83, h:stripH, fontFace:'Calibri', valign:'middle' });
    }
  });

  addFooter(slide, FOOTER);
}

// ── write ─────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: 'Airbnb_FundUtilisation_Slides5to11.pptx' })
  .then(name => console.log('Wrote', name));
