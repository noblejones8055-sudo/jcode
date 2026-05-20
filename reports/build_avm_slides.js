const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"

const GREEN  = '1B5E20';
const CREAM  = 'F5EDD8';
const PANEL  = 'EDE0C4';
const YELLOW = 'F5C518';
const WHITE  = 'FFFFFF';
const DARK   = '2E2E2E';

const W = 13.33;
const H = 7.5;

// ── helpers ──────────────────────────────────────────────────────────────────
function addBackground(slide) {
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:CREAM}, line:{color:CREAM} });
}
function addTitle(slide, title) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x:3.5, y:0.22, w:6.33, h:0.72,
    fill:{color:CREAM}, line:{color:GREEN, width:2}, rectRadius:0.15,
  });
  slide.addText(title, {
    x:3.5, y:0.22, w:6.33, h:0.72,
    fontSize:22, bold:true, color:GREEN, fontFace:'Calibri', align:'center', valign:'middle',
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

// ── SLIDE 6 — About the Project ──────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'About the Project');

  // Yellow banner
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.75);
  slide.addText('Your partnership powered work across four locations — from remote mountain outposts to busy urban MRFs', {
    x:0.65, y:1.15, w:12.03, h:0.75,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // 2x2 location grid (left side, 4 cards)
  const locs = [
    {
      col:0, row:0,
      name:'Govind Wildlife Sanctuary, Uttarakhand',
      desc:'A remote, resource-scarce region inside a protected wildlife reserve — where every basic amenity makes a tangible difference to our team on the ground.',
    },
    {
      col:1, row:0,
      name:'Harrawala MRF, Dehradun',
      desc:'Our flagship Material Recovery Facility — where dry and wet waste is collected, sorted, and processed every single day, keeping it out of landfills.',
    },
    {
      col:0, row:1,
      name:'Nagrota, Himachal Pradesh',
      desc:'A new project site where the work began with understanding the ground reality — a baseline study to design an evidence-backed waste management programme.',
    },
    {
      col:1, row:1,
      name:'Shimla, Himachal Pradesh',
      desc:'Field operations spanning multiple villages, panchayats, and partner sites — where our team travels regularly to keep the work moving.',
    },
  ];

  const cardW = 4.9, cardH = 2.2, sx = 0.5, sy = 2.1, gap = 0.15;
  locs.forEach(l => {
    const x = sx + l.col * (cardW + gap);
    const y = sy + l.row * (cardH + gap);
    addBeigePanel(slide, x, y, cardW, cardH);
    slide.addText(l.name, { x:x+0.2, y:y+0.12, w:cardW-0.4, h:0.4, fontSize:13, bold:true, color:GREEN, fontFace:'Calibri' });
    slide.addText(l.desc, { x:x+0.2, y:y+0.58, w:cardW-0.4, h:cardH-0.7, fontSize:11.5, color:DARK, fontFace:'Calibri', valign:'top' });
  });

  // Right green panel — How Support Was Used
  addGreenPanel(slide, 10.68, 2.1, 2.15, 4.57);
  slide.addText([
    { text:'How Your\nSupport Was Used\n\n', options:{bold:true, fontSize:13, color:YELLOW, breakLine:false} },
    { text:'Programmatic\nCapital Expenses\n', options:{bold:true, fontSize:12, color:WHITE, breakLine:false} },
    { text:'Equipment & infrastructure\nto enable operations\n\n', options:{fontSize:11, color:WHITE, italic:true, breakLine:false} },
    { text:'Programmatic\nOperational Expenses\n', options:{bold:true, fontSize:12, color:WHITE, breakLine:false} },
    { text:'Day-to-day costs to\ndeliver impact on the ground', options:{fontSize:11, color:WHITE, italic:true} },
  ], { x:10.83, y:2.2, w:1.85, h:4.35, fontFace:'Calibri', valign:'top' });

  addFooter(slide, 'Govind  |  Harrawala MRF  |  Nagrota  |  Shimla  |  FY 2025–26');
}

// ── SLIDE 7 — Govind: Work in the Wild ───────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Govind Wildlife Sanctuary — Work in the Wild');

  // Yellow location context strip
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Inside a protected wildlife reserve, at high altitude, across 4 remote villages — our team is building waste management systems where no systems existed before.', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left narrative panel — the work
  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'What\'s Happening on the Ground\n\n', options:{bold:true, fontSize:16, color:GREEN} },
    { text:'The Govind Wildlife Sanctuary project covers four villages — Gaichwan, Netwar, Doni, and Satta — deep inside one of Uttarakhand\'s most ecologically sensitive protected zones. These are places where waste management infrastructure has historically been non-existent.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Our teams collect waste across these villages, segregate it on-site, and transport tonnes of material out of the sanctuary — often across rough mountain roads — to our Harrawala MRF in Dehradun for scientific processing.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'This year, 30 trekking companies committed to the Trek It Back campaign, bringing trail waste down from popular routes like Kedarkantha rather than leaving it in fragile ecosystems. Community-led Waste Banks are being established, and local cleanliness committees are taking ownership of operations.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Every kilogram collected here protects a forest, a water source, or a wildlife habitat that millions depend on.', options:{fontSize:12.5, color:DARK, italic:true} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  // Right stats panel — green
  addGreenPanel(slide, 8.55, 2.0, 4.28, 3.35);
  const govindStats = [
    { num:'5',         label:'Panchayats with active\nwaste management systems' },
    { num:'10,000+',   label:'kg waste collected &\ntransported out of the sanctuary' },
    { num:'30',        label:'Trekking companies\nonboarded for trail waste' },
    { num:'4',         label:'Villages covered inside\nGovind Wildlife Sanctuary' },
  ];
  govindStats.forEach((s, i) => {
    slide.addText(s.num,   { x:8.72, y:2.08 + i*0.82, w:1.5,  h:0.45, fontSize:22, bold:true, color:YELLOW, fontFace:'Calibri', align:'center' });
    slide.addText(s.label, { x:10.1, y:2.08 + i*0.82, w:2.55, h:0.45, fontSize:10.5, color:WHITE, fontFace:'Calibri', valign:'middle' });
  });

  // AVM contribution callout strip — thin, at the bottom of right column
  addYellowPanel(slide, 8.55, 5.5, 4.28, 1.25);
  slide.addText([
    { text:'Powered by your support\n', options:{bold:true, fontSize:12, color:GREEN} },
    { text:'Laptop  ·  Water Purifier', options:{fontSize:12, color:DARK} },
    { text:'\nEssential tools that keep our remote team connected, healthy, and operational.', options:{fontSize:10.5, color:DARK, italic:true} },
  ], { x:8.68, y:5.55, w:4.0, h:1.15, fontFace:'Calibri', valign:'top' });

  addFooter(slide, 'Govind Wildlife Sanctuary  |  Gaichwan · Netwar · Doni · Satta  |  Uttarakhand');
}

// ── SLIDE 8 — Harrawala MRF: Dehradun's Waste Processing Engine ──────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Harrawala MRF — Dehradun\'s Waste Processing Engine');

  // Yellow banner
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Every day at Harrawala, waste that would end up in a landfill is given a second life — sorted, processed, and sent back into the recycling chain.', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left narrative panel
  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'What\'s Happening on the Ground\n\n', options:{bold:true, fontSize:16, color:GREEN} },
    { text:'The Harrawala Material Recovery Facility is our flagship urban operations hub in Dehradun. Every month, tonnes of dry and wet waste from Ward 97 and beyond arrive here — and leave as sorted recyclables, compost, and recoverable material.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Green Workers segregate waste by category, process organic matter through in-vessel composting, and dispatch recyclables to verified market partners. The MRF now handles up to 1,200 kg of wet waste per day and has converted multi-layered plastics — once a disposal burden — into a revenue stream.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'The facility also serves as a learning centre. This year alone, it hosted Indian Forest Service officers, Dehradun\'s Mayor, urban planning engineers, and delegations from Mizoram — all coming to see how a community-linked MRF actually works.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Ward 97 households pay monthly user fees and trust the system — averaging ₹1.63 lakh per month — a sign of how deeply this MRF is embedded in the community.', options:{fontSize:12.5, color:DARK, italic:true} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  // Right stats panel — green
  addGreenPanel(slide, 8.55, 2.0, 4.28, 3.35);
  const mrfStats = [
    { num:'519+ MT',    label:'waste diverted from\nlandfills in Q3 alone' },
    { num:'1,200 kg',   label:'wet waste processed\nper day at the MRF' },
    { num:'1,472',      label:'households served\nin Ward 97, Dehradun' },
    { num:'₹1.63L',     label:'avg monthly user fees\ncollected from the community' },
  ];
  mrfStats.forEach((s, i) => {
    slide.addText(s.num,   { x:8.72, y:2.08 + i*0.82, w:1.7,  h:0.45, fontSize:20, bold:true, color:YELLOW, fontFace:'Calibri', align:'center' });
    slide.addText(s.label, { x:10.3, y:2.08 + i*0.82, w:2.35, h:0.45, fontSize:10.5, color:WHITE, fontFace:'Calibri', valign:'middle' });
  });

  // AVM contribution callout — yellow strip bottom right
  addYellowPanel(slide, 8.55, 5.5, 4.28, 1.25);
  slide.addText([
    { text:'Powered by your support\n', options:{bold:true, fontSize:12, color:GREEN} },
    { text:'High-Pressure Water Pump  ·  Industrial Cooler  ·  Office Furniture', options:{fontSize:11, color:DARK} },
    { text:'\nInfrastructure that keeps the MRF hygienic, humane, and operational year-round.', options:{fontSize:10.5, color:DARK, italic:true} },
  ], { x:8.68, y:5.55, w:4.0, h:1.15, fontFace:'Calibri', valign:'top' });

  addFooter(slide, 'Harrawala MRF  |  Ward 97, Dehradun  |  Programmatic Capital Expenses');
}

// ── SLIDE 9 — Nagrota: Laying the Groundwork ─────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Nagrota — Laying the Groundwork');

  // Left beige narrative panel
  addBeigePanel(slide, 0.5, 1.15, 7.5, 5.6);
  slide.addText([
    { text:'Every meaningful intervention begins with understanding the ground reality.\n\n', options:{bold:true, fontSize:16, color:GREEN} },
    { text:'Nagrota, Kangra is one of our newest project sites in Himachal Pradesh. Before any collection systems, infrastructure, or community programmes could begin, we needed to know the truth of what was happening on the ground.\n\n', options:{fontSize:13, color:DARK} },
    { text:'Your support enabled us to conduct a ', options:{fontSize:13, color:DARK} },
    { text:'Baseline Survey', options:{bold:true, fontSize:13, color:GREEN} },
    { text:' — the first and most critical step in designing an effective waste management programme for the region.\n\n', options:{fontSize:13, color:DARK} },
    { text:'The baseline survey allowed our team to:\n', options:{bold:true, fontSize:13, color:GREEN} },
  ], { x:0.75, y:1.28, w:7.0, h:3.3, fontFace:'Calibri', valign:'top' });

  const nagrota_points = [
    'Map current waste generation and disposal patterns across the area',
    'Understand community awareness, attitudes, and waste behaviour',
    'Identify critical gaps in existing waste infrastructure',
    'Engage with local stakeholders, Panchayats, and authorities',
    'Build a data-backed foundation for all future interventions',
  ];
  nagrota_points.forEach((pt, i) => {
    slide.addText(`• ${pt}`, {
      x:0.85, y:4.65 + i * 0.4, w:6.9, h:0.38,
      fontSize:12.5, color:DARK, fontFace:'Calibri',
    });
  });

  // Right green "Why It Matters" panel
  addGreenPanel(slide, 8.28, 1.15, 4.55, 5.6);
  slide.addText([
    { text:'Why It Matters\n\n', options:{bold:true, fontSize:18, color:YELLOW, breakLine:false} },
    { text:'Without a baseline, every intervention is guesswork.\n\n', options:{bold:true, fontSize:14, color:WHITE, italic:true, breakLine:false} },
    { text:'Your support helped us start Nagrota the right way — with evidence, understanding, and a clear roadmap for what comes next.\n\n', options:{fontSize:13, color:WHITE, breakLine:false} },
    { text:'This is how Waste Warriors builds systems that last: grounded in data, shaped by community realities, and designed for long-term ownership.\n\n', options:{fontSize:13, color:WHITE, breakLine:false} },
    { text:'Nagrota is now part of our growing footprint in Himachal Pradesh — and this baseline is the foundation everything else will be built on.', options:{fontSize:13, color:WHITE} },
  ], { x:8.5, y:1.3, w:4.1, h:5.25, fontFace:'Calibri', valign:'top' });

  addFooter(slide, 'Programmatic Operational Expenses  |  Nagrota, Kangra, Himachal Pradesh');
}

// ── SLIDE 10 — Shimla: Keeping Field Operations Mobile ───────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Shimla — Keeping Field Operations Mobile');

  // Yellow banner
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.75);
  slide.addText('Field mobility is what turns plans into action on the ground', {
    x:0.65, y:1.15, w:12.03, h:0.75,
    fontSize:16, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left beige narrative panel
  addBeigePanel(slide, 0.5, 2.1, 7.5, 4.65);
  slide.addText([
    { text:'Field Travel & Fuel Support\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Across Shimla, our work spans multiple villages, panchayats, and partner sites — spread across steep, often difficult terrain.\n\n', options:{fontSize:13, color:DARK} },
    { text:'Reaching these locations consistently — for site visits, community engagement, government liaison, baseline surveys, and team coordination — depends entirely on reliable field mobility.\n\n', options:{fontSize:13, color:DARK} },
    { text:'Your support covered field travel and fuel costs for the Shimla team, ensuring our work never stops at the office door — it goes wherever the work needs us to be.', options:{fontSize:13, color:DARK} },
  ], { x:0.75, y:2.22, w:7.0, h:4.4, fontFace:'Calibri', valign:'top' });

  // Right yellow panel — What This Enabled
  addYellowPanel(slide, 8.28, 2.1, 4.55, 4.65);
  slide.addText('What This Enabled', {
    x:8.45, y:2.22, w:4.2, h:0.45,
    fontSize:15, bold:true, color:GREEN, fontFace:'Calibri',
  });

  const shimlaPoints = [
    'Regular site visits for the Shimla Baseline Survey covering 1,807 households across 4 Panchayats',
    'Community meetings and door-to-door awareness drives across scattered settlements',
    'Coordination with government stakeholders, Block Development Officers, and Panchayat leaders',
    'On-ground monitoring of interventions and segregation practices',
    'Installation support for the Banka Shimla structure — built from 235 kg of recovered waste, turning the MRF into a visible landmark for responsible consumption',
    'Continuity of field operations through monsoon-damaged roads and remote terrain',
  ];
  shimlaPoints.forEach((pt, i) => {
    slide.addText(`• ${pt}`, {
      x:8.42, y:2.82 + i * 0.63, w:4.22, h:0.6,
      fontSize:11.5, color:DARK, fontFace:'Calibri', valign:'top',
    });
  });

  addFooter(slide, 'Programmatic Operational Expenses  |  Shimla, Himachal Pradesh');
}

// ── SLIDE 11 — What Your Support Made Possible ───────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'What Your Support Made Possible');

  slide.addText('Across four locations and through both capital and operational support, your partnership shaped real, on-the-ground impact this year.', {
    x:0.5, y:1.1, w:12.33, h:0.8,
    fontSize:14, color:DARK, fontFace:'Calibri', align:'center', italic:true,
  });

  const wrap = [
    {
      col:0, row:0,
      bg:GREEN, numColor:YELLOW, txtColor:WHITE,
      location:'Govind Wildlife Sanctuary',
      headline:'Equipped a remote team\nwith essential amenities',
      stats:'Laptop + Water Purifier installed\nfor a team working in one of IHR\'s\nmost resource-scarce postings',
    },
    {
      col:1, row:0,
      bg:PANEL, numColor:GREEN, txtColor:DARK,
      location:'Harrawala MRF, Dehradun',
      headline:'Strengthened daily operations\nat our Material Recovery Facility',
      stats:'High-pressure pump · Industrial cooler · Office furniture\nFacility processed 519+ MT of waste in Q3 alone',
    },
    {
      col:0, row:1,
      bg:PANEL, numColor:GREEN, txtColor:DARK,
      location:'Nagrota, Himachal Pradesh',
      headline:'Built the evidence base\nfor a new intervention',
      stats:'Baseline Survey conducted\nData-backed foundation laid for\na new SWM programme in Kangra',
    },
    {
      col:1, row:1,
      bg:GREEN, numColor:YELLOW, txtColor:WHITE,
      location:'Shimla, Himachal Pradesh',
      headline:'Kept field operations moving\nacross project sites',
      stats:'Fuel & field travel covered\n1,807 households surveyed\nBanka Shimla installation enabled',
    },
  ];

  const cW = 5.9, cH = 2.25, sx = 0.5, sy = 2.1, gx = 0.63, gy = 0.25;
  wrap.forEach(c => {
    const x = sx + c.col * (cW + gx);
    const y = sy + c.row * (cH + gy);
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w:cW, h:cH, fill:{color:c.bg}, line:{color:c.bg}, rectRadius:0.1 });
    slide.addText(c.location, { x:x+0.25, y:y+0.12, w:cW-0.5, h:0.3, fontSize:11.5, bold:true, color:c.numColor, fontFace:'Calibri' });
    slide.addText(c.headline, { x:x+0.25, y:y+0.45, w:cW-0.5, h:0.55, fontSize:13.5, bold:true, color:c.txtColor, fontFace:'Calibri' });
    slide.addText(c.stats,    { x:x+0.25, y:y+1.05, w:cW-0.5, h:1.05, fontSize:11.5, color:c.txtColor, fontFace:'Calibri', valign:'top' });
  });

  addFooter(slide, 'From all of us at Waste Warriors — thank you for making this possible  |  FY 2025–26');
}

// ── SLIDE 12 — Key Challenges & Bottlenecks ──────────────────────────────────
{
  const slide = pptx.addSlide();

  // Dark green full-bleed background for visual contrast
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:GREEN}, line:{color:GREEN} });

  // Title pill — inverted: yellow fill, green border, green text
  slide.addShape(pptx.ShapeType.roundRect, {
    x:2.8, y:0.28, w:7.73, h:0.78,
    fill:{color:YELLOW}, line:{color:CREAM, width:2.5}, rectRadius:0.18,
  });
  slide.addText('Key Challenges & Bottlenecks', {
    x:2.8, y:0.28, w:7.73, h:0.78,
    fontSize:24, bold:true, color:GREEN, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Subtitle / hook
  slide.addText('The realities our teams navigate every day — and where sustained partnership matters most', {
    x:0.5, y:1.18, w:12.33, h:0.42,
    fontSize:13, italic:true, color:CREAM, fontFace:'Calibri', align:'center', valign:'middle',
  });

  const challenges = [
    {
      num:'01',
      title:'Remote Terrain & Broken Lifelines',
      body:'In Govind, poor roads, weak network coverage, and monsoon-damaged Waste Banks routinely cut our teams off — sometimes leaving frontline workers without phone signal for days at a time.',
    },
    {
      num:'02',
      title:'Retaining a Frontline Workforce',
      body:'Extreme physical demands and resource-scarce postings make it difficult to attract and retain team members. The loss of long-serving Green Workers leaves gaps that are hard to fill.',
    },
    {
      num:'03',
      title:'The Economics of Low-Value Waste',
      body:'Multi-layered plastics and reject waste historically cost more to process than they earn back. Operational margins remain thin, and recovery depends on continuous market-building.',
    },
    {
      num:'04',
      title:'Reaching Scattered Mountain Communities',
      body:'In Shimla, our teams travel entirely on foot across steep terrain to reach households spread across remote panchayats — making consistent field coverage one of our biggest operational lifts.',
    },
    {
      num:'05',
      title:'Long-Term Funding Continuity',
      body:'Building durable waste systems requires multi-year horizons. Year-on-year funding uncertainty makes it harder to plan, retain teams, and scale interventions that are already showing results.',
    },
  ];

  // 5 horizontal strips with yellow number accent on left + cream card on right
  const stripX = 0.5;
  const stripW = 12.33;
  const stripH = 1.0;
  const stripStart = 1.7;
  const stripGap = 0.05;

  challenges.forEach((c, i) => {
    const y = stripStart + i * (stripH + stripGap);

    // Cream/beige card body
    slide.addShape(pptx.ShapeType.roundRect, {
      x: stripX, y, w: stripW, h: stripH,
      fill:{color:CREAM}, line:{color:CREAM}, rectRadius:0.06,
    });

    // Yellow vertical accent block on the left holding the number
    slide.addShape(pptx.ShapeType.rect, {
      x: stripX, y, w: 1.4, h: stripH,
      fill:{color:YELLOW}, line:{color:YELLOW},
    });

    // Giant number in green on yellow
    slide.addText(c.num, {
      x: stripX, y, w: 1.4, h: stripH,
      fontSize: 44, bold: true, color: GREEN, fontFace:'Calibri',
      align:'center', valign:'middle',
    });

    // Title (bold green) + body (dark) on the cream card
    slide.addText(c.title, {
      x: stripX + 1.6, y: y + 0.08, w: stripW - 1.8, h: 0.36,
      fontSize: 15, bold: true, color: GREEN, fontFace:'Calibri', valign:'middle',
    });
    slide.addText(c.body, {
      x: stripX + 1.6, y: y + 0.44, w: stripW - 1.8, h: stripH - 0.5,
      fontSize: 12, color: DARK, fontFace:'Calibri', valign:'top',
    });
  });

  // Footer — yellow strip on green
  slide.addShape(pptx.ShapeType.rect, { x:0, y:H-0.55, w:W, h:0.55, fill:{color:YELLOW}, line:{color:YELLOW} });
  slide.addText('Your continued partnership is what helps us turn these challenges into systems that last', {
    x:0.3, y:H-0.55, w:W-0.6, h:0.55,
    fontSize:13, bold:true, italic:true, color:GREEN, fontFace:'Calibri', align:'center', valign:'middle',
  });
}

pptx.writeFile({ fileName: 'AVM_FundUtilisation_Slides6to12.pptx' })
  .then(() => console.log('DONE — AVM_FundUtilisation_Slides6to12.pptx'))
  .catch(e => { console.error(e); process.exit(1); });
