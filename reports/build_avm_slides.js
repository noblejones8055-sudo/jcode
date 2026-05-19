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

// ── SLIDE 7 — Govind: Bringing the Basics to a Resource-Scarce Region ────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Govind — Bringing the Basics to a Resource-Scarce Region');

  // Yellow context strip
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.82);
  slide.addText('Govind is one of our most remote project locations — limited amenities, difficult terrain, and poor connectivity make retention and daily operations a real challenge. With your support, we helped close that gap.', {
    x:0.65, y:1.15, w:12.03, h:0.82,
    fontSize:13, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left card — Laptop
  addBeigePanel(slide, 0.5, 2.18, 6.1, 4.52);
  slide.addText('Laptop', { x:0.75, y:2.3, w:5.6, h:0.45, fontSize:18, bold:true, color:GREEN, fontFace:'Calibri' });
  slide.addText([
    { text:'In a region where connectivity and equipment are hard to come by, a laptop has become the backbone of daily work for the Govind team.\n\n', options:{fontSize:13, color:DARK} },
    { text:'It is used for:\n', options:{bold:true, fontSize:13, color:GREEN} },
  ], { x:0.75, y:2.82, w:5.6, h:1.2, fontFace:'Calibri', valign:'top' });

  const laptopPoints = [
    'Data analysis and field reporting',
    'Monitoring & evaluation tracking',
    'Field documentation and record-keeping',
    'Virtual coordination with the regional team',
    'Communications and project planning',
  ];
  laptopPoints.forEach((pt, i) => {
    slide.addText(`• ${pt}`, {
      x:0.85, y:4.08 + i * 0.42, w:5.4, h:0.4,
      fontSize:13, color:DARK, fontFace:'Calibri',
    });
  });

  // Right card — Water Purifier
  addGreenPanel(slide, 6.85, 2.18, 6.0, 4.52);
  slide.addText('Water Purifier', { x:7.1, y:2.3, w:5.5, h:0.45, fontSize:18, bold:true, color:YELLOW, fontFace:'Calibri' });
  slide.addText([
    { text:'Clean drinking water is not a given in remote locations like Govind. A water purifier installed at the team\'s office means:\n\n', options:{fontSize:13, color:WHITE} },
  ], { x:7.1, y:2.82, w:5.5, h:0.9, fontFace:'Calibri', valign:'top' });

  const waterPoints = [
    { label:'Safe, clean drinking water', sub:'every day — no compromise.' },
    { label:'Reduced risk of waterborne illness', sub:'protecting team health in an area with limited medical access.' },
    { label:'A more comfortable workspace', sub:'one less daily hardship for team members living and working in the field.' },
    { label:'Better retention', sub:'small amenities like this signal care for people, making it easier to attract and keep team members in a difficult posting.' },
  ];
  waterPoints.forEach((pt, i) => {
    slide.addText([
      { text:`• ${pt.label}: `, options:{bold:true, color:YELLOW} },
      { text:pt.sub, options:{color:WHITE} },
    ], { x:7.1, y:3.9 + i * 0.65, w:5.5, h:0.6, fontSize:12.5, fontFace:'Calibri', valign:'top' });
  });

  addFooter(slide, 'Programmatic Capital Expenses  |  Govind Wildlife Sanctuary, Uttarakhand');
}

// ── SLIDE 8 — Harrawala MRF: Strengthening Day-to-Day Operations ─────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Harrawala MRF — Strengthening Day-to-Day Operations');

  // Yellow banner
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.78);
  slide.addText('Your support equipped our Material Recovery Facility with the infrastructure needed to keep operations running smoothly — for the workers, and for the waste.', {
    x:0.65, y:1.15, w:12.03, h:0.78,
    fontSize:13.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // 3-column cards
  const cards = [
    {
      x: 0.5,
      title: 'High-Pressure\nWater Pump',
      titleColor: GREEN,
      bg: PANEL,
      border: GREEN,
      textColor: DARK,
      body: 'A high-pressure water pump keeps the Harrawala MRF hygienic. It enables thorough cleaning of waste-handling areas, vehicles, and processing equipment — a critical requirement when dealing with mixed, organic, and dry waste streams every day.\n\nHygiene at an MRF is not optional — it protects Green Workers, prevents contamination, and maintains the quality of recyclables.',
    },
    {
      x: 4.61,
      title: 'Industrial Cooler',
      titleColor: YELLOW,
      bg: GREEN,
      border: GREEN,
      textColor: WHITE,
      body: 'Dehradun summers are intense. Green Workers at the Harrawala MRF handle heavy physical labour — sorting, loading, and processing waste — under demanding conditions.\n\nAn industrial cooler at the facility provides much-needed relief during peak heat, reducing physical strain and health risks, and making the MRF a more dignified and sustainable place to work.\n\nIn Q3 alone, the MRF diverted over 519 MT of waste — none of that is possible without a workforce that can sustain the effort.',
    },
    {
      x: 8.72,
      title: 'Office Furniture',
      titleColor: GREEN,
      bg: PANEL,
      border: GREEN,
      textColor: DARK,
      body: 'A functioning office needs basic infrastructure. Furniture — desks, chairs, and seating for common areas — gives the MRF\'s coordination team a proper, dignified workspace.\n\nThis matters: the Harrawala MRF is not just a processing facility. It serves as a learning hub, hosting visits from government officials, IFS officers, urban planning engineers, and institutional delegations. A well-equipped space reflects the professionalism and credibility of the work happening here.',
    },
  ];

  cards.forEach(c => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x:c.x, y:2.15, w:3.91, h:4.55,
      fill:{color:c.bg}, line:{color:c.border, width:1.5}, rectRadius:0.1,
    });
    slide.addText(c.title, {
      x:c.x+0.2, y:2.25, w:3.5, h:0.65,
      fontSize:15, bold:true, color:c.titleColor, fontFace:'Calibri',
    });
    slide.addText(c.body, {
      x:c.x+0.2, y:2.98, w:3.5, h:3.55,
      fontSize:11.5, color:c.textColor, fontFace:'Calibri', valign:'top',
    });
  });

  addFooter(slide, 'Programmatic Capital Expenses  |  Harrawala MRF, Dehradun');
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
