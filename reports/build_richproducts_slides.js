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
    x:2.5, y:0.22, w:8.33, h:0.72,
    fill:{color:CREAM}, line:{color:GREEN, width:2}, rectRadius:0.15,
  });
  slide.addText(title, {
    x:2.5, y:0.22, w:8.33, h:0.72,
    fontSize:21, bold:true, color:GREEN, fontFace:'Calibri', align:'center', valign:'middle',
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
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Project Mighty Mountains — building community-owned waste management systems across rural Dharamshala', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left beige narrative
  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Pyaare Pahad Programme\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Dharamshala\'s rural belt sits at the foothills of the Dhauladhar range — once-quiet villages now transformed by tourism, homestays, and a steady rise in consumption. With that change has come a waste crisis that local infrastructure was never built for.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Waste Warriors has been working in this region since 2022 to set up door-to-door collection, segregation, user-fee systems, and circular waste linkages — all anchored in community ownership.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Through your partnership, we were able to keep this work running end-to-end this year — from the people who lead it on the ground to the operational backbone that makes everyday delivery possible.', options:{fontSize:12.5, color:DARK, italic:true} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  // Right green panel — How Your Support Was Used
  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'How Your Support\nWas Used\n\n', options:{bold:true, fontSize:14, color:YELLOW, breakLine:false} },
    { text:'• Project Team Salaries\n', options:{fontSize:12, color:WHITE} },
    { text:'• Local Entrepreneur Model — Impact Funding\n', options:{fontSize:12, color:WHITE} },
    { text:'• Segregation Centre — Rent & Utilities\n', options:{fontSize:12, color:WHITE} },
    { text:'• Low-Value Waste Transportation\n', options:{fontSize:12, color:WHITE} },
    { text:'• Day-to-Day Project Overheads\n', options:{fontSize:12, color:WHITE} },
    { text:'• Repair & Maintenance\n', options:{fontSize:12, color:WHITE} },
    { text:'• IT & Communication Tools\n\n', options:{fontSize:12, color:WHITE} },
    { text:'Together, these inputs kept the entire DHMR programme running — for the team, the community, and the waste.', options:{fontSize:11, italic:true, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, 'Project Mighty Mountains  |  Dharamshala Rural, Himachal Pradesh  |  FY 2025–26');
}

// ── SLIDE 7 — The Programme on the Ground ────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'The Programme on the Ground');

  // Yellow banner
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('A community-led waste management system that now reaches thousands of households across rural Dharamshala', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left beige narrative
  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'Where We Work\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Our team is active across ', options:{fontSize:12.5, color:DARK} },
    { text:'6 gram panchayats covering roughly 20 villages', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' in the Dharamshala block of Kangra district — out of 27 panchayats in the block, leaving substantial room to scale.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Together, these panchayats represent ', options:{fontSize:12.5, color:DARK} },
    { text:'3,588 active service units', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' — a mix of households and commercial establishments — all part of an organised door-to-door collection system tied to user fees, community engagement, and onward processing.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Awareness on the Ground\n\n', options:{bold:true, fontSize:14, color:GREEN} },
    { text:'Through the year, our team ran clean-up drives, school programmes, menstrual hygiene workshops, and community meetings — engaging over ', options:{fontSize:12.5, color:DARK} },
    { text:'1,200 people', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:'. The recently launched ', options:{fontSize:12.5, color:DARK} },
    { text:'Swachh Vyavsay, Swasth Vyavsay', options:{bold:true, italic:true, fontSize:12.5, color:GREEN} },
    { text:' campaign brought ', options:{fontSize:12.5, color:DARK} },
    { text:'60+ commercial properties', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' into the conversation — extending the programme beyond households into the local economy.', options:{fontSize:12.5, color:DARK} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  // Right green stats panel
  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  const stats = [
    { num:'6',       label:'Gram Panchayats\nunder active operations' },
    { num:'~20',     label:'Villages reached across\nthe Dharamshala block' },
    { num:'3,588',   label:'Service units served\n(households + commercial)' },
    { num:'1,200+',  label:'Community members engaged\nthrough drives & workshops' },
    { num:'60+',     label:'Commercial properties under\nSwachh Vyavsay campaign' },
  ];
  stats.forEach((s, i) => {
    slide.addText(s.num,   { x:8.72, y:2.1 + i*0.88, w:1.5,  h:0.5, fontSize:24, bold:true, color:YELLOW, fontFace:'Calibri', align:'center', valign:'middle' });
    slide.addText(s.label, { x:10.1, y:2.1 + i*0.88, w:2.6,  h:0.7, fontSize:10.5, color:WHITE, fontFace:'Calibri', valign:'middle' });
  });

  addFooter(slide, 'Dharamshala Rural  |  6 Panchayats  |  3,588 Units  |  ~20 Villages');
}

// ── SLIDE 8 — The Local Entrepreneur Model ───────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'The Local Entrepreneur Model');

  // Yellow banner
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Building local ownership of waste systems — and your Impact Funding is what makes it work in the early years', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // 3 horizontal strips
  const strips = [
    {
      label: 'How the Model Works',
      body: 'Day-to-day waste collection, user-fee collection, and transportation in the 6 panchayats are run by a Local Entrepreneur — a small business owner from the region — rather than directly by Waste Warriors. The panchayats pay a per-household rate for the service; the LE provides the labour, vehicles, and supervisors. Our role is to enable, monitor, and support.',
    },
    {
      label: 'Why the Model Matters',
      body: 'This shifts ownership of operations into local hands. It creates entrepreneurial income, builds long-term institutional capacity within the community, and ensures the system can keep running even when external partners step back. It is how waste management becomes a permanent feature of the local economy — not a project that ends.',
    },
    {
      label: 'How Your Impact Funding Helps',
      body: 'In rural Himalayan settings, recyclable sales rarely cover the full cost of running collection and processing. The Impact Funding Support enabled through your partnership bridges that gap — covering the difference between what panchayats and households can pay, and what it actually costs to keep operations sustainable. Without it, the model would not survive its early, formative years.',
    },
  ];
  strips.forEach((s, i) => {
    const y = 2.05 + i * 1.55;
    addBeigePanel(slide, 0.5, y, 12.33, 1.42);
    slide.addText(s.label, { x:0.75, y:y+0.08, w:3.6,  h:0.45, fontSize:14, bold:true, color:GREEN, fontFace:'Calibri', valign:'middle' });
    slide.addText(s.body,  { x:4.45, y:y+0.05, w:8.2,  h:1.32, fontSize:12, color:DARK, fontFace:'Calibri', valign:'middle' });
  });

  addFooter(slide, 'Local Entrepreneur Model  |  Impact Funding Support  |  Dharamshala Rural');
}

// ── SLIDE 9 — The People & Operations Behind the Project ─────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'The People & Operations Behind the Project');

  // Yellow banner
  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Made possible by your partnership — the team, the infrastructure, and the everyday inputs that keep DHMR running', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // 2x3 grid of cards
  const items = [
    {
      col:0, row:0, bg:PANEL, titleColor:GREEN, textColor:DARK,
      label:'Project Team Salaries',
      body:'The Project Manager (part-time), Project Coordinator, and Senior Management leads who direct strategy, oversee field operations, manage donor reporting, and keep the project on track.',
    },
    {
      col:1, row:0, bg:GREEN, titleColor:YELLOW, textColor:WHITE,
      label:'Segregation Centre',
      body:'The rent and utilities for our dry waste processing space — the physical home of segregation, sorting, and onward dispatch. Without a stable processing site, no collection system can sustain itself.',
    },
    {
      col:2, row:0, bg:PANEL, titleColor:GREEN, textColor:DARK,
      label:'Low-Value Waste Transportation',
      body:'Low-value plastics and multi-layered packaging — the hardest waste streams to recycle locally — were transported out to authorised recyclers, ensuring nothing was burnt, buried, or left behind.',
    },
    {
      col:0, row:1, bg:GREEN, titleColor:YELLOW, textColor:WHITE,
      label:'Day-to-Day Project Overheads',
      body:'The unseen but essential running costs — operational consumables, on-ground supplies, courier and communications, and the small day-to-day expenditures that quietly keep a field project alive.',
    },
    {
      col:1, row:1, bg:PANEL, titleColor:GREEN, textColor:DARK,
      label:'Repair & Maintenance',
      body:'Keeping vehicles, collection equipment, and processing infrastructure operational year-round. Downtime in a rural waste system means waste piling up — repair and maintenance is what prevents that.',
    },
    {
      col:2, row:1, bg:GREEN, titleColor:YELLOW, textColor:WHITE,
      label:'IT & Communication Tools',
      body:'The devices that connect our team to the field, the funder, and to each other — enabling data collection, real-time coordination, donor reporting, and timely community communications.',
    },
  ];

  const cardW = 4.05, cardH = 2.25, sx = 0.4, sy = 2.05, gx = 0.13, gy = 0.18;
  items.forEach(c => {
    const x = sx + c.col * (cardW + gx);
    const y = sy + c.row * (cardH + gy);
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y, w:cardW, h:cardH,
      fill:{color:c.bg}, line:{color:GREEN, width:1.2}, rectRadius:0.1,
    });
    slide.addText(c.label, { x:x+0.18, y:y+0.12, w:cardW-0.36, h:0.4, fontSize:13, bold:true, color:c.titleColor, fontFace:'Calibri' });
    slide.addText(c.body,  { x:x+0.18, y:y+0.55, w:cardW-0.36, h:cardH-0.65, fontSize:10.5, color:c.textColor, fontFace:'Calibri', valign:'top' });
  });

  addFooter(slide, 'Salaries  |  Operations  |  Infrastructure  |  Transportation  |  IT  |  Maintenance');
}

// ── SLIDE 10 — What This Made Possible ───────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'What This Made Possible');

  slide.addText('Through your partnership, the DHMR programme converted everyday inputs into measurable on-ground impact this year', {
    x:0.5, y:1.1, w:12.33, h:0.7,
    fontSize:13.5, italic:true, color:DARK, fontFace:'Calibri', align:'center',
  });

  const tiles = [
    { num:'237 MT',   label:'of dry waste collected\nand processed' },
    { num:'9,278 kg', label:'of wet waste\ncollected and processed' },
    { num:'3,588',    label:'service units linked to\norganised waste collection' },
    { num:'1,200+',   label:'community members engaged\nthrough drives & workshops' },
    { num:'60+',      label:'commercial properties under\nSwachh Vyavsay, Swasth Vyavsay' },
    { num:'6 GPs',    label:'covering ~20 villages\nin the Dharamshala block' },
  ];
  tiles.forEach((t, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 4.28, y = 2.0 + row * 2.25;
    const bg = (i % 2 === 0) ? GREEN : YELLOW;
    const numColor = (i % 2 === 0) ? YELLOW : GREEN;
    const txtColor = (i % 2 === 0) ? WHITE : DARK;
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w:4.08, h:2.0, fill:{color:bg}, line:{color:bg}, rectRadius:0.1 });
    slide.addText(t.num,   { x:x+0.15, y:y+0.18, w:3.78, h:0.78, fontSize:30, bold:true, color:numColor, fontFace:'Calibri', align:'center' });
    slide.addText(t.label, { x:x+0.15, y:y+1.05, w:3.78, h:0.85, fontSize:12, color:txtColor, fontFace:'Calibri', align:'center', valign:'top' });
  });

  addFooter(slide, 'Powered by your partnership  |  Project Mighty Mountains  |  FY 2025–26');
}

// ── SLIDE 11 — Key Challenges & Bottlenecks (Q3-derived) ─────────────────────
{
  const slide = pptx.addSlide();

  // Dark green BG
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:W, h:H, fill:{color:GREEN}, line:{color:GREEN} });

  // Inverted yellow title pill
  slide.addShape(pptx.ShapeType.roundRect, {
    x:2.8, y:0.28, w:7.73, h:0.78,
    fill:{color:YELLOW}, line:{color:CREAM, width:2.5}, rectRadius:0.18,
  });
  slide.addText('Key Challenges & Bottlenecks', {
    x:2.8, y:0.28, w:7.73, h:0.78,
    fontSize:24, bold:true, color:GREEN, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Subtitle
  slide.addText('Honest reflections from Q3 — what still tests the durability of the DHMR programme', {
    x:0.5, y:1.18, w:12.33, h:0.42,
    fontSize:13, italic:true, color:CREAM, fontFace:'Calibri', align:'center', valign:'middle',
  });

  const challenges = [
    {
      num:'01',
      title:'Closing the Cost Recovery Gap',
      body:'In our strongest panchayat this quarter (Sheela Butehad), user fees reached 100% — but processing costs were only ~60% recovered. The remaining gap is what Impact Funding bridges; sustained partnership is still essential.',
    },
    {
      num:'02',
      title:'Managing Wet Waste at Scale',
      body:'Wet waste solutions had to be designed from scratch this quarter — a rapid digester and a local piggery together handle up to 300 kg daily, but sustaining and replicating this across all panchayats remains a real operational lift.',
    },
    {
      num:'03',
      title:'Building Supervisory Capacity Locally',
      body:'Ecopreneur-led supervision had visible gaps in fee collection accuracy and financial oversight. Replacing them with two Community Resource Persons strengthened the system — but local supervisory talent remains scarce and slow to build.',
    },
    {
      num:'04',
      title:'Sustaining Community Engagement',
      body:'Behaviour change in rural waste systems is not linear. Awareness fatigue, segregation slippage, and inconsistent participation continue to require constant reinforcement through drives, workshops, and on-ground presence.',
    },
    {
      num:'05',
      title:'Ensuring Funding Continuity for the LE Model',
      body:'The Local Entrepreneur model needs multi-year horizons to stabilise. Funding fluctuations — even temporary ones — risk unwinding hard-won community trust and operational momentum already in place.',
    },
  ];

  const stripX = 0.5, stripW = 12.33, stripH = 1.0, stripStart = 1.7, stripGap = 0.05;

  challenges.forEach((c, i) => {
    const y = stripStart + i * (stripH + stripGap);

    slide.addShape(pptx.ShapeType.roundRect, {
      x:stripX, y, w:stripW, h:stripH,
      fill:{color:CREAM}, line:{color:CREAM}, rectRadius:0.06,
    });
    slide.addShape(pptx.ShapeType.rect, {
      x:stripX, y, w:1.4, h:stripH,
      fill:{color:YELLOW}, line:{color:YELLOW},
    });
    slide.addText(c.num, {
      x:stripX, y, w:1.4, h:stripH,
      fontSize:44, bold:true, color:GREEN, fontFace:'Calibri', align:'center', valign:'middle',
    });
    slide.addText(c.title, {
      x:stripX+1.6, y:y+0.08, w:stripW-1.8, h:0.36,
      fontSize:15, bold:true, color:GREEN, fontFace:'Calibri', valign:'middle',
    });
    slide.addText(c.body, {
      x:stripX+1.6, y:y+0.44, w:stripW-1.8, h:stripH-0.5,
      fontSize:11.5, color:DARK, fontFace:'Calibri', valign:'top',
    });
  });

  // Yellow footer
  slide.addShape(pptx.ShapeType.rect, { x:0, y:H-0.55, w:W, h:0.55, fill:{color:YELLOW}, line:{color:YELLOW} });
  slide.addText('Your continued partnership is what helps us turn these bottlenecks into systems that last', {
    x:0.3, y:H-0.55, w:W-0.6, h:0.55,
    fontSize:13, bold:true, italic:true, color:GREEN, fontFace:'Calibri', align:'center', valign:'middle',
  });
}

pptx.writeFile({ fileName: 'RichProducts_FundUtilisation_Slides6to11.pptx' })
  .then(() => console.log('DONE — RichProducts_FundUtilisation_Slides6to11.pptx'))
  .catch(e => { console.error(e); process.exit(1); });
