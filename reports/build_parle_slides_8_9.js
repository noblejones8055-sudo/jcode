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

const FOOTER = 'Project Pir Panjal  |  Parle Biscuits Pvt. Ltd. × Waste Warriors Society  |  FY 2025–26';

// ── SLIDE 8 — Dharamshala Smart City (Swachh Survekshan Partnership) ─────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Dharamshala Smart City — Swachh Survekshan Partnership');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Working alongside the Dharamshala Municipal Corporation to fix waste at its source — across all 17 wards', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left beige narrative panel
  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Work on the Ground\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Mixed waste is hard to process. The real fix has to happen ', options:{fontSize:12.5, color:DARK} },
    { text:'at the source', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:'. Made possible by your partnership, we are working hand-in-hand with the Dharamshala Municipal Corporation to strengthen source segregation, improve collection discipline, and support the city\'s push for a stronger Swachh Survekshan 2026 ranking.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Our team engages daily with households, commercial establishments, sanitation workers, and contractor drivers — closing feedback loops between the citizen, the collection vehicle, and the ULB. The 27 collection vehicles operated by the third-party contractor are now monitored on the ground for consistency and corrective action.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Behaviour change is being driven through Nukkad Nataks, wall murals, Swachhata Shapaths, and Swachhata Ambassadors — while a digital dashboard is helping the Corporation see, in real time, what\'s working and where to intervene.', options:{fontSize:12.5, color:DARK} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  // Right green panel — Key Levers
  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Key Levers\n\n', options:{bold:true, fontSize:15, color:YELLOW} },
    { text:'17 wards\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'citywide IEC & monitoring\n\n', options:{fontSize:11, color:CREAM} },
    { text:'27 collection vehicles\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'third-party fleet monitored\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Swachh Survekshan 2026\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'aligned with national indicators\n\n', options:{fontSize:11, color:CREAM} },
    { text:'IEC Toolkit\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'Nukkad Nataks • Wall Murals\nSwachhata Shapaths\nSwachhata Ambassadors\n\n', options:{fontSize:11, color:CREAM} },
    { text:'Digital Dashboard\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'real-time data for the ULB', options:{fontSize:11, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 9 — Dehradun Model Ward: Wet Waste Processing Shed ─────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Dehradun Model Ward — Wet Waste Processing Shed');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Scaling up wet waste processing capacity at the Harrawala MRF', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Four beige strips on the left (stat strips)
  const strips = [
    'Approximately 1,200 kg of wet waste is now received and processed daily through collection operations.',
    'Earlier, the wet waste unit at Harrawala MRF had a processing capacity of only 700–800 kg per day.',
    'With the development of the new processing shed, the additional waste generated is now being efficiently managed and processed.',
    'Powered by your support, this also helped co-fund employee salaries and meet ongoing project requirements.',
  ];
  const stripH = 1.05;
  const stripY0 = 2.05;
  const stripGap = 0.15;
  const stripW = 7.8;
  strips.forEach((text, i) => {
    const y = stripY0 + i * (stripH + stripGap);
    addBeigePanel(slide, 0.5, y, stripW, stripH);
    slide.addText(text, {
      x:0.75, y:y, w:stripW-0.5, h:stripH,
      fontSize:13, color:DARK, fontFace:'Calibri', valign:'middle',
    });
  });

  // Right yellow callout
  addYellowPanel(slide, 8.6, 2.05, 4.23, 3.6);
  slide.addText([
    { text:'Community Engagement\n\n', options:{bold:true, fontSize:14, color:GREEN} },
    { text:'Alongside the new shed, the MRF & Model Ward team conducted ', options:{fontSize:12, color:DARK} },
    { text:'5 awareness workshops', options:{bold:true, fontSize:12, color:GREEN} },
    { text:' on wet and dry waste management — engaging ', options:{fontSize:12, color:DARK} },
    { text:'340 community members', options:{bold:true, fontSize:12, color:GREEN} },
    { text:' and reinforcing source segregation through active participation.', options:{fontSize:12, color:DARK} },
  ], { x:8.78, y:2.18, w:3.9, h:3.4, fontFace:'Calibri', valign:'top' });

  // Right green callout (smaller, below yellow) — capacity uplift highlight
  addGreenPanel(slide, 8.6, 5.8, 4.23, 0.95);
  slide.addText([
    { text:'Capacity uplift:\n', options:{bold:true, fontSize:12, color:YELLOW} },
    { text:'700–800 kg/day  →  1,200 kg/day', options:{bold:true, fontSize:14, color:WHITE} },
  ], { x:8.78, y:5.85, w:3.9, h:0.85, fontFace:'Calibri', align:'center', valign:'middle' });

  addFooter(slide, FOOTER);
}

pptx.writeFile({ fileName: 'Parle_Slides_8and9_Preview.pptx' })
  .then(name => console.log('Wrote', name));
