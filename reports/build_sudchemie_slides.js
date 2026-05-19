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
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: CREAM }, line: { color: CREAM } });
}
function addTitle(slide, title) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 3.5, y: 0.22, w: 6.33, h: 0.72,
    fill: { color: CREAM }, line: { color: GREEN, width: 2 }, rectRadius: 0.15,
  });
  slide.addText(title, {
    x: 3.5, y: 0.22, w: 6.33, h: 0.72,
    fontSize: 22, bold: true, color: GREEN, fontFace: 'Calibri', align: 'center', valign: 'middle',
  });
}
function addFooter(slide, text) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: H - 0.72, w: W, h: 0.72, fill: { color: GREEN }, line: { color: GREEN } });
  slide.addText(text, {
    x: 0.3, y: H - 0.72, w: W - 0.6, h: 0.72,
    fontSize: 13, color: WHITE, fontFace: 'Calibri', align: 'center', valign: 'middle',
  });
}
function addBeigePanel(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, fill: { color: PANEL }, line: { color: GREEN, width: 1.5 }, rectRadius: 0.1,
  });
}
function addYellowPanel(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: YELLOW }, line: { color: YELLOW } });
}
function addGreenPanel(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: GREEN }, line: { color: GREEN } });
}

// ── SLIDE 5 — About the Project ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'About the Project');

  addBeigePanel(slide, 0.5, 1.15, 7.8, 5.55);
  slide.addText([
    { text: 'Project Catalyst', options: { bold: true, fontSize: 18, color: GREEN, breakLine: true } },
    { text: '\n', options: { fontSize: 8 } },
    { text: "Project Catalyst is Waste Warriors' ongoing initiative to support solid waste management operations at the Harrawala Material Recovery Facility (MRF) in Dehradun.\n\n", options: { fontSize: 14, color: DARK } },
    { text: 'The Harrawala MRF is a key facility where waste is collected, sorted, and processed — keeping it out of landfills and putting it back into the recycling chain.\n\n', options: { fontSize: 14, color: DARK } },
    { text: 'Your contribution to Waste Warriors made it possible for us to keep this facility running and to support the people who make it work every day.', options: { fontSize: 14, color: DARK } },
  ], { x: 0.75, y: 1.3, w: 7.3, h: 5.2, fontFace: 'Calibri', valign: 'top' });

  addYellowPanel(slide, 8.55, 1.15, 4.28, 5.55);
  slide.addText([
    { text: 'Location', options: { bold: true, fontSize: 15, color: GREEN, breakLine: true } },
    { text: 'Harrawala MRF, Dehradun', options: { fontSize: 14, color: DARK, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: 'FY', options: { bold: true, fontSize: 15, color: GREEN, breakLine: true } },
    { text: '2025 – 26', options: { fontSize: 14, color: DARK, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: 'Implemented by', options: { bold: true, fontSize: 15, color: GREEN, breakLine: true } },
    { text: 'Waste Warriors Society', options: { fontSize: 14, color: DARK } },
  ], { x: 8.75, y: 1.35, w: 3.85, h: 5.2, fontFace: 'Calibri', valign: 'top' });

  addFooter(slide, 'Waste Warriors Society  |  Harrawala MRF, Dehradun');
}

// ── SLIDE 6 — The Green Workers ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'The Green Workers');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.85);
  slide.addText('The heart of our operations — trained workers who collect, sort, and process waste every single day', {
    x: 0.6, y: 1.15, w: 12.13, h: 0.85,
    fontSize: 15, bold: true, color: DARK, fontFace: 'Calibri', align: 'center', valign: 'middle',
  });

  const boxes = [
    { x: 0.5,  label: 'Who They Are', body: 'Green Workers are trained waste management professionals who work at the Harrawala MRF. They are the ones who make clean waste processing possible — day in, day out.' },
    { x: 4.91, label: 'What They Do', body: 'They collect waste from the community, segregate it into categories, and ensure it is processed and sent for recycling. Their work directly keeps waste out of landfills.' },
    { x: 9.32, label: 'Why They Matter', body: 'Without Green Workers, none of this is possible. Your contribution ensured they had the support, stability, and care needed to continue their essential work.' },
  ];
  boxes.forEach(b => {
    addBeigePanel(slide, b.x, 2.2, 3.91, 4.5);
    slide.addText(b.label, { x: b.x + 0.2, y: 2.35, w: 3.5, h: 0.45, fontSize: 15, bold: true, color: GREEN, fontFace: 'Calibri' });
    slide.addText(b.body,  { x: b.x + 0.2, y: 2.9,  w: 3.5, h: 3.65, fontSize: 13, color: DARK, fontFace: 'Calibri', valign: 'top' });
  });

  addFooter(slide, 'Harrawala MRF, Dehradun');
}

// ── SLIDE 7 — Dignified Livelihoods ─────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Dignified Livelihoods');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.8);
  slide.addText('Your contribution ensured our Green Workers had fair wages, daily nutrition, and financial protection', {
    x: 0.6, y: 1.15, w: 12.13, h: 0.8,
    fontSize: 14, bold: true, color: DARK, fontFace: 'Calibri', align: 'center', valign: 'middle',
  });

  const cards = [
    { col:0, row:0, bg:PANEL, border:GREEN, titleColor:GREEN,  textColor:DARK,  label:'Fair Wages',   body:'Salaries were paid to Green Workers at the Harrawala MRF. A stable income gave them the dignity and financial security to focus on their work without worry.' },
    { col:1, row:0, bg:GREEN, border:GREEN, titleColor:YELLOW, textColor:WHITE, label:'Daily Meals',  body:'A nutritious meal was provided to every Green Worker at the MRF each working day — keeping them healthy, energised, and motivated throughout their demanding shifts.' },
    { col:0, row:1, bg:GREEN, border:GREEN, titleColor:YELLOW, textColor:WHITE, label:'Gratuity',     body:'Long-serving Green Workers received gratuity as a recognition of their commitment and service. Your contribution helped us honour this obligation and reward loyalty.' },
    { col:1, row:1, bg:PANEL, border:GREEN, titleColor:GREEN,  textColor:DARK,  label:'Health Cover', body:'Green Workers are exposed to health risks every day. Your contribution helped cover part of their health insurance, ensuring access to medical care when they need it most.' },
  ];

  const cardW = 5.9, cardH = 2.6, sx = 0.5, sy = 2.1, gx = 0.63, gy = 0.3;
  cards.forEach(c => {
    const x = sx + c.col*(cardW+gx);
    const y = sy + c.row*(cardH+gy);
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w:cardW, h:cardH, fill:{color:c.bg}, line:{color:c.border,width:1.5}, rectRadius:0.1 });
    slide.addText(c.label, { x:x+0.25, y:y+0.18, w:cardW-0.5, h:0.4,        fontSize:15,   bold:true, color:c.titleColor, fontFace:'Calibri' });
    slide.addText(c.body,  { x:x+0.25, y:y+0.65, w:cardW-0.5, h:cardH-0.75, fontSize:12.5, color:c.textColor, fontFace:'Calibri', valign:'top' });
  });

  addFooter(slide, 'Fair Wages  |  Daily Meals  |  Gratuity  |  Health Cover  |  Harrawala MRF, Dehradun');
}

// ── SLIDE 8 — Wet Waste & Composting ────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Wet Waste & Composting');

  addBeigePanel(slide, 0.5, 1.15, 7.8, 5.55);
  slide.addText([
    { text: 'Your contribution kept composting operations running at the MRF.\n\n', options: { bold: true, fontSize: 16, color: GREEN } },
    { text: 'The Harrawala MRF processes both dry and wet waste. For wet waste — food scraps and organic matter — the facility uses an in-vessel composting process to convert it into useful compost.\n\n', options: { fontSize: 13, color: DARK } },
    { text: 'Your contribution funded materials like sawdust, which are essential to the composting process. Sawdust helps manage moisture, reduce odour, and speed up decomposition — keeping the composting unit running smoothly.\n\n', options: { fontSize: 13, color: DARK } },
    { text: 'The compost produced is given back to the community as a small token of appreciation for segregating their waste at home — closing the loop and encouraging continued participation.', options: { fontSize: 13, color: DARK } },
  ], { x: 0.75, y: 1.3, w: 7.3, h: 5.2, fontFace: 'Calibri', valign: 'top' });

  addGreenPanel(slide, 8.55, 1.15, 4.28, 5.55);
  slide.addText([
    { text: '25 MT\n', options: { bold: true, fontSize: 36, color: YELLOW } },
    { text: 'of wet waste processed into compost\n\n', options: { fontSize: 13, color: WHITE } },
    { text: 'Compost returned to the community of Ward 97, Dehradun', options: { fontSize: 13, color: WHITE, italic: true } },
  ], { x: 8.75, y: 1.35, w: 3.85, h: 5.2, fontFace: 'Calibri', valign: 'middle', align: 'center' });

  addFooter(slide, 'Wet Waste Processing  |  Sawdust for Composting  |  Harrawala MRF, Dehradun');
}

// ── SLIDE 9 — On-Ground Coordination ────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'On-Ground Coordination');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.85);
  slide.addText('Your contribution funded a dedicated coordinator who kept everything running at the MRF', {
    x: 0.6, y: 1.15, w: 12.13, h: 0.85,
    fontSize: 15, bold: true, color: DARK, fontFace: 'Calibri', align: 'center', valign: 'middle',
  });

  const strips = [
    { label: 'Day-to-Day Operations',        body: 'The Project Coordinator oversaw all daily activities at the Harrawala MRF — ensuring waste collection, sorting, and processing happened smoothly and on schedule.' },
    { label: 'Supporting Green Workers',      body: 'The Coordinator worked directly with Green Workers on the ground — providing guidance, resolving issues, and making sure the team had what they needed to do their jobs well.' },
    { label: 'Keeping the Project on Track',  body: 'From tracking outputs to managing logistics, the Coordinator ensured that every aspect of the project was running as planned — and that your contribution was put to the best possible use.' },
  ];
  strips.forEach((s, i) => {
    const y = 2.2 + i * 1.35;
    addBeigePanel(slide, 0.5, y, 12.33, 1.2);
    slide.addText(s.label, { x: 0.75, y: y + 0.05, w: 3.5,  h: 0.45, fontSize: 14, bold: true, color: GREEN, fontFace: 'Calibri', valign: 'middle' });
    slide.addText(s.body,  { x: 4.4,  y: y + 0.05, w: 8.2,  h: 1.1,  fontSize: 13, color: DARK, fontFace: 'Calibri', valign: 'middle' });
  });

  addFooter(slide, 'Project Coordinator  |  Harrawala MRF, Dehradun');
}

// ── SLIDE 10 — What Your Support Made Possible ──────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'What Your Support Made Possible');

  slide.addText('Because of your contribution, Waste Warriors was able to keep the Harrawala MRF running, support the people who work there, and process waste that would otherwise have ended up in a landfill.', {
    x: 0.5, y: 1.1, w: 12.33, h: 0.9,
    fontSize: 14, color: DARK, fontFace: 'Calibri', align: 'center', italic: true,
  });

  const items = [
    { num: '175 MT+', label: 'of dry & wet waste\ndiverted from landfills' },
    { num: '25 MT',   label: 'of wet waste converted\ninto compost' },
    { num: '35+',     label: 'waste workers & staff\nworking safely at MRF' },
    { num: '1,472',   label: 'families directly served\nin Ward 97, Dehradun' },
    { num: '91%',     label: 'of dry waste processed\nsent for recycling' },
    { num: '1 Year',  label: 'of uninterrupted\nMRF operations' },
  ];
  items.forEach((item, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 4.28, y = 2.2 + row * 2.1;
    const bg = (i % 2 === 0) ? GREEN : YELLOW;
    const numColor = (i % 2 === 0) ? YELLOW : GREEN;
    const txtColor = (i % 2 === 0) ? WHITE : DARK;
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 4.08, h: 1.85, fill: { color: bg }, line: { color: bg }, rectRadius: 0.1 });
    slide.addText(item.num,   { x: x + 0.15, y: y + 0.12, w: 3.78, h: 0.7, fontSize: 28, bold: true, color: numColor, fontFace: 'Calibri', align: 'center' });
    slide.addText(item.label, { x: x + 0.15, y: y + 0.82, w: 3.78, h: 0.9, fontSize: 12, color: txtColor, fontFace: 'Calibri', align: 'center', valign: 'top' });
  });

  addFooter(slide, 'Thank you for making this possible  |  Waste Warriors Society  |  FY 2025–26');
}

pptx.writeFile({ fileName: 'SudChemie_FundUtilisation_Slides5to11.pptx' })
  .then(() => console.log('DONE'))
  .catch(e => { console.error(e); process.exit(1); });
