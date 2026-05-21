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

const FOOTER = 'Project Pir Panjal  |  Parle Biscuits Pvt. Ltd. × Waste Warriors Society  |  FY 2025–26';

// ── SLIDE 5 — About the Project ──────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'About the Project');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Strengthening waste systems across five Himalayan landscapes — from rural panchayats to Ramsar wetlands', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left beige narrative
  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'A Multi-Site Programme\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Waste Warriors works across the eco-sensitive belts of Himachal Pradesh and Uttarakhand — high-altitude, high-footfall, and high-stakes geographies where waste systems are still being built from the ground up.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Thanks to your support, this year your partnership powered our work across ', options:{fontSize:12.5, color:DARK} },
    { text:'five distinct intervention sites', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:' — each at a different stage of maturity, each anchored in community ownership, and each contributing to a shared goal of cleaner, climate-resilient Himalayan communities.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'You enabled this. We are the hands on the ground.', options:{fontSize:12.5, color:DARK, italic:true} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  // Right green panel — Where Your Support Travels
  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Where Your\nSupport Travels\n\n', options:{bold:true, fontSize:15, color:YELLOW} },
    { text:'• Dharamshala Rural\n   (Pyaare Pahad)\n\n', options:{fontSize:12, color:WHITE} },
    { text:'• Dharamshala Smart City\n   (Urban MRF)\n\n', options:{fontSize:12, color:WHITE} },
    { text:'• Dehradun Model Ward\n   Program (Harrawala)\n\n', options:{fontSize:12, color:WHITE} },
    { text:'• Special Projects\n   (Renukaji & Asan)\n\n', options:{fontSize:12, color:WHITE} },
    { text:'• Senior Management\n   Involvement', options:{fontSize:12, color:WHITE} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 6 — How Your Support Was Used ──────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'How Your Support Was Used');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Powering the backbone of ground operations across every project location', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // 5 horizontal strips (alternating beige & green)
  const strips = [
    { label:'Salaries & Project Teams', desc:'Paying the people who run the work — coordinators, field staff, and community mobilisers across every site.', color:'beige' },
    { label:'Accommodation & Project Rentals', desc:'Keeping our field bases, segregation centres, and team housing operational across remote and rural locations.', color:'green' },
    { label:'IEC Materials', desc:'Posters, banners, signages, wall murals, workshop kits — the everyday tools that turn awareness into action.', color:'beige' },
    { label:'Refreshments & Community Engagement', desc:'Tea, meals, and hospitality during clean-up drives, Chai pe Charchas, exposure visits, and stakeholder meetings.', color:'green' },
    { label:'Senior Management Involvement', desc:'Strategic oversight, policy advocacy, partnership building, and field reviews from our leadership team.', color:'beige' },
  ];
  const stripH = 0.78;
  const stripY0 = 2.05;
  const stripGap = 0.18;
  strips.forEach((s, i) => {
    const y = stripY0 + i * (stripH + stripGap);
    if (s.color === 'beige') {
      addBeigePanel(slide, 0.5, y, 12.33, stripH);
      slide.addText([
        { text:s.label + '   ', options:{bold:true, fontSize:13.5, color:GREEN} },
        { text:s.desc, options:{fontSize:12, color:DARK} },
      ], { x:0.75, y:y, w:11.83, h:stripH, fontFace:'Calibri', valign:'middle' });
    } else {
      addGreenPanel(slide, 0.5, y, 12.33, stripH);
      slide.addText([
        { text:s.label + '   ', options:{bold:true, fontSize:13.5, color:YELLOW} },
        { text:s.desc, options:{fontSize:12, color:WHITE} },
      ], { x:0.75, y:y, w:11.83, h:stripH, fontFace:'Calibri', valign:'middle' });
    }
  });

  addFooter(slide, FOOTER);
}

// ── SLIDE 7 — DHMR / Pyaare Pahad ────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Dharamshala Rural — Pyaare Pahad');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Community-led waste systems taking root across six gram panchayats', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Work on the Ground\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Made possible by your partnership, our team continued building one of our most mature rural waste systems — door-to-door collection, user fee discipline, wet waste solutions, and growing community ownership.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Local Entrepreneurs now run day-to-day operations, women\'s groups have become first-line champions of the system, and Panchayat representatives are joining collection rounds themselves — a clear sign that the work is becoming theirs, not ours.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Wet waste solutions through rapid digesters and local piggeries now cover all seven panchayats, closing a long-standing gap in the system.', options:{fontSize:12.5, color:DARK} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Highlights\n\n', options:{bold:true, fontSize:15, color:YELLOW} },
    { text:'6 gram panchayats\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'~20 villages covered\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'3,588 active service units\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'households + commercial\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Paying units 33% → 39%\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'growing community trust\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'2 Local Entrepreneurs\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'running daily operations\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Wet waste covered\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'across all 7 panchayats', options:{fontSize:11.5, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 8 — DHMU / Dharamshala Smart City ──────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Dharamshala Smart City — Urban MRF');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('A growing engine for material recovery in the Himalayas', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Work on the Ground\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'With your backing, our Dharamshala Urban Material Recovery Facility scaled up its capacity to process some of the hardest waste streams in the system — multi-layered plastics, low-value packaging, and non-recyclables that would otherwise end up in landfills or eco-sensitive rivers.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'A new baler machine improved storage and operational efficiency, and the facility secured its formal Consent-to-Establish — a regulatory milestone that moves the MRF from preparatory work to legally authorised, long-term operations.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Our Ecodaan tourism programme deepened in McLeodganj, Bhagsu, and Dharamkot — embedding sustainability into the visitor experience.', options:{fontSize:12.5, color:DARK} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Highlights\n\n', options:{bold:true, fontSize:15, color:YELLOW} },
    { text:'184+ MT dry waste\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'diverted from landfills\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'17,000+ kg MLP\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'multi-layered plastic recovered\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Consent-to-Establish\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'regulatory milestone secured\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'New baler installed\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'storage & efficiency upgrade\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Ecodaan\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'800+ tourists engaged', options:{fontSize:11.5, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 9 — DDN MWP / Harrawala ────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Dehradun Model Ward Program');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('An urban model where doorstep collection, user fees, and an MRF come together as one system', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:13.5, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'The Work on the Ground\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Powered by your support, Ward 97 in Harrawala continues to demonstrate how consistent doorstep collection, growing user-fee discipline, and a fully functioning Material Recovery Facility can come together as one self-sustaining urban model.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'This year, the facility addressed a longstanding economic challenge in low-value plastics by onboarding Vardhman Industries — turning multi-layered plastic from a freight cost into a revenue stream. Composting capacity expanded significantly, and senior officials including the Dehradun Mayor and IFS officers visited the MRF to study the model.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'The ward also became a learning hub — hosting visiting officials, students, and policy delegations from across India.', options:{fontSize:12.5, color:DARK} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Highlights\n\n', options:{bold:true, fontSize:15, color:YELLOW} },
    { text:'241+ MT collected\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'from Ward 97 (Q1 alone)\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'~₹1.63L / month\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'user fees, growing steadily\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'1,200 kg / day\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'wet waste processing capacity\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'MLP at ₹5 / kg\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'new revenue stream unlocked\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Mayor & IFS\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'official visits & endorsement', options:{fontSize:11.5, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 10 — Special Projects: Renukaji & Asan ─────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Special Projects — Renukaji & Asan');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Protecting Himalayan ecosystems of deep ecological and cultural value', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  // Left panel — Renukaji
  addBeigePanel(slide, 0.5, 2.0, 6.1, 4.75);
  slide.addText([
    { text:'Renukaji Wetlands\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'India\'s smallest Ramsar site\n\n', options:{italic:true, fontSize:11.5, color:DARK} },
    { text:'Made possible by your partnership, we continued building a community-led waste management system at one of India\'s most ecologically sensitive wetland sites — home to over 100 bird species and a major pilgrimage destination.\n\n', options:{fontSize:12, color:DARK} },
    { text:'Highlights this year:\n', options:{bold:true, fontSize:12.5, color:GREEN} },
    { text:'• Household participation grew ', options:{fontSize:12, color:DARK} },
    { text:'4x (37 → 200)\n', options:{bold:true, fontSize:12, color:GREEN} },
    { text:'• Door-to-door collection now active across ', options:{fontSize:12, color:DARK} },
    { text:'two panchayats\n', options:{bold:true, fontSize:12, color:GREEN} },
    { text:'• Renuka Fair: ', options:{fontSize:12, color:DARK} },
    { text:'10.2 MT waste managed, 5.6 MT processed\n', options:{bold:true, fontSize:12, color:GREEN} },
    { text:'• Dedicated waste storage sites in ', options:{fontSize:12, color:DARK} },
    { text:'Dadahu & Khala Kiyar', options:{bold:true, fontSize:12, color:GREEN} },
  ], { x:0.72, y:2.12, w:5.7, h:4.5, fontFace:'Calibri', valign:'top' });

  // Right panel — Asan (green, smaller content as it's emerging)
  addGreenPanel(slide, 6.75, 2.0, 6.08, 4.75);
  slide.addText([
    { text:'Asan\n', options:{bold:true, fontSize:17, color:YELLOW} },
    { text:'An emerging chapter\n\n', options:{italic:true, fontSize:11.5, color:CREAM} },
    { text:'With your backing, we are extending our community-led model to the Asan landscape — an ecologically rich and culturally significant zone where the need for organised waste systems is growing rapidly.\n\n', options:{fontSize:12, color:WHITE} },
    { text:'This is early-stage work — laying the groundwork, building relationships with local stakeholders, and applying lessons from our Renukaji intervention to ensure that the system here is built right from the start.\n\n', options:{fontSize:12, color:WHITE} },
    { text:'In the coming year, this work is expected to scale meaningfully — and your partnership is helping us hold space for it.', options:{italic:true, fontSize:12, color:CREAM} },
  ], { x:6.97, y:2.12, w:5.68, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── SLIDE 11 — Senior Management Involvement ─────────────────────────────────
{
  const slide = pptx.addSlide();
  addBackground(slide);
  addTitle(slide, 'Senior Management Involvement');

  addYellowPanel(slide, 0.5, 1.15, 12.33, 0.72);
  slide.addText('Strategic oversight that holds the work together across geographies', {
    x:0.65, y:1.15, w:12.03, h:0.72,
    fontSize:14, bold:true, color:DARK, fontFace:'Calibri', align:'center', valign:'middle',
  });

  addBeigePanel(slide, 0.5, 2.0, 7.8, 4.75);
  slide.addText([
    { text:'Why This Matters\n\n', options:{bold:true, fontSize:17, color:GREEN} },
    { text:'Field teams deliver the work, but it is senior leadership that holds the larger system together — building institutional partnerships, navigating policy, steering long-term direction, and ensuring that quality, accountability, and learning travel across every project site.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'Thanks to your support, this layer of strategic oversight remained active and well-resourced through the year — strengthening our state and district relationships, guiding new geographies, and creating roadmaps for long-term sustainability.\n\n', options:{fontSize:12.5, color:DARK} },
    { text:'You enabled the foundation that lets the work continue, scale, and stay grounded.', options:{fontSize:12.5, color:DARK, italic:true} },
  ], { x:0.72, y:2.12, w:7.38, h:4.5, fontFace:'Calibri', valign:'top' });

  addGreenPanel(slide, 8.55, 2.0, 4.28, 4.75);
  slide.addText([
    { text:'Leadership Impact\n\n', options:{bold:true, fontSize:15, color:YELLOW} },
    { text:'Strategic MoU\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'with HP Rural Department\nfor next 3 years\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Sustainability Roadmaps\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'for rural project geographies\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Policy & Stakeholder\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'liaison across districts & states\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Cross-Site Field Reviews\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'and quality oversight\n\n', options:{fontSize:11.5, color:CREAM} },
    { text:'Partner & Donor\n', options:{bold:true, fontSize:12.5, color:WHITE} },
    { text:'engagement and reporting', options:{fontSize:11.5, color:CREAM} },
  ], { x:8.75, y:2.15, w:3.9, h:4.5, fontFace:'Calibri', valign:'top' });

  addFooter(slide, FOOTER);
}

// ── write file ───────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: 'Parle_FundUtilisation_Slides5to11.pptx' })
  .then(name => console.log('Wrote', name));
