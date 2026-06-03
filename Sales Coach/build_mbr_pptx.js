// April 2026 NAM MBR — PPTX builder (rebuildable from this script)
const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'April 2026 NAM MBR — Thomas R Lindsey';
pres.author = 'Thomas R Lindsey';

const C = {
  bg: '0F1117', surface: '1A1D27', surface2: '222632', border: '2A2E3B',
  text: 'E4E6ED', muted: '8B8FA3', white: 'FFFFFF',
  accent: '6C5CE7', green: '00B894', orange: 'FDCB6E', red: 'E17055', blue: '74B9FF', navy: '0D2340'
};
const F = { head: 'Calibri', body: 'Calibri' };

pres.defineSlideMaster({
  title: 'DARK',
  background: { color: C.bg },
  objects: [
    { text: { text: 'April NAM MBR · Thomas R Lindsey · 05.04.2026', options: { x: 0.4, y: 7.15, w: 8, h: 0.25, fontFace: F.body, fontSize: 8, color: C.muted } } },
    { text: { text: 'SUPERMETRICS', options: { x: 11.4, y: 7.15, w: 1.8, h: 0.25, fontFace: F.head, fontSize: 8, color: C.muted, align: 'right', bold: true, charSpacing: 2 } } }
  ]
});

const slideTitle = (slide, num, title) => {
  slide.addText(`SLIDE ${num}`, { x: 0.4, y: 0.3, w: 6, h: 0.3, fontFace: F.body, fontSize: 9, color: C.muted, bold: true, charSpacing: 4 });
  slide.addText(title, { x: 0.4, y: 0.55, w: 12.5, h: 0.7, fontFace: F.head, fontSize: 28, bold: true, color: C.white });
};

// ===== S1 TITLE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  s.background = { color: C.navy };
  s.addText('MONTHLY BUSINESS REVIEW', { x: 0.7, y: 1.8, w: 12, h: 0.4, fontFace: F.body, fontSize: 13, color: 'A0B0D0', bold: true, charSpacing: 5 });
  s.addText('April NAM MBR', { x: 0.7, y: 2.3, w: 12, h: 1.6, fontFace: F.head, fontSize: 70, bold: true, color: C.white });
  s.addText('Thomas R Lindsey  ·  05.04.2026  ·  Supermetrics', { x: 0.7, y: 4.0, w: 12, h: 0.4, fontFace: F.body, fontSize: 16, color: 'A0B0D0' });
  s.addShape('roundRect', { x: 0.7, y: 4.7, w: 9.5, h: 0.55, fill: { color: '4A2520' }, line: { color: C.red, width: 1 }, rectRadius: 0.27 });
  s.addText('⚠  Worst revenue month in some time  ·  37% to quota  ·  pipe gen at 103%', { x: 0.85, y: 4.72, w: 9.2, h: 0.5, fontFace: F.body, fontSize: 13, color: C.orange, bold: true });
}

// ===== S2 OVERVIEW =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 2, 'Apr Overview — Results I Planned to Achieve');
  s.addText('REVENUE GOALS', { x: 0.4, y: 1.45, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 3 });
  const card = (sl, x, y, w, h, accent, label, big, sub) => {
    sl.addShape('rect', { x, y, w, h, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    sl.addShape('rect', { x, y, w: 0.08, h, fill: { color: accent }, line: { width: 0 } });
    sl.addText(label, { x: x+0.25, y: y+0.12, w: w-0.4, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, bold: true, charSpacing: 2 });
    sl.addText(big, { x: x+0.25, y: y+0.45, w: w-0.4, h: 0.6, fontFace: F.head, fontSize: 26, bold: true, color: accent });
    sl.addText(sub, { x: x+0.25, y: y+1.05, w: w-0.4, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted });
  };
  card(s, 0.4, 1.8, 6.2, 1.5, C.red, 'APRIL FC (start of month)', '$12,988 / $46,750 = 28%', 'Forecast at start was already short of quota');
  card(s, 6.8, 1.8, 6.2, 1.5, C.red, 'ACTUAL REVENUE', '$17,213 / $46,750 = 37%', 'Worst revenue month in some time. Wins all sub-$5.5k');
  s.addText('PIPELINE GOALS', { x: 0.4, y: 3.55, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 3 });
  card(s, 0.4, 3.9, 6.2, 1.5, C.green, 'APRIL PIPELINE GOAL (start of month)', '$187,000', '4× monthly quota');
  card(s, 6.8, 3.9, 6.2, 1.5, C.green, 'ACTUAL PIPELINE CREATED', '$193,100 / 103%', '23 deals · first month over goal in 2026');
  s.addShape('roundRect', { x: 0.4, y: 5.65, w: 12.55, h: 1.3, fill: { color: '3A2520' }, line: { color: C.red, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The headline: ', options: { bold: true, color: C.red } },
    { text: 'Pipe gen worked. Conversion didn\'t. April delivered the most pipeline of the year and the worst revenue of the year in the same month — that gap is the story of this MBR, and MEDDICC is how we close it.', options: { color: C.text } }
  ], { x: 0.6, y: 5.75, w: 12.15, h: 1.1, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S3 EXECUTED PLAN =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 3, 'April Monthly Plan — What I Committed vs What I Executed');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 4.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('MITs (COMMITTED)', { x: 0.6, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '⚠  TOFU TOFU TOFU — PARTIAL\n', options: { color: C.orange, bold: true, fontSize: 12 } },
    { text: '    333 contacts enrolled vs 400 target (83%) — 153 in HubSpot cadence + ~180 enrolled offline · 23 deals created (103% pipe goal)\n\n', options: { color: C.muted, fontSize: 10 } },
    { text: '✗  MEDDICC / MAPs — FAIL\n', options: { color: C.red, bold: true, fontSize: 12 } },
    { text: '    Docs were not built for $15k+ deals by 04/10. No MAPs in flight.\n', options: { color: C.muted, fontSize: 10 } },
    { text: '    Effective immediately: MAPs non-negotiable on every deal $15k+ — automation already in place.\n\n', options: { color: C.text, fontSize: 10, italic: true } },
    { text: '✗  Disqualification — FAIL\n', options: { color: C.red, bold: true, fontSize: 12 } },
    { text: '    4 fast-fails sub-30d (Speed Sport, Oversight, Fusable, Chefman). MNP died at 425 days.', options: { color: C.muted, fontSize: 10 } }
  ], { x: 0.6, y: 1.9, w: 5.8, h: 3.4, fontFace: F.body });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 4.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('ACTIONS & ACTIVITIES', { x: 7.0, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '⚠  100 contacts/wk (50 Exp / 50 OB)\n', options: { color: C.orange, bold: true, fontSize: 12 } },
    { text: '    ~83 enrolled/wk avg (incl. offline) — close to target; pipe gen beat goal\n\n', options: { color: C.muted, fontSize: 10 } },
    { text: '✓  SUPER Framework / Deal Doc\n', options: { color: C.green, bold: true, fontSize: 12 } },
    { text: '    Bump threshold to $8k+ deals (down from $5k) to save time — automation already covers this today\n\n', options: { color: C.muted, fontSize: 10 } },
    { text: '✓  Daily structure / discipline\n', options: { color: C.green, bold: true, fontSize: 12 } },
    { text: '    21 Call Prep blocks · 16 Cadence Actions · 14 Opp Check-ins', options: { color: C.muted, fontSize: 10 } }
  ], { x: 7.0, y: 1.9, w: 5.8, h: 3.4, fontFace: F.body });

  s.addText('COACHING THEME — NAVIGATING UP', { x: 0.4, y: 5.55, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '✓  Tracer ($100k) — multi-threaded successfully · qualification moving · now considered Qualified\n', options: { color: C.green, bold: true, fontSize: 11 } },
    { text: '✗  Multi-threading with value — did NOT happen on Uniphore ($40k loss)\n', options: { color: C.red, bold: true, fontSize: 11 } },
    { text: '⚠  Making the ask (EB, cell, exec sponsor) — still inconsistent. Single-threaded on too many $10k+ deals.\n', options: { color: C.orange, bold: true, fontSize: 11 } },
    { text: '    Hard next steps + make the ask — every meeting must end with a confirmed next step on calendar AND the explicit ask (EB / cell / exec sponsor / signature path)', options: { color: C.muted, fontSize: 10 } }
  ], { x: 0.4, y: 5.9, w: 12.6, h: 1.2, fontFace: F.body });
}

const col3 = (slide, items) => {
  const w = 4.05, gap = 0.2, startX = 0.4;
  items.forEach((item, i) => {
    const x = startX + i * (w + gap);
    slide.addShape('rect', { x, y: 1.4, w, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    slide.addShape('rect', { x, y: 1.4, w: 0.06, h: 4.5, fill: { color: item.accent }, line: { width: 0 } });
    slide.addText(item.title, { x: x + 0.2, y: 1.55, w: w - 0.3, h: 0.4, fontFace: F.body, fontSize: 11, color: item.accent, bold: true, charSpacing: 2 });
    slide.addText(item.body, { x: x + 0.2, y: 2.0, w: w - 0.3, h: 3.8, fontFace: F.body, fontSize: 10.5 });
  });
};

// ===== S4 REVIEW =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 4, 'April Review');
  col3(s, [
    { accent: C.green, title: 'WHAT I DID WELL', body: [
      { text: '✓  Pipe gen at 103% — 23 deals, $193k\n\n', options: { color: C.text, fontSize: 10.5 } },
      { text: '✓  Mid-market mix improving — ACV creeping up. ', options: { color: C.green, bold: true } },
      { text: '6 deals at $10k+ credit (DTCC, StayTuned, Revibe, Upwork, AMN, DS18, BlueDoor) — meaningful step up from prior months.\n\n', options: { color: C.muted } },
      { text: '✓  MCP/AI positioning landed (Paxos closed in 3 days; Upwork, Dale Howard, IA all MCP plays)\n\n', options: { color: C.text } },
      { text: '✓  Daily structure held — Call Prep, Cadence, Opp Check-ins every business day\n\n', options: { color: C.text } },
      { text: '✓  Account-mgmt saves: Hawaiian backfill same-day, BBC de-escalation, Spectrum CSM handoff, PRH renewal recovery\n\n', options: { color: C.text } },
      { text: '✓  Calligo partner kickoff was a strong session — 5 co-built campaign plays in flight (next slide)', options: { color: C.text } }
    ] },
    { accent: C.red, title: 'WHERE I NEED TO IMPROVE', body: [
      { text: '✗  Uniphore $40k — no EB conversation booked before pricing → lost. ', options: { color: C.red, bold: true } },
      { text: 'This is the MEDDICC headline lesson. Will not repeat.\n\n', options: { color: C.muted } },
      { text: '✗  Disqualification — 4 sub-30d fast-fails entered pipe before fit confirmed\n\n', options: { color: C.text } },
      { text: '✗  Large-deal stalling — 11 Qualified deals worth $74k didn\'t progress\n\n', options: { color: C.text } },
      { text: '✗  Email-to-call ratio 3.5:1 — too email-heavy on $10k+ deals', options: { color: C.text } }
    ] },
    { accent: C.accent, title: 'KEY DEAL LEARNINGS', body: [
      { text: 'Uniphore $40k autopsy: ', options: { bold: true, color: C.text } },
      { text: 'shallow discovery + no EB + 69-day cycle = the headline miss. Bringing 1-pager to 1:1.\n\n', options: { color: C.muted } },
      { text: 'Q1 carryover graveyard cost $85k ', options: { bold: true, color: C.text } },
      { text: 'on Apr 1 alone (Downstream, Equitable, Poco Bero, Tellus, IRIS). No more month-end dispositioning.\n\n', options: { color: C.muted } },
      { text: 'MCP win pattern is repeatable: ', options: { bold: true, color: C.text } },
      { text: 'existing customer + clear use case + AI positioning = sub-30-day close.\n\n', options: { color: C.muted } },
      { text: 'Calligo partnership unlocked ', options: { bold: true, color: C.text } },
      { text: '— strong kickoff w/ 5 campaign plays. Honest read: low confidence in Nick Mishko (slow, not a deal-driver) — I\'ll need to drive the motion.', options: { color: C.muted } }
    ] }
  ]);
}

// ===== S4b CALLIGO =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '4b', 'Calligo Partnership — Kickoff Output');
  s.addText('Strong kickoff session · 5 co-built campaign plays. Partnership is the lever to move OB% from 9% toward 25%+ in May/June.', { x: 0.4, y: 1.3, w: 12.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted });
  const plays = [
    { accent: C.accent, title: '1 · AI ROADMAP CAMPAIGN', body: [
      { text: 'Target: ', options: { bold: true, color: C.text, fontSize: 10 } },
      { text: 'existing customers focused on AI readiness + automation.\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'Offer:\n', options: { color: C.muted, bold: true, fontSize: 9 } },
      { text: '• Half-day AI roadmap session\n• Identify practical AI use cases\n• Prioritize where marketing data supports automation\n• Connect SM pipelines to Calligo\'s automation work\n\n', options: { color: C.muted, fontSize: 9 } },
      { text: '"You already have access to the data. The next question is what workflows can be automated on top of it."', options: { color: C.muted, italic: true, fontSize: 9 } } ] },
    { accent: C.blue, title: '2 · SNOWFLAKE EXPANSION', body: [
      { text: 'Target: ', options: { bold: true, color: C.text, fontSize: 10 } },
      { text: 'SM customers using Snowflake or considering a warehouse.\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'Offer:\n', options: { color: C.muted, bold: true, fontSize: 9 } },
      { text: '• Review current SM setup\n• Identify warehouse / data modeling gaps\n• Calligo builds the staging / modeling layer\n• Expand SM usage into the DWH motion', options: { color: C.muted, fontSize: 9 } } ] },
    { accent: C.green, title: '3 · SPORTS & ENTERTAINMENT', body: [
      { text: 'Target: ', options: { bold: true, color: C.text, fontSize: 10 } },
      { text: 'MLS teams · national sports orgs · ticketing · venue operators · entertainment groups.\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '"Supermetrics centralizes marketing and performance data. Calligo helps structure it into a warehouse and build reporting, AI, and automation layers on top."\n\n', options: { color: C.muted, italic: true, fontSize: 9 } },
      { text: 'Use current opportunities as proof points.', options: { color: C.muted, fontSize: 9 } } ] },
    { accent: C.orange, title: '4 · CUSTOMER RE-ENGAGEMENT', body: [
      { text: 'Target customers with:\n', options: { bold: true, color: C.text, fontSize: 10 } },
      { text: '• Low usage\n• Stalled expansion\n• Known reporting complexity\n• Past lost opportunities\n• Unclear warehouse strategy\n• Heavy manual reporting workflows\n\n', options: { color: C.muted, fontSize: 9 } },
      { text: 'Use Calligo as the reason to restart the conversation.', options: { color: C.muted, fontSize: 9 } } ] },
    { accent: C.accent, title: '5 · LINKEDIN AWARENESS', body: [
      { text: 'Light co-branded LI campaign — post themes:\n\n', options: { bold: true, color: C.text, fontSize: 10 } },
      { text: '• AI readiness for marketing teams\n• Moving beyond dashboards\n• Automating campaign decisions\n• Building a clean marketing data foundation\n• Supermetrics + Calligo customer use cases', options: { color: C.muted, fontSize: 9 } } ] },
    { accent: C.orange, title: '⚠  HONEST READ ON THE COUNTERPART', body: [
      { text: 'Low confidence in Nick Mishko ', options: { bold: true, color: C.text, fontSize: 10 } },
      { text: 'at Calligo — dry, slow, not a go-getter, doesn\'t think outside the box to get deals done.\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'So what: ', options: { bold: true, color: C.muted, fontSize: 9 } },
      { text: 'The plays are good and worth executing. But I\'ll need to drive the motion — set timelines, push meetings, own follow-ups. Don\'t sit on partner-led pace.', options: { color: C.muted, fontSize: 9 } } ] }
  ];
  const cw = 4.05, ch = 2.4, gap = 0.2, sx = 0.4, sy = 1.7;
  plays.forEach((p, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = sx + col * (cw + gap), y = sy + row * (ch + gap);
    s.addShape('rect', { x, y, w: cw, h: ch, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y, w: 0.05, h: ch, fill: { color: p.accent }, line: { width: 0 } });
    s.addText(p.title, { x: x + 0.15, y: y + 0.1, w: cw - 0.25, h: 0.3, fontFace: F.body, fontSize: 10, bold: true, color: p.accent, charSpacing: 1 });
    s.addText(p.body, { x: x + 0.15, y: y + 0.45, w: cw - 0.25, h: ch - 0.55, fontFace: F.body, valign: 'top' });
  });
}

// ===== S5 PIPELINE GEN =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 5, 'April Pipeline Gen Recap');
  const headers = ['Source', 'Count', '% of Pipe', 'Value', '% Value', 'ACV'];
  const rows = [
    ['Inbound', '12', '52%', '$95,000', '49.2%', '$7,917'],
    ['Outbound', '2', '9%', '$18,000', '9.3%', '$9,000'],
    ['Expansion', '9', '39%', '$80,100', '41.5%', '$8,900'],
    ['TOTAL', '23', '100%', '$193,100', '100%', '$8,396']
  ];
  const tbl = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11 } })),
    ...rows.map((r, i) => r.map((c, j) => ({
      text: c,
      options: { bold: i === 3, color: i === 3 ? C.white : C.text, fill: { color: i === 3 ? C.surface2 : C.surface }, fontSize: 12,
        align: j === 0 ? 'left' : (j === 3 || j === 5 ? 'right' : 'center') }
    })))
  ];
  s.addTable(tbl, { x: 0.4, y: 1.5, w: 12.55, colW: [2.2, 1.55, 2.0, 2.5, 2.0, 2.3], rowH: 0.45, fontFace: F.body, border: { type: 'solid', pt: 0.5, color: C.border } });
  s.addShape('rect', { x: 0.4, y: 4.1, w: 12.55, h: 0.5, fill: { color: '0D3D2E' }, line: { color: C.green, width: 0.5 } });
  s.addText([
    { text: 'Pipeline Goal: ', options: { bold: true, color: C.text } },
    { text: '$187,000  ', options: { color: C.text } },
    { text: '·  103%', options: { bold: true, color: C.green } },
    { text: '  ·  first month over goal in 2026', options: { color: C.muted } }
  ], { x: 0.6, y: 4.15, w: 12.2, h: 0.4, fontFace: F.body, fontSize: 13, valign: 'middle' });
  s.addShape('roundRect', { x: 0.4, y: 5.1, w: 12.55, h: 1.6, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'So what: ', options: { bold: true, color: C.orange } },
    { text: 'Pipeline machine is working — but mix is heavily Inbound + Expansion (~91%). ', options: { color: C.text } },
    { text: 'Outbound contributed just ~9%', options: { bold: true, color: C.text } },
    { text: ' — the May daily TOFU discipline (slides 9–11) is where Outbound % needs to climb.\n\n', options: { color: C.text } },
    { text: 'Note: source split is approximate pending HubSpot validation.', options: { italic: true, color: C.muted, fontSize: 11 } }
  ], { x: 0.6, y: 5.2, w: 12.2, h: 1.4, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S6 YTD =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 6, 'YTD 2026 Pipeline Trend');
  const headers = ['Month', 'Deals', 'Pipe Value', 'Goal', '% to Goal', 'ACV'];
  const rows = [
    ['January', '9', '$80,572', '$160,824', '50%', '$8,952', C.red],
    ['February', '18', '$145,292', '$160,824', '90%', '$8,072', C.orange],
    ['March', '16', '$155,853', '$176,000', '89%', '$9,741', C.orange],
    ['April', '23', '$193,100', '$187,000', '103%', '$8,396', C.green]
  ];
  const tbl = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11 } })),
    ...rows.map((r, i) => r.slice(0, 6).map((c, j) => ({
      text: c, options: { bold: i === 3 || j === 4, color: j === 4 ? r[6] : (i === 3 ? C.white : C.text),
        fill: { color: i === 3 ? '0D3D2E' : C.surface }, fontSize: 13,
        align: j === 0 ? 'left' : (j === 2 || j === 3 || j === 5 ? 'right' : 'center') }
    })))
  ];
  s.addTable(tbl, { x: 0.4, y: 1.5, w: 12.55, colW: [2.5, 1.5, 2.5, 2.5, 1.8, 1.75], rowH: 0.5, fontFace: F.body, border: { type: 'solid', pt: 0.5, color: C.border } });
  s.addShape('roundRect', { x: 0.4, y: 5.0, w: 12.55, h: 1.6, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Trend: ', options: { bold: true, color: C.accent } },
    { text: '50% → 90% → 89% → ', options: { color: C.text } },
    { text: '103%', options: { bold: true, color: C.green } },
    { text: '. Pipeline gen is on a clean upward arc. The arc is real — what\'s missing is the conversion arc behind it.', options: { color: C.text } }
  ], { x: 0.6, y: 5.15, w: 12.2, h: 1.3, fontFace: F.body, fontSize: 14, valign: 'middle' });
}

// ===== S7 EXECUTION =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 7, 'April Closed Deal Execution — The Real Story');
  const stats = [
    { label: 'WIN RATE', val: '24%', sub: '5 won / 21 closed', color: C.orange },
    { label: 'CLOSED WON', val: '$17.2k', sub: 'all sub-$5.5k', color: C.green },
    { label: 'CLOSED LOST (Y1 ACV)', val: '$177k', sub: '16 deals · 10× the wins', color: C.red },
    { label: 'NET CLOSED ACTIVITY', val: '−$160k', sub: 'Won minus lost', color: C.red }
  ];
  stats.forEach((st, i) => {
    const x = 0.4 + i * 3.2;
    s.addShape('rect', { x, y: 1.4, w: 3.0, h: 1.3, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addText(st.label, { x: x + 0.15, y: 1.5, w: 2.8, h: 0.3, fontFace: F.body, fontSize: 9, color: C.muted, bold: true, charSpacing: 1 });
    s.addText(st.val, { x: x + 0.15, y: 1.8, w: 2.8, h: 0.6, fontFace: F.head, fontSize: 28, bold: true, color: st.color });
    s.addText(st.sub, { x: x + 0.15, y: 2.4, w: 2.8, h: 0.25, fontFace: F.body, fontSize: 9, color: C.muted });
  });
  s.addText('3 STORIES INSIDE THE LOSSES', { x: 0.4, y: 2.95, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  const tbl = [
    [{ text: 'Story', options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11 } },
     { text: 'Deals', options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11 } },
     { text: 'Lost ACV', options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11, align: 'right' } },
     { text: 'So what', options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11 } }],
    [{ text: 'Q1 carryover cleanup (Apr 1)', options: { bold: true, color: C.text, fontSize: 11, fill: { color: C.surface } } },
     { text: 'Downstream (390d), Equitable (112d), Poco Bero, Tellus, IRIS', options: { color: C.text, fontSize: 10, fill: { color: C.surface } } },
     { text: '$85,200', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } },
     { text: 'Should have died in March. Inflated forecast, masked stalls.', options: { color: C.muted, fontSize: 10, fill: { color: C.surface } } }],
    [{ text: 'Mid-cycle big-deal loss', options: { bold: true, color: C.text, fontSize: 11, fill: { color: C.surface } } },
     { text: 'Uniphore ($40k, 69d), Comcast (137d), MNP (425d)', options: { color: C.text, fontSize: 10, fill: { color: C.surface } } },
     { text: '$60,000', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } },
     { text: 'Uniphore is the coachable miss. MEDDICC red flag: no EB.', options: { color: C.muted, fontSize: 10, fill: { color: C.surface } } }],
    [{ text: 'Fast disqualification failures', options: { bold: true, color: C.text, fontSize: 11, fill: { color: C.surface } } },
     { text: 'Speed Sport (11d), Oversight (11d), Fusable (24d), Chefman (26d)', options: { color: C.text, fontSize: 10, fill: { color: C.surface } } },
     { text: '$21,900', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } },
     { text: 'Got into pipe before fit confirmed. MEDDICC card at Day 1 prevents this.', options: { color: C.muted, fontSize: 10, fill: { color: C.surface } } }]
  ];
  s.addTable(tbl, { x: 0.4, y: 3.3, w: 12.55, colW: [2.8, 4.0, 1.5, 4.25], rowH: 0.6, fontFace: F.body, border: { type: 'solid', pt: 0.5, color: C.border } });
  s.addShape('roundRect', { x: 0.4, y: 6.2, w: 12.55, h: 0.85, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'April wins were all small + fast (Paxos 3d, DQ 29d, Cousin Labs 28d). ', options: { bold: true, color: C.orange } },
    { text: 'Zero $10k+ wins. Every $10k+ deal that closed in April closed lost. May focus: protect the $10k+ pipe with MEDDICC.', options: { color: C.text } }
  ], { x: 0.6, y: 6.27, w: 12.2, h: 0.7, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S8 MITs =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 8, 'May MITs');
  const mits = [
    { accent: C.accent, num: '1', title: 'MEDDICC OPERATIONALIZED', body: [
      { text: '80% of Stage 2+ deals scored. Weakest dimension named per deal. Doc artifacts produced (scorecard, value narrative, MAP, exec summary). MAPs ', options: { color: C.text, fontSize: 13 } },
      { text: 'non-negotiable', options: { bold: true, color: C.text, fontSize: 13 } },
      { text: ' on every $15k+ deal · SUPER framework on every $8k+ deal — both already automated.\n\n', options: { color: C.text, fontSize: 13 } },
      { text: 'The headline MIT — April proved deal management is the gap. Disqualification-at-Day-14 lives inside this (MEDDICC card at Day 1 = the gate; 60-day rule = the floor).', options: { italic: true, color: C.muted, fontSize: 11 } } ] },
    { accent: C.blue, num: '2', title: 'TOFU AT THE ACTIVITY LEVEL', body: [
      { text: 'Daily and measurable, not a vibe. Per-day enrollment, emails, dials, LinkedIn, net-new accounts, booked meetings — all tracked. See slides 9–11 for the math and the playbook.\n\n', options: { color: C.text, fontSize: 13 } },
      { text: 'April hit 103% pipe gen at ~83% of contact-enrollment target. May tightens both — but with realistic numbers, not aspirational ones.', options: { italic: true, color: C.muted, fontSize: 11 } } ] }
  ];
  mits.forEach((m, i) => {
    const x = 0.4 + i * 6.4;
    s.addShape('rect', { x, y: 1.5, w: 6.2, h: 5.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y: 1.5, w: 0.08, h: 5.0, fill: { color: m.accent }, line: { width: 0 } });
    s.addShape('ellipse', { x: x + 0.35, y: 1.75, w: 0.8, h: 0.8, fill: { color: m.accent }, line: { width: 0 } });
    s.addText(m.num, { x: x + 0.35, y: 1.78, w: 0.8, h: 0.75, fontFace: F.head, fontSize: 32, bold: true, color: C.white, align: 'center' });
    s.addText(m.title, { x: x + 1.3, y: 1.85, w: 4.7, h: 0.6, fontFace: F.head, fontSize: 17, bold: true, color: m.accent, charSpacing: 1, valign: 'middle' });
    s.addText(m.body, { x: x + 0.35, y: 2.8, w: 5.65, h: 3.5, fontFace: F.body, valign: 'top' });
  });
}

// ===== S9 FUNNEL MATH =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 9, 'May Funnel Math — Top-Down Derivation');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('LOCKED INPUTS', { x: 0.6, y: 1.5, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '• May quota: ', options: { color: C.text, fontSize: 12 } },
    { text: '$46,750\n', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '• 4× pipeline goal: ', options: { color: C.text, fontSize: 12 } },
    { text: '$187,000\n', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '• Plan target: ', options: { color: C.text, fontSize: 12 } },
    { text: '$200,000', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: ' (+7% overage for slippage)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• ACV baseline: closed-won avg last 6 months ≈ ', options: { color: C.text, fontSize: 12 } },
    { text: '$11,000', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '\n   (better anchor than April\'s depressed $2.9k closed-won ACV)', options: { italic: true, color: C.muted, fontSize: 10 } }
  ], { x: 0.6, y: 1.85, w: 5.8, h: 1.7, fontFace: F.body });
  s.addText('BACKED-OUT TARGETS', { x: 0.6, y: 3.7, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '• Deals needed: ', options: { color: C.text, fontSize: 12 } },
    { text: '~18', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: ' ($200k ÷ $11k ACV)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Booked meetings: ', options: { color: C.text, fontSize: 12 } },
    { text: '~60', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: ' (~33% mtg → deal · slightly below April\'s 72)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Contacts enrolled: ', options: { color: C.text, fontSize: 12 } },
    { text: '~330', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: ' (matches April\'s true 333 incl. offline)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Account base: ', options: { color: C.text, fontSize: 12 } },
    { text: '~70 net-new', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: ' (lift OB% from 9% toward 25%)', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 4.05, w: 5.8, h: 1.8, fontFace: F.body });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('APRIL BASELINE → MAY TARGET', { x: 7.0, y: 1.5, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  const dataRows = [
    ['Contacts enrolled (incl. offline)', '333', '330', 'flat', C.text],
    ['Calls logged', '280', '300', '+7%', C.green],
    ['Emails sent', '979', '1,000', 'flat', C.text],
    ['Meetings', '72', '60', '−17%', C.orange],
    ['New deals', '23', '~18', 'fewer · larger', C.orange],
    ['Pipe generated', '$193k', '$200k', '+4%', C.green]
  ].map((r, i) => [
    { text: r[0], options: { color: C.text, fontSize: 11, bold: i === 5, fill: { color: C.surface2 } } },
    { text: r[1], options: { color: C.text, fontSize: 11, align: 'right', bold: i === 5, fill: { color: C.surface2 } } },
    { text: r[2], options: { color: C.text, fontSize: 11, align: 'right', bold: i === 5, fill: { color: C.surface2 } } },
    { text: r[3], options: { color: r[4], fontSize: 11, align: 'right', bold: true, fill: { color: C.surface2 } } }
  ]);
  const tblData = [
    [{ text: 'Activity', options: { bold: true, color: C.muted, fontSize: 10, fill: { color: C.surface } } },
     { text: 'Apr', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } },
     { text: 'May', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } },
     { text: 'Δ', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } }],
    ...dataRows
  ];
  s.addTable(tblData, { x: 7.0, y: 1.85, w: 5.8, colW: [2.6, 0.9, 0.9, 1.4], rowH: 0.4, fontFace: F.body, border: { type: 'solid', pt: 0.5, color: C.border } });
  s.addText('Realistic over aspirational. 25 deals + 80 mtgs was a hustle in April — not repeatable as the plan. Shift focus: fewer, larger, mid-market deals at $11k+ ACV.', { x: 7.0, y: 5.1, w: 5.8, h: 0.7, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
  s.addShape('roundRect', { x: 0.4, y: 6.1, w: 12.55, h: 0.85, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The shift: ', options: { bold: true, color: C.accent } },
    { text: 'April proved volume is doable. May proves quality is doable. Same pipe goal, fewer deals, higher ACV — and MEDDICC discipline behind every $10k+ deal so they actually close.', options: { color: C.text } }
  ], { x: 0.6, y: 6.18, w: 12.2, h: 0.7, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S10 DAILY ACTIVITY (UPDATED — 18 working days) =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 10, 'May Daily Activity Targets');
  s.addText('18 working days in May (21 minus 3 PTO). Same monthly targets, consolidated into days actually worked. If the daily numbers hit, the monthly goal is mathematical.', { x: 0.4, y: 1.3, w: 12.6, h: 0.4, fontFace: F.body, fontSize: 11, color: C.muted });
  const headers = ['Activity', 'Monthly', 'Weekly (avg)', 'Daily (18d)', 'Notes'];
  const rows = [
    ['New contacts enrolled', '330', '~83', '~18/day', '~9 OB + ~9 Exp · email-verified before enroll', true],
    ['Emails sent (cadence steps)', '1,000', '~250', '~56/day', 'Cadence queue worked AM', false],
    ['Calls / dials', '300', '~75', '~17/day', 'Live conversations on $10k+ deals — replace email', false],
    ['LinkedIn touches', '~210', '~52', '~12/day', '5 connect + 3 InMail + 3 comments', false],
    ['Net-new accounts worked', '~70', '~17', '~4/day', 'Mon list-build + daily research/contact-pull', false],
    ['Booked meetings', '60', '~15', '~3-4/day', 'Down from April\'s 72 — quality over volume', false],
    ['Net-new deals created', '~18', '~4-5', '~1/day', 'Fewer, larger · target $11k+ ACV mid-market', true]
  ];
  const tblD = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fontSize: 11, fill: { color: C.surface2 }, align: h === 'Activity' || h === 'Notes' ? 'left' : 'right' } })),
    ...rows.map(r => [
      { text: r[0], options: { bold: r[5], color: C.text, fontSize: 12, fill: { color: r[5] ? C.surface2 : C.surface } } },
      { text: r[1], options: { color: C.text, fontSize: 12, align: 'right', fill: { color: r[5] ? C.surface2 : C.surface } } },
      { text: r[2], options: { color: C.text, fontSize: 12, align: 'right', fill: { color: r[5] ? C.surface2 : C.surface } } },
      { text: r[3], options: { color: C.green, fontSize: 12, align: 'right', bold: true, fill: { color: r[5] ? C.surface2 : C.surface } } },
      { text: r[4], options: { color: C.muted, fontSize: 10, fill: { color: r[5] ? C.surface2 : C.surface } } }
    ])
  ];
  s.addTable(tblD, { x: 0.4, y: 1.8, w: 12.55, colW: [3.0, 1.5, 1.4, 1.6, 5.05], rowH: 0.45, fontFace: F.body, border: { type: 'solid', pt: 0.5, color: C.border } });
  s.addShape('roundRect', { x: 0.4, y: 5.95, w: 12.55, h: 1.0, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The "no exceptions" lines: ', options: { bold: true, color: C.accent } },
    { text: '18 contacts enrolled/day + 4 net-new accounts/day on every working day. ', options: { color: C.text } },
    { text: '3 PTO days × the daily target = the catch-up that has to happen on either side of time off — not slipped to next month.', options: { bold: true, color: C.text } }
  ], { x: 0.6, y: 6.05, w: 12.2, h: 0.85, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S11 PLAYBOOK =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 11, 'How We Work the Activities — Calendar Blocks, Not Vibes');
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 12.55, h: 0.9, fill: { color: '231D3D' }, line: { color: C.accent, width: 1.5 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Commit: ', options: { bold: true, color: C.accent, fontSize: 13 } },
    { text: 'Rework my calendar in ', options: { color: C.text, fontSize: 13 } },
    { text: 'Reclaim', options: { bold: true, color: C.text, fontSize: 13 } },
    { text: ' by Mon 5/5 to enforce the daily output for each week. Every block below becomes a recurring Reclaim task w/ a habit window — auto-rescheduled around customer mtgs but never dropped. The calendar is the contract.', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 1.4, w: 12.2, h: 0.75, fontFace: F.body, valign: 'middle' });
  const blocks = [
    ['Account list build', 'Mondays · 60 min', 'Pull 20 net-new accounts: ZoomInfo + intent signals + 2025 closed-lost rework. Drop into HubSpot with vertical tag + assigned cadence.'],
    ['Contact enrollment', 'Daily AM · 30 min', '~18 contacts/day · ~9 OB + ~9 Exp (consolidated for 18 working days). Verify email in Sales Nav before enroll (April BRP had >50% bounce).'],
    ['Cadence actions', 'Daily AM · 60 min', 'Work the queue first thing: ~56 emails + ~17 dials due today.'],
    ['LinkedIn block', 'Daily midday · 20 min', '5 connect requests + 3 InMails + 3 thoughtful comments on target-account posts. Tagged as activities in HubSpot.'],
    ['Mid-funnel block', 'Daily PM · 60 min', 'One live touch on every Qualified / POV deal stagnant 7+ days. No more pure-email follow-ups on $10k+ deals.'],
    ['Friday hygiene block', 'Weekly · 90 min', 'Disposition every deal with zero activity in 14 days. Update MEDDICC card on every Stage 2+ deal. No more April-1 cleanup days.'],
    ['Closed-lost rework', 'Tuesdays · 45 min', '5 accounts/wk from 2025 closed-lost · restart cadence with new angle (usually MCP/AI).']
  ];
  const blkH = 0.62, blkY = 2.4;
  blocks.forEach((b, i) => {
    const y = blkY + i * blkH;
    s.addText([
      { text: '■  ', options: { color: C.accent, fontSize: 14 } },
      { text: b[0], options: { bold: true, color: C.text, fontSize: 13 } },
      { text: '   —   ', options: { color: C.muted, fontSize: 12 } },
      { text: b[1], options: { color: C.blue, fontSize: 11, bold: true } }
    ], { x: 0.4, y, w: 12.6, h: 0.3, fontFace: F.body });
    s.addText(b[2], { x: 0.7, y: y + 0.3, w: 12.3, h: 0.32, fontFace: F.body, fontSize: 10, color: C.muted });
  });
}

// ===== S12 TRUE PIPELINE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 12, 'May True Pipeline — $168k Open · 16 Deals');
  const headers = ['Account', 'Opportunity', 'ERR', 'Stage', '% Close', 'Weighted'];
  const dealRows = [
    ['Scotiabank', 'MIC | KCG | $40k', '$39,000', 'Proposal', '82%', '$32,000', C.orange, true],
    ['AMBSE', 'MIC | $15k', '$15,000', 'Proof of Value', '50%', '$7,500', C.accent, true],
    ['DTCC', 'MIC? | $25k', '$25,000', 'Qualified', '30%', '$7,500', C.blue, false],
    ['StayTuned', 'DACT | $20k', '$20,000', 'Qualification', '20%', '$4,000', C.blue, false],
    ['Revibe', 'MIC | $20k', '$20,000', 'Proof of Value', '35%', '$7,000', C.accent, false],
    ['Upwork', 'MIC/MCP | $15k', '$15,000', 'Proof of Value', '40%', '$6,000', C.accent, false],
    ['Grafton', 'DWH | $15k', '$15,000', 'Qualified', '25%', '$3,750', C.blue, false],
    ['AMN', 'SNO | $12k', '$12,000', 'Qualified', '30%', '$3,600', C.blue, false],
    ['DS18', 'MIC | $12k', '$12,000', 'Qualified', '25%', '$3,000', C.blue, false],
    ['Future Builds', 'API | $12k', '$12,000', 'Proposal', '40%', '$4,800', C.orange, false],
    ['BlueDoor', 'MIC | $10k', '$10,000', 'Proof of Value', '45%', '$4,500', C.accent, false],
    ['WELL', 'MIC | $4k', '$4,000', 'Qualified', '35%', '$1,400', C.blue, false],
    ['Furnished Finder', 'MCP | $4k', '$3,800', 'Qualified', '75%', '$2,850', C.blue, false],
    ['Other (3)', 'Growth/MCP/Starter', '$6,000', '—', '—', '$1,500', C.muted, false],
    ['TOTAL', '16 deals', '$208,800', '—', '~42%', '$89,400', C.green, true]
  ];
  const tbl = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fontSize: 10, fill: { color: C.surface2 }, align: h === 'Account' || h === 'Opportunity' || h === 'Stage' ? 'left' : 'right' } })),
    ...dealRows.map((r, i) => [
      { text: r[0], options: { bold: r[7], color: C.text, fontSize: 11, fill: { color: i === dealRows.length - 1 ? '0D3D2E' : C.surface } } },
      { text: r[1], options: { color: C.text, fontSize: 11, fill: { color: i === dealRows.length - 1 ? '0D3D2E' : C.surface } } },
      { text: r[2], options: { color: C.text, fontSize: 11, align: 'right', fill: { color: i === dealRows.length - 1 ? '0D3D2E' : C.surface } } },
      { text: r[3], options: { color: r[6], fontSize: 10, bold: true, fill: { color: i === dealRows.length - 1 ? '0D3D2E' : C.surface } } },
      { text: r[4], options: { color: C.text, fontSize: 11, align: 'right', fill: { color: i === dealRows.length - 1 ? '0D3D2E' : C.surface } } },
      { text: r[5], options: { color: i === dealRows.length - 1 ? C.green : C.text, bold: r[7], fontSize: 11, align: 'right', fill: { color: i === dealRows.length - 1 ? '0D3D2E' : C.surface } } }
    ])
  ];
  s.addTable(tbl, { x: 0.4, y: 1.4, w: 12.55, colW: [2.0, 2.5, 1.4, 2.4, 1.4, 2.85], rowH: 0.34, fontFace: F.body, border: { type: 'solid', pt: 0.4, color: C.border } });
}

// ===== S13 FORECAST =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 13, 'May Forecast');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 8.0, h: 5.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('FORECAST (HIGH-CONFIDENCE)', { x: 0.6, y: 1.5, w: 7.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.green, bold: true, charSpacing: 2 });
  s.addText([
    { text: '■  Scotiabank — $32k weighted (82%) ', options: { bold: true, color: C.green, fontSize: 12 } },
    { text: '· final security review · wider-team call 5/15 · pushing security/procurement call this week or next\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '■  Furnished Finder — $2.9k (75%) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '· MCP closeout\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  WELL — $1.4k (35%)', options: { bold: true, color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 1.85, w: 7.6, h: 1.5, fontFace: F.body });
  s.addText('SWING DEALS', { x: 0.6, y: 3.45, w: 7.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.orange, bold: true, charSpacing: 2 });
  s.addText([
    { text: '■  AMBSE ($15k) — main swing ', options: { bold: true, color: C.accent, fontSize: 12 } },
    { text: '· POV active · MEDDICC focus this week\n', options: { color: C.text, fontSize: 11 } },
    { text: '■  Revibe ($20k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— POV active · pain re-discovery needed\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  StayTuned ($20k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— MEDDICC = build value narrative + ROI by 5/12\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  Upwork ($15k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— POV active · multi-thread / champion test\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  BlueDoor ($10k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— POV · pulling June go-live forward\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  Future Builds ($12k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— Proposal · wild card\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  AMN / Grafton / DS18 ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— $10k+ Qualified pool', options: { color: C.muted, fontSize: 11 } }
  ], { x: 0.6, y: 3.8, w: 7.6, h: 3.0, fontFace: F.body });

  s.addShape('rect', { x: 8.6, y: 1.4, w: 4.4, h: 5.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('FORECAST / OPEN PIPELINE', { x: 8.8, y: 1.5, w: 4.0, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, bold: true, charSpacing: 2 });
  const fcRows = [
    ['Total Quota', '$46,750', false, C.text],
    ['Confirmed FC', '$44,800', true, C.green],
    ['Weighted FC', '$74,400', true, C.green],
    ['Open Pipeline', '$168,000', false, C.text],
    ['Weighted Open', '$74,400', false, C.text],
    ['Gap to Quota', '−$27,700', true, C.green],
    ['# Open Deals', '16', false, C.text],
    ['Closed Won (MTD)', '$0', false, C.text],
    ['Closed Lost (MTD)', '$0', false, C.text]
  ];
  const fcT = fcRows.map(r => [
    { text: r[0], options: { color: C.text, fontSize: 11, bold: r[2], fill: { color: C.surface } } },
    { text: r[1], options: { color: r[3], fontSize: 11, align: 'right', bold: r[2], fill: { color: C.surface } } }
  ]);
  s.addTable(fcT, { x: 8.8, y: 1.85, w: 4.0, colW: [2.4, 1.6], rowH: 0.36, fontFace: F.body, border: { type: 'solid', pt: 0.4, color: C.border } });
  s.addText('Quota covered on paper from Day 1. Coverage doesn\'t equal cash — MEDDICC is what converts coverage into revenue.', { x: 8.8, y: 5.4, w: 4.0, h: 1.3, fontFace: F.body, fontSize: 10, color: C.muted, italic: true, valign: 'top' });
}

// ===== S13b MEDDICC =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '13b', 'MEDDICC — How May Deals Get Driven');
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 12.55, h: 0.9, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'April had pipe but couldn\'t close. ', options: { color: C.text } },
    { text: 'The conversion lever is MEDDICC discipline', options: { bold: true, color: C.orange } },
    { text: ' — operationalize it, don\'t just name-check it. This is the bridge between "pipeline at 103%" and "revenue at 37%": deal management was the missing piece, MEDDICC is how we fix it.', options: { color: C.text } }
  ], { x: 0.6, y: 1.4, w: 12.2, h: 0.75, fontFace: F.body, fontSize: 12, valign: 'middle' });
  const cols = [
    { title: 'THE 8 DIMENSIONS — IN A SUPERMETRICS DEAL', body: [
      { text: 'M', options: { bold: true, color: C.accent, fontSize: 11 } },
      { text: 'etrics — quantified value: "saves 10 hrs/wk", "$X attributable revenue"\n', options: { color: C.text, fontSize: 10 } },
      { text: 'E', options: { bold: true, color: C.accent, fontSize: 11 } },
      { text: 'conomic Buyer — confirmed name + budget authority + direct conversation booked\n', options: { color: C.text, fontSize: 10 } },
      { text: 'D', options: { bold: true, color: C.accent, fontSize: 11 } },
      { text: 'ecision Criteria — written list of what they need to see to say yes\n', options: { color: C.text, fontSize: 10 } },
      { text: 'D', options: { bold: true, color: C.accent, fontSize: 11 } },
      { text: 'ecision Process — steps from now → contract, with owners + dates\n', options: { color: C.text, fontSize: 10 } },
      { text: 'P', options: { bold: true, color: C.accent, fontSize: 11 } },
      { text: 'aper Process — procurement, legal, security, finance routes mapped\n', options: { color: C.text, fontSize: 10 } },
      { text: 'I', options: { bold: true, color: C.accent, fontSize: 11 } },
      { text: 'dentify Pain — quantified, urgent, owned, top-3 priority\n', options: { color: C.text, fontSize: 10 } },
      { text: 'C', options: { bold: true, color: C.accent, fontSize: 11 } },
      { text: 'hampion — willing to sell internally, has capital, tested\n', options: { color: C.text, fontSize: 10 } },
      { text: 'C', options: { bold: true, color: C.accent, fontSize: 11 } },
      { text: 'ompetition — incumbent or alternatives named; "do nothing" included', options: { color: C.text, fontSize: 10 } } ] },
    { title: 'DOCUMENTS WE PRODUCE PER DEAL', body: [
      { text: 'MEDDICC Scorecard ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: '— HubSpot deal field or Google Doc · R/Y/G on each of 8 dims · updated every 14 days\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Value Narrative (1-page) ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: '— pain → impact → metrics → recommendation. Coaching Goal 1.\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Mutual Action Plan (MAP) ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: '— shared with the buyer · dates + owners through close · re-introduce after March\'s MSP failure.\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Executive Summary ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: '— for $20k+ deals · the vehicle for the EB conversation · Coaching Goal 3.', options: { color: C.muted, fontSize: 10 } } ] },
    { title: 'WHAT CHANGES IN CONVERSATIONS', body: [
      { text: 'Discovery call: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'must surface 1 quantified Metric + named Pain owner.\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '2nd call: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Champion test ("would you bring this to your VP?") + Decision Criteria asked explicitly.\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Proposal stage: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'MAP shared as a working doc, not after the fact.\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Every $10k+ deal: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'direct EB conversation booked before sending pricing. April lost Uniphore $40k partly because we never met the EB.', options: { color: C.muted, fontSize: 10 } } ] }
  ];
  cols.forEach((col, i) => {
    const x = 0.4 + i * 4.25, y = 2.4;
    s.addShape('rect', { x, y, w: 4.05, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addText(col.title, { x: x + 0.15, y: y + 0.12, w: 3.8, h: 0.5, fontFace: F.body, fontSize: 10, color: C.accent, bold: true, charSpacing: 1 });
    s.addText(col.body, { x: x + 0.15, y: y + 0.65, w: 3.8, h: 3.7, fontFace: F.body, valign: 'top' });
  });
}

// ===== S13c CADENCE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '13c', 'MEDDICC Scorecard Cadence — Operational Rhythm');
  const tbl = [
    [{ text: 'When', options: { bold: true, color: C.muted, fontSize: 11, fill: { color: C.surface2 } } },
     { text: 'What', options: { bold: true, color: C.muted, fontSize: 11, fill: { color: C.surface2 } } }],
    ...[
      ['Day 1 of any new deal', 'Open MEDDICC card · rate every dimension R/Y/G'],
      ['Every 14 days', 'Update card · flag any dimension that hasn\'t moved'],
      ['Weekly 1:1 with Andrew', 'Walk through 1 deal card dimension by dimension'],
      ['Stage progression gate', 'Cannot move to Proposal without ≥6/8 green-or-yellow'],
      ['60-day rule', '$10k+ deal at 60d with reds in EB / Pain / Decision Criteria → disqualify']
    ].map(r => [
      { text: r[0], options: { bold: true, color: C.text, fontSize: 13, fill: { color: C.surface } } },
      { text: r[1], options: { color: C.text, fontSize: 13, fill: { color: C.surface } } }
    ])
  ];
  s.addTable(tbl, { x: 0.4, y: 1.5, w: 12.55, colW: [4.0, 8.55], rowH: 0.65, fontFace: F.body, border: { type: 'solid', pt: 0.5, color: C.border } });
  s.addShape('roundRect', { x: 0.4, y: 5.6, w: 12.55, h: 1.0, fill: { color: '231D3D' }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Target by 5/31: ', options: { bold: true, color: C.accent } },
    { text: '80% of Stage 2+ deals have a completed MEDDICC scorecard with no unaddressed reds. Matches the SMART metric in the 30-day coaching plan (Goal 2).', options: { color: C.text } }
  ], { x: 0.6, y: 5.7, w: 12.2, h: 0.85, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S13d PER DEAL MEDDICC =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '13d', 'May MEDDICC Targets — Per Active Deal');
  s.addText('Each major open deal · which dimension is the red-to-fix · the May action.', { x: 0.4, y: 1.3, w: 12.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted });
  const headers = ['Deal', 'ERR', 'Weakest Dimension', 'May Action', 'Due'];
  const rows = [
    ['Scotiabank', '$39k', 'Decision Process (procurement)', 'Lock signature path with King · working session Wk 1', '5/9', C.orange],
    ['DTCC', '$25k', 'Economic Buyer', 'Identify + book EB intro · build Exec Summary', '5/9', C.red],
    ['StayTuned', '$20k', 'Metrics', 'Value narrative + ROI calc · share back to champion', '5/12', C.orange],
    ['Revibe', '$20k', 'Pain (urgency)', 'Re-discover · what\'s the trigger? Without urgency, disqualify', '5/15', C.red],
    ['Upwork', '$15k', 'Champion', 'Test champion w/ "bring to VP" ask · multi-thread', '5/15', C.orange],
    ['Grafton', '$15k', 'Decision Criteria', 'Get written list of must-haves this week', '5/9', C.orange],
    ['AMN', '$12k', 'Decision Criteria', 'Get written list · verify SNO scope', '5/15', C.orange],
    ['DS18', '$12k', 'Economic Buyer', 'EB intro from current champion', '5/16', C.red],
    ['BlueDoor', '$10k', 'Pain', 'Quantify pain or disqualify by 60d', '5/22', C.orange]
  ];
  const tbl = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fontSize: 10, fill: { color: C.surface2 }, align: h === 'ERR' ? 'right' : 'left' } })),
    ...rows.map(r => [
      { text: r[0], options: { bold: true, color: C.text, fontSize: 11, fill: { color: C.surface } } },
      { text: r[1], options: { color: C.text, fontSize: 11, align: 'right', fill: { color: C.surface } } },
      { text: r[2], options: { color: r[5], fontSize: 11, bold: true, fill: { color: C.surface } } },
      { text: r[3], options: { color: C.text, fontSize: 10, fill: { color: C.surface } } },
      { text: r[4], options: { color: C.text, fontSize: 11, bold: true, fill: { color: C.surface } } }
    ])
  ];
  s.addTable(tbl, { x: 0.4, y: 1.7, w: 12.55, colW: [1.8, 1.0, 3.3, 5.55, 0.9], rowH: 0.42, fontFace: F.body, border: { type: 'solid', pt: 0.4, color: C.border } });
  s.addShape('roundRect', { x: 0.4, y: 6.0, w: 12.55, h: 0.85, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Pattern: ', options: { bold: true, color: C.orange } },
    { text: '4 of 9 are weak on EB or Decision Criteria — same gaps that killed Uniphore. Fix these in May or repeat April in May.', options: { color: C.text } }
  ], { x: 0.6, y: 6.08, w: 12.2, h: 0.7, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S14 TOP DEAL CLOSING =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 14, 'Top Deal Closing Plans');
  const deals = [
    { accent: C.green, badge: 'FORECAST', title: 'Scotiabank — $39k · Proposal · 82%', sub: 'Target Close: 05/29/2026 · timeline looks OK', body: [
      { text: 'Stage: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Final security review\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'Wider-team call locked for 5/15\n', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'This week / next: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'trying to set up direct call with security + procurement team for clear line to close by EOM\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'MEDDICC red: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Decision Process — procurement path\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'Top action: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Lock the security/procurement working session', options: { color: C.muted, fontSize: 11 } } ] },
    { accent: C.accent, badge: 'MAIN SWING', title: 'AMBSE — $15k · POV · 50%', sub: 'Target Close: 05/29/2026', body: [
      { text: 'Stage: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Proof of Value active\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'MEDDICC red: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'EB / Decision Criteria — confirm both this week\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'Top action: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Drive POV success criteria → proposal · book EB intro · MAP shared by 5/12\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'Timeline: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: '5/9 POV checkpoint · 5/13 EB conversation · 5/16 proposal · 5/22 agreement · 5/29 close', options: { color: C.muted, fontSize: 11 } } ] },
    { accent: C.orange, badge: 'SWING / PULL-FORWARD', title: 'BlueDoor — $10k · POV · 45%', sub: 'Buyer\'s natural close: June · Pull-forward target: 05/30', body: [
      { text: 'Stage: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'POV active · looking to move forward in June\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'Strategy: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Pull contract signature into May for May credit + still hit June go-live\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'MEDDICC red: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Pain (urgency to act in May)\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'Top action: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Make the ask — "any reason we can\'t paper this in May to lock the go-live and EOM incentive?"', options: { color: C.muted, fontSize: 11 } } ] },
    { accent: C.blue, badge: 'WILD CARD', title: 'Future Builds — $12k · Proposal · 40%', sub: 'Target Close: 05/29/2026 (stretch)', body: [
      { text: 'Stage: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Proposal · API motion\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'MEDDICC red: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Decision Process — verify signer + paper path\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'Top action: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Confirm signer + paper path this week · push for signature in May\n', options: { color: C.muted, fontSize: 11 } },
      { text: 'Why wild card: ', options: { bold: true, color: C.text, fontSize: 11 } },
      { text: 'Could close in May if paper moves fast; otherwise slips to June', options: { color: C.muted, fontSize: 11 } } ] }
  ];
  const cw = 6.2, ch = 2.4, gap = 0.15;
  deals.forEach((d, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.4 + col * (cw + gap), y = 1.4 + row * (ch + gap);
    s.addShape('rect', { x, y, w: cw, h: ch, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y, w: 0.06, h: ch, fill: { color: d.accent }, line: { width: 0 } });
    s.addText(d.title, { x: x + 0.18, y: y + 0.1, w: cw - 1.6, h: 0.35, fontFace: F.body, fontSize: 12, bold: true, color: d.accent });
    s.addShape('roundRect', { x: x + cw - 1.5, y: y + 0.12, w: 1.35, h: 0.3, fill: { color: d.accent }, line: { width: 0 }, rectRadius: 0.04 });
    s.addText(d.badge, { x: x + cw - 1.5, y: y + 0.13, w: 1.35, h: 0.28, fontFace: F.body, fontSize: 8, bold: true, color: C.white, align: 'center', charSpacing: 1, valign: 'middle' });
    s.addText(d.sub, { x: x + 0.18, y: y + 0.45, w: cw - 0.3, h: 0.25, fontFace: F.body, fontSize: 9.5, color: C.muted, italic: true });
    s.addText(d.body, { x: x + 0.18, y: y + 0.75, w: cw - 0.3, h: ch - 0.85, fontFace: F.body, valign: 'top' });
  });
  s.addShape('roundRect', { x: 0.4, y: 6.4, w: 12.55, h: 0.6, fill: { color: '0D3D2E' }, line: { color: C.green, width: 1 }, rectRadius: 0.08 });
  s.addText([
    { text: 'Coverage math: ', options: { bold: true, color: C.green } },
    { text: 'Scotiabank (forecast) + AMBSE (swing) + BlueDoor (pull-forward) + Future Builds (wild card) = up to ', options: { color: C.text } },
    { text: '$76k of close potential', options: { bold: true, color: C.text } },
    { text: ' in May vs $46.75k quota. Even half of these landing puts the month in the green.', options: { color: C.text } }
  ], { x: 0.6, y: 6.45, w: 12.2, h: 0.5, fontFace: F.body, fontSize: 11.5, valign: 'middle' });
}

// ===== S15 COACHING =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 15, 'Coaching Focus — Rebuilt Around MEDDICC');
  s.addText([
    { text: 'The 30-day coaching plan had 3 SMART goals: Value Narrative, MEDDPICC, Executive Summary. April\'s data confirms ', options: { color: C.muted } },
    { text: 'MEDDICC is THE coaching focus', options: { bold: true, color: C.text } },
    { text: ' — pipe gen is fine, deal management isn\'t.', options: { color: C.muted } }
  ], { x: 0.4, y: 1.3, w: 12.6, h: 0.45, fontFace: F.body, fontSize: 12 });
  const goals = [
    { accent: C.accent, label: 'PRIMARY', title: 'MEDDICC Scorecard',
      bullets: ['80% of Stage 2+ deals scored by 5/31', '1 deal card brought to every weekly 1:1 (Andrew challenges yellows + reds)', 'R/Y/G updated every 14 days, on every Stage 2+ deal'] },
    { accent: C.blue, label: 'SECONDARY', title: 'Value Narrative',
      bullets: ['Every $10k+ deal has a 1-page narrative before Proposal stage', 'Quantified Metric + named Pain owner = required', 'Carry forward Coaching Goal 1 from 30-day plan'] },
    { accent: C.green, label: 'SECONDARY', title: 'Executive Summaries',
      bullets: ['Top 3 May deals: AMBSE, StayTuned, Revibe (and Scotiabank as needed)', 'The vehicle for the EB conversation', 'Bring all 3 to 1:1 on 5/19'] },
    { accent: C.orange, label: 'NEW THEME', title: 'Disqualification at Day 14',
      bullets: ['April\'s lesson: Speed Sport, Oversight, Fusable, Chefman all entered pipe before fit confirmed', 'MEDDICC scorecard at Day 1 = the gate', '60-day rule on $10k+ deals = the floor'] }
  ];
  const cw = 6.2, ch = 2.0, gap = 0.15;
  goals.forEach((g, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.4 + col * (cw + gap), y = 1.85 + row * (ch + gap);
    s.addShape('rect', { x, y, w: cw, h: ch, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y, w: 0.06, h: ch, fill: { color: g.accent }, line: { width: 0 } });
    s.addText(g.label, { x: x + 0.18, y: y + 0.1, w: 2, h: 0.25, fontFace: F.body, fontSize: 9, color: g.accent, bold: true, charSpacing: 2 });
    s.addText(g.title, { x: x + 0.18, y: y + 0.35, w: cw - 0.3, h: 0.4, fontFace: F.head, fontSize: 17, bold: true, color: g.accent });
    const bulletText = g.bullets.flatMap(b => [
      { text: '• ', options: { color: g.accent, fontSize: 11 } },
      { text: b + '\n', options: { color: C.text, fontSize: 11 } }
    ]);
    s.addText(bulletText, { x: x + 0.18, y: y + 0.85, w: cw - 0.3, h: ch - 0.95, fontFace: F.body, valign: 'top' });
  });
  s.addShape('roundRect', { x: 0.4, y: 6.05, w: 12.55, h: 0.95, fill: { color: '231D3D' }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Carried forward from March: ', options: { bold: true, color: C.accent } },
    { text: 'Discovery, Value Selling, Enterprise Motions — but reframed as MEDDICC mechanics. Discovery = surfacing M+I+C. Enterprise Motions = mapping P + Decision Process. Same skills, sharper container.', options: { color: C.text } }
  ], { x: 0.6, y: 6.15, w: 12.2, h: 0.8, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S16 OPENING =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  s.background = { color: C.navy };
  s.addText('📝 1:1 OPENING LINE', { x: 0.7, y: 0.6, w: 12, h: 0.4, fontFace: F.body, fontSize: 14, color: C.green, bold: true, charSpacing: 4 });
  s.addText([
    { text: '"April had a split personality — pipe gen at ', options: { color: C.white, italic: true } },
    { text: '103% on 23 deals ($193k)', options: { bold: true, color: C.white, italic: true } },
    { text: ', but revenue at ', options: { color: C.white, italic: true } },
    { text: '37% / $17.2k', options: { bold: true, color: C.white, italic: true } },
    { text: ' against quota. The TOFU machine works; the conversion machine broke. Uniphore $40k is the headline loss — autopsy in hand.\n\nThe May plan is two things: ', options: { color: C.white, italic: true } },
    { text: 'realistic activity targets', options: { bold: true, color: C.white, italic: true } },
    { text: ' consolidated into 18 working days (18 contacts, ~56 emails, 17 dials, 12 LinkedIn, 3-4 meetings booked per working day), and ', options: { color: C.white, italic: true } },
    { text: 'MEDDICC operationalized', options: { bold: true, color: C.white, italic: true } },
    { text: ' — every Stage 2+ deal scored, weakest dimension named, doc artifacts produced.\n\nMay starts ', options: { color: C.white, italic: true } },
    { text: 'over-covered on quota', options: { bold: true, color: C.green, italic: true } },
    { text: '; the job is converting the weighted pipe to closed."', options: { color: C.white, italic: true } }
  ], { x: 1.0, y: 1.5, w: 11.3, h: 5, fontFace: F.body, fontSize: 17, lineSpacing: 30 });
}

pres.writeFile({ fileName: '/Users/thom/Documents/Personal/Code Projects/Sales Coach/April_2026_NAM_MBR.pptx' })
  .then(f => console.log('Wrote:', f));
