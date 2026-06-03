// May 2026 NAM MBR — PPTX builder (rebuildable from this script)
// Mirrors May_2026_NAM_MBR.html
const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'May 2026 NAM MBR — Thomas R Lindsey';
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
    { text: { text: 'May NAM MBR · Thomas R Lindsey · 06.03.2026', options: { x: 0.4, y: 7.15, w: 8, h: 0.25, fontFace: F.body, fontSize: 8, color: C.muted } } },
    { text: { text: 'SUPERMETRICS', options: { x: 11.4, y: 7.15, w: 1.8, h: 0.25, fontFace: F.head, fontSize: 8, color: C.muted, align: 'right', bold: true, charSpacing: 2 } } }
  ]
});

const slideTitle = (slide, num, title) => {
  slide.addText(`SLIDE ${num}`, { x: 0.4, y: 0.3, w: 6, h: 0.3, fontFace: F.body, fontSize: 9, color: C.muted, bold: true, charSpacing: 4 });
  slide.addText(title, { x: 0.4, y: 0.55, w: 12.5, h: 0.7, fontFace: F.head, fontSize: 28, bold: true, color: C.white });
};

const card = (sl, x, y, w, h, accent, label, big, sub) => {
  sl.addShape('rect', { x, y, w, h, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  sl.addShape('rect', { x, y, w: 0.08, h, fill: { color: accent }, line: { width: 0 } });
  sl.addText(label, { x: x + 0.25, y: y + 0.12, w: w - 0.4, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, bold: true, charSpacing: 2 });
  sl.addText(big, { x: x + 0.25, y: y + 0.45, w: w - 0.4, h: 0.6, fontFace: F.head, fontSize: 26, bold: true, color: accent });
  sl.addText(sub, { x: x + 0.25, y: y + 1.05, w: w - 0.4, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted });
};

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

const tableOpts = (x, y, w, colW, rowH) => ({ x, y, w, colW, rowH, fontFace: F.body, border: { type: 'solid', pt: 0.5, color: C.border } });

// ===== S1 TITLE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  s.background = { color: C.navy };
  s.addText('MONTHLY BUSINESS REVIEW', { x: 0.7, y: 1.8, w: 12, h: 0.4, fontFace: F.body, fontSize: 13, color: 'A0B0D0', bold: true, charSpacing: 5 });
  s.addText('May NAM MBR', { x: 0.7, y: 2.3, w: 12, h: 1.6, fontFace: F.head, fontSize: 70, bold: true, color: C.white });
  s.addText('Thomas R Lindsey  ·  06.03.2026  ·  Supermetrics', { x: 0.7, y: 4.0, w: 12, h: 0.4, fontFace: F.body, fontSize: 16, color: 'A0B0D0' });
  s.addShape('roundRect', { x: 0.7, y: 4.7, w: 10.5, h: 0.55, fill: { color: '143A2E' }, line: { color: C.green, width: 1 }, rectRadius: 0.27 });
  s.addText('✓  Revenue recovered to 86%  ·  Scotiabank closed $35k  ·  pipe gen fell to ~54%', { x: 0.85, y: 4.72, w: 10.2, h: 0.5, fontFace: F.body, fontSize: 13, color: C.green, bold: true });
}

// ===== S2 OVERVIEW =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 2, 'May Overview — The Mirror Image of April');
  s.addText('REVENUE GOALS', { x: 0.4, y: 1.45, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 3 });
  card(s, 0.4, 1.8, 6.2, 1.5, C.red, 'APRIL REVENUE (LAST MONTH)', '$17,213 / 37%', 'Worst revenue month in some time');
  card(s, 6.8, 1.8, 6.2, 1.5, C.green, 'MAY ACTUAL REVENUE', '$40,379 / 86%', 'Best revenue month in some time · Scotiabank $35k anchor');
  s.addText('PIPELINE GOALS', { x: 0.4, y: 3.55, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 3 });
  card(s, 0.4, 3.9, 6.2, 1.5, C.green, 'APRIL PIPELINE CREATED', '$193,100 / 103%', '23 deals · first month over goal in 2026');
  card(s, 6.8, 3.9, 6.2, 1.5, C.red, 'MAY PIPELINE CREATED', '~$101,000 / ~54%', '20 deals · ~$5k avg ACV · calls collapsed to 39');
  s.addShape('roundRect', { x: 0.4, y: 5.65, w: 12.55, h: 1.3, fill: { color: '143A2E' }, line: { color: C.green, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The headline: ', options: { bold: true, color: C.green } },
    { text: 'April delivered the most pipeline and the worst revenue of the year. May flipped both — revenue recovered to 86%, conversion worked (Scotiabank). The new gap is the inverse of April\'s: the top of the funnel went quiet. ', options: { color: C.text } },
    { text: 'June has to do both motions at once.', options: { bold: true, color: C.orange } }
  ], { x: 0.6, y: 5.75, w: 12.15, h: 1.1, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S3 APRIL PLAN PLAYED OUT =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 3, 'April Plan — How It Played Out in May');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 4.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('MAY MITs (FROM APRIL PLAN)', { x: 0.6, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '✓  MEDDICC Operationalized — WORKED (where applied)\n', options: { color: C.green, bold: true, fontSize: 12 } },
    { text: '    Scotiabank $35k closed the MEDDICC way — EB + engineering engaged, technical call, security review navigated. The plan\'s #1 named deal converted.\n\n', options: { color: C.muted, fontSize: 10 } },
    { text: '✗  TOFU at the activity level — FAIL\n', options: { color: C.red, bold: true, fontSize: 12 } },
    { text: '    Calls fell to 39 (from 280). The daily numbers were the plan; they didn\'t get hit. Pipe gen dropped to ~54%.', options: { color: C.muted, fontSize: 10 } }
  ], { x: 0.6, y: 1.9, w: 5.8, h: 3.4, fontFace: F.body });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 4.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('THE TRADE-OFF', { x: 7.0, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '~  Call Prep blocks held (~18) — but didn\'t convert to dials (39)\n', options: { color: C.orange, bold: true, fontSize: 12 } },
    { text: '    The block was on the calendar; the calls weren\'t made inside it. The single most coachable item.\n\n', options: { color: C.muted, fontSize: 10 } },
    { text: '✓  MEDDICC depth on Scotiabank — multi-threaded, security-led, EB engaged\n\n', options: { color: C.green, bold: true, fontSize: 11 } },
    { text: '✗  MEDDICC breadth missing — DTCC ($25k) & Industry Dive ($20k) lost without the same motion', options: { color: C.red, bold: true, fontSize: 11 } }
  ], { x: 7.0, y: 1.9, w: 5.8, h: 3.4, fontFace: F.body });

  s.addShape('roundRect', { x: 0.4, y: 5.65, w: 12.55, h: 1.3, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The one-sentence read: ', options: { bold: true, color: C.accent } },
    { text: 'May was a closing sprint — it converted the flagship deal the April plan was built around, but it did so by coming off the phones. ', options: { color: C.text } },
    { text: 'A sprint is not a system — June has to run conversion and prospecting in parallel.', options: { bold: true, color: C.orange } }
  ], { x: 0.6, y: 5.75, w: 12.15, h: 1.1, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S4 REVIEW =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 4, 'May Review');
  col3(s, [
    { accent: C.green, title: 'WHAT I DID WELL', body: [
      { text: '✓  Scotiabank closed — $35k. ', options: { color: C.green, bold: true } },
      { text: 'The April anchor, won via the full MEDDICC motion.\n\n', options: { color: C.muted } },
      { text: '✓  Revenue recovery 37% → 86% — best revenue month in some time\n\n', options: { color: C.text } },
      { text: '✓  Flor ($5.4k) — second fast MCP/Claude close; AI pattern still repeatable\n\n', options: { color: C.text } },
      { text: '✓  CRM hygiene improved — SFDC blocks doubled (~8)\n\n', options: { color: C.text } },
      { text: '✓  Fast disqualification working — Graff Auto (9d), S88 (11d)\n\n', options: { color: C.text } },
      { text: '✓  Contact creation held (177) — raw material for June is there', options: { color: C.text } }
    ] },
    { accent: C.red, title: 'WHERE I NEED TO IMPROVE', body: [
      { text: '✗  Calls collapsed — 280 → 39. ', options: { color: C.red, bold: true } },
      { text: '18 Call Prep blocks, 39 dials. The execution gap inside the block.\n\n', options: { color: C.muted } },
      { text: '✗  Pipe gen fell to ~54% — 20 small deals, two biggest lost same-month\n\n', options: { color: C.text } },
      { text: '✗  MEDDICC not applied broadly — DTCC $25k & Industry Dive $20k lost without it\n\n', options: { color: C.text } },
      { text: '✗  Cadence discipline lapsed week of 5/18–5/22', options: { color: C.text } }
    ] },
    { accent: C.accent, title: 'KEY DEAL LEARNINGS', body: [
      { text: 'Scotiabank win pattern: ', options: { bold: true, color: C.text } },
      { text: 'EB + technical/engineering multi-thread + early security docs = the enterprise motion. Replicable.\n\n', options: { color: C.muted } },
      { text: 'Industry Dive ($20k, 262 days): ', options: { bold: true, color: C.text } },
      { text: 'the stale-deal death the 60-day rule was meant to prevent.\n\n', options: { color: C.muted } },
      { text: 'DTCC ($25k, 28 days): ', options: { bold: true, color: C.text } },
      { text: 'never got the EB/technical motion that saved Scotiabank. Same lesson, opposite outcome.\n\n', options: { color: C.muted } },
      { text: 'The volume bill: ', options: { bold: true, color: C.text } },
      { text: 'one month off the phones is invisible in May\'s revenue but shows up in July\'s pipeline.', options: { color: C.muted } }
    ] }
  ]);
}

// ===== S4b SCOTIABANK =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '4b', 'Scotiabank — Anatomy of the MEDDICC Win');
  s.addText('$35k closed · 421-day cycle · Cross-sell · the deal the April plan was built around. The template for every enterprise deal going forward.', { x: 0.4, y: 1.3, w: 12.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted });
  const cols = [
    { accent: C.green, title: 'WHAT WE DID RIGHT', body: [
      { text: '✓  Multi-threaded late-stage into their engineering org — Director of Engineering + Lead & Senior Solutions Architects\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '✓  Ran a technical SE call with a written prep brief (5/27) — anticipated SFTP vs API vs Hub/BigQuery objections\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '✓  Got security/DR docs into their vendor review early (4/24–4/28, confirmed received)\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '✓  Kept the EB + decision path live through signature', options: { color: C.text, fontSize: 10 } }
    ] },
    { accent: C.blue, title: 'WHY IT CONVERTED', body: [
      { text: '• Their engineering team tried to revalidate the ingestion method right before signature — and we were ready\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '• The security story ("we drop files inside your perimeter") fit a regulated bank\'s posture\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '• Google partner motion gave an extra lever on the BigQuery option\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '• The April Uniphore lesson (no EB → lost) was inverted here (EB engaged → won)', options: { color: C.text, fontSize: 10, bold: true } }
    ] },
    { accent: C.green, title: 'THE REPLICABLE PLAYBOOK', body: [
      { text: '• Late-stage = multi-thread into technical/engineering, not just the champion\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '• Every technical call gets a prep brief\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '• Push security/legal docs early, before they\'re asked for\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'So what: ', options: { color: C.green, fontSize: 10, bold: true } },
      { text: 'now the standard for Upwork, Grafton, AMBSE, BlueDoor, BT1D — June\'s $10k+ pipe', options: { color: C.muted, fontSize: 10 } }
    ] }
  ];
  const cw = 4.05, gap = 0.2, sx = 0.4, sy = 1.7, ch = 4.9;
  cols.forEach((p, i) => {
    const x = sx + i * (cw + gap);
    s.addShape('rect', { x, y: sy, w: cw, h: ch, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y: sy, w: 0.05, h: ch, fill: { color: p.accent }, line: { width: 0 } });
    s.addText(p.title, { x: x + 0.15, y: sy + 0.12, w: cw - 0.25, h: 0.3, fontFace: F.body, fontSize: 10, bold: true, color: p.accent, charSpacing: 1 });
    s.addText(p.body, { x: x + 0.15, y: sy + 0.5, w: cw - 0.25, h: ch - 0.6, fontFace: F.body, valign: 'top' });
  });
}

// ===== S5 PIPELINE GEN =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 5, 'May Pipeline Gen Recap');
  const headers = ['Type', 'Count', 'Value (ERR)', 'ACV', 'Notes'];
  const rows = [
    ['New Business', '13', '~$63,000', '~$4,800', 'Mostly Growth/Starter; BT1D $19k standout'],
    ['Cross-sell', '3', '~$14,000', '~$4,700', 'Tracer (Vayner), Charter One, TaskRabbit'],
    ['Upsell', '2', '~$24,000', '~$12,000', 'Chefman, Maxio — both lost same month'],
    ['Small / Growth', '2', '~$2,000', '~$1,000', 'Starter / Growth'],
    ['TOTAL', '20', '~$101,000', '~$5,050', '—']
  ];
  const tbl = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11, align: (h === 'Type' || h === 'Notes') ? 'left' : (h === 'Count' ? 'center' : 'right') } })),
    ...rows.map((r, i) => r.map((c, j) => ({
      text: c,
      options: { bold: i === 4, color: i === 4 ? C.white : C.text, fill: { color: i === 4 ? C.surface2 : C.surface }, fontSize: 11,
        align: j === 0 || j === 4 ? 'left' : (j === 1 ? 'center' : 'right') }
    })))
  ];
  s.addTable(tbl, tableOpts(0.4, 1.5, 12.55, [2.2, 1.2, 2.0, 1.6, 5.55], 0.45));
  s.addShape('rect', { x: 0.4, y: 4.0, w: 12.55, h: 0.5, fill: { color: '3A2520' }, line: { color: C.red, width: 0.5 } });
  s.addText([
    { text: 'Pipeline Goal: ', options: { bold: true, color: C.text } },
    { text: '$187,000  ', options: { color: C.text } },
    { text: '·  ~54%', options: { bold: true, color: C.red } },
    { text: '  ·  arc broke after April\'s 103%', options: { color: C.muted } }
  ], { x: 0.6, y: 4.05, w: 12.2, h: 0.4, fontFace: F.body, fontSize: 13, valign: 'middle' });
  s.addShape('roundRect', { x: 0.4, y: 5.0, w: 12.55, h: 1.6, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'So what: ', options: { bold: true, color: C.orange } },
    { text: 'Net durable new pipeline is even thinner than $101k — the two biggest "new" deals (Diablo $18k, Maxio $12k) were closed lost within the same month. Average new-deal ACV fell to ~$5k (April: $8.4k).\n\n', options: { color: C.text } },
    { text: 'Value is deal-name-encoded ERR from HubSpot; the SF AE-Automation credit used in April isn\'t available this cycle — but the drop is real and large.', options: { italic: true, color: C.muted, fontSize: 11 } }
  ], { x: 0.6, y: 5.1, w: 12.2, h: 1.4, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S6 YTD FLIP =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 6, 'YTD 2026 — The Flip');
  // Pipe gen table
  s.addText('PIPELINE GEN % TO GOAL', { x: 0.4, y: 1.45, w: 6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  const pg = [
    ['Month', 'Deals', '% Goal'],
    ['January', '9', '50%', C.red],
    ['February', '18', '90%', C.orange],
    ['March', '16', '89%', C.orange],
    ['April', '23', '103%', C.green],
    ['May', '20', '~54%', C.red]
  ];
  const pgT = [
    pg[0].map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11, align: h === 'Month' ? 'left' : 'center' } })),
    ...pg.slice(1).map((r, i) => r.slice(0, 3).map((c, j) => ({
      text: c, options: { bold: i === 4 || j === 2, color: j === 2 ? r[3] : (i === 4 ? C.white : C.text), fill: { color: i === 4 ? '3A2520' : C.surface }, fontSize: 12, align: j === 0 ? 'left' : 'center' }
    })))
  ];
  s.addTable(pgT, tableOpts(0.4, 1.8, 6.1, [2.6, 1.7, 1.8], 0.45));
  s.addText('50 → 90 → 89 → 103 → 54.  The pipe-gen arc broke.', { x: 0.4, y: 4.7, w: 6.1, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
  // Revenue table
  s.addText('REVENUE % TO QUOTA', { x: 6.9, y: 1.45, w: 6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  const rv = [
    ['Month', 'Revenue', '% Quota'],
    ['April', '$17,213', '37%', C.red],
    ['May', '$40,379', '86%', C.green]
  ];
  const rvT = [
    rv[0].map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11, align: h === 'Month' ? 'left' : (h === 'Revenue' ? 'right' : 'center') } })),
    ...rv.slice(1).map((r, i) => r.slice(0, 3).map((c, j) => ({
      text: c, options: { bold: i === 1 || j === 2, color: j === 2 ? r[3] : (i === 1 ? C.white : C.text), fill: { color: i === 1 ? '143A2E' : C.surface }, fontSize: 12, align: j === 0 ? 'left' : (j === 1 ? 'right' : 'center') }
    })))
  ];
  s.addTable(rvT, tableOpts(6.9, 1.8, 6.1, [2.3, 2.2, 1.6], 0.45));
  s.addText('37 → 86.  The revenue arc turned up — as pipe gen turned down.', { x: 6.9, y: 3.3, w: 6.1, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });

  s.addShape('roundRect', { x: 0.4, y: 5.2, w: 12.55, h: 1.4, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The pattern: ', options: { bold: true, color: C.accent } },
    { text: 'these two lines moved in opposite directions, two months running. That\'s not coincidence — it\'s a focus problem. ', options: { color: C.text } },
    { text: 'The goal isn\'t to pick one line; it\'s to lift both at once.', options: { bold: true, color: C.green } }
  ], { x: 0.6, y: 5.3, w: 12.2, h: 1.2, fontFace: F.body, fontSize: 14, valign: 'middle' });
}

// ===== S7 CLOSED EXECUTION =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 7, 'May Closed Deal Execution — The Real Story');
  const stats = [
    { label: 'REVENUE (WON)', val: '$40.4k', sub: '86% to quota · 2 deals', color: C.green },
    { label: 'BIGGEST WIN', val: '$35k', sub: 'Scotiabank · MEDDICC', color: C.green },
    { label: 'CLOSED LOST (Y1 ACV)', val: '~$125k', sub: '18 deals', color: C.red },
    { label: 'NET CLOSED ACTIVITY', val: '−$85k', sub: 'Won minus lost', color: C.red }
  ];
  stats.forEach((st, i) => {
    const x = 0.4 + i * 3.2;
    s.addShape('rect', { x, y: 1.4, w: 3.0, h: 1.3, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addText(st.label, { x: x + 0.15, y: 1.5, w: 2.8, h: 0.3, fontFace: F.body, fontSize: 9, color: C.muted, bold: true, charSpacing: 1 });
    s.addText(st.val, { x: x + 0.15, y: 1.8, w: 2.8, h: 0.6, fontFace: F.head, fontSize: 28, bold: true, color: st.color });
    s.addText(st.sub, { x: x + 0.15, y: 2.4, w: 2.8, h: 0.25, fontFace: F.body, fontSize: 9, color: C.muted });
  });
  s.addText('THE LOSSES THAT MATTERED', { x: 0.4, y: 2.95, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  const mk = (t, c, b, sz) => ({ text: t, options: { color: c, bold: !!b, fontSize: sz || 10, fill: { color: C.surface } } });
  const tbl = [
    [mk('Story', C.muted, true, 11), mk('Deals', C.muted, true, 11), { text: 'Lost ACV', options: { color: C.muted, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface2 } } }, mk('So what', C.muted, true, 11)].map(o => (o.options.fill = { color: C.surface2 }, o)),
    [mk('Big mid-funnel deaths', C.text, true, 11), mk('DTCC ($25k,28d), Industry Dive ($20k,262d), Diablo ($18k), Maxio ($12k), Future Builds ($12k)', C.text, false, 10), { text: '~$87,000', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } }, mk('In the April pipeline. No EB/technical motion ran — the opposite of Scotiabank.', C.muted, false, 10)],
    [mk('Stale carryover', C.text, true, 11), mk('Manhattan ($8k,84d), Promega ($5k,68d), Vooks (61d), Organicare ($2k,99d)', C.text, false, 10), { text: '~$17,000', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } }, mk('Long-cycle, no real next step. 60-day rule should have caught these.', C.muted, false, 10)],
    [mk('Fast disqualifications (good)', C.text, true, 11), mk('S88 (11d), Graff Auto (9d), Chefman (0d), Parc Safari (0d)', C.text, false, 10), { text: '~$7,000', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } }, mk('Died fast — the right behavior. Disqualification muscle developing.', C.muted, false, 10)]
  ];
  s.addTable(tbl, tableOpts(0.4, 3.3, 12.55, [2.8, 4.6, 1.4, 3.75], 0.62));
  s.addShape('roundRect', { x: 0.4, y: 6.05, w: 12.55, h: 0.9, fill: { color: '143A2E' }, line: { color: C.green, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The split inside the month: ', options: { bold: true, color: C.green } },
    { text: 'the one deal that got the full MEDDICC motion (Scotiabank) won $35k; the two biggest that didn\'t (DTCC, Industry Dive) lost $45k. ', options: { color: C.text } },
    { text: 'June\'s job: give every $10k+ deal the Scotiabank treatment.', options: { bold: true, color: C.orange } }
  ], { x: 0.6, y: 6.12, w: 12.2, h: 0.75, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S8 JUNE MITs =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 8, 'June MITs');
  const mits = [
    { accent: C.blue, num: '1', title: 'RESTART THE DIAL ENGINE', body: [
      { text: 'Convert Call Prep blocks into ', options: { color: C.text, fontSize: 13 } },
      { text: '≥15 dials every working day', options: { bold: true, color: C.text, fontSize: 13 } },
      { text: ' — 200+ calls for June. The structure already held in May (~18 blocks); the failure was execution inside the block (39 dials). Rebuild Cadence discipline — no skipped weeks like 5/18–5/22.\n\n', options: { color: C.text, fontSize: 13 } },
      { text: 'The headline MIT — May\'s revenue was bought by coming off the phones. June refuses that trade.', options: { italic: true, color: C.muted, fontSize: 11 } } ] },
    { accent: C.accent, num: '2', title: 'MEDDICC BREADTH — EVERY $10k+ DEAL', body: [
      { text: 'Scotiabank proved the motion. Now make it the standard, not the exception: EB engaged before pricing, technical multi-thread, security docs early, MAP shared — on Upwork, Grafton, AMBSE, BlueDoor, BT1D. Disqualify-at-day-60 enforced (the Industry Dive lesson).\n\n', options: { color: C.text, fontSize: 13 } },
      { text: 'May proved depth on one deal. June proves breadth across the portfolio.', options: { italic: true, color: C.muted, fontSize: 11 } } ] }
  ];
  mits.forEach((m, i) => {
    const x = 0.4 + i * 6.4;
    s.addShape('rect', { x, y: 1.5, w: 6.2, h: 5.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y: 1.5, w: 0.08, h: 5.0, fill: { color: m.accent }, line: { width: 0 } });
    s.addShape('ellipse', { x: x + 0.35, y: 1.75, w: 0.8, h: 0.8, fill: { color: m.accent }, line: { width: 0 } });
    s.addText(m.num, { x: x + 0.35, y: 1.78, w: 0.8, h: 0.75, fontFace: F.head, fontSize: 32, bold: true, color: C.white, align: 'center' });
    s.addText(m.title, { x: x + 1.3, y: 1.85, w: 4.7, h: 0.6, fontFace: F.head, fontSize: 16, bold: true, color: m.accent, charSpacing: 1, valign: 'middle' });
    s.addText(m.body, { x: x + 0.35, y: 2.8, w: 5.65, h: 3.5, fontFace: F.body, valign: 'top' });
  });
}

// ===== S9 FUNNEL MATH =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 9, 'June Funnel Math — Top-Down Derivation');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('LOCKED INPUTS', { x: 0.6, y: 1.5, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '• June quota: ', options: { color: C.text, fontSize: 12 } }, { text: '$46,750', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' (unchanged)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• 4× pipeline goal: ', options: { color: C.text, fontSize: 12 } }, { text: '$187,000\n', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '• Plan target: ', options: { color: C.text, fontSize: 12 } }, { text: '$200,000', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' (+7% buffer)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• ACV baseline: ', options: { color: C.text, fontSize: 12 } }, { text: '$11,000', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' (6-mo avg, anchored by Scotiabank $35k)', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 1.85, w: 5.8, h: 1.7, fontFace: F.body });
  s.addText('BACKED-OUT TARGETS', { x: 0.6, y: 3.7, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '• Deals needed: ', options: { color: C.text, fontSize: 12 } }, { text: '~18', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' ($200k ÷ $11k)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Booked meetings: ', options: { color: C.text, fontSize: 12 } }, { text: '~60', options: { bold: true, color: C.text, fontSize: 12 } }, { text: '\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Contacts enrolled: ', options: { color: C.text, fontSize: 12 } }, { text: '~330', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' (May held 177 — push back up)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Calls: ', options: { color: C.text, fontSize: 12 } }, { text: '~300', options: { bold: true, color: C.green, fontSize: 12 } }, { text: ' — the explicit fix for the May miss (39 → 300)', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 4.05, w: 5.8, h: 1.8, fontFace: F.body });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('MAY BASELINE → JUNE TARGET', { x: 7.0, y: 1.5, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  const dataRows = [
    ['Calls logged', '39', '300', '+669%', C.green, true],
    ['Contacts created', '177', '330', '+86%', C.green, false],
    ['Emails sent', '567', '1,000', '+76%', C.green, false],
    ['Meetings', '54', '60', '+11%', C.text, false],
    ['New deals', '20', '~18', 'fewer · larger', C.orange, false],
    ['Pipe generated', '~$101k', '$200k', '+98%', C.green, true]
  ].map(r => [
    { text: r[0], options: { color: C.text, fontSize: 11, bold: r[5], fill: { color: r[5] ? C.surface2 : C.surface } } },
    { text: r[1], options: { color: C.text, fontSize: 11, align: 'right', bold: r[5], fill: { color: r[5] ? C.surface2 : C.surface } } },
    { text: r[2], options: { color: C.text, fontSize: 11, align: 'right', bold: r[5], fill: { color: r[5] ? C.surface2 : C.surface } } },
    { text: r[3], options: { color: r[4], fontSize: 11, align: 'right', bold: true, fill: { color: r[5] ? C.surface2 : C.surface } } }
  ]);
  const tblData = [
    [{ text: 'Activity', options: { bold: true, color: C.muted, fontSize: 10, fill: { color: C.surface } } },
     { text: 'May', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } },
     { text: 'June', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } },
     { text: 'Δ', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } }],
    ...dataRows
  ];
  s.addTable(tblData, tableOpts(7.0, 1.85, 5.8, [2.4, 1.0, 1.0, 1.4], 0.42));
  s.addText('The big lever is calls. May\'s 39 is the anomaly to fix; everything else follows from putting dials back inside the Call Prep blocks that already exist.', { x: 7.0, y: 5.1, w: 5.8, h: 0.7, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
  s.addShape('roundRect', { x: 0.4, y: 6.1, w: 12.55, h: 0.85, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The shift: ', options: { bold: true, color: C.accent } },
    { text: 'April proved volume. May proved conversion. ', options: { color: C.text } },
    { text: 'June proves both', options: { bold: true, color: C.green } },
    { text: ' — the dial engine back on, with Scotiabank-grade MEDDICC behind every $10k+ deal.', options: { color: C.text } }
  ], { x: 0.6, y: 6.18, w: 12.2, h: 0.7, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S10 DAILY ACTIVITY =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 10, 'June Daily Activity Targets');
  s.addText('21 working days in June (Jun 1–30; Juneteenth 6/19 holiday). The May lesson: a block on the calendar is not a dial made. June counts dials, not blocks.', { x: 0.4, y: 1.3, w: 12.6, h: 0.4, fontFace: F.body, fontSize: 11, color: C.muted });
  const headers = ['Activity', 'Monthly', 'Weekly (avg)', 'Daily (21d)', 'Notes'];
  const rows = [
    ['Calls / dials', '300', '~75', '~15/day', 'The May miss (39). Made inside Call Prep blocks', true],
    ['New contacts enrolled', '330', '~83', '~16/day', 'Email-verified before enroll', false],
    ['Emails sent (cadence)', '1,000', '~250', '~48/day', 'Cadence queue worked AM', false],
    ['LinkedIn touches', '~210', '~52', '~10/day', '5 connect + 3 InMail + 3 comments', false],
    ['Net-new accounts worked', '~70', '~17', '~4/day', 'Mon list-build to refill funnel', false],
    ['Booked meetings', '60', '~15', '~3/day', 'Balance closing + new', false],
    ['Net-new deals created', '~18', '~4-5', '~1/day', 'Fewer, larger · $11k+ ACV', false]
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
  s.addTable(tblD, tableOpts(0.4, 1.8, 12.55, [3.0, 1.5, 1.4, 1.6, 5.05], 0.45));
  s.addShape('roundRect', { x: 0.4, y: 5.95, w: 12.55, h: 1.0, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The "no exceptions" line: ', options: { bold: true, color: C.orange } },
    { text: '15 dials on every working day. ', options: { bold: true, color: C.text } },
    { text: 'May had the blocks (~18 Call Prep) and missed the dials (39 total). The metric that matters is calls made, not calendar coverage.', options: { color: C.text } }
  ], { x: 0.6, y: 6.05, w: 12.2, h: 0.85, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S11 PLAYBOOK =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 11, 'How We Work the Activities — Dials Inside the Blocks');
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 12.55, h: 0.9, fill: { color: '231D3D' }, line: { color: C.accent, width: 1.5 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Commit: ', options: { bold: true, color: C.accent, fontSize: 13 } },
    { text: 'The Reclaim blocks already exist and held in May. The June change is ', options: { color: C.text, fontSize: 13 } },
    { text: 'accountability inside the block', options: { bold: true, color: C.text, fontSize: 13 } },
    { text: ' — every Call Prep block ends with dials logged in HubSpot, tracked daily, reviewed Friday. In June the dial count is the contract.', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 1.4, w: 12.2, h: 0.75, fontFace: F.body, valign: 'middle' });
  const blocks = [
    ['Account list build', 'Mondays · 60 min', 'Pull 20 net-new accounts to refill the funnel May thinned: ZoomInfo + intent + 2025 closed-lost rework.'],
    ['Contact enrollment', 'Daily AM · 30 min', '~16 contacts/day · verify email before enroll.'],
    ['Cadence + dials', 'Daily AM · 60 min', '~48 emails + ~15 dials due today — the dials are the part that slipped in May.'],
    ['LinkedIn block', 'Daily midday · 20 min', '5 connect + 3 InMail + 3 comments on target-account posts.'],
    ['Mid-funnel block', 'Daily PM · 60 min', 'Scotiabank-style motion on every $10k+ deal: EB, technical multi-thread, security docs early.'],
    ['Friday hygiene + MEDDICC', 'Weekly · 90 min', 'Disqualify anything 60+ days with no next step (Industry Dive rule). Update MEDDICC card on every Stage 2+ deal.'],
    ['Closed-lost rework', 'Tuesdays · 45 min', '5 accounts/wk from 2025 closed-lost · restart with MCP/AI angle.']
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
  slideTitle(s, 12, 'June True Pipeline — ~$167k Open (June-Close)');
  const headers = ['Account', 'Opportunity', 'ERR', 'Stage', '% Close', 'Weighted'];
  const dealRows = [
    ['US Soccer', 'MIC+DWH/DA | $40k', '$40,000', 'Qualified', '30%', '$12,000', C.blue, true],
    ['BT1D', 'S3 | $19k', '$19,000', 'Qualified', '30%', '$5,700', C.blue, true],
    ['Upwork', 'MIC/MCP | $15k', '$15,000', 'Proof of Value', '45%', '$6,750', C.accent, false],
    ['Grafton', 'DWH | $15k', '$15,000', 'Qualified', '30%', '$4,500', C.blue, false],
    ['AMBSE', 'MIC | $15k', '$15,000', 'Qualified', '45%', '$6,750', C.blue, false],
    ['BlueDoor', 'MIC | $10k', '$10,000', 'Qualified', '40%', '$4,000', C.blue, false],
    ['Liberty', 'Copilot | $6k', '$6,000', 'Qualified', '35%', '$2,100', C.blue, false],
    ['Charter One', 'API | $6k', '$6,000', 'Qualified', '35%', '$2,100', C.blue, false],
    ['TaskRabbit', 'Gem Agent | $5k', '$5,000', 'Qualified', '35%', '$1,750', C.blue, false],
    ['Furnished Finder', 'MCP | $4k', '$4,000', 'Qualified', '70%', '$2,800', C.blue, false],
    ['Kraken', 'MCP | $3.5k', '$3,500', 'Qualified', '50%', '$1,750', C.blue, false],
    ['Other (~13)', 'Growth/MCP/Starter $2-5k', '~$28,500', '—', '~30%', '~$8,500', C.muted, false],
    ['TOTAL', '~25 June-close deals', '~$167,000', '—', '~35%', '~$58,700', C.green, true]
  ];
  const tbl = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fontSize: 10, fill: { color: C.surface2 }, align: h === 'Account' || h === 'Opportunity' || h === 'Stage' ? 'left' : 'right' } })),
    ...dealRows.map((r, i) => {
      const last = i === dealRows.length - 1;
      const bg = last ? '143A2E' : C.surface;
      return [
        { text: r[0], options: { bold: r[7], color: C.text, fontSize: 11, fill: { color: bg } } },
        { text: r[1], options: { color: C.text, fontSize: 11, fill: { color: bg } } },
        { text: r[2], options: { color: C.text, fontSize: 11, align: 'right', fill: { color: bg } } },
        { text: r[3], options: { color: r[6], fontSize: 10, bold: true, fill: { color: bg } } },
        { text: r[4], options: { color: C.text, fontSize: 11, align: 'right', fill: { color: bg } } },
        { text: r[5], options: { color: last ? C.green : C.text, bold: r[7], fontSize: 11, align: 'right', fill: { color: bg } } }
      ];
    })
  ];
  s.addTable(tbl, tableOpts(0.4, 1.4, 12.55, [2.0, 2.7, 1.4, 2.2, 1.4, 2.85], 0.345));
  s.addText('No single Scotiabank-sized anchor this month — June is a portfolio of $10–19k deals. That\'s why MEDDICC breadth is the MIT: depth on one deal won\'t carry the month.', { x: 0.4, y: 6.4, w: 12.6, h: 0.4, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
}

// ===== S13 FORECAST =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 13, 'June Forecast');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 8.0, h: 5.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('FORECAST (HIGH-CONFIDENCE)', { x: 0.6, y: 1.5, w: 7.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.green, bold: true, charSpacing: 2 });
  s.addText([
    { text: '■  Furnished Finder — $2.8k (70%) ', options: { bold: true, color: C.green, fontSize: 12 } },
    { text: '· MCP closeout\n', options: { color: C.text, fontSize: 11 } },
    { text: '■  Kraken — $1.75k (50%) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '· MCP, early-June close', options: { color: C.muted, fontSize: 11 } }
  ], { x: 0.6, y: 1.85, w: 7.6, h: 1.1, fontFace: F.body });
  s.addText('SWING DEALS (THE MONTH IS MADE HERE)', { x: 0.6, y: 3.0, w: 7.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.orange, bold: true, charSpacing: 2 });
  s.addText([
    { text: '■  US Soccer ($40k) — overdue, must re-date or disqualify by 6/5. ', options: { bold: true, color: C.red, fontSize: 12 } },
    { text: 'Biggest single swing.\n', options: { color: C.text, fontSize: 11 } },
    { text: '■  AMBSE ($15k) — carried main swing ', options: { bold: true, color: C.accent, fontSize: 12 } },
    { text: '· EB conversation before pricing\n', options: { color: C.text, fontSize: 11 } },
    { text: '■  Upwork ($15k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— POV · champion test / multi-thread\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  Grafton ($15k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— get written decision criteria\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  BT1D ($19k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— new $19k deal · build value narrative early\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  BlueDoor ($10k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— POV · quantify pain or disqualify by 60d', options: { color: C.muted, fontSize: 11 } }
  ], { x: 0.6, y: 3.35, w: 7.6, h: 2.4, fontFace: F.body });
  s.addShape('roundRect', { x: 0.6, y: 5.9, w: 7.6, h: 0.85, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.08 });
  s.addText([
    { text: 'The narrative: ', options: { bold: true, color: C.orange } },
    { text: 'June opens at $0 closed with no Scotiabank-sized lock in sight. Coverage is a portfolio of mid-size deals — converts only if MEDDICC runs on all of them. And the dials must come back, or July is the next April.', options: { color: C.text } }
  ], { x: 0.75, y: 5.96, w: 7.3, h: 0.73, fontFace: F.body, fontSize: 10.5, valign: 'middle' });

  s.addShape('rect', { x: 8.6, y: 1.4, w: 4.4, h: 5.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('FORECAST / OPEN PIPELINE', { x: 8.8, y: 1.5, w: 4.0, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, bold: true, charSpacing: 2 });
  const fcRows = [
    ['Total Quota', '$46,750', false, C.text],
    ['Closed Won (MTD)', '$0', false, C.text],
    ['Open Pipeline (Jun)', '~$167,000', false, C.text],
    ['Weighted Open', '~$58,700', true, C.green],
    ['# Open Deals', '~39', false, C.text],
    ['Coverage vs Quota', '~1.26×', true, C.green]
  ];
  let fy = 1.95;
  fcRows.forEach(r => {
    s.addText(r[0], { x: 8.8, y: fy, w: 2.5, h: 0.35, fontFace: F.body, fontSize: 11, color: C.muted });
    s.addText(r[1], { x: 11.0, y: fy, w: 1.85, h: 0.35, fontFace: F.body, fontSize: 12, color: r[3], bold: r[2], align: 'right' });
    fy += 0.5;
  });
  s.addText('Weighted pipe covers quota ~1.26× — thinner cushion than May\'s Scotiabank-anchored start. Every $10k+ deal matters; none can be left without the MEDDICC motion.', { x: 8.8, y: 5.1, w: 4.0, h: 1.6, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
}

// ===== S13b MEDDICC REPLICATE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '13b', 'MEDDICC — Replicate the Scotiabank Motion');
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 12.55, h: 0.85, fill: { color: '143A2E' }, line: { color: C.green, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'May proved MEDDICC converts flagship deals. The June job is breadth: ', options: { bold: true, color: C.green, fontSize: 12 } },
    { text: 'run the exact Scotiabank motion on every $10k+ deal — not just the one that got attention. Depth won May\'s revenue; breadth wins June\'s portfolio.', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 1.38, w: 12.2, h: 0.7, fontFace: F.body, valign: 'middle' });
  const cols = [
    { accent: C.green, title: 'WHAT WON SCOTIABANK', body: [
      { text: 'E', options: { bold: true, color: C.green, fontSize: 11 } }, { text: 'conomic Buyer — engaged + decision path live through signature\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'Technical multi-thread', options: { bold: true, color: C.text, fontSize: 11 } }, { text: ' — engineering org, not just champion\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'P', options: { bold: true, color: C.green, fontSize: 11 } }, { text: 'aper Process — security/DR docs early, confirmed received\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'Prep brief', options: { bold: true, color: C.text, fontSize: 11 } }, { text: ' before the technical call — objections anticipated', options: { color: C.muted, fontSize: 10 } }
    ] },
    { accent: C.red, title: 'WHAT LOST DTCC + INDUSTRY DIVE', body: [
      { text: '✗  No EB / technical motion ran (DTCC, 28d)\n\n', options: { color: C.red, fontSize: 10, bold: true } },
      { text: '✗  262-day stale death, no next step (Industry Dive)\n\n', options: { color: C.red, fontSize: 10, bold: true } },
      { text: '✗  Pure-email follow-up on $20k+ deals\n\n', options: { color: C.red, fontSize: 10, bold: true } },
      { text: '✗  60-day disqualification rule not enforced', options: { color: C.red, fontSize: 10, bold: true } }
    ] },
    { accent: C.blue, title: 'JUNE OPERATING RHYTHM', body: [
      { text: '• Day 1: open MEDDICC card, rate R/Y/G\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '• Every $10k+ deal: EB booked before pricing\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '• Every technical call: written prep brief\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '• Friday: disqualify anything 60d+ with no next step', options: { color: C.text, fontSize: 10 } }
    ] }
  ];
  const cw = 4.05, gap = 0.2, sx = 0.4, sy = 2.35, ch = 4.2;
  cols.forEach((p, i) => {
    const x = sx + i * (cw + gap);
    s.addShape('rect', { x, y: sy, w: cw, h: ch, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y: sy, w: 0.05, h: ch, fill: { color: p.accent }, line: { width: 0 } });
    s.addText(p.title, { x: x + 0.15, y: sy + 0.12, w: cw - 0.25, h: 0.3, fontFace: F.body, fontSize: 10, bold: true, color: p.accent, charSpacing: 1 });
    s.addText(p.body, { x: x + 0.15, y: sy + 0.5, w: cw - 0.25, h: ch - 0.6, fontFace: F.body, valign: 'top' });
  });
}

// ===== S13d PER-DEAL MEDDICC =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '13d', 'June MEDDICC Targets — Per Active Deal');
  s.addText('Each major open deal · the red-to-fix dimension · the June action. Plus two loss autopsies carried from May.', { x: 0.4, y: 1.3, w: 12.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted });
  const headers = ['Deal', 'ERR', 'Weakest Dimension', 'June Action', 'Due'];
  const rows = [
    ['US Soccer', '$40k', 'Decision Process', 'Real next step or re-date / disqualify', '6/5', false],
    ['BT1D', '$19k', 'Metrics', 'Value narrative + ROI on new deal', '6/12', false],
    ['Upwork', '$15k', 'Champion', 'Multi-thread / champion test', '6/12', false],
    ['Grafton', '$15k', 'Decision Criteria', 'Get written must-haves', '6/12', false],
    ['AMBSE', '$15k', 'Economic Buyer', 'EB conversation before pricing', '6/10', false],
    ['BlueDoor', '$10k', 'Pain (urgency)', 'Quantify or disqualify by 60d', '6/18', false],
    ['DTCC (autopsy)', 'lost', 'Economic Buyer', '1-pager: why no EB motion ran', '6/6', true],
    ['Industry Dive (autopsy)', 'lost', 'Decision Process', '1-pager: the 262-day stale death', '6/6', true]
  ];
  const tbl = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fontSize: 11, fill: { color: C.surface2 }, align: h === 'ERR' || h === 'Due' ? 'right' : 'left' } })),
    ...rows.map(r => {
      const bg = r[5] ? '3A2520' : C.surface;
      return [
        { text: r[0], options: { color: C.text, fontSize: 11, bold: r[5], fill: { color: bg } } },
        { text: r[1], options: { color: r[5] ? C.red : C.text, fontSize: 11, align: 'right', fill: { color: bg } } },
        { text: r[2], options: { color: C.text, fontSize: 11, fill: { color: bg } } },
        { text: r[3], options: { color: C.muted, fontSize: 11, fill: { color: bg } } },
        { text: r[4], options: { color: C.text, fontSize: 11, align: 'right', fill: { color: bg } } }
      ];
    })
  ];
  s.addTable(tbl, tableOpts(0.4, 1.75, 12.55, [2.6, 1.0, 2.6, 4.95, 1.4], 0.5));
  s.addShape('roundRect', { x: 0.4, y: 6.0, w: 12.55, h: 0.9, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Pattern: ', options: { bold: true, color: C.accent } },
    { text: 'the deals that won (Scotiabank) had EB + technical engagement; the deals that lost (DTCC, Industry Dive) didn\'t. ', options: { color: C.text } },
    { text: 'Same lesson, second month — June is about not learning it a third time.', options: { bold: true, color: C.green } }
  ], { x: 0.6, y: 6.07, w: 12.2, h: 0.75, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S15 COACHING =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 15, 'Coaching Focus — June');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 2.6, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 0.4, y: 1.4, w: 0.08, h: 2.6, fill: { color: C.blue }, line: { width: 0 } });
  s.addText('PRIMARY — DIAL DISCIPLINE', { x: 0.6, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 12, color: C.blue, bold: true, charSpacing: 1 });
  s.addText([
    { text: '• Blocks → ≥15 dials/day, logged in HubSpot, reviewed weekly\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '• The May gap was execution inside the block, not the calendar\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '• Walk one day\'s call log in each 1:1', options: { color: C.text, fontSize: 11 } }
  ], { x: 0.6, y: 2.0, w: 5.8, h: 1.9, fontFace: F.body });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 2.6, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 6.8, y: 1.4, w: 0.08, h: 2.6, fill: { color: C.accent }, line: { width: 0 } });
  s.addText('SECONDARY — MEDDICC BREADTH', { x: 7.0, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 12, color: C.accent, bold: true, charSpacing: 1 });
  s.addText([
    { text: '• Full Scotiabank motion on every $10k+ late-stage deal\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '• 1 deal card walked through every weekly 1:1\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '• Disqualification-at-day-60 enforced (Industry Dive lesson)', options: { color: C.text, fontSize: 11 } }
  ], { x: 7.0, y: 2.0, w: 5.8, h: 1.9, fontFace: F.body });

  s.addText('THEME & CARRY-FORWARD', { x: 0.4, y: 4.2, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '✓  New theme — Sustain, don\'t sprint. ', options: { color: C.green, bold: true, fontSize: 12 } },
    { text: 'May was a closing sprint; June must run conversion and prospecting in parallel.\n\n', options: { color: C.text, fontSize: 12 } },
    { text: '•  Carried forward: ', options: { color: C.accent, bold: true, fontSize: 12 } },
    { text: 'Value Narratives on $10k+ deals · Exec Summaries on top 3 · Discovery depth (the Scotiabank standard).\n\n', options: { color: C.text, fontSize: 12 } },
    { text: '•  The one-line goal: ', options: { color: C.accent, bold: true, fontSize: 12 } },
    { text: 'lift the revenue line AND the pipe-gen line in the same month — the thing neither April nor May managed.', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.4, y: 4.55, w: 12.6, h: 2.3, fontFace: F.body });
}

const OUT = 'May_2026_NAM_MBR.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Wrote ' + OUT)).catch(e => { console.error(e); process.exit(1); });
