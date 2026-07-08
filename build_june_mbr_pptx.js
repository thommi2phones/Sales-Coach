// June 2026 NAM MBR — PPTX builder (rebuildable from this script)
// Mirrors the May_2026_NAM_MBR structure; June numbers from
// June_2026_Performance_Review.md and June_2026_NAM_MBR_Brief.md
const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'June 2026 NAM MBR — Thomas R Lindsey';
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
    { text: { text: 'June NAM MBR · Thomas R Lindsey · 07.08.2026', options: { x: 0.4, y: 7.15, w: 8, h: 0.25, fontFace: F.body, fontSize: 8, color: C.muted } } },
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
    slide.addShape('rect', { x, y: 1.4, w, h: 5.45, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    slide.addShape('rect', { x, y: 1.4, w: 0.06, h: 5.45, fill: { color: item.accent }, line: { width: 0 } });
    slide.addText(item.title, { x: x + 0.2, y: 1.58, w: w - 0.3, h: 0.4, fontFace: F.body, fontSize: 12.5, color: item.accent, bold: true, charSpacing: 2 });
    slide.addText(item.body, { x: x + 0.2, y: 2.15, w: w - 0.3, h: 4.55, fontFace: F.body, fontSize: 12.5, valign: 'top' });
  });
};

const tableOpts = (x, y, w, colW, rowH) => ({ x, y, w, colW, rowH, fontFace: F.body, border: { type: 'solid', pt: 0.5, color: C.border } });

// ===== S1 TITLE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  s.background = { color: C.navy };
  s.addText('MONTHLY BUSINESS REVIEW', { x: 0.7, y: 1.8, w: 12, h: 0.4, fontFace: F.body, fontSize: 13, color: 'A0B0D0', bold: true, charSpacing: 5 });
  s.addText('June NAM MBR', { x: 0.7, y: 2.3, w: 12, h: 1.6, fontFace: F.head, fontSize: 70, bold: true, color: C.white });
  s.addText('Thomas R Lindsey  ·  07.08.2026  ·  Supermetrics', { x: 0.7, y: 4.0, w: 12, h: 0.4, fontFace: F.body, fontSize: 16, color: 'A0B0D0' });
  s.addShape('roundRect', { x: 0.7, y: 4.7, w: 11.6, h: 0.55, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.27 });
  s.addText('✓  Revenue held 87% — 2nd straight month  ·  pipe-gen recovered ~96%  ·  but 0 logged dials & every $10k+ deal closed lost', { x: 0.85, y: 4.72, w: 11.3, h: 0.5, fontFace: F.body, fontSize: 12.5, color: C.orange, bold: true });
}

// ===== S2 OVERVIEW =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 2, 'June Overview — Held the Line, Showed the Cracks');
  s.addText('REVENUE GOALS', { x: 0.4, y: 1.45, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 3 });
  card(s, 0.4, 1.8, 6.2, 1.5, C.green, 'MAY REVENUE (LAST MONTH)', '$40,379 / 86%', 'Best revenue month in some time · Scotiabank $35k anchor');
  card(s, 6.8, 1.8, 6.2, 1.5, C.green, 'JUNE ACTUAL REVENUE', '$40,605 / 87%', 'Second straight strong month · but AMBSE $10k + 11 deals ≤$4.9k');
  s.addText('PIPELINE GOALS', { x: 0.4, y: 3.55, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 3 });
  card(s, 0.4, 3.9, 6.2, 1.5, C.red, 'MAY PIPELINE CREATED', '~$101,000 / ~54%', '20 deals · ~$5k avg ACV · calls collapsed to 39');
  card(s, 6.8, 3.9, 6.2, 1.5, C.green, 'JUNE PIPELINE CREATED', '~$179,000 / ~96%', '23 deals · ~$7.8k avg · recovered — but inbound/expansion, 0 dials');
  s.addShape('roundRect', { x: 0.4, y: 5.65, w: 12.55, h: 1.3, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The headline: ', options: { bold: true, color: C.orange } },
    { text: 'On the scorecard June is the "both engines on" month — revenue held 87% for a second straight month, pipe-gen recovered to ~96%. Underneath it is fragile: the number came from breadth + inbound + one carryover, ', options: { color: C.text } },
    { text: 'every $10k+ deal that closed, closed lost, and it was delivered on ZERO logged dials.', options: { bold: true, color: C.red } }
  ], { x: 0.6, y: 5.75, w: 12.15, h: 1.1, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S3 MAY PLAN PLAYED OUT =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 3, 'May Plan — How It Played Out in June');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 4.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('JUNE MITs (FROM MAY PLAN)', { x: 0.6, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '~  MIT #1 Restart the dial engine — SUPERSEDED BY A BOOK CHANGE\n', options: { color: C.orange, bold: true, fontSize: 14 } },
    { text: 'Mid-month the entire book shifted and the motion changed. Cold outbound / TOFU became non-critical — effort outside the Bronze base was futile as the role moved. Calls went 39 → 0, but that is a deprioritization, not a miss — these metrics don\'t really count this month.\n\n', options: { color: C.muted, fontSize: 12.5 } },
    { text: '✓  MIT #2 MEDDICC on $10k+ — TOOK HOLD IN CALL PREP\n', options: { color: C.green, bold: true, fontSize: 14 } },
    { text: 'The Call Prep blocks became genuinely useful for framing MEDDICC-based discovery questions — the questioning discipline that now carries into the account motion. Full breadth across every $10k+ deal didn\'t land, but the habit did.', options: { color: C.muted, fontSize: 12.5 } }
  ], { x: 0.6, y: 2.0, w: 5.8, h: 3.5, fontFace: F.body, valign: 'top' });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 4.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('WHAT HAPPENED INSTEAD', { x: 7.0, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '✓  Revenue held 87% + pipe-gen recovered ~96%\n', options: { color: C.green, bold: true, fontSize: 14 } },
    { text: 'Both lines healthy — on breadth (11 wins ≤$4.9k), one carryover (AMBSE $10k), inbound demos and expansion. The expansion motion is the tell for what comes next.\n\n', options: { color: C.muted, fontSize: 12.5 } },
    { text: '~  The motion changed underneath the plan\n', options: { color: C.orange, bold: true, fontSize: 14 } },
    { text: 'The book restructured mid-month — cold outbound deprioritized, focus shifting to the installed base. The May plan\'s dial + MEDDICC-breadth targets were overtaken by the KAM transition now underway.', options: { color: C.muted, fontSize: 12.5 } }
  ], { x: 7.0, y: 2.0, w: 5.8, h: 3.5, fontFace: F.body, valign: 'top' });

  s.addShape('roundRect', { x: 0.4, y: 5.65, w: 12.55, h: 1.3, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The one-sentence read: ', options: { bold: true, color: C.accent } },
    { text: 'June held revenue at 87% and recovered pipe-gen — and mid-month the book and role changed, retiring the outbound motion the May plan was built on. ', options: { color: C.text } },
    { text: 'July is a transition month: a new KAM book, not a dial sprint.', options: { bold: true, color: C.orange } }
  ], { x: 0.6, y: 5.75, w: 12.15, h: 1.1, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S4 REVIEW =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 4, 'June Review');
  col3(s, [
    { accent: C.green, title: 'WHAT I DID WELL', body: [
      { text: '✓  Two straight months near quota — 86% → 87%. ', options: { color: C.green, bold: true } },
      { text: 'The April slump is firmly behind; back-to-back is durable.\n\n', options: { color: C.muted } },
      { text: '✓  Pipe-gen recovered to ~96% — ~$179k / 23 deals / ~$7.8k avg after May\'s ~54% scare\n\n', options: { color: C.text } },
      { text: '✓  AMBSE finally closed ($10k) — the April "main swing" / May carryover is off the board\n\n', options: { color: C.text } },
      { text: '✓  Expansion motion working — BRP (300-acct), Tracer, YMCA, Promega, CFG\n\n', options: { color: C.text } },
      { text: '✓  Meetings held up (62, up from 54) — inbound demand real and being worked', options: { color: C.text } }
    ] },
    { accent: C.red, title: 'WHERE I NEED TO IMPROVE', body: [
      { text: '~  Outbound wound down with the book change — ', options: { color: C.orange, bold: true } },
      { text: 'calls 39 → 0 as cold TOFU was deprioritized. Not a miss, but the KAM motion still needs its own activity rhythm.\n\n', options: { color: C.muted } },
      { text: '✗  Big deals still leaked — ', options: { color: C.red, bold: true } },
      { text: 'the $10k+ deals that reached a close were mostly overdue / dead ops that never got the full motion. Don\'t carry zombies into the new book.\n\n', options: { color: C.muted } },
      { text: '✗  Day-60 hygiene not enforced — ', options: { color: C.red, bold: true } },
      { text: 'stale deals inflated pipe, then died. Clean the book on day one of the KAM transition.\n\n', options: { color: C.muted } },
      { text: '~  Small-deal breadth won\'t scale — ', options: { color: C.orange, bold: true } },
      { text: 'a month of ≤$5k deals hit 87%, but the gold base is where the real ACV lives.', options: { color: C.muted } }
    ] },
    { accent: C.accent, title: 'KEY DEAL LEARNINGS', body: [
      { text: 'The composition tells the story: ', options: { bold: true, color: C.text } },
      { text: '11 of 12 wins were ≤$4.9k; the only $10k win (AMBSE) was already on the board.\n\n', options: { color: C.muted } },
      { text: 'Expansion is the durable engine: ', options: { bold: true, color: C.text } },
      { text: 'BRP, Tracer, YMCA, Promega, CFG all grew inside the installed base — the exact motion the new KAM book is built to run.\n\n', options: { color: C.muted } },
      { text: 'MEDDICC questioning took hold: ', options: { bold: true, color: C.text } },
      { text: 'Call Prep sharpened discovery on live deals — a habit to carry into account planning.\n\n', options: { color: C.muted } },
      { text: 'The $10k+ lesson still stands: ', options: { bold: true, color: C.text } },
      { text: 'big deals convert only with the full EB / multi-thread motion (Scotiabank proof) — now aimed at the gold base.', options: { color: C.muted } }
    ] }
  ]);
}

// ===== S5 PIPELINE GEN =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 5, 'June Pipeline Gen Recap');
  const headers = ['Type', 'Count', 'Value (ERR)', 'ACV', 'Notes'];
  const rows = [
    ['New Business (inbound)', '13', '~$95,000', '~$7,300', 'Mostly Growth/MCP; Shubert $12k, TaxAct $10k'],
    ['Cross-sell / Upsell', '7', '~$60,000', '~$8,600', 'BRP DACT $25k, RLE $20k, WestGate $7.5k, YMCA'],
    ['Large DACT new-biz', '3', '~$24,000', '~$8,000', 'Tadi $25k, MS Buildon $25k — early, prone to stall'],
    ['TOTAL', '23', '~$179,000', '~$7,800', '—']
  ];
  const tbl = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11, align: (h === 'Type' || h === 'Notes') ? 'left' : (h === 'Count' ? 'center' : 'right') } })),
    ...rows.map((r, i) => r.map((c, j) => ({
      text: c,
      options: { bold: i === 3, color: i === 3 ? C.white : C.text, fill: { color: i === 3 ? C.surface2 : C.surface }, fontSize: 11,
        align: j === 0 || j === 4 ? 'left' : (j === 1 ? 'center' : 'right') }
    })))
  ];
  s.addTable(tbl, tableOpts(0.4, 1.5, 12.55, [2.2, 1.2, 2.0, 1.6, 5.55], 0.45));
  s.addShape('rect', { x: 0.4, y: 3.85, w: 12.55, h: 0.5, fill: { color: '143A2E' }, line: { color: C.green, width: 0.5 } });
  s.addText([
    { text: 'Pipeline Goal: ', options: { bold: true, color: C.text } },
    { text: '$187,000  ', options: { color: C.text } },
    { text: '·  ~96%', options: { bold: true, color: C.green } },
    { text: '  ·  recovered from April\'s 103% dip through May\'s ~54%', options: { color: C.muted } }
  ], { x: 0.6, y: 3.9, w: 12.2, h: 0.4, fontFace: F.body, fontSize: 13, valign: 'middle' });
  s.addShape('roundRect', { x: 0.4, y: 4.9, w: 12.55, h: 1.7, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'So what: ', options: { bold: true, color: C.orange } },
    { text: 'The ~96% is genuinely good and avg ACV is back to ~$7.8k (May: ~$5k). But read where it came from — almost none is outbound-sourced (calls = 0). It\'s inbound new-business plus expansion (BRP, Tracer, RLE, WestGate, YMCA). Three of the biggest are $25k DACT deals (Tadi, MS Buildon, BRP) — early-stage and, on this account\'s history, exactly the profile that stalls at Qualification.\n\n', options: { color: C.text } },
    { text: 'The recovery is real but not durable or self-generated — the inbound machine working, not the rep prospecting. If inbound softens, there is no outbound floor.', options: { italic: true, color: C.muted, fontSize: 11 } }
  ], { x: 0.6, y: 5.0, w: 12.2, h: 1.5, fontFace: F.body, fontSize: 13, valign: 'middle' });
}

// ===== S6 YTD =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 6, 'YTD Pace & the Pipe → Revenue Correlation');
  // Pipe gen % table (left)
  s.addText('PIPELINE GEN % TO GOAL', { x: 0.4, y: 1.4, w: 6, h: 0.3, fontFace: F.body, fontSize: 12, color: C.muted, bold: true, charSpacing: 2 });
  const pg = [
    ['Month', 'Deals', '% Goal'],
    ['January', '9', '50%', C.red],
    ['February', '18', '90%', C.orange],
    ['March', '16', '89%', C.orange],
    ['April', '23', '103%', C.green],
    ['May', '20', '~54%', C.red],
    ['June', '23', '~96%', C.green]
  ];
  const pgT = [
    pg[0].map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 12, align: h === 'Month' ? 'left' : 'center' } })),
    ...pg.slice(1).map((r, i) => r.slice(0, 3).map((c, j) => ({
      text: c, options: { bold: i === 5 || j === 2, color: j === 2 ? r[3] : (i === 5 ? C.white : C.text), fill: { color: i === 5 ? '143A2E' : C.surface }, fontSize: 12.5, align: j === 0 ? 'left' : 'center' }
    })))
  ];
  s.addTable(pgT, tableOpts(0.4, 1.75, 6.1, [2.6, 1.7, 1.8], 0.35));
  // Revenue % table (right)
  s.addText('REVENUE % TO QUOTA — ALL MONTHS', { x: 6.9, y: 1.4, w: 6, h: 0.3, fontFace: F.body, fontSize: 12, color: C.muted, bold: true, charSpacing: 2 });
  const rv = [
    ['Month', 'Revenue', '% Quota'],
    ['January', '$52,590', '131%', C.green, false],
    ['February', '$43,667', '109%', C.green, false],
    ['March', '$58,465', '133%', C.green, false],
    ['April', '$17,213', '37%', C.red, false],
    ['May', '$40,379', '86%', C.green, false],
    ['June', '$40,605', '87%', C.green, false],
    ['YTD (Jan–Jun)', '$252,919', '96% pace', C.orange, true]
  ];
  const rvT = [
    rv[0].map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 12, align: h === 'Month' ? 'left' : (h === 'Revenue' ? 'right' : 'center') } })),
    ...rv.slice(1).map((r) => r.slice(0, 3).map((c, j) => ({
      text: c, options: { bold: r[4] || j === 2, color: j === 2 ? r[3] : (r[4] ? C.white : C.text), fill: { color: r[4] ? '3D2E0D' : C.surface }, fontSize: 12.5, align: j === 0 ? 'left' : (j === 1 ? 'right' : 'center') }
    })))
  ];
  s.addTable(rvT, tableOpts(6.9, 1.75, 6.1, [2.3, 2.2, 1.6], 0.31));

  // Bottom-left: correlation & lead time
  s.addShape('rect', { x: 0.4, y: 4.5, w: 6.2, h: 2.45, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 0.4, y: 4.5, w: 0.06, h: 2.45, fill: { color: C.accent }, line: { width: 0 } });
  s.addText('PIPE → REVENUE: WHAT WE\'RE ACTUALLY TRACKING', { x: 0.6, y: 4.6, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 12, color: C.accent, bold: true, charSpacing: 1 });
  s.addText([
    { text: 'Correlation is weak at every lag ', options: { color: C.text, fontSize: 12.5 } },
    { text: '(r = 0.01–0.28, 18 mo)', options: { bold: true, color: C.orange, fontSize: 12.5 } },
    { text: ' — pipeline is lumpy, no clean fixed lead time.\n', options: { color: C.text, fontSize: 12.5 } },
    { text: 'Stable ratio: ', options: { bold: true, color: C.green, fontSize: 12.5 } },
    { text: '~29% of trailing-3-mo pipe', options: { bold: true, color: C.text, fontSize: 12.5 } },
    { text: ' converts to revenue.\n', options: { color: C.text, fontSize: 12.5 } },
    { text: 'So what: ', options: { bold: true, color: C.text, fontSize: 12.5 } },
    { text: 'pipe-gen % is a loose predictor — expansion + renewal timing (the KAM motion) moves revenue more than raw new-pipe volume.', options: { color: C.muted, fontSize: 12.5 } }
  ], { x: 0.6, y: 4.95, w: 5.8, h: 1.95, fontFace: F.body, valign: 'top' });

  // Bottom-right: Jul-Oct projection
  s.addShape('rect', { x: 6.8, y: 4.5, w: 6.2, h: 2.45, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 6.8, y: 4.5, w: 0.06, h: 2.45, fill: { color: C.green }, line: { width: 0 } });
  s.addText('PROJECTED ATTAINMENT — JUL → OCT', { x: 7.0, y: 4.6, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 12, color: C.green, bold: true, charSpacing: 1 });
  const proj = [
    ['Month', 'Quota', 'Proj Rev', 'Pace'],
    ['July', '$23.4k', '$18.8k conf', '81%', C.green],
    ['August', '$46.8k', '~$40k', '~86%', C.orange],
    ['September', '$46.8k', '~$40k+', '~86%+', C.orange],
    ['October', '$46.8k', '~$40k+', '~86%+', C.orange]
  ];
  const projT = [
    proj[0].map(h => ({ text: h, options: { bold: true, color: C.muted, fontSize: 11, fill: { color: C.surface }, align: h === 'Month' ? 'left' : 'right' } })),
    ...proj.slice(1).map(r => [
      { text: r[0], options: { color: C.text, fontSize: 11.5, fill: { color: C.surface2 } } },
      { text: r[1], options: { color: C.text, fontSize: 11.5, align: 'right', fill: { color: C.surface2 } } },
      { text: r[2], options: { color: C.text, fontSize: 11.5, align: 'right', fill: { color: C.surface2 } } },
      { text: r[3], options: { color: r[4], fontSize: 11.5, align: 'right', bold: true, fill: { color: C.surface2 } } }
    ])
  ];
  s.addTable(projT, tableOpts(7.0, 4.95, 5.8, [1.7, 1.2, 1.5, 1.4], 0.28));
  s.addText('Jul = confirmed FC ($18.8k, 81%) with weighted upside to $46.1k (197%). Aug–Oct pace ~85% at the current ~$43k/mo pipe × 29% yield — the KAM book\'s expansion + renewals must close the ~15% gap (Sep–Oct carry seasonal upside; both strong in 2025).', { x: 7.0, y: 6.42, w: 5.8, h: 0.5, fontFace: F.body, fontSize: 9.5, color: C.muted, italic: true });
}

// ===== S7 CLOSED EXECUTION =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 7, 'June Closed Deal Execution — The Real Story');
  const stats = [
    { label: 'REVENUE (WON)', val: '$40.6k', sub: '87% to quota · 12 deals', color: C.green },
    { label: 'BIGGEST WIN', val: '$10k', sub: 'AMBSE · carryover, not new', color: C.orange },
    { label: 'CLOSED LOST (Y1 ACV)', val: '~$81k', sub: '8 deals · mostly stale/dead ops', color: C.red },
    { label: 'NET CLOSED ACTIVITY', val: '−$40k', sub: 'Won minus lost', color: C.red }
  ];
  stats.forEach((st, i) => {
    const x = 0.4 + i * 3.2;
    s.addShape('rect', { x, y: 1.4, w: 3.0, h: 1.3, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addText(st.label, { x: x + 0.15, y: 1.5, w: 2.8, h: 0.3, fontFace: F.body, fontSize: 9, color: C.muted, bold: true, charSpacing: 1 });
    s.addText(st.val, { x: x + 0.15, y: 1.8, w: 2.8, h: 0.6, fontFace: F.head, fontSize: 28, bold: true, color: st.color });
    s.addText(st.sub, { x: x + 0.15, y: 2.4, w: 2.8, h: 0.25, fontFace: F.body, fontSize: 9, color: C.muted });
  });
  s.addText('THE LOSSES THAT MATTERED', { x: 0.4, y: 2.95, w: 12, h: 0.3, fontFace: F.body, fontSize: 12.5, color: C.muted, bold: true, charSpacing: 2 });
  const mk = (t, c, b, sz) => ({ text: t, options: { color: c, bold: !!b, fontSize: sz || 12.5, fill: { color: C.surface } } });
  const tbl = [
    [mk('Story', C.muted, true, 13), mk('Deals', C.muted, true, 13), { text: 'Lost ACV', options: { color: C.muted, bold: true, fontSize: 13, align: 'right', fill: { color: C.surface2 } } }, mk('So what', C.muted, true, 13)].map(o => (o.options.fill = { color: C.surface2 }, o)),
    [mk('Stale-deal deaths (300+ days)', C.text, true, 13), mk('Unilever Prestige ($40k, created 8/2025), AdRoll ($13k, created 6/2025)', C.text, false, 12), { text: '~$53,000', options: { color: C.red, bold: true, fontSize: 13, align: 'right', fill: { color: C.surface } } }, mk('~10–12 mo cycles, no real next step. The Industry Dive death, repeated.', C.muted, false, 12)],
    [mk('$10k+ that never got the motion', C.text, true, 13), mk('AMN ($12k, in the April pipe) — overdue / dead op', C.text, false, 12), { text: '~$12,000', options: { color: C.red, bold: true, fontSize: 13, align: 'right', fill: { color: C.surface } } }, mk('MEDDICC breadth never reached it; no EB. A dead op, not a live-deal loss.', C.muted, false, 12)],
    [mk('Mid / small cleanup', C.text, true, 13), mk('Charter One ($6k), Furnished Finder ($4k), Poppy ($2k), PSA ($2k), AHS ($2k)', C.text, false, 12), { text: '~$16,000', options: { color: C.red, bold: true, fontSize: 13, align: 'right', fill: { color: C.surface } } }, mk('Long-cycle small deals dispositioned — the healthy part of the number.', C.muted, false, 12)]
  ];
  s.addTable(tbl, tableOpts(0.4, 3.3, 12.55, [2.8, 4.6, 1.4, 3.75], 0.66));
  s.addShape('roundRect', { x: 0.4, y: 6.05, w: 12.55, h: 0.9, fill: { color: '3A2520' }, line: { color: C.red, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The read: ', options: { bold: true, color: C.red } },
    { text: 'the $10k+ that closed lost were mostly overdue / dead ops (Unilever, AdRoll, AMN) that never got the full motion — not live deals fumbled. (Celebrands excluded — ~$1k closed, split across two ops.) ', options: { color: C.text } },
    { text: 'The KAM job: give the gold base the Scotiabank-grade motion so the real ACV converts.', options: { bold: true, color: C.orange } }
  ], { x: 0.6, y: 6.12, w: 12.2, h: 0.75, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S8 JULY MITs =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 8, 'July MIT — Get a Handle on the New KAM Book');
  // Context banner — the motion change
  s.addShape('roundRect', { x: 0.4, y: 1.35, w: 12.55, h: 1.15, fill: { color: '231D3D' }, line: { color: C.accent, width: 1.5 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The motion is changing. ', options: { bold: true, color: C.accent, fontSize: 13 } },
    { text: 'We are reverting from the all-bound AE motion (one AE owns inbound + outbound + expansion) back to the 2024 customer-growth model I led. I move from manager to ', options: { color: C.text, fontSize: 13 } },
    { text: 'Key Account Manager (KAM)', options: { bold: true, color: C.text, fontSize: 13 } },
    { text: ' on the brand side — owning a defined book of accounts, not chasing net-new cold.', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 1.45, w: 12.15, h: 0.95, fontFace: F.body, valign: 'middle' });

  // Single MIT card
  s.addShape('rect', { x: 0.4, y: 2.7, w: 12.55, h: 3.15, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 0.4, y: 2.7, w: 0.08, h: 3.15, fill: { color: C.blue }, line: { width: 0 } });
  s.addShape('ellipse', { x: 0.75, y: 2.95, w: 0.8, h: 0.8, fill: { color: C.blue }, line: { width: 0 } });
  s.addText('1', { x: 0.75, y: 2.98, w: 0.8, h: 0.75, fontFace: F.head, fontSize: 32, bold: true, color: C.white, align: 'center' });
  s.addText('LEARN THE BOOK & BUILD THE PATH TO QUOTA', { x: 1.7, y: 3.0, w: 11.0, h: 0.5, fontFace: F.head, fontSize: 17, bold: true, color: C.blue, valign: 'middle' });
  s.addText([
    { text: 'The book: ', options: { bold: true, color: C.text, fontSize: 13 } },
    { text: '~75 Gold accounts + ~70 Scaled / Bronze = ~145 accounts. ', options: { color: C.text, fontSize: 13 } },
    { text: 'The whole of July is understanding this book cold and finding where the opportunity is to make quota for the rest of the year.\n\n', options: { color: C.muted, fontSize: 12 } },
    { text: '•  Map & tier the Gold base — health, spend, whitespace, renewal timing\n', options: { color: C.text, fontSize: 12 } },
    { text: '•  Build account plans for the top Gold accounts (biggest ARR / nearest renewals first)\n', options: { color: C.text, fontSize: 12 } },
    { text: '•  Surface expansion + renewal opportunity — e.g. Hawaiian Airlines $20k MIC+DWH renewal\n', options: { color: C.text, fontSize: 12 } },
    { text: '•  Identify the 2–3 plays with the clearest path to revenue in H2\n', options: { color: C.text, fontSize: 12 } },
    { text: '•  No cold-call focus — the motion is retain, deepen, and expand the installed base', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.75, y: 3.75, w: 11.9, h: 2.05, fontFace: F.body, valign: 'top' });

  s.addShape('roundRect', { x: 0.4, y: 6.05, w: 12.55, h: 0.9, fill: { color: '13392D' }, line: { color: C.green, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'This is the only MIT this month. ', options: { bold: true, color: C.green, fontSize: 13 } },
    { text: 'July is a ramp: success is a clear, tiered book strategy and coverage of the Gold base — not an activity or dial count. Get the book right and H2 quota has a plan behind it.', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 6.12, w: 12.15, h: 0.75, fontFace: F.body, valign: 'middle' });
}

// ===== S9 FUNNEL MATH =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 9, 'July Plan — KAM Book & Path to Quota');
  // Left: the book & the math
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('THE BOOK & THE MATH', { x: 0.6, y: 1.5, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '• July revenue quota: ', options: { color: C.text, fontSize: 13 } }, { text: '$23,400', options: { bold: true, color: C.orange, fontSize: 13 } }, { text: ' (reduced; full $46,750 resumes Aug)\n', options: { color: C.text, fontSize: 13 } },
    { text: '• The book: ', options: { color: C.text, fontSize: 13 } }, { text: '~145 accounts', options: { bold: true, color: C.text, fontSize: 13 } }, { text: ' — 75 Gold + 70 Scaled/Bronze\n', options: { color: C.text, fontSize: 13 } },
    { text: '• Open July pipe: ', options: { color: C.text, fontSize: 13 } }, { text: '$106.9k', options: { bold: true, color: C.green, fontSize: 13 } }, { text: ' · 457% coverage · $46.1k weighted (197%)\n', options: { color: C.text, fontSize: 13 } },
    { text: '• The July math is ', options: { color: C.text, fontSize: 13 } }, { text: 'book coverage', options: { bold: true, color: C.blue, fontSize: 13 } }, { text: ', not cold-funnel throughput — cover the Gold base and surface the expansion/renewal that funds H2 quota.', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 1.9, w: 5.8, h: 2.1, fontFace: F.body, valign: 'top' });
  s.addText('PATH TO H2 QUOTA', { x: 0.6, y: 4.05, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 12, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '• Retain the renewals in the book (e.g. Hawaiian Airlines $20k)\n', options: { color: C.text, fontSize: 13 } },
    { text: '• Expand the installed base — June\'s expansion wins (BRP, Tracer, YMCA) are the template\n', options: { color: C.text, fontSize: 13 } },
    { text: '• Convert the $106.9k open with MEDDICC discipline on every $10k+', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 4.4, w: 5.8, h: 1.45, fontFace: F.body, valign: 'top' });

  // Right: June activity baseline
  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('JUNE ACTIVITY BASELINE (HubSpot / Command Center)', { x: 7.0, y: 1.5, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 12, color: C.muted, bold: true, charSpacing: 1 });
  const actRows = [
    ['Meetings', '62', 'Healthy inbound + customer load (~2.8/day)', C.green],
    ['Emails logged', '394', 'Down, still the primary written channel', C.text],
    ['New contacts', '90', 'Lighter — top-of-list enrollment slowed', C.text],
    ['Calls / dials', '0', 'Cold dials deprioritized — intentional', C.orange]
  ];
  const actT = [
    [{ text: 'Activity', options: { bold: true, color: C.muted, fontSize: 12, fill: { color: C.surface } } },
     { text: 'June', options: { bold: true, color: C.muted, fontSize: 12, align: 'right', fill: { color: C.surface } } },
     { text: 'Read', options: { bold: true, color: C.muted, fontSize: 12, fill: { color: C.surface } } }],
    ...actRows.map(r => [
      { text: r[0], options: { color: C.text, fontSize: 12.5, fill: { color: C.surface2 } } },
      { text: r[1], options: { color: r[3], fontSize: 12.5, align: 'right', bold: true, fill: { color: C.surface2 } } },
      { text: r[2], options: { color: C.muted, fontSize: 11.5, fill: { color: C.surface2 } } }
    ])
  ];
  s.addTable(actT, tableOpts(7.0, 1.95, 5.8, [1.7, 0.8, 3.3], 0.48));
  s.addText([
    { text: 'The read: ', options: { bold: true, color: C.blue, fontSize: 13 } },
    { text: 'June ran on inbound demand + expansion — meetings healthy, cold outbound at zero by design. In the KAM motion that is the point: the KPI shifts from dials to account coverage, retention, and expansion pipeline.', options: { color: C.text, fontSize: 13 } }
  ], { x: 7.0, y: 4.55, w: 5.8, h: 1.75, fontFace: F.body, valign: 'top' });

  s.addShape('roundRect', { x: 0.4, y: 6.1, w: 12.55, h: 0.85, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The shift: ', options: { bold: true, color: C.accent } },
    { text: 'the old top-down funnel math (quota → pipe → dials) retires with the AE motion. ', options: { color: C.text } },
    { text: 'July\'s math is book coverage', options: { bold: true, color: C.blue } },
    { text: ' — 75 Gold accounts mapped and planned, expansion/renewal surfaced, the $106.9k open converted.', options: { color: C.text } }
  ], { x: 0.6, y: 6.18, w: 12.2, h: 0.7, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S10 DAILY ACTIVITY =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 10, 'July KAM Activity Plan — Working the Book');
  s.addText('The KAM motion replaces cold-outbound targets with book coverage. 22 working days in July (Jul 3 holiday). The measure this month is the Gold base mapped, planned, and touched — not dials.', { x: 0.4, y: 1.3, w: 12.6, h: 0.4, fontFace: F.body, fontSize: 11, color: C.muted });
  const headers = ['KAM Activity', 'Monthly', 'Weekly (avg)', 'Notes'];
  const rows = [
    ['Gold accounts reviewed & tiered', '75', '~19', 'Cover the full Gold base at least once — health, spend, whitespace', true],
    ['Account plans built', '~20', '~5', 'Top Gold accounts first (biggest ARR / nearest renewals)', false],
    ['Customer meetings / QBRs', '~40', '~10', 'Retention + expansion conversations across the book', false],
    ['Expansion / renewal ops surfaced', '~10', '~2-3', 'Whitespace + renewal timing (e.g. Hawaiian Airlines $20k)', false],
    ['Open pipeline advanced', '$106.9k', 'weekly', 'MEDDICC discipline on every $10k+ open deal', false],
    ['Cold dials', '0', '—', 'Deprioritized in the KAM motion — not a July KPI', false]
  ];
  const tblD = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.muted, fontSize: 11, fill: { color: C.surface2 }, align: h === 'KAM Activity' || h === 'Notes' ? 'left' : 'right' } })),
    ...rows.map(r => [
      { text: r[0], options: { bold: r[4], color: C.text, fontSize: 12, fill: { color: r[4] ? C.surface2 : C.surface } } },
      { text: r[1], options: { color: C.text, fontSize: 12, align: 'right', fill: { color: r[4] ? C.surface2 : C.surface } } },
      { text: r[2], options: { color: C.blue, fontSize: 12, align: 'right', bold: true, fill: { color: r[4] ? C.surface2 : C.surface } } },
      { text: r[3], options: { color: C.muted, fontSize: 10, fill: { color: r[4] ? C.surface2 : C.surface } } }
    ])
  ];
  s.addTable(tblD, tableOpts(0.4, 1.8, 12.55, [3.4, 1.5, 1.6, 6.05], 0.5));
  s.addShape('roundRect', { x: 0.4, y: 5.95, w: 12.55, h: 1.0, fill: { color: '13392D' }, line: { color: C.green, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The month-end bar: ', options: { bold: true, color: C.green } },
    { text: 'every Gold account touched and tiered, the top accounts with a written plan, and a named path to H2 quota. ', options: { color: C.text } },
    { text: 'Book coverage is the metric — not calls.', options: { bold: true, color: C.text } }
  ], { x: 0.6, y: 6.05, w: 12.2, h: 0.85, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S11 PLAYBOOK =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 11, 'How I\'ll Work the Book — Weekly KAM Rhythm');
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 12.55, h: 0.9, fill: { color: '231D3D' }, line: { color: C.accent, width: 1.5 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The rhythm shifts ', options: { bold: true, color: C.accent, fontSize: 13 } },
    { text: 'from a cold-dial cadence to a book-management operating system — a repeatable week for covering the Gold base, advancing renewals + expansion, and keeping the book clean. Same discipline the Reclaim blocks built, pointed at the installed base.', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 1.4, w: 12.2, h: 0.75, fontFace: F.body, valign: 'middle' });
  const blocks = [
    ['Monday book review', 'Mondays · 60 min', 'Rank Gold accounts by ARR, renewal timing, and whitespace; set the week\'s target accounts and the plays to run.'],
    ['Account planning', 'Daily AM · 60 min', 'Build / update account plans for top Gold accounts — stakeholders, value realized, expansion paths.'],
    ['Customer conversations', 'Daily · 2–3 mtgs', 'QBRs, check-ins, and expansion discovery across the book — retention + growth, not net-new cold.'],
    ['Pipeline advancement', 'Daily PM · 45 min', 'MEDDICC on every $10k+ open op (RLE, Upwork, HA) — EB, next step, drive it forward.'],
    ['Renewal watch', 'Weekly · 30 min', 'Track upcoming renewals (e.g. Hawaiian Airlines $20k, Q3 2027) — no surprise churn, plan each early.'],
    ['Friday hygiene + MEDDICC', 'Weekly · 90 min', 'Day-60 disqualification sweep; update MEDDICC cards on every open op; keep the book clean.']
  ];
  const blkH = 0.72, blkY = 2.4;
  blocks.forEach((b, i) => {
    const y = blkY + i * blkH;
    s.addText([
      { text: '■  ', options: { color: C.accent, fontSize: 14 } },
      { text: b[0], options: { bold: true, color: C.text, fontSize: 13 } },
      { text: '   —   ', options: { color: C.muted, fontSize: 12 } },
      { text: b[1], options: { color: C.blue, fontSize: 11, bold: true } }
    ], { x: 0.4, y, w: 12.6, h: 0.3, fontFace: F.body });
    s.addText(b[2], { x: 0.7, y: y + 0.3, w: 12.3, h: 0.36, fontFace: F.body, fontSize: 10, color: C.muted });
  });
}

// ===== S12 TRUE PIPELINE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 12, 'July Forecast & Pipeline — Command Center · Jul 8');
  const fc = [
    { label: 'JULY QUOTA (REDUCED)', val: '$23.4k', sub: 'full $46.75k resumes Aug', color: C.orange },
    { label: 'CONFIRMED FC', val: '$18.8k', sub: '81% of quota · 4 committed', color: C.green },
    { label: 'WEIGHTED FC', val: '$46.1k', sub: '197% of quota', color: C.green },
    { label: 'OPEN PIPELINE', val: '$106.9k', sub: '457% · 22 July-close', color: C.text }
  ];
  fc.forEach((c, i) => {
    const x = 0.4 + i * 3.19;
    s.addShape('rect', { x, y: 1.35, w: 3.0, h: 1.05, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addText(c.label, { x: x + 0.15, y: 1.42, w: 2.8, h: 0.25, fontFace: F.body, fontSize: 8.5, color: C.muted, bold: true, charSpacing: 1 });
    s.addText(c.val, { x: x + 0.15, y: 1.64, w: 2.8, h: 0.5, fontFace: F.head, fontSize: 22, bold: true, color: c.color });
    s.addText(c.sub, { x: x + 0.15, y: 2.14, w: 2.8, h: 0.22, fontFace: F.body, fontSize: 8, color: C.muted });
  });
  const headers = ['Account', 'Opportunity', 'ERR', 'Stage', '% Close', 'Weighted'];
  const dealRows = [
    ['Shopify', 'X Ads | $2k', '$1,750', 'Committed', '90%', '$1,600', C.green, true],
    ['Upwork', 'MIC/MCP | $15k', '$11,500', 'Agreement · legal/DPA', '85%', '$9,800', C.green, true],
    ['Shronk', 'Growth | $2k', '$2,124', 'Committed', '75%', '$1,600', C.green, true],
    ['Kraken', 'MCP | $3.5k', '$3,468', 'Committed', '70%', '$2,400', C.green, true],
    ['HA', 'Gemini Agent | $7.5k', '$7,500', 'Qualification', '65%', '$4,900', C.blue, false],
    ['RLE', 'BQ | $20k', '$20,000', 'Proof of Value · main swing', '50%', '$10,000', C.accent, true],
    ['Other (~16)', 'Growth/MCP/Claude $2-5k', '~$60,500', '—', '~26%', '~$15,800', C.muted, false],
    ['TOTAL', '22 July-close deals', '$106,900', '—', '43%', '$46,100', C.green, true]
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
  s.addTable(tbl, tableOpts(0.4, 2.55, 12.55, [1.7, 2.7, 1.3, 2.55, 1.3, 3.0], 0.34));
  s.addShape('rect', { x: 0.4, y: 5.75, w: 12.55, h: 0.5, fill: { color: '3A2520' }, line: { color: C.red, width: 0.5 } });
  s.addText([
    { text: 'Swing / overdue (not in the $106.9k July-close): ', options: { bold: true, color: C.red } },
    { text: 'US Soccer $40k (overdue 5/28), BlueDoor $10k (overdue 6/25), IA $5k (overdue 5/28) — re-date or disqualify', options: { color: C.text } }
  ], { x: 0.6, y: 5.8, w: 12.2, h: 0.4, fontFace: F.body, fontSize: 11, valign: 'middle' });
  s.addShape('roundRect', { x: 0.4, y: 6.35, w: 12.55, h: 0.65, fill: { color: '143A2E' }, line: { color: C.green, width: 1 }, rectRadius: 0.08 });
  s.addText([
    { text: 'The read: ', options: { bold: true, color: C.green } },
    { text: 'July opens 197% weighted-covered with $18.8k already committed (81%) on the reduced $23.4k quota — coverage is not the risk. A no-anchor portfolio (largest weighted RLE $10k) converts on MEDDICC breadth; Upwork (85%, legal/DPA) is the strongest signal. HubSpot|MCP moved to Aug.', options: { color: C.text } }
  ], { x: 0.6, y: 6.4, w: 12.2, h: 0.6, fontFace: F.body, fontSize: 10.5, valign: 'middle' });
}

// ===== S13 HAWAIIAN AIRLINES ACCOUNT PLAN =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 13, 'Closing / Account Plan — Hawaiian Airlines (Gold)');
  s.addText('The marquee Gold-account plan. The other open ops are smaller (Growth/MCP $2–5k) and run on the standard cadence — no standalone plan needed.', { x: 0.4, y: 1.3, w: 12.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted });
  const cards = [
    { accent: C.orange, title: 'ACCOUNT SNAPSHOT', body: [
      { text: 'Tier: ', options: { bold: true, color: C.text, fontSize: 11 } }, { text: 'Gold · established, expanding customer\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Products: ', options: { bold: true, color: C.text, fontSize: 11 } }, { text: 'MIC + DWH (managed ingestion + data warehouse)\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Installed base: ', options: { bold: true, color: C.text, fontSize: 11 } }, { text: '$40k SNO (2023) + $20k MIC+DWH (won 2025) — proven, renewing account\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Open op: ', options: { bold: true, color: C.green, fontSize: 11 } }, { text: '$20k MIC+DWH renewal · Q3 2027 · 40% · created 9/2025', options: { color: C.muted, fontSize: 10 } }
    ] },
    { accent: C.green, title: 'THE PLAN — RETAIN + EXPAND', body: [
      { text: '•  Retain: ', options: { bold: true, color: C.text, fontSize: 11 } }, { text: 'lock the $20k Q3-2027 renewal early — map decision process + value realized before it hits the danger zone\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '•  Expand: ', options: { bold: true, color: C.text, fontSize: 11 } }, { text: 'DWH growth, added sources / seats — the June expansion template (BRP, Tracer) applied here\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '•  Multi-thread: ', options: { bold: true, color: C.text, fontSize: 11 } }, { text: 'EB + technical / data-eng stakeholders, not one champion\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '•  QBR: ', options: { bold: true, color: C.text, fontSize: 11 } }, { text: 'value review — quantify ingestion / warehouse ROI delivered to date', options: { color: C.muted, fontSize: 10 } }
    ] },
    { accent: C.accent, title: 'MEDDICC + NEXT STEPS', body: [
      { text: 'E', options: { bold: true, color: C.accent, fontSize: 11 } }, { text: 'B — confirm the economic buyer for the 2027 renewal\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'M', options: { bold: true, color: C.accent, fontSize: 11 } }, { text: 'etrics — quantify data value delivered (usage, time saved)\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'D', options: { bold: true, color: C.accent, fontSize: 11 } }, { text: 'ecision Process — map the renewal approval path\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'C', options: { bold: true, color: C.accent, fontSize: 11 } }, { text: 'hampion — confirm + coach the internal advocate\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'Next steps: ', options: { bold: true, color: C.green, fontSize: 11 } }, { text: 'book QBR (Jul) · account plan doc (Jul) · renewal strategy locked by Q1 2027', options: { color: C.muted, fontSize: 10 } }
    ] }
  ];
  const cw = 4.05, gap = 0.2, sx = 0.4, sy = 1.75, ch = 4.4;
  cards.forEach((p, i) => {
    const x = sx + i * (cw + gap);
    s.addShape('rect', { x, y: sy, w: cw, h: ch, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y: sy, w: 0.05, h: ch, fill: { color: p.accent }, line: { width: 0 } });
    s.addText(p.title, { x: x + 0.15, y: sy + 0.12, w: cw - 0.25, h: 0.3, fontFace: F.body, fontSize: 10, bold: true, color: p.accent, charSpacing: 1 });
    s.addText(p.body, { x: x + 0.15, y: sy + 0.5, w: cw - 0.25, h: ch - 0.6, fontFace: F.body, valign: 'top' });
  });
  s.addShape('roundRect', { x: 0.4, y: 6.35, w: 12.55, h: 0.65, fill: { color: '13392D' }, line: { color: C.green, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The KAM template: ', options: { bold: true, color: C.green } },
    { text: 'retain the renewal, expand the base, multi-thread the account. Hawaiian is the model — the rest of the Gold book gets the same treatment through July.', options: { color: C.text } }
  ], { x: 0.6, y: 6.42, w: 12.2, h: 0.5, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S13b MEDDICC REPLICATE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '13b', 'Appendix / Leave-Behind — MEDDICC Proven on Scotiabank');
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 12.55, h: 0.85, fill: { color: '143A2E' }, line: { color: C.green, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Scotiabank proved MEDDICC converts flagship deals. June forgot it — ', options: { bold: true, color: C.green, fontSize: 12 } },
    { text: 'every $10k+ deal died stale, unqualified, or without an EB. The July job is breadth: run the exact Scotiabank motion on every $10k+ deal, not just whatever inbound sends.', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 1.38, w: 12.2, h: 0.7, fontFace: F.body, valign: 'middle' });
  const cols = [
    { accent: C.green, title: 'WHAT WON SCOTIABANK', body: [
      { text: 'E', options: { bold: true, color: C.green, fontSize: 13 } }, { text: 'conomic Buyer — engaged + decision path live through signature\n\n', options: { color: C.text, fontSize: 12.5 } },
      { text: 'Technical multi-thread', options: { bold: true, color: C.text, fontSize: 13 } }, { text: ' — engineering org, not just champion\n\n', options: { color: C.muted, fontSize: 12.5 } },
      { text: 'P', options: { bold: true, color: C.green, fontSize: 13 } }, { text: 'aper Process — security/DR docs early, confirmed received\n\n', options: { color: C.text, fontSize: 12.5 } },
      { text: 'Prep brief', options: { bold: true, color: C.text, fontSize: 13 } }, { text: ' before the technical call — objections anticipated', options: { color: C.muted, fontSize: 12.5 } }
    ] },
    { accent: C.red, title: 'WHAT LOST JUNE\'S $10k+ DEALS', body: [
      { text: '✗  No EB motion ran (AMN — sat live for months)\n\n', options: { color: C.red, fontSize: 12.5, bold: true } },
      { text: '✗  ~10–12 month stale deaths, no next step (Unilever, AdRoll)\n\n', options: { color: C.red, fontSize: 12.5, bold: true } },
      { text: '✗  Died unqualified in ~10 days (Celebrands)\n\n', options: { color: C.red, fontSize: 12.5, bold: true } },
      { text: '✗  Day-60 disqualification rule not enforced', options: { color: C.red, fontSize: 12.5, bold: true } }
    ] },
    { accent: C.blue, title: 'THE MEDDICC OPERATING RHYTHM', body: [
      { text: '• Day 1: open MEDDICC card, rate R/Y/G\n\n', options: { color: C.text, fontSize: 12.5 } },
      { text: '• Every $10k+ deal: EB booked before pricing\n\n', options: { color: C.text, fontSize: 12.5 } },
      { text: '• Every technical call: written prep brief\n\n', options: { color: C.text, fontSize: 12.5 } },
      { text: '• Friday: disqualify anything 60d+ with no next step', options: { color: C.text, fontSize: 12.5 } }
    ] }
  ];
  const cw = 4.05, gap = 0.2, sx = 0.4, sy = 2.3, ch = 4.0;
  cols.forEach((p, i) => {
    const x = sx + i * (cw + gap);
    s.addShape('rect', { x, y: sy, w: cw, h: ch, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y: sy, w: 0.05, h: ch, fill: { color: p.accent }, line: { width: 0 } });
    s.addText(p.title, { x: x + 0.15, y: sy + 0.15, w: cw - 0.25, h: 0.3, fontFace: F.body, fontSize: 12, bold: true, color: p.accent, charSpacing: 1 });
    s.addText(p.body, { x: x + 0.15, y: sy + 0.6, w: cw - 0.25, h: ch - 0.7, fontFace: F.body, valign: 'top' });
  });
}

// ===== S15 COACHING =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 15, 'Coaching Focus — July');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 2.6, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 0.4, y: 1.4, w: 0.08, h: 2.6, fill: { color: C.blue }, line: { width: 0 } });
  s.addText('PRIMARY — RAMP INTO THE KAM MOTION', { x: 0.6, y: 1.58, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 13, color: C.blue, bold: true, charSpacing: 1 });
  s.addText([
    { text: '• Learn the new book cold — 75 Gold + 70 Scaled/Bronze (~145 accounts)\n\n', options: { color: C.text, fontSize: 13 } },
    { text: '• Build account plans for the top Gold accounts — ARR + renewal-timing first\n\n', options: { color: C.text, fontSize: 13 } },
    { text: '• Name the path to full-year quota from the installed base', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 2.05, w: 5.8, h: 1.85, fontFace: F.body, valign: 'top' });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 2.6, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 6.8, y: 1.4, w: 0.08, h: 2.6, fill: { color: C.accent }, line: { width: 0 } });
  s.addText('SECONDARY — RETAIN & EXPAND THE BASE', { x: 7.0, y: 1.58, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 13, color: C.accent, bold: true, charSpacing: 1 });
  s.addText([
    { text: '• Renewals first — Hawaiian Airlines $20k (Q3 2027) and every dated renewal, planned early\n\n', options: { color: C.text, fontSize: 13 } },
    { text: '• Expand the June-proven way — BRP, Tracer, YMCA are the template\n\n', options: { color: C.text, fontSize: 13 } },
    { text: '• MEDDICC discipline on the open $10k+ ops (RLE, Upwork, HA)', options: { color: C.text, fontSize: 13 } }
  ], { x: 7.0, y: 2.05, w: 5.8, h: 1.85, fontFace: F.body, valign: 'top' });

  s.addText('THEME & CARRY-FORWARD', { x: 0.4, y: 4.25, w: 12, h: 0.3, fontFace: F.body, fontSize: 12.5, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '✓  New theme — the motion change. ', options: { color: C.green, bold: true, fontSize: 13.5 } },
    { text: 'From the all-bound AE model to KAM (the 2024 customer-growth motion). July is a ramp, not a sprint — the deliverable is a book strategy, not an activity count.\n\n', options: { color: C.text, fontSize: 13.5 } },
    { text: '•  Carried forward: ', options: { color: C.accent, bold: true, fontSize: 13.5 } },
    { text: 'the MEDDICC questioning habit from Call Prep, and day-60 hygiene to clean the book on entry (no inherited zombies).\n\n', options: { color: C.text, fontSize: 13.5 } },
    { text: '•  The one-line goal: ', options: { color: C.accent, bold: true, fontSize: 13.5 } },
    { text: 'get the book right — a tiered Gold-base strategy with a named path to H2 quota. Coverage, not dials.', options: { color: C.text, fontSize: 13.5 } }
  ], { x: 0.4, y: 4.6, w: 12.6, h: 2.3, fontFace: F.body, valign: 'top' });
}

const OUT = 'June_2026_NAM_MBR.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Wrote ' + OUT)).catch(e => { console.error(e); process.exit(1); });
