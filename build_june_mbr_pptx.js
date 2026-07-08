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
    { text: '✗  MIT #1 Restart the dial engine — TOTAL FAIL\n', options: { color: C.red, bold: true, fontSize: 12 } },
    { text: '    The #1 fix was ≥15 dials/day. Calls went 39 → 0 — the wrong direction, against ~22 Call Prep blocks Reclaim scheduled. The block-to-dial gap went from partial (May) to total (June).\n\n', options: { color: C.muted, fontSize: 10 } },
    { text: '✗  MIT #2 MEDDICC breadth on every $10k+ deal — FAIL\n', options: { color: C.red, bold: true, fontSize: 12 } },
    { text: '    Every $10k+ deal that reached a close closed LOST: Unilever $40k, AdRoll $13k, AMN $12k, Celebrands $12k. The Scotiabank motion did not reach the big pipe.', options: { color: C.muted, fontSize: 10 } }
  ], { x: 0.6, y: 1.9, w: 5.8, h: 3.4, fontFace: F.body });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 4.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('WHAT HAPPENED INSTEAD', { x: 7.0, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '✓  Revenue held 87% + pipe-gen recovered ~96%\n', options: { color: C.green, bold: true, fontSize: 12 } },
    { text: '    Both lines looked healthy — but on breadth (11 wins ≤$4.9k), one carryover (AMBSE $10k), inbound demos and expansion, not the two committed fixes.\n\n', options: { color: C.muted, fontSize: 10 } },
    { text: '~  The structure held; the execution didn\'t\n', options: { color: C.orange, bold: true, fontSize: 12 } },
    { text: '    ~22 Call Prep blocks on the calendar, 0 dials inside them. Scheduling the intention is not making the call — proven a second, more complete time.', options: { color: C.muted, fontSize: 10 } }
  ], { x: 7.0, y: 1.9, w: 5.8, h: 3.4, fontFace: F.body });

  s.addShape('roundRect', { x: 0.4, y: 5.65, w: 12.55, h: 1.3, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The one-sentence read: ', options: { bold: true, color: C.accent } },
    { text: 'June didn\'t fix May\'s gap — it papered over it with volume and expansion. A green revenue number is hiding a stone-cold outbound engine and a $10k+ pipe leaking out the bottom. ', options: { color: C.text } },
    { text: 'That\'s a Q3 revenue risk, not a recovery.', options: { bold: true, color: C.orange } }
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
      { text: '✗  Calls went to ZERO — 39 → 0. ', options: { color: C.red, bold: true } },
      { text: '~22 Call Prep blocks, 0 dials. The #1 MIT moved backward.\n\n', options: { color: C.muted } },
      { text: '✗  Every $10k+ deal closed LOST — Unilever $40k, AdRoll $13k, AMN $12k, Celebrands $12k\n\n', options: { color: C.text } },
      { text: '✗  Stale deals still dying at 300+ days — day-60 disqualification not enforced\n\n', options: { color: C.text } },
      { text: '✗  Self-generated pipe near zero — contacts halved (90); no outbound floor under the number', options: { color: C.text } }
    ] },
    { accent: C.accent, title: 'KEY DEAL LEARNINGS', body: [
      { text: 'The composition tells the story: ', options: { bold: true, color: C.text } },
      { text: '11 of 12 wins were ≤$4.9k; the only $10k win (AMBSE) was already on the board.\n\n', options: { color: C.muted } },
      { text: 'Unilever ($40k, created 8/2025): ', options: { bold: true, color: C.text } },
      { text: 'the ~10-month stale death the day-60 rule was meant to prevent.\n\n', options: { color: C.muted } },
      { text: 'AMN ($12k): ', options: { bold: true, color: C.text } },
      { text: 'sat as live pipe for months, never got an EB — the DTCC lesson, repeated.\n\n', options: { color: C.muted } },
      { text: 'The volume bill: ', options: { bold: true, color: C.text } },
      { text: 'two months off the phones is invisible in June\'s revenue but is the Q3 pipeline risk.', options: { color: C.muted } }
    ] }
  ]);
}

// ===== S4b THE $10k+ LOSSES =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '4b', 'The $10k+ Losses — Anatomy of the Pattern');
  s.addText('Every deal ≥$10k that closed in June closed lost — Unilever $40k, AdRoll $13k, AMN $12k, Celebrands $12k. The inverse of May\'s Scotiabank win. The same lesson, third month.', { x: 0.4, y: 1.3, w: 12.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted });
  const cols = [
    { accent: C.red, title: 'WHAT WENT WRONG', body: [
      { text: '✗  Unilever Prestige ($40k) — created 8/2025, ~10-month cycle that never had a real next step\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '✗  AdRoll ($13k) — created 6/2025, another ~12-month zombie that died loud\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '✗  AMN ($12k) — named in the April pipeline, sat as live pipe for months with no EB\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '✗  Celebrands ($12k) — entered and died in ~10 days, never qualified', options: { color: C.text, fontSize: 10 } }
    ] },
    { accent: C.orange, title: 'WHY THEY DIED', body: [
      { text: '• No EB / technical multi-thread — the motion that saved Scotiabank never ran\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '• Day-60 disqualification not enforced — zombies inflated pipe, then imploded\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '• Email-only follow-up on $10k+ deals — email doesn\'t advance or save the big deals\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: '• The rep was reacting to inbound, not driving the late-stage motion on the big pipe', options: { color: C.text, fontSize: 10, bold: true } }
    ] },
    { accent: C.green, title: 'THE REPLICABLE FIX', body: [
      { text: '• EB engaged before pricing on every $10k+ deal\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '• Prep brief + technical multi-thread — the Scotiabank standard\n\n', options: { color: C.text, fontSize: 10 } },
      { text: '• Security/legal docs pushed early; day-60 disqualification swept every Friday\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'So what: ', options: { color: C.green, fontSize: 10, bold: true } },
      { text: 'now the standard for Upwork, RLE, Shubert, TaxAct, BlueDoor — July\'s $10k+ pipe', options: { color: C.muted, fontSize: 10 } }
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
  slideTitle(s, 6, 'YTD 2026 — The Recovery, and the Catch');
  // Pipe gen table
  s.addText('PIPELINE GEN % TO GOAL', { x: 0.4, y: 1.45, w: 6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
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
    pg[0].map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11, align: h === 'Month' ? 'left' : 'center' } })),
    ...pg.slice(1).map((r, i) => r.slice(0, 3).map((c, j) => ({
      text: c, options: { bold: i === 5 || j === 2, color: j === 2 ? r[3] : (i === 5 ? C.white : C.text), fill: { color: i === 5 ? '143A2E' : C.surface }, fontSize: 12, align: j === 0 ? 'left' : 'center' }
    })))
  ];
  s.addTable(pgT, tableOpts(0.4, 1.8, 6.1, [2.6, 1.7, 1.8], 0.42));
  s.addText('50 → 90 → 89 → 103 → 54 → 96.  The pipe-gen arc recovered.', { x: 0.4, y: 4.95, w: 6.1, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
  // Revenue table
  s.addText('REVENUE % TO QUOTA', { x: 6.9, y: 1.45, w: 6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  const rv = [
    ['Month', 'Revenue', '% Quota'],
    ['April', '$17,213', '37%', C.red],
    ['May', '$40,379', '86%', C.green],
    ['June', '$40,605', '87%', C.green]
  ];
  const rvT = [
    rv[0].map(h => ({ text: h, options: { bold: true, color: C.muted, fill: { color: C.surface2 }, fontSize: 11, align: h === 'Month' ? 'left' : (h === 'Revenue' ? 'right' : 'center') } })),
    ...rv.slice(1).map((r, i) => r.slice(0, 3).map((c, j) => ({
      text: c, options: { bold: i === 2 || j === 2, color: j === 2 ? r[3] : (i === 2 ? C.white : C.text), fill: { color: i === 2 ? '143A2E' : C.surface }, fontSize: 12, align: j === 0 ? 'left' : (j === 1 ? 'right' : 'center') }
    })))
  ];
  s.addTable(rvT, tableOpts(6.9, 1.8, 6.1, [2.3, 2.2, 1.6], 0.42));
  s.addText('37 → 86 → 87.  Two straight months near quota — but calls: 280 → 39 → 0.', { x: 6.9, y: 3.15, w: 6.1, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });

  s.addShape('roundRect', { x: 0.4, y: 5.5, w: 12.55, h: 1.3, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The pattern: ', options: { bold: true, color: C.accent } },
    { text: 'both headline lines are green — revenue held, pipe-gen recovered. But the calls line (280 → 39 → 0) kept breaking, and every $10k+ close was a loss. ', options: { color: C.text } },
    { text: 'The scorecard is green; the engine underneath is cold.', options: { bold: true, color: C.orange } }
  ], { x: 0.6, y: 5.6, w: 12.2, h: 1.1, fontFace: F.body, fontSize: 14, valign: 'middle' });
}

// ===== S7 CLOSED EXECUTION =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 7, 'June Closed Deal Execution — The Real Story');
  const stats = [
    { label: 'REVENUE (WON)', val: '$40.6k', sub: '87% to quota · 12 deals', color: C.green },
    { label: 'BIGGEST WIN', val: '$10k', sub: 'AMBSE · carryover, not new', color: C.orange },
    { label: 'CLOSED LOST (Y1 ACV)', val: '~$93k', sub: '9 deals · every $10k+ lost', color: C.red },
    { label: 'NET CLOSED ACTIVITY', val: '−$52k', sub: 'Won minus lost', color: C.red }
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
    [mk('Stale-deal deaths (300+ days)', C.text, true, 11), mk('Unilever Prestige ($40k, created 8/2025), AdRoll ($13k, created 6/2025)', C.text, false, 10), { text: '~$53,000', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } }, mk('~10–12 mo cycles, no real next step. The Industry Dive death, repeated.', C.muted, false, 10)],
    [mk('$10k+ that never got the motion', C.text, true, 11), mk('AMN ($12k, in the April pipe), Celebrands ($12k, died in ~10d)', C.text, false, 10), { text: '~$24,000', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } }, mk('MEDDICC breadth (May MIT) never reached them. No EB; Celebrands died unqualified.', C.muted, false, 10)],
    [mk('Mid / small cleanup', C.text, true, 11), mk('Charter One ($6k), Furnished Finder ($4k), Poppy ($2k), PSA ($2k), AHS ($2k)', C.text, false, 10), { text: '~$16,000', options: { color: C.red, bold: true, fontSize: 11, align: 'right', fill: { color: C.surface } } }, mk('Long-cycle small deals dispositioned — the healthy part of the number.', C.muted, false, 10)]
  ];
  s.addTable(tbl, tableOpts(0.4, 3.3, 12.55, [2.8, 4.6, 1.4, 3.75], 0.62));
  s.addShape('roundRect', { x: 0.4, y: 6.05, w: 12.55, h: 0.9, fill: { color: '3A2520' }, line: { color: C.red, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The lesson, three months old: ', options: { bold: true, color: C.red } },
    { text: 'the $10k+ pipe converts only when the full MEDDICC/EB/security motion runs on it (Scotiabank $35k), and it dies when it doesn\'t (Unilever, AdRoll, AMN, Celebrands). ', options: { color: C.text } },
    { text: 'July\'s job: give every $10k+ deal the Scotiabank treatment.', options: { bold: true, color: C.orange } }
  ], { x: 0.6, y: 6.12, w: 12.2, h: 0.75, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S8 JULY MITs =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 8, 'July MITs');
  const mits = [
    { accent: C.blue, num: '1', title: 'TURN THE PHONES ON — MEASURED', body: [
      { text: 'Dials have been cold two months (39 → 0). The Call Prep block clearly isn\'t enough — pair every block with a ', options: { color: C.text, fontSize: 13 } },
      { text: '≥15 logged-dial/day target', options: { bold: true, color: C.text, fontSize: 13 } },
      { text: ', a same-day dial count, and a weekly 1:1 accountability check. 300+ dials for July. This is now non-negotiable and escalated.\n\n', options: { color: C.text, fontSize: 13 } },
      { text: 'The headline MIT — June hit the number on inbound while the outbound engine stayed stone cold. July refuses that.', options: { italic: true, color: C.muted, fontSize: 11 } } ] },
    { accent: C.accent, num: '2', title: 'CONVERT THE $10k+ PIPE — STOP LIVING ON <$5k BREADTH', body: [
      { text: 'June hit 87% on 11 small deals + one carryover while every $10k+ deal died. Run the full Scotiabank-style MEDDICC/EB/security motion across Upwork, RLE, Shubert, TaxAct, BlueDoor. Enforce day-60 disqualification (the Unilever/AdRoll lesson).\n\n', options: { color: C.text, fontSize: 13 } },
      { text: 'June proved I can hit the number on breadth. July must prove I can close big deals again — quality and self-generated volume at once.', options: { italic: true, color: C.muted, fontSize: 11 } } ] }
  ];
  mits.forEach((m, i) => {
    const x = 0.4 + i * 6.4;
    s.addShape('rect', { x, y: 1.5, w: 6.2, h: 5.0, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
    s.addShape('rect', { x, y: 1.5, w: 0.08, h: 5.0, fill: { color: m.accent }, line: { width: 0 } });
    s.addShape('ellipse', { x: x + 0.35, y: 1.75, w: 0.8, h: 0.8, fill: { color: m.accent }, line: { width: 0 } });
    s.addText(m.num, { x: x + 0.35, y: 1.78, w: 0.8, h: 0.75, fontFace: F.head, fontSize: 32, bold: true, color: C.white, align: 'center' });
    s.addText(m.title, { x: x + 1.3, y: 1.8, w: 4.7, h: 0.85, fontFace: F.head, fontSize: 15, bold: true, color: m.accent, charSpacing: 1, valign: 'middle' });
    s.addText(m.body, { x: x + 0.35, y: 2.8, w: 5.65, h: 3.5, fontFace: F.body, valign: 'top' });
  });
}

// ===== S9 FUNNEL MATH =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 9, 'July Funnel Math — Top-Down Derivation');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('LOCKED INPUTS', { x: 0.6, y: 1.5, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '• July revenue quota: ', options: { color: C.text, fontSize: 12 } }, { text: '$23,400', options: { bold: true, color: C.orange, fontSize: 12 } }, { text: ' (reduced — July only; full $46,750 resumes Aug)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Pipeline goal: ', options: { color: C.text, fontSize: 12 } }, { text: '$187,000', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' (held at full-month 4×)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Plan target: ', options: { color: C.text, fontSize: 12 } }, { text: '$200,000', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' (+7% buffer)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• ACV baseline: ', options: { color: C.text, fontSize: 12 } }, { text: '$11,000', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' (6-mo avg) — but June\'s actual win ACV was ~$3.4k; the gap IS the coaching point', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 1.85, w: 5.8, h: 1.7, fontFace: F.body });
  s.addText('BACKED-OUT TARGETS', { x: 0.6, y: 3.7, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '• Deals needed: ', options: { color: C.text, fontSize: 12 } }, { text: '~18', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' ($200k ÷ $11k)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Booked meetings: ', options: { color: C.text, fontSize: 12 } }, { text: '~60', options: { bold: true, color: C.text, fontSize: 12 } }, { text: '\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Contacts enrolled: ', options: { color: C.text, fontSize: 12 } }, { text: '~330', options: { bold: true, color: C.text, fontSize: 12 } }, { text: ' (June halved to 90 — push back up)\n', options: { color: C.text, fontSize: 12 } },
    { text: '• Calls: ', options: { color: C.text, fontSize: 12 } }, { text: '~300', options: { bold: true, color: C.green, fontSize: 12 } }, { text: ' — the explicit fix for the June miss (0 → 300)', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 4.05, w: 5.8, h: 1.8, fontFace: F.body });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 4.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('JUNE BASELINE → JULY TARGET', { x: 7.0, y: 1.5, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  const dataRows = [
    ['Calls logged', '0', '300', 'from zero', C.green, true],
    ['Contacts created', '90', '330', '+267%', C.green, false],
    ['Emails sent', '394', '1,000', '+154%', C.green, false],
    ['Meetings', '62', '60', '~flat', C.text, false],
    ['New deals', '23', '~18', 'fewer · larger', C.orange, false],
    ['Pipe generated', '~$179k', '$200k', '+12%', C.green, true]
  ].map(r => [
    { text: r[0], options: { color: C.text, fontSize: 11, bold: r[5], fill: { color: r[5] ? C.surface2 : C.surface } } },
    { text: r[1], options: { color: C.text, fontSize: 11, align: 'right', bold: r[5], fill: { color: r[5] ? C.surface2 : C.surface } } },
    { text: r[2], options: { color: C.text, fontSize: 11, align: 'right', bold: r[5], fill: { color: r[5] ? C.surface2 : C.surface } } },
    { text: r[3], options: { color: r[4], fontSize: 11, align: 'right', bold: true, fill: { color: r[5] ? C.surface2 : C.surface } } }
  ]);
  const tblData = [
    [{ text: 'Activity', options: { bold: true, color: C.muted, fontSize: 10, fill: { color: C.surface } } },
     { text: 'June', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } },
     { text: 'July', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } },
     { text: 'Δ', options: { bold: true, color: C.muted, fontSize: 10, align: 'right', fill: { color: C.surface } } }],
    ...dataRows
  ];
  s.addTable(tblData, tableOpts(7.0, 1.85, 5.8, [2.4, 1.0, 1.0, 1.4], 0.42));
  s.addText('The big lever is calls. June\'s 0 is the number to fix; everything else follows from putting dials back inside the ~22 Call Prep blocks that already exist and go unused.', { x: 7.0, y: 5.1, w: 5.8, h: 0.7, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
  s.addShape('roundRect', { x: 0.4, y: 6.1, w: 12.55, h: 0.85, fill: { color: C.surface2 }, line: { color: C.accent, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'The shift: ', options: { bold: true, color: C.accent } },
    { text: 'April proved volume. May proved conversion. June proved breadth+inbound. ', options: { color: C.text } },
    { text: 'July proves the outbound floor', options: { bold: true, color: C.green } },
    { text: ' — the dial engine back on, with Scotiabank-grade MEDDICC behind every $10k+ deal.', options: { color: C.text } }
  ], { x: 0.6, y: 6.18, w: 12.2, h: 0.7, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S10 DAILY ACTIVITY =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 10, 'July Daily Activity Targets');
  s.addText('22 working days in July (Jul 1–31; Jul 3 holiday for July 4 observed). The June lesson: a block on the calendar is not a dial made — zero dials on 22 blocks proves the block alone does nothing. July counts dials, not blocks.', { x: 0.4, y: 1.3, w: 12.6, h: 0.4, fontFace: F.body, fontSize: 11, color: C.muted });
  const headers = ['Activity', 'Monthly', 'Weekly (avg)', 'Daily (22d)', 'Notes'];
  const rows = [
    ['Calls / dials', '300', '~75', '~15/day', 'The June miss (0). Logged inside Call Prep blocks', true],
    ['New contacts enrolled', '330', '~83', '~15/day', 'Email-verified before enroll', false],
    ['Emails sent (cadence)', '1,000', '~250', '~46/day', 'Cadence queue worked AM', false],
    ['LinkedIn touches', '~210', '~52', '~10/day', '5 connect + 3 InMail + 3 comments', false],
    ['Net-new accounts worked', '~70', '~17', '~3-4/day', 'Mon list-build to put an outbound floor down', false],
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
    { text: '15 logged dials on every working day. ', options: { bold: true, color: C.text } },
    { text: 'June had the blocks (~22 Call Prep) and made 0 dials. The metric that matters is calls logged, not calendar coverage — zero is not an option.', options: { color: C.text } }
  ], { x: 0.6, y: 6.05, w: 12.2, h: 0.85, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S11 PLAYBOOK =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 11, 'How We Work the Activities — Dials Inside the Blocks');
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 12.55, h: 0.9, fill: { color: '231D3D' }, line: { color: C.accent, width: 1.5 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Commit: ', options: { bold: true, color: C.accent, fontSize: 13 } },
    { text: 'The Reclaim blocks already exist and held all June. The July change is ', options: { color: C.text, fontSize: 13 } },
    { text: 'accountability inside the block', options: { bold: true, color: C.text, fontSize: 13 } },
    { text: ' — every Call Prep block ends with dials logged in HubSpot, counted same-day, reviewed weekly in the 1:1. In July the dial count is the contract.', options: { color: C.text, fontSize: 13 } }
  ], { x: 0.6, y: 1.4, w: 12.2, h: 0.75, fontFace: F.body, valign: 'middle' });
  const blocks = [
    ['Account list build', 'Mondays · 60 min', 'Pull 20 net-new accounts to put an outbound floor under the inbound number: ZoomInfo + intent + 2025 closed-lost rework.'],
    ['Contact enrollment', 'Daily AM · 30 min', '~15 contacts/day · verify email before enroll (June halved to 90).'],
    ['Cadence + dials', 'Daily AM · 60 min', '~46 emails + ≥15 dials due today — LOGGED. The dials are the part that hit zero in June.'],
    ['LinkedIn block', 'Daily midday · 20 min', '5 connect + 3 InMail + 3 comments on target-account posts.'],
    ['Mid-funnel block', 'Daily PM · 60 min', 'Scotiabank-style motion on every $10k+ deal: EB, technical multi-thread, security docs early.'],
    ['Friday hygiene + MEDDICC', 'Weekly · 90 min', 'Day-60 disqualification sweep — every deal 60+ days with no next step (Unilever/AdRoll rule). Update MEDDICC card on every Stage 2+ deal.'],
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
  slideTitle(s, 12, 'July True Pipeline — $106.9k Open (Command Center · Jul 8)');
  const headers = ['Account', 'Opportunity', 'ERR', 'Stage', '% Close', 'Weighted'];
  const dealRows = [
    ['Shopify', 'X Ads | $2k', '$1,750', 'Committed', '90%', '$1,600', C.green, true],
    ['Upwork', 'MIC/MCP | $15k', '$11,500', 'Agreement · legal/DPA', '85%', '$9,800', C.green, true],
    ['Shronk', 'Growth | $2k', '$2,124', 'Committed', '75%', '$1,600', C.green, true],
    ['Kraken', 'MCP | $3.5k', '$3,468', 'Pending Sales Ops reopen', '70%', '$2,400', C.orange, true],
    ['HA', 'Gemini Agent | $7.5k', '$7,500', 'Qualification', '65%', '$4,900', C.blue, false],
    ['TaskRabbit', 'Gem Agent | $5k', '$5,000', 'Qualification', '56%', '$2,800', C.blue, false],
    ['RLE', 'BQ | $20k', '$20,000', 'Proof of Value · main swing', '50%', '$10,000', C.accent, true],
    ['Other (~15)', 'Growth/MCP/Claude $2-4k', '~$55,500', '—', '~23%', '~$13,000', C.muted, false],
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
  s.addTable(tbl, tableOpts(0.4, 1.4, 12.55, [2.0, 2.7, 1.4, 2.2, 1.4, 2.85], 0.42));
  s.addShape('rect', { x: 0.4, y: 5.55, w: 12.55, h: 0.55, fill: { color: '3A2520' }, line: { color: C.red, width: 0.5 } });
  s.addText([
    { text: 'Swing / overdue (re-date or disqualify · not in the $106.9k July-close): ', options: { bold: true, color: C.red } },
    { text: 'US Soccer $40k (overdue 5/28), BlueDoor $10k (overdue 6/25), IA $5k (overdue 5/28)', options: { color: C.text } }
  ], { x: 0.6, y: 5.6, w: 12.2, h: 0.45, fontFace: F.body, fontSize: 12, valign: 'middle' });
  s.addText('Command-center reconciled (Jul 8): $18.8k already committed (81% of the reduced $23.4k July quota), $46.1k weighted (197%). Coverage is not the risk — it\'s a no-anchor portfolio of $7–20k deals. MEDDICC breadth on RLE, Upwork, HA is the MIT. Kraken pending Sales Ops reopen; HubSpot|MCP moved to Aug.', { x: 0.4, y: 6.25, w: 12.6, h: 0.6, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
}

// ===== S13 FORECAST =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 13, 'July Forecast');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 8.0, h: 5.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('FORECAST (HIGH-CONFIDENCE)', { x: 0.6, y: 1.5, w: 7.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.green, bold: true, charSpacing: 2 });
  s.addText([
    { text: '■  Upwork — $9.8k (85%) ', options: { bold: true, color: C.green, fontSize: 12 } },
    { text: '· in Agreement, moving through legal/DPA — strongest close signal (Gmail thread through 6/29)\n', options: { color: C.text, fontSize: 11 } },
    { text: '■  Shopify $1.6k (90%) · Shronk $1.6k (75%) · Kraken $2.4k (70%) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '· the other committed deals · $18.8k confirmed FC total (81% of quota). Kraken pending Sales Ops reopen.', options: { color: C.muted, fontSize: 11 } }
  ], { x: 0.6, y: 1.85, w: 7.6, h: 1.1, fontFace: F.body });
  s.addText('SWING DEALS (THE MONTH IS MADE HERE)', { x: 0.6, y: 3.0, w: 7.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.orange, bold: true, charSpacing: 2 });
  s.addText([
    { text: '■  RLE ($20k, 50%) — new POV, main swing ', options: { bold: true, color: C.accent, fontSize: 12 } },
    { text: '· $10k weighted · value narrative + ROI on the POV\n', options: { color: C.text, fontSize: 11 } },
    { text: '■  HA ($7.5k, 65%) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— Gemini Agent · drive qualification → proposal\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  Shubert ($12k, 10%) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— early · EB conversation before pricing\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  WestGate ($7.5k, 10%) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— POV expansion · quantify value\n', options: { color: C.muted, fontSize: 11 } },
    { text: '■  US Soccer ($40k) — overdue 5/28, re-date or disqualify by 7/11 ', options: { bold: true, color: C.red, fontSize: 12 } },
    { text: '· not in the $106.9k July-close\n', options: { color: C.text, fontSize: 11 } },
    { text: '■  BlueDoor ($10k) ', options: { bold: true, color: C.text, fontSize: 12 } },
    { text: '— overdue 6/25 · quantify pain or disqualify by day-60', options: { color: C.muted, fontSize: 11 } }
  ], { x: 0.6, y: 3.35, w: 7.6, h: 2.4, fontFace: F.body });
  s.addShape('roundRect', { x: 0.6, y: 5.9, w: 7.6, h: 0.85, fill: { color: '3D2E0D' }, line: { color: C.orange, width: 1 }, rectRadius: 0.08 });
  s.addText([
    { text: 'The narrative: ', options: { bold: true, color: C.orange } },
    { text: 'July opens $0 closed but 197% weighted-covered ($46.1k) on the reduced $23.4k quota — $18.8k already committed (81%). Coverage is not the risk; conversion is. A no-anchor portfolio converts only if MEDDICC runs on every $10k+, and the dials must come back — or a green forecast hides a cold-outbound Q3.', options: { color: C.text } }
  ], { x: 0.75, y: 5.96, w: 7.3, h: 0.73, fontFace: F.body, fontSize: 10.5, valign: 'middle' });

  s.addShape('rect', { x: 8.6, y: 1.4, w: 4.4, h: 5.5, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addText('FORECAST / OPEN PIPELINE', { x: 8.8, y: 1.5, w: 4.0, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, bold: true, charSpacing: 2 });
  const fcRows = [
    ['July Quota (reduced)', '$23,400', true, C.orange],
    ['Closed Won (MTD)', '$0', false, C.text],
    ['Confirmed FC', '$18,800 · 81%', false, C.green],
    ['Weighted FC', '$46,100 · 197%', true, C.green],
    ['Open Pipeline', '$106,900 · 457%', false, C.text],
    ['# Open Deals', '22', false, C.text]
  ];
  let fy = 1.95;
  fcRows.forEach(r => {
    s.addText(r[0], { x: 8.8, y: fy, w: 2.5, h: 0.35, fontFace: F.body, fontSize: 11, color: C.muted });
    s.addText(r[1], { x: 11.0, y: fy, w: 1.85, h: 0.35, fontFace: F.body, fontSize: 12, color: r[3], bold: r[2], align: 'right' });
    fy += 0.5;
  });
  s.addText('Weighted pipe covers the reduced July quota 1.97× and $18.8k is already committed (81%). Coverage is healthy — the risk is the no-anchor, all-inbound mix. Every $10k+ deal still needs the MEDDICC motion. (Full $46,750 quota resumes Aug.)', { x: 8.8, y: 5.1, w: 4.0, h: 1.6, fontFace: F.body, fontSize: 10, color: C.muted, italic: true });
}

// ===== S13b MEDDICC REPLICATE =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, '13b', 'MEDDICC — Proven on Scotiabank, Missing in June');
  s.addShape('roundRect', { x: 0.4, y: 1.3, w: 12.55, h: 0.85, fill: { color: '143A2E' }, line: { color: C.green, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Scotiabank proved MEDDICC converts flagship deals. June forgot it — ', options: { bold: true, color: C.green, fontSize: 12 } },
    { text: 'every $10k+ deal died stale, unqualified, or without an EB. The July job is breadth: run the exact Scotiabank motion on every $10k+ deal, not just whatever inbound sends.', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.6, y: 1.38, w: 12.2, h: 0.7, fontFace: F.body, valign: 'middle' });
  const cols = [
    { accent: C.green, title: 'WHAT WON SCOTIABANK', body: [
      { text: 'E', options: { bold: true, color: C.green, fontSize: 11 } }, { text: 'conomic Buyer — engaged + decision path live through signature\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'Technical multi-thread', options: { bold: true, color: C.text, fontSize: 11 } }, { text: ' — engineering org, not just champion\n\n', options: { color: C.muted, fontSize: 10 } },
      { text: 'P', options: { bold: true, color: C.green, fontSize: 11 } }, { text: 'aper Process — security/DR docs early, confirmed received\n\n', options: { color: C.text, fontSize: 10 } },
      { text: 'Prep brief', options: { bold: true, color: C.text, fontSize: 11 } }, { text: ' before the technical call — objections anticipated', options: { color: C.muted, fontSize: 10 } }
    ] },
    { accent: C.red, title: 'WHAT LOST JUNE\'S $10k+ DEALS', body: [
      { text: '✗  No EB motion ran (AMN — sat live for months)\n\n', options: { color: C.red, fontSize: 10, bold: true } },
      { text: '✗  ~10–12 month stale deaths, no next step (Unilever, AdRoll)\n\n', options: { color: C.red, fontSize: 10, bold: true } },
      { text: '✗  Died unqualified in ~10 days (Celebrands)\n\n', options: { color: C.red, fontSize: 10, bold: true } },
      { text: '✗  Day-60 disqualification rule not enforced', options: { color: C.red, fontSize: 10, bold: true } }
    ] },
    { accent: C.blue, title: 'JULY OPERATING RHYTHM', body: [
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
  slideTitle(s, '13d', 'July MEDDICC Targets — Per Active Deal');
  s.addText('Each major open deal · the red-to-fix dimension · the July action. Plus two loss autopsies carried from June.', { x: 0.4, y: 1.3, w: 12.6, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted });
  const headers = ['Deal', 'ERR', 'Weakest Dimension', 'July Action', 'Due'];
  const rows = [
    ['Upwork', '$15k', 'Paper Process', 'Close legal/DPA — drive to signature', '7/15', false],
    ['RLE', '$20k', 'Metrics', 'Value narrative + ROI on the POV', '7/17', false],
    ['US Soccer', '$40k', 'Decision Process', 'Real next step or re-date / disqualify (overdue 5/28)', '7/11', false],
    ['Shubert', '$12k', 'Economic Buyer', 'EB conversation before pricing', '7/17', false],
    ['TaxAct', '$10k', 'Champion', 'Multi-thread / champion test', '7/17', false],
    ['BlueDoor', '$10k', 'Pain (urgency)', 'Quantify or disqualify by day-60', '7/18', false],
    ['Unilever (autopsy)', 'lost $40k', 'Decision Process', '1-pager: the ~10-month stale death', '7/9', true],
    ['AMN (autopsy)', 'lost $12k', 'Economic Buyer', '1-pager: the SNO deal that never got an EB', '7/9', true]
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
  s.addShape('roundRect', { x: 0.4, y: 6.0, w: 12.55, h: 0.9, fill: { color: '3A2520' }, line: { color: C.red, width: 1 }, rectRadius: 0.1 });
  s.addText([
    { text: 'Pattern: ', options: { bold: true, color: C.red } },
    { text: 'the deals that win (Scotiabank, AMBSE eventually) get multi-threaded and driven; the deals that lose (Unilever, AdRoll, AMN) sit until they die. ', options: { color: C.text } },
    { text: 'Same lesson, third month — July is about not learning it a fourth time.', options: { bold: true, color: C.orange } }
  ], { x: 0.6, y: 6.07, w: 12.2, h: 0.75, fontFace: F.body, fontSize: 12, valign: 'middle' });
}

// ===== S15 COACHING =====
{ const s = pres.addSlide({ masterName: 'DARK' });
  slideTitle(s, 15, 'Coaching Focus — July');
  s.addShape('rect', { x: 0.4, y: 1.4, w: 6.2, h: 2.6, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 0.4, y: 1.4, w: 0.08, h: 2.6, fill: { color: C.blue }, line: { width: 0 } });
  s.addText('PRIMARY — DIAL DISCIPLINE (ESCALATED)', { x: 0.6, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 12, color: C.blue, bold: true, charSpacing: 1 });
  s.addText([
    { text: '• Now a two-month failure (39 → 0) — block + logged-dial target + weekly review\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '• ≥15 dials/day, logged in HubSpot, counted same-day. Non-negotiable\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '• Walk one day\'s call log in each 1:1', options: { color: C.text, fontSize: 11 } }
  ], { x: 0.6, y: 2.0, w: 5.8, h: 1.9, fontFace: F.body });

  s.addShape('rect', { x: 6.8, y: 1.4, w: 6.2, h: 2.6, fill: { color: C.surface2 }, line: { color: C.border, width: 0.5 } });
  s.addShape('rect', { x: 6.8, y: 1.4, w: 0.08, h: 2.6, fill: { color: C.accent }, line: { width: 0 } });
  s.addText('SECONDARY — MEDDICC BREADTH', { x: 7.0, y: 1.55, w: 5.8, h: 0.3, fontFace: F.body, fontSize: 12, color: C.accent, bold: true, charSpacing: 1 });
  s.addText([
    { text: '• Full Scotiabank motion on every $10k+ late-stage deal, not just inbound\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '• 1 deal card walked through every weekly 1:1\n\n', options: { color: C.text, fontSize: 11 } },
    { text: '• Day-60 disqualification enforced (Unilever/AdRoll lesson)', options: { color: C.text, fontSize: 11 } }
  ], { x: 7.0, y: 2.0, w: 5.8, h: 1.9, fontFace: F.body });

  s.addText('THEME & CARRY-FORWARD', { x: 0.4, y: 4.2, w: 12, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted, bold: true, charSpacing: 2 });
  s.addText([
    { text: '✓  New theme — Day-60 disqualification. ', options: { color: C.green, bold: true, fontSize: 12 } },
    { text: 'Unilever ($40k) and AdRoll ($13k) are June\'s stale deaths — kill zombies before they inflate and then implode the pipe.\n\n', options: { color: C.text, fontSize: 12 } },
    { text: '•  Carried forward: ', options: { color: C.accent, bold: true, fontSize: 12 } },
    { text: 'Value Narratives on $10k+ deals · Exec Summaries on top 3 · "sustain, don\'t sprint" — two motions at once.\n\n', options: { color: C.text, fontSize: 12 } },
    { text: '•  The one-line goal: ', options: { color: C.accent, bold: true, fontSize: 12 } },
    { text: 'put an outbound floor under a green June and close big deals again — don\'t let a green number hide a cold-outbound Q3.', options: { color: C.text, fontSize: 12 } }
  ], { x: 0.4, y: 4.55, w: 12.6, h: 2.3, fontFace: F.body });
}

const OUT = 'June_2026_NAM_MBR.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Wrote ' + OUT)).catch(e => { console.error(e); process.exit(1); });
