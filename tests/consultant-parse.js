#!/usr/bin/env node
// The consultant parser, on a synthetic grid. Nothing real is needed: the
// parser reads only numPages and a list of {str, x, y}, so the fixture is that
// list — see tests/README.md for why no real roster is committed.
//
//   node tests/consultant-parse.js
//
// The column x's below deliberately do NOT match the profile's pdf_columns.
// They are modelled on a real export whose grid sits ~35px to the left, where
// the fixed coordinates read the Meetings column as duty slot 3, the Leave
// column as Meetings and the Call column as Leave — so every on-call day was
// recorded as annual leave. The header row is the anchor that fixes it.
const { chromium } = require('playwright');
const path = require('path');
const APP = 'file://' + path.join(__dirname, '..', 'index.html');

let failed = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(`${ok ? 'ok  ' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`        want ${JSON.stringify(want)}\n        got  ${JSON.stringify(got)}`);
};

// March 2027: the 1st is a Monday, so 2–5 are Tue–Fri and none is a holiday.
const HEADER_Y = 124;
const ITEMS = [
  ['Consultant duty Roster March 2027', 56,  84],
  // Header — the anchor. Content in each column sits a little to its left.
  ['Day', 80, HEADER_Y], ['Weekday', 133, HEADER_Y],
  ['1', 267, HEADER_Y], ['2', 328, HEADER_Y], ['3', 389, HEADER_Y],
  ['Meetings etc.', 433, HEADER_Y], ['Leave', 504, HEADER_Y], ['Call', 568, HEADER_Y],
  ['2nd', 629, HEADER_Y],
  // A row immediately under the header: it must not be skipped.
  ['2', 84, 136], ['Tuesday', 117, 136],
  ['Alpha', 239, 136], ['Bravo', 300, 136],
  ['Clin Gov Alpha & Bravo', 422, 136],          // a meeting, not a duty
  ['Alpha', 544, 136],                            // on call
  // Two consultants sharing one duty cell.
  ['3', 84, 148], ['Wednesday', 117, 148],
  ['Alpha & Bravo', 361, 148], ['Charlie', 544, 148],
  // Leave, which the fixed coordinates used to read as Meetings.
  ['4', 84, 160], ['Thursday', 117, 160],
  ['Charlie', 239, 160], ['Delta Leave', 483, 160],
  // The other spelling of a shared cell.
  ['5', 84, 172], ['Friday', 117, 172], ['Alpha / Bravo', 300, 172],
];

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(800);

  const parsed = await page.evaluate(async items => {
    const H = 842;
    const real = window.pdfjsLib;
    window.pdfjsLib = { __proto__: real, getDocument: () => ({ promise: Promise.resolve({
      numPages: 1,
      getPage: () => Promise.resolve({
        getViewport: () => ({ height: H, width: 1190 }),
        getTextContent: () => Promise.resolve({
          items: items.map(([s, x, y]) => ({ str: s, transform: [0,0,0,0, x, H - y] })) }),
      })
    })})};
    try {
      const d = await parseConsultantRosterPDF(new ArrayBuffer(0), VHW_FALLBACK_PROFILE);
      for (const day of d.days) day.month = 2;           // March
      const per = {};
      for (const n of ['Alpha','Bravo','Charlie','Delta'])
        per[n] = getConsultantShifts(d, n, 2, VHW_FALLBACK_PROFILE, 2027);
      return { doctors: [...d.doctors].sort(), dates: d.days.map(x => x.date), per };
    } finally { window.pdfjsLib = real; }
  }, ITEMS);

  // The row directly under the header is data, not chrome.
  check('every row is read, starting with the one under the header',
        parsed.dates, [2, 3, 4, 5]);

  // A meeting is not a person, and a shared cell is two people. Delta is
  // absent on purpose: the staff list is built from the duty and call columns
  // only, so a consultant who appears nowhere but the Leave column all month
  // cannot be picked. Their leave is still read correctly once they are
  // selected — see the Delta check further down — so this is a gap in the
  // list, not in the parse.
  check('the staff list is who was on duty, and nothing else',
        parsed.doctors, ['Alpha', 'Bravo', 'Charlie']);

  const type = (who, day) => (parsed.per[who][day] || {}).typeLabel || null;
  const band = (who, day) => {
    const s = parsed.per[who][day] || {};
    return [s.nf, s.nt, s.ot1f, s.ot1t, s.ot2f, s.ot2t];
  };

  // Day 2: Alpha is in slot 1 and on call; Bravo is in slot 2; the Clin Gov
  // meeting names both of them and must count for neither.
  check('slot 1 plus call is an on-call day', type('Alpha', 2), 'On Call - Weekday');
  check('with normal hours, OT on-site, and OT off-site overnight',
        band('Alpha', 2), ['07H30','15H30','15H30','16H30','16H30','07H30']);
  check('a duty slot without call is a normal day', type('Bravo', 2), 'Consultant Day - 07H30');
  check('and it carries no overnight OT', band('Bravo', 2).slice(4), ['', '']);
  check('the meeting gives Charlie nothing', type('Charlie', 2), null);

  // Day 3: one cell, two consultants — both are on duty.
  check('both names in one duty cell are credited (&)',
        [type('Alpha', 3), type('Bravo', 3)],
        ['Consultant Day - 07H30', 'Consultant Day - 07H30']);
  check('and the call column that day is Charlie', type('Charlie', 3), 'On Call - Weekday');

  // Day 4: the Leave column, which the fixed coordinates read as Meetings.
  check('leave is read as leave', type('Delta', 4), 'Leave - Annual');
  check('and Charlie still has his duty day', type('Charlie', 4), 'Consultant Day - 07H30');

  // Day 5: the slash spelling of a shared cell.
  check('both names in one duty cell are credited (/)',
        [type('Alpha', 5), type('Bravo', 5)],
        ['Consultant Day - 07H30', 'Consultant Day - 07H30']);

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nthe consultant grid is read by its own header');
  process.exit(failed ? 1 : 0);
})();
