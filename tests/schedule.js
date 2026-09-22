#!/usr/bin/env node
// The schedule table's time boxes: what can be typed into them, what follows
// what, and what a call hands over to the morning after it.
//
//   node tests/schedule.js
//
// No file is needed. The table is driven straight from state.editedShifts,
// which is what a parse fills in, so a synthetic month exercises exactly the
// same code the real one does.
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

// March 2027: the 1st is a Monday, so the 3rd, 4th and 5th are Wednesday to
// Friday and nothing in the run is a weekend or a public holiday. Alpha is on
// call on the 3rd and on an ordinary duty day on the 4th and 5th, which is the
// shape the handover is about.
const HEADER_Y = 124;
const GRID = [
  ['Consultant duty Roster March 2027', 56, 84],
  ['Day', 80, HEADER_Y], ['Weekday', 133, HEADER_Y],
  ['1', 267, HEADER_Y], ['2', 328, HEADER_Y], ['3', 389, HEADER_Y],
  ['Meetings etc.', 433, HEADER_Y], ['Leave', 504, HEADER_Y], ['Call', 568, HEADER_Y],
  ['3', 84, 136], ['Wednesday', 117, 136], ['Alpha', 239, 136], ['Alpha', 544, 136],
  ['4', 84, 148], ['Thursday',  117, 148], ['Alpha', 239, 148], ['Bravo', 544, 148],
  ['5', 84, 160], ['Friday',    117, 160], ['Alpha', 239, 160], ['Bravo', 544, 160],
  ['6', 84, 172], ['Saturday',  117, 172], ['Bravo', 239, 172], ['Bravo', 544, 172],
];

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage({ viewport: { width: 1280, height: 1000 } });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(1200);

  // Parse the grid and draw the month the way the app does, so the table under
  // test is the real one: buildPreview fills editedShifts itself from the
  // consultant data, and seeding it beforehand would only be overwritten.
  const draw = () => page.evaluate(async items => {
    const H = 842;
    const real = window.pdfjsLib;
    window.pdfjsLib = { __proto__: real, getDocument: () => ({ promise: Promise.resolve({
      numPages: 1,
      getPage: () => Promise.resolve({
        getViewport: () => ({ height: H, width: 1190 }),
        getTextContent: () => Promise.resolve({
          items: items.map(([s, x, y]) => ({ str: s, width: 0,
                                             transform: [0,0,0,0, x, H - y] })) }),
      })
    })})};
    try {
      activeProfile = VHW_FALLBACK_PROFILE;
      state.parsedFiles = []; state.rosterData = null;
      state.consultantFiles = [new File([new Uint8Array([37])],
        'EC Consultant Roster - March 2027.pdf', { type: 'application/pdf' })];
      state.consultantFile = state.consultantFiles[0];
      // The whole upload path, not just the parser: it is what opens step 2
      // and fills the month dropdown, and the table cannot be typed into from
      // step 1 — every step but the current one carries the hidden attribute.
      await parseAndStoreConsultantRoster();
      const m = parseInt(document.getElementById('monthSelect').value);
      const y = parseInt(document.getElementById('yearInput').value);
      state.selectedDoctor = 'Alpha';
      state.previewMonth = m; state.previewYear = y;
      buildPreview('Alpha', m, y);
      wizGo(2);
    } finally { window.pdfjsLib = real; }
  }, GRID);

  await page.evaluate(() => showEcSelected());
  await draw();
  await page.waitForTimeout(600);

  const box = (d, f) => `.time-edit[data-day="${d}"][data-field="${f}"]`;
  const val = (d, f) => page.evaluate(s => {
    const i = document.querySelector(s); return i ? i.value : null; }, box(d, f));
  const row = async d => ({
    nf: await val(d,'nf'), nt: await val(d,'nt'), ot1f: await val(d,'ot1f'),
    ot1t: await val(d,'ot1t'), ot2f: await val(d,'ot2f'), ot2t: await val(d,'ot2t') });
  const enter = async (d, f, text) => {
    await page.fill(box(d, f), text);
    await page.keyboard.press('Tab');
    await page.waitForTimeout(350);
  };

  check('the month is drawn', await row(3),
        { nf:'07H30', nt:'15H30', ot1f:'15H30', ot1t:'16H30', ot2f:'16H30', ot2t:'07H30' });

  // ── the two "from" boxes are derived, not typed ─────────────────────────
  check('OT1 From and OT2 From are read-only, the other four are not',
        await page.evaluate(() => ['nf','nt','ot1f','ot1t','ot2f','ot2t'].map(f => {
          const i = document.querySelector(`.time-edit[data-day="3"][data-field="${f}"]`);
          return f + (i.readOnly ? ':ro' : ':edit');
        })),
        ['nf:edit','nt:edit','ot1f:ro','ot1t:edit','ot2f:ro','ot2t:edit']);
  check('and they are out of the tab order, so Tab goes Norm To → OT1 To',
        await page.evaluate(() => ['ot1f','ot2f'].map(f =>
          document.querySelector(`.time-edit[data-day="3"][data-field="${f}"]`).tabIndex)),
        [-1, -1]);

  // Moving the end of a band moves the start of the next one with it.
  await enter(3, 'nt', '16H00');
  check('moving Norm To moves OT1 From', (await row(3)).ot1f, '16H00');
  await enter(3, 'ot1t', '17H00');
  check('and moving OT1 To moves OT2 From', (await row(3)).ot2f, '17H00');

  // ── an hour on its own is on the hour ───────────────────────────────────
  await enter(5, 'nf', '7');
  check('typing 7 is 07H00', (await row(5)).nf, '07H00');
  check('and the eight-hour day follows it', (await row(5)).nt, '15H00');
  await enter(5, 'nf', '08H');
  check('so is 08H', (await row(5)).nf, '08H00');
  await enter(5, 'nf', '0930');
  check('while four digits still mean what they did', (await row(5)).nf, '09H30');

  // ── the clear button ────────────────────────────────────────────────────
  check('every editable box carries one, and no read-only box does',
        await page.evaluate(() =>
          [...document.querySelectorAll('tr[data-day="3"] .time-clear')].map(b => b.dataset.field)),
        ['nf','nt','ot1t','ot2t']);
  // It is drawn only on hover or focus: six columns over thirty-one rows is
  // 186 crosses if they are all there at rest.
  check('and it stays out of sight until the box is pointed at',
        await page.isVisible('tr[data-day="3"] .time-clear[data-field="ot2t"]'), false);
  // Focus counts as pointing at it, which is what makes the button reachable
  // on a touch screen — and is the half of the rule a headless run can drive,
  // since :hover cannot be dispatched.
  await page.focus(box(3, 'ot2t'));
  await page.waitForTimeout(200);
  check('then it is there', await page.isVisible('tr[data-day="3"] .time-clear[data-field="ot2t"]'), true);
  await page.evaluate(() =>
    document.querySelector('tr[data-day="3"] .time-clear[data-field="ot2t"]').click());
  await page.waitForTimeout(300);
  check('clearing the end of a band empties it and gives its start back',
        [(await row(3)).ot2t, (await row(3)).ot2f], ['', '']);

  // ── the overnight handover ──────────────────────────────────────────────
  // Put day 3 back on call and take the two directions in turn.
  await draw();
  await page.waitForTimeout(500);

  await enter(3, 'ot2t', '08H00');
  check('a call ending later starts the next morning later',
        [(await row(3)).ot2t, (await row(4)).nf], ['08H00', '08H00']);
  check('and that morning is still eight hours long', (await row(4)).nt, '16H00');

  await enter(4, 'nf', '07H00');
  check('coming in early pulls the call it followed back',
        [(await row(4)).nf, (await row(3)).ot2t], ['07H00', '07H00']);
  check('so the two can never overlap', await page.evaluate(() => {
    const es = state.editedShifts;
    return es[3].ot2t === es[4].nf;
  }), true);

  // Day 5 follows an ordinary duty day, not a call, so there is nothing to
  // hand over and nothing to write back.
  await enter(5, 'nf', '06H00');
  check('a day after an ordinary duty day hands over nothing',
        [(await row(4)).ot2t || '', (await row(4)).ot2f || ''], ['', '']);

  // ── the row with an undo button is the same height as the rest ──────────
  // An inline-flex box takes its baseline from its first flex item, and the
  // undo button's only item is a masked ::before with no text in it — so it
  // had no baseline of its own, was set on its bottom margin edge, and added
  // its descender to the row. Days 3 and 5 are both ordinary shift rows; only
  // the first has been edited.
  await draw();
  await page.waitForTimeout(500);
  await enter(3, 'nt', '16H00');
  check('every row stands the same height, undo button or not',
        await page.evaluate(() => {
          const h = d => +document.querySelector(`tr[data-day="${d}"]`).getBoundingClientRect().height.toFixed(1);
          const u = d => !!document.querySelector(`tr[data-day="${d}"] .row-undo`);
          return { edited: u(3), untouched: u(5), same: h(3) === h(5) };
        }),
        { edited: true, untouched: false, same: true });

  // ── the action column's head is the pencil ──────────────────────────────
  check('the action column is headed by an icon, not the word',
        await page.evaluate(() => {
          const h = document.querySelector('.preview-table thead .act-head');
          if (!h) return 'no icon';
          return [h.closest('th').textContent.trim(), h.getAttribute('aria-label'),
                  getComputedStyle(h).maskImage.includes('svg')];
        }),
        ['', 'Actions', true]);

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nthe schedule’s time boxes behave');
  process.exit(failed ? 1 : 0);
})();
