#!/usr/bin/env node
// A consultant on-call roster is a roster on its own: the staff list, the
// month and step 2 all come from it, so step 1 must count as complete.
//
//   node tests/consultant.js
//   CONSULTANT_ROSTER="/path/to/Consultant Duty Roster May 2026.pdf" node tests/consultant.js
//
// The first half needs no file — it drives the gate directly. The second half
// runs the real upload when CONSULTANT_ROSTER points at one; no consultant
// roster is committed here, for the same reason no real roster is (see
// tests/README.md). The month and year come off the sheet's own title line,
// with the file name only as a fallback — so a roster renamed without its
// month still lands in the right month.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');

const APP = 'file://' + path.join(__dirname, '..', 'index.html');
const ROSTER = process.env.CONSULTANT_ROSTER;

let failed = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(`${ok ? 'ok  ' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`        want ${JSON.stringify(want)}\n        got  ${JSON.stringify(got)}`);
};

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(1200);

  // ── the gate itself, no file needed ──────────────────────────────────────
  check('nothing uploaded: step 1 is blocked',
        await page.evaluate(() => [wizExtracted(), wizBlockedReason(1) !== null]), [false, true]);

  check('consultant data alone completes step 1', await page.evaluate(() => {
    state.consultantData = { days: [{ date: 1, month: 4 }], doctors: new Set(['Cloete']) };
    return [wizExtracted(), wizBlockedReason(1)];
  }), [true, null]);

  check('and clearing it blocks step 1 again', await page.evaluate(() => {
    state.consultantData = null;
    return [wizExtracted(), wizBlockedReason(1) !== null];
  }), [false, true]);

  // ── the month, on a file whose name does not carry one ───────────────────
  // The reported bug: "Consultant Duty Roster 2026 Shared.pdf" is an April
  // roster, and with only the file name to go on it was filed under the month
  // it happened to be opened in. getConsultantShifts drops every day whose
  // month is not the selected one, so the preview then showed that roster's
  // duties against a different month's dates. No PDF is needed to prove it:
  // the parser reads nothing but numPages and a list of {str, x, y}.
  const TITLED = [
    ['Consultant duty Roster April 2026', 56, 84],
    ['Day', 80, 124], ['Weekday', 133, 124],
    ['1', 267, 124], ['2', 328, 124], ['3', 389, 124],
    ['Meetings etc.', 433, 124], ['Leave', 504, 124], ['Call', 568, 124],
    ['1', 84, 136], ['Wednesday', 117, 136], ['Alpha', 239, 136], ['Bravo', 544, 136],
    ['2', 84, 148], ['Thursday', 117, 148], ['Bravo', 239, 148],
  ];
  // Three files, three months, named so that nothing about the name says which.
  const MONTH_GRIDS = ['April', 'May', 'June'].map(m => [
    [`Consultant duty Roster ${m} 2026`, 56, 84],
    ['Day', 80, 124], ['Weekday', 133, 124],
    ['1', 267, 124], ['2', 328, 124], ['3', 389, 124],
    ['Meetings etc.', 433, 124], ['Leave', 504, 124], ['Call', 568, 124],
    ['1', 84, 136], ['Wednesday', 117, 136], ['Alpha', 239, 136], ['Bravo', 544, 136],
  ]);

  check('a roster with no month in its file name still lands in its own month',
        await page.evaluate(async items => {
    const H = 842;
    const real = window.pdfjsLib;
    window.pdfjsLib = { __proto__: real, getDocument: () => ({ promise: Promise.resolve({
      numPages: 1,
      getPage: () => Promise.resolve({
        getViewport: () => ({ height: H, width: 1190 }),
        getTextContent: () => Promise.resolve({
          items: items.map(([t, x, y]) => ({ str: t, transform: [0,0,0,0, x, H - y] })) }),
      })
    })})};
    try {
      activeProfile = VHW_FALLBACK_PROFILE;
      state.parsedFiles = []; state.rosterData = null;
      state.consultantFiles = [new File([new Uint8Array([37])],
        'Consultant Duty Roster 2026 Shared.pdf', { type: 'application/pdf' })];
      await parseAndStoreConsultantRoster();
      return [document.getElementById('monthSelect').value,
              document.getElementById('yearInput').value,
              [...new Set(state.rosterData.days.map(d => d.month))]];
    } finally { window.pdfjsLib = real; }
  }, TITLED), ['3', '2026', [3]]);

  // ── the file count in the status line ────────────────────────────────────
  // It said "across 1 file(s)" however many were queued: it counted
  // state.consultantFile, the single-file field, and the consultant-only
  // branch had the 1 written into the string. Nine months of roster uploaded,
  // one file reported.
  check('the status line counts every consultant file that parsed',
        await page.evaluate(async items => {
    const H = 842;
    const real = window.pdfjsLib;
    window.pdfjsLib = { __proto__: real, getDocument: () => ({ promise: Promise.resolve({
      numPages: 1,
      getPage: () => Promise.resolve({
        getViewport: () => ({ height: H, width: 1190 }),
        getTextContent: () => Promise.resolve({
          items: items.map(([t, x, y]) => ({ str: t, transform: [0,0,0,0, x, H - y] })) }),
      })
    })})};
    try {
      activeProfile = VHW_FALLBACK_PROFILE;
      state.parsedFiles = []; state.pendingFiles = []; state.rosterData = null;
      state.consultantFiles = ['April','May','June'].map(m =>
        new File([new Uint8Array([37])], `EC Consultant Roster - ${m} 2026.pdf`,
                 { type: 'application/pdf' }));
      state.consultantFile = state.consultantFiles[0];
      const btn = document.getElementById('parseBtn');
      btn.disabled = false;
      btn.click();
      await new Promise(r => setTimeout(r, 1500));
      return [state.consultantFileCount,
              /across 3 files\b/.test(document.getElementById('parseStatus').textContent)];
    } finally { window.pdfjsLib = real; }
  }, TITLED), [3, true]);

  // ── the viewer opens on the month that is on screen ──────────────────────
  // Rebuilding the picker's options resets the selection to the first file, so
  // nine months of consultant roster always opened on whichever sorted first,
  // however long you had been reading another one — and a panel there to be
  // checked against the schedule showing the wrong month is worse than none.
  check('the viewer opens on the previewed month, not the first file',
        await page.evaluate(async grid => {
    const H = 842;
    const real = window.pdfjsLib;
    // One grid per file, in the order they are parsed, so the three files
    // genuinely hold three different months.
    let call = 0;
    window.pdfjsLib = { __proto__: real, getDocument: () => {
      const items = grid[Math.min(call++, grid.length - 1)];
      return { promise: Promise.resolve({
        numPages: 1,
        getPage: () => Promise.resolve({
          getViewport: () => ({ height: H, width: 1190 }),
          getTextContent: () => Promise.resolve({
            items: items.map(([t, x, y]) => ({ str: t, width: 0,
                                               transform: [0,0,0,0, x, H - y] })) }),
        })
      })};
    }};
    try {
      activeProfile = VHW_FALLBACK_PROFILE;
      state.parsedFiles = []; state.rosterData = null;
      // Names deliberately in a different order from the months inside them,
      // so matching on the name alone could not pass this.
      state.consultantFiles = ['one','two','three'].map(n =>
        new File([new Uint8Array([37])], `roster-${n}.pdf`, { type: 'application/pdf' }));
      state.consultantFile = state.consultantFiles[0];
      await parseAndStoreConsultantRoster();
      // Read the middle one: May, the second file.
      state.previewMonth = 4; state.previewYear = 2026;
      state.editedShifts = { 3: { nf: '07H30' } };
      rosterViewOpen();
      const pick = document.getElementById('rosterViewPick');
      return [pick.selectedIndex, pick.value,
              (state.consultantFileMonths || {})['roster-two.pdf']];
    } finally { window.pdfjsLib = real; }
  }, MONTH_GRIDS), [1, 'roster-two.pdf', 4]);

  // ── the real thing ───────────────────────────────────────────────────────
  if (!ROSTER || !fs.existsSync(ROSTER)) {
    console.log('skip  set CONSULTANT_ROSTER=/path/to/a/consultant/roster.pdf for the upload checks');
  } else {
    await page.evaluate(() => { if (typeof showEcSelected === 'function') showEcSelected(); });
    await page.setInputFiles('#consultantFile', ROSTER);
    await page.waitForTimeout(600);
    await page.click('#parseBtn');
    await page.waitForTimeout(8000);

    check('the consultant roster parsed',
          await page.evaluate(() => (state.consultantData?.doctors?.size || 0) > 0), true);
    check('a staff list was built from it',
          await page.evaluate(() => document.querySelectorAll('.doctor-chip').length > 0), true);
    check('the month and year were detected',
          await page.evaluate(() => [document.getElementById('monthSelect').value !== '',
                                     /^\d{4}$/.test(document.getElementById('yearInput').value)]),
          [true, true]);
    check('step 2 is open', await page.evaluate(() =>
          getComputedStyle(document.getElementById('step2')).display !== 'none'), true);
    // The reported symptom: everything above was already true and this was not.
    check('Continue to review is enabled', await page.evaluate(() => {
      const b = document.querySelector('#sec-1 [data-wiz="next"]');
      return b ? b.disabled : 'no button';
    }), false);

    // The file the schedule came from must be viewable. On a consultant-only
    // upload it is the only file there is, and View roster file used to sit
    // disabled over a roster the app had just drawn a month from.
    check('the uploaded roster can be viewed back',
          await page.evaluate(() => rosterViewFiles().length > 0), true);
    check('and the View roster file button is enabled',
          await page.evaluate(() => {
            const b = document.getElementById('viewRosterBtn');
            return b ? b.disabled : 'missing';
          }), false);

    // And the schedule it leads to is real.
    const first = await page.evaluate(() => {
      const c = [...document.querySelectorAll('.doctor-chip')]
        .map(e => e.dataset.name).filter(n => !n.includes('&'));
      return c[0] || null;
    });
    if (first) {
      // Click the chip rather than setting state: the chip's handler is what
      // runs checkReady, and Preview stays disabled without it.
      await page.evaluate(n => document.querySelector(`.doctor-chip[data-name="${n}"]`).click(), first);
      await page.waitForTimeout(400);
      check('Preview is enabled once a consultant is picked',
            await page.evaluate(() => document.getElementById('previewBtn').disabled), false);
      await page.evaluate(() => document.getElementById('previewBtn').click());
      await page.waitForTimeout(1500);
      check(`previewing ${JSON.stringify(first)} gives a month of rows`,
            await page.evaluate(() => document.querySelectorAll('#previewArea tr[data-day]').length >= 28), true);
      check('with duty days filled in from the consultant roster',
            await page.evaluate(() => Object.keys(state.editedShifts || {}).length > 0), true);
    }
  }

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\na consultant roster stands on its own');
  process.exit(failed ? 1 : 0);
})();
