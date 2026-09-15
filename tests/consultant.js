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
// tests/README.md). Note the month and year are read from the FILE NAME, so a
// renamed file falls back to today's month — keep the month name in it.
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
    check('the month and year came off the filename',
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
