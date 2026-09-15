#!/usr/bin/env node
// The roster viewer's find row must stay on screen while the pages scroll
// under it — stepping through matches is exactly when the next-match button
// is wanted.
//
//   ROSTER=/path/to/a/roster.pdf node tests/rosterview.js
//
// Needs a real PDF: the panel renders pages, and there is nothing to scroll
// without them. Skips with a note when ROSTER is unset.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');

const APP = 'file://' + path.join(__dirname, '..', 'index.html');
const ROSTER = process.env.ROSTER;

let failed = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(`${ok ? 'ok  ' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`        want ${JSON.stringify(want)}\n        got  ${JSON.stringify(got)}`);
};

(async () => {
  if (!ROSTER || !fs.existsSync(ROSTER)) {
    console.log('skip  set ROSTER=/path/to/a/roster.pdf to run the viewer checks');
    process.exit(0);
  }
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(1200);
  await page.evaluate(() => { if (typeof showEcSelected === 'function') showEcSelected(); });
  await page.setInputFiles('#rosterFile', ROSTER);
  await page.waitForTimeout(400);
  await page.click('#parseBtn');
  await page.waitForTimeout(7000);

  await page.click('#hdrViewBtn');
  await page.waitForTimeout(3000);
  check('the viewer opened', await page.isVisible('#rosterViewOverlay'), true);

  const body = '#rosterViewOverlay .modal-body';
  const scrollHeight = await page.evaluate(s => document.querySelector(s).scrollHeight, body);
  const clientHeight = await page.evaluate(s => document.querySelector(s).clientHeight, body);
  check('there is more than a panel-full of pages to scroll',
        scrollHeight > clientHeight + 200, true);

  // Where the find row sits relative to the scrolling box, before and after.
  const offset = () => page.evaluate(s => {
    const b = document.querySelector(s);
    const t = document.querySelector('#rosterViewOverlay .rv-tools');
    if (!t) return 'no find row';   // report it, rather than throwing on null
    return Math.round(t.getBoundingClientRect().top - b.getBoundingClientRect().top);
  }, body);

  const before = await offset();
  await page.evaluate(s => { document.querySelector(s).scrollTop = 1200; }, body);
  await page.waitForTimeout(400);
  const after = await offset();
  check('the find row does not move when the pages scroll', [before, after], [0, 0]);

  check('the search box is still on screen', await page.isVisible('#rosterViewFind'), true);
  check('and so are both step buttons',
        [await page.isVisible('#rosterViewPrev'), await page.isVisible('#rosterViewNext')],
        [true, true]);

  // It has to be opaque, or the pages show through as they pass under it.
  check('the bar is opaque', await page.evaluate(() => {
    const bg = getComputedStyle(document.querySelector('#rosterViewOverlay .rv-tools')).backgroundColor;
    return /^rgba\(/.test(bg) ? bg : 'opaque';
  }), 'opaque');

  // Stepping to a match must not leave it hidden behind the bar.
  await page.fill('#rosterViewFind', 'a');
  await page.waitForTimeout(600);
  const n = await page.evaluate(() => document.querySelectorAll('#rosterViewBody .rv-marks i').length);
  if (n > 0) {
    await page.click('#rosterViewNext');
    await page.waitForTimeout(600);
    await page.click('#rosterViewNext');
    await page.waitForTimeout(600);
    check('the current match is below the bar, not behind it', await page.evaluate(() => {
      const cur = document.querySelector('#rosterViewBody .rv-marks i.is-current');
      if (!cur) return 'no current match';
      const t = document.querySelector('#rosterViewOverlay .rv-tools').getBoundingClientRect();
      return cur.getBoundingClientRect().top >= t.bottom - 1;
    }), true);
  } else {
    console.log('note  no matches for the probe query; skipped the step check');
  }

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nthe find row stays put while the pages scroll');
  process.exit(failed ? 1 : 0);
})();
