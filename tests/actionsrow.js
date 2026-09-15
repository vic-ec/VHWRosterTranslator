#!/usr/bin/env node
// Extract data and Clear all sit under whichever upload zone holds the files:
// the department zone by default and whenever it has something, the consultant
// zone when that is the only one with a file.
//
//   node tests/actionsrow.js
//
// No real roster needed — nothing is parsed, only queued, so a tiny stub PDF
// is enough to put a file in each zone.
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

const STUB = { name: 'stub.pdf', mimeType: 'application/pdf',
               buffer: Buffer.from('%PDF-1.4\n%%EOF\n') };

// Which zone is the row sitting under? Compare left edges, which is what the
// eye actually reads — not the class, which would only restate the code.
const side = p => p.evaluate(() => {
  const row = document.querySelector('.two > .actions-row').getBoundingClientRect();
  const cols = [...document.querySelectorAll('.two > .upload-zone-col')]
    .filter(c => c.offsetParent !== null)
    .map(c => c.getBoundingClientRect());
  if (!cols.length) return 'no zones';
  const near = cols.reduce((a, b) => Math.abs(b.left - row.left) < Math.abs(a.left - row.left) ? b : a);
  return { under: cols.indexOf(near) === 0 ? 'department' : 'consultant',
           alignedLeft: Math.abs(near.left - row.left) < 2,
           sameWidth: Math.abs(near.width - row.width) < 2 };
});

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  for (const [label, vp, twoCols] of [
    ['desktop', { width: 1280, height: 900 }, true],
    ['phone',   { width: 390,  height: 844 }, false],
  ]) {
    const page = await browser.newPage({ viewport: vp });
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
    await page.goto(APP);
    await page.waitForTimeout(1200);
    // The consultant zone only exists for a consultant profile.
    await page.evaluate(() => {
      if (typeof showEcSelected === 'function') showEcSelected();
      document.getElementById('consultantZoneWrap').style.display = 'block';
      if (typeof syncActionsSide === 'function') syncActionsSide();
    });
    await page.waitForTimeout(300);

    check(`${label}: nothing uploaded — under the department zone`,
          (await side(page)).under, 'department');

    await page.setInputFiles('#consultantFile', STUB);
    await page.waitForTimeout(400);
    const withCons = await side(page);
    check(`${label}: consultant roster only — ${twoCols ? 'moves right' : 'stays put, one column'}`,
          withCons.under, twoCols ? 'consultant' : 'department');
    check(`${label}: still flush with that zone, same width`,
          [withCons.alignedLeft, withCons.sameWidth], [true, true]);

    await page.setInputFiles('#rosterFile', STUB);
    await page.waitForTimeout(400);
    check(`${label}: both uploaded — the department roster wins`,
          (await side(page)).under, 'department');

    await page.evaluate(() => { state.pendingFiles = []; renderFileList(); });
    await page.waitForTimeout(300);
    check(`${label}: department file removed — back to the consultant zone`,
          (await side(page)).under, twoCols ? 'consultant' : 'department');

    // Hiding the zone has to bring the buttons back with it.
    await page.evaluate(() => {
      document.getElementById('consultantZoneWrap').style.display = 'none';
      syncActionsSide();
    });
    await page.waitForTimeout(200);
    check(`${label}: consultant zone hidden — back to the department zone`,
          (await side(page)).under, 'department');

    if (errors.length) { failed++; console.log(`\n${label} page errors: ` + errors.join(' | ')); }
    await page.context().close();
  }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nthe actions row follows the files');
  process.exit(failed ? 1 : 0);
})();
