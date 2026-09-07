#!/usr/bin/env node
// Replay every fixture through the parser in index.html and compare against
// the expectations recorded beside it.
//
//   node tests/run.js            check every fixture
//   node tests/run.js --update   re-record expectations from current behaviour
//
// Needs Playwright's Chromium; set CHROME to point at a specific binary.
// The fixtures carry no PDF — see tests/README.md for why, and for what these
// tests do and do not prove.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');

const DIR = path.join(__dirname, 'fixtures');
const UPDATE = process.argv.includes('--update');
const APP = 'file://' + path.join(__dirname, '..', 'index.html');

const replay = (page, fx) => page.evaluate(async fx => {
  const real = window.pdfjsLib;
  // pdf.js exports non-writable properties, so swap the whole namespace.
  window.pdfjsLib = { __proto__: real, getDocument: () => ({ promise: Promise.resolve({
    numPages: fx.pages.length,
    getPage: n => Promise.resolve({ getTextContent: () => Promise.resolve({
      items: fx.pages[n - 1].map(([s, x, y]) => ({ str: s, transform: [0, 0, 0, 0, x, y] })) }) })
  })})};
  try {
    const r = await parseRosterPDF(new ArrayBuffer(0));
    const shifts = {};
    for (const d of r.days) for (const n of (d.allNames || [])) shifts[n] = (shifts[n] || 0) + 1;
    return { days: r.days.length, staff: [...r.doctors].sort(), shifts };
  } finally { window.pdfjsLib = real; }
}, fx);

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  // The catalogue fetch is the only network the app makes; block it so a test
  // run is offline and deterministic.
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(800);

  let failed = 0;
  for (const file of fs.readdirSync(DIR).filter(f => f.endsWith('.json')).sort()) {
    const p = path.join(DIR, file);
    const fx = JSON.parse(fs.readFileSync(p, 'utf8'));
    const got = await replay(page, fx);
    if (UPDATE) {
      fx.expect = got;
      fs.writeFileSync(p, JSON.stringify(fx, null, 0).replace(/^{/, '{\n') + '\n');
      console.log(`recorded  ${file}  ${got.days} days, ${got.staff.length} staff`);
      continue;
    }
    const want = fx.expect;
    if (!want) { console.log(`no expectations  ${file} — run with --update`); failed++; continue; }
    const diffs = [];
    if (got.days !== want.days) diffs.push(`days ${want.days} → ${got.days}`);
    if (String(got.staff) !== String(want.staff)) {
      const lost = want.staff.filter(n => !got.staff.includes(n));
      const gained = got.staff.filter(n => !want.staff.includes(n));
      diffs.push(`staff${lost.length ? ' lost ' + lost.join(',') : ''}${gained.length ? ' gained ' + gained.join(',') : ''}`);
    }
    for (const n of new Set([...Object.keys(want.shifts), ...Object.keys(got.shifts)]))
      if ((want.shifts[n] || 0) !== (got.shifts[n] || 0))
        diffs.push(`${n} ${want.shifts[n] || 0} → ${got.shifts[n] || 0}`);
    if (diffs.length) { failed++; console.log(`FAIL  ${file}\n        ${diffs.join('\n        ')}`); }
    else console.log(`ok    ${file}  ${got.days} days, ${got.staff.length} staff, ${Object.values(got.shifts).reduce((a, c) => a + c, 0)} shifts`);
  }
  if (errors.length) { console.log('\npage errors: ' + errors.join(' | ')); failed++; }
  await browser.close();
  if (!UPDATE) console.log(failed ? `\n${failed} failing` : '\nall fixtures match');
  process.exit(failed && !UPDATE ? 1 : 0);
})();
