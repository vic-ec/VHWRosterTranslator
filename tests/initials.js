#!/usr/bin/env node
// Two doctors who share a surname are told apart on the roster by a first
// initial. A PDF may set that initial in the same text item as the surname, or
// as an item of its own — both must reach the parser as two distinct people.
//
//   node tests/initials.js
//
// Built by mutating a real fixture rather than by adding a synthetic one, so
// the layout, coordinates and every other name stay exactly as a real roster
// has them; only one surname is duplicated under two initials.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');

const APP = 'file://' + path.join(__dirname, '..', 'index.html');
const FX = path.join(__dirname, 'fixtures', 'ec-roster-2024.json');
// The most-rostered name in that fixture, so the split shows up on many rows.
const SHARED = 'Netew';

let failed = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(`${ok ? 'ok  ' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`        want ${JSON.stringify(want)}\n        got  ${JSON.stringify(got)}`);
};

const replay = (page, fx) => page.evaluate(async fx => {
  const real = window.pdfjsLib;
  window.pdfjsLib = { __proto__: real, getDocument: () => ({ promise: Promise.resolve({
    numPages: fx.pages.length,
    getPage: n => Promise.resolve({ getTextContent: () => Promise.resolve({
      items: fx.pages[n-1].map(([s,x,y]) => ({ str:s, transform:[0,0,0,0,x,y] })) }) })
  })})};
  try {
    const r = await parseRosterPDF(new ArrayBuffer(0));
    const shifts = {};
    for (const d of r.days) for (const n of (d.allNames || [])) shifts[n] = (shifts[n] || 0) + 1;
    return { staff: [...r.doctors].sort(), shifts };
  } finally { window.pdfjsLib = real; }
}, fx);

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(800);

  const base = JSON.parse(fs.readFileSync(FX, 'utf8'));
  const before = await replay(page, base);
  const shared = before.shifts[SHARED] || 0;
  check(`the fixture has a ${SHARED} to split`, shared > 0, true);

  // Alternate the initial so the shifts divide between the two doctors.
  const both = [`M. ${SHARED}`, `J. ${SHARED}`];
  for (const [label, mutate] of [
    ['one item, "M. Name"', (pg, n) => {
      for (const it of pg) if (it[0] === SHARED) it[0] = n() ;
    }],
    ['two items, "M." then "Name"', (pg, n) => {
      const out = [];
      for (const it of pg) {
        if (it[0] === SHARED) out.push([n().split(' ')[0], it[1] - 14, it[2]]);
        out.push(it);
      }
      pg.length = 0; pg.push(...out);
    }],
  ]) {
    const fx = JSON.parse(JSON.stringify(base));
    let i = 0;
    const next = () => both[i++ % 2];
    for (const pg of fx.pages) mutate(pg, next);
    const after = await replay(page, fx);

    check(`${label}: both doctors appear`,
          after.staff.filter(s => s.includes(SHARED)).sort(), [...both].sort());
    check(`${label}: one more name than before`,
          after.staff.length, before.staff.length + 1);
    check(`${label}: their shifts add up to the original`,
          (after.shifts[both[0]] || 0) + (after.shifts[both[1]] || 0), shared);
    check(`${label}: the bare surname is gone`,
          after.shifts[SHARED] || 0, 0);
    // Nobody else moved.
    const others = o => Object.fromEntries(Object.entries(o).filter(([k]) => !k.includes(SHARED)));
    check(`${label}: every other doctor is untouched`,
          others(after.shifts), others(before.shifts));
  }

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nboth initial layouts keep the two doctors apart');
  process.exit(failed ? 1 : 0);
})();
