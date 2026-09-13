#!/usr/bin/env node
// Generate Z1(a) documents and read them back out of the .docx, asserting each
// leave type lands in the row the printed form actually has for it.
//
//   node tests/z1a.js            check
//   node tests/z1a.js --update   re-record the roster-path snapshot
//
// Needs Playwright's Chromium; set CHROME to point at a specific binary.
// docx.js and JSZip are both already in the page, so the unzip happens there
// and node needs nothing beyond Playwright.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const crypto = require('crypto');

const APP = 'file://' + path.join(__dirname, '..', 'index.html');
const SNAP = path.join(__dirname, 'z1a-roster-snapshot.txt');
const UPDATE = process.argv.includes('--update');

const BASE = { firstName:'Anna', surname:'Bester', persal:'12345678',
  signatureDate:'20/12/2026', supervisorName:'Paul Xafis',
  component:'Emergency Medicine', addressDuringLeave:'12 Test Road, Wynberg' };

const docxXml = (page, extra) => page.evaluate(async ([base, extra]) => {
  const blob = await generateZ1ADocx({ ...base, ...extra });
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  return zip.file('word/document.xml').async('string');
}, [BASE, extra]);

// Every table row as an array of its cells' text. Split on the closing angle
// bracket: '<w:tc' alone also matches <w:tcPr>, <w:tcW>, <w:tcBorders> and
// <w:tcMar>, which shreds every cell into fragments.
const cellsOf = xml => xml.split('<w:tr>').slice(1).map(tr =>
  tr.split('<w:tc>').slice(1).map(tc =>
    (tc.match(/<w:t(?: [^>]*)?>[^<]*/g) || []).map(s => s.replace(/^<w:t(?: [^>]*)?>/, '')).join('')));
const rowFor = (xml, label) => cellsOf(xml).find(c => c[0] === label);
// A label can now carry one row per leave period, so the filled ones are what
// matters; blanks are the printed form's own empty lines.
const rowsFor = (xml, label) => cellsOf(xml)
  .filter(c => c[0] === label && (c[1] || c[2] || c[3]))
  .map(c => c.slice(1, 4));

let failed = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(`${ok ? 'ok  ' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`        want ${JSON.stringify(want)}\n        got  ${JSON.stringify(got)}`);
};

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(800);

  // 1. The case editedShifts cannot express at all: a period crossing New Year.
  let xml = await docxXml(page, { leaveRows: [
    { type:'Leave - Annual', startDate:'28/12/2026', endDate:'05/01/2027', count:5 } ] });
  check('annual, crossing the year', (rowFor(xml,'Annual Leave')||[]).slice(0,4),
        ['Annual Leave','28/12/2026','05/01/2027','5']);

  // 2. Study is a Special Leave, and must name itself on the Specify line.
  xml = await docxXml(page, { leaveRows: [
    { type:'Leave - Study', startDate:'02/03/2027', endDate:'06/03/2027', count:5 } ] });
  check('study fills the Special Leave row',
        (rowFor(xml,'Special Leave ((Provide supporting evidence)')||[]).slice(0,4),
        ['Special Leave ((Provide supporting evidence)','02/03/2027','06/03/2027','5']);
  check('study names itself on the Specify line',
        (rowFor(xml,'Specify Type of Special Leave')||[])[1],
        'Study (preparation and examinations)');

  // 3. Maternity is the one leave type rendered by calRow.
  xml = await docxXml(page, { leaveRows: [
    { type:'Leave - Maternity', startDate:'01/03/2027', endDate:'30/06/2027', count:4 } ] });
  const mat = rowFor(xml,'Maternity Leave (Provide supporting evidence))') || [];
  check('maternity start/end', [mat[1], mat[2]], ['01/03/2027','30/06/2027']);
  check('maternity months in the last cell', mat[mat.length-1], '4');

  // 4. Unpaid sits under the calendar-days header but is a leaveRow4.
  xml = await docxXml(page, { leaveRows: [
    { type:'Leave - Unpaid', startDate:'01/02/2027', endDate:'28/02/2027', count:28 } ] });
  check('unpaid', (rowFor(xml,'Unpaid Leave (Provide motivation)')||[]).slice(0,4),
        ['Unpaid Leave (Provide motivation)','01/02/2027','28/02/2027','28']);

  // 5. Two types share the Special row, but they are still two periods: one
  //    merged 03→12 span would claim six consecutive days that were not taken.
  //    They come out in date order regardless of the order passed in.
  xml = await docxXml(page, { leaveRows: [
    { type:'Leave - Special', startDate:'10/05/2027', endDate:'12/05/2027', count:3, specify:'Bereavement' },
    { type:'Leave - Study',   startDate:'03/05/2027', endDate:'05/05/2027', count:3 } ] });
  check('special and study each get their own row, in date order',
        rowsFor(xml,'Special Leave ((Provide supporting evidence)'),
        [['03/05/2027','05/05/2027','3'], ['10/05/2027','12/05/2027','3']]);
  check('both reasons on the Specify line',
        (rowFor(xml,'Specify Type of Special Leave')||[])[1],
        'Bereavement; Study (preparation and examinations)');

  // 6. The other six types each reach their own row.
  for (const [type, label] of [
    ['Leave - Sick','Normal Sick Leave (Provide supporting evidence when applicable)'],
    ['Leave - Family Responsibility','Family Responsibility Leave (Provide supporting evidence)'],
    ['Leave - Prenatal','Pre-natal Leave (Provide supporting evidence)'],
    ['Leave - Paternity','Paternity Leave (Provide supporting evidence)'],
  ]) {
    const x = await docxXml(page, { leaveRows: [
      { type, startDate:'01/06/2027', endDate:'03/06/2027', count:3 } ] });
    check(type, (rowFor(x, label)||[]).slice(1,4), ['01/06/2027','03/06/2027','3']);
  }

  // 7. Both Yes/No declarations are No, on both routes. This is the assertion
  //    that should have existed from the start — without it, nobody could tell
  //    from the code whether the box was ticked, and the comment in RIGHT_YN
  //    was the only evidence either way.
  for (const [route, extra] of [
    ['leave-only', { leaveRows: [{ type:'Leave - Annual', startDate:'01/02/2027', endDate:'05/02/2027', count:5 }] }],
    ['roster',     { month:6, year:2026, editedShifts: { 3:{ typeLabel:'Leave - Annual' } } }],
  ]) {
    const x = await docxXml(page, extra);
    for (const label of ['Shift Worker','Casual Employee']) {
      const row = cellsOf(x).find(c => c.includes(label)) || [];
      const i = row.indexOf(label);
      check(`${label} declares No (${route})`, row.slice(i, i+5),
            [label, 'Yes', '', 'No', '\u2713']);
    }
  }

  // 8. The roster path must be untouched by any of the above. One leave type
  //    only, so the snapshot stays meaningful when the merging loop changes.
  const rosterXml = await docxXml(page, { month: 6, year: 2026, editedShifts: {
    3: { typeLabel:'Leave - Annual' }, 4: { typeLabel:'Leave - Annual' },
    7: { typeLabel:'Leave - Annual' },
    9: { typeLabel:'WD Shift - 08H00' } } });
  const hash = crypto.createHash('sha256').update(rosterXml).digest('hex');
  if (UPDATE) {
    fs.writeFileSync(SNAP, hash + '\n');
    console.log('recorded roster-path snapshot ' + hash.slice(0,16));
  } else if (!fs.existsSync(SNAP)) {
    console.log('no roster snapshot — run with --update'); failed++;
  } else {
    check('roster path unchanged (sha256 of word/document.xml)',
          hash, fs.readFileSync(SNAP,'utf8').trim());
  }
  // Separate blocks are separate periods. The fixture is leave on the 3rd and
  // 4th and again on the 7th, with Monday the 6th worked in between — one row
  // spanning 03→07 would read as five days off instead of three.
  const annualRows = x => rowsFor(x, 'Annual Leave');
  check('a worked day splits one leave into two rows', annualRows(rosterXml),
        [['03/07/2026','04/07/2026','2'], ['07/07/2026','07/07/2026','1']]);

  // ...but a gap the doctor would not have worked anyway does not split it,
  // or every fortnight's leave would arrive as two separate applications.
  const L = { typeLabel: 'Leave - Annual' };
  const fortnight = await docxXml(page, { month: 6, year: 2026, editedShifts: {
    6:L, 7:L, 8:L, 9:L, 10:L, 13:L, 14:L, 15:L, 16:L, 17:L } });
  check('a bridged weekend keeps one row', annualRows(fortnight),
        [['06/07/2026','17/07/2026','10']]);

  // The case this was reported for.
  const farApart = await docxXml(page, { month: 6, year: 2026, editedShifts: { 3:L, 27:L } });
  check('the 3rd and the 27th are two single days, not a 25-day span',
        annualRows(farApart), [['03/07/2026','03/07/2026','1'], ['27/07/2026','27/07/2026','1']]);

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nall Z1(a) checks pass');
  process.exit(failed && !UPDATE ? 1 : 0);
})();
