#!/usr/bin/env node
// A roster prefixes a first initial to tell two doctors of the same surname
// apart. That initial is a first name, and must land in the first-name box —
// without disturbing any other shape of name.
//
//   node tests/names.js
//
// Needs Playwright's Chromium; set CHROME to point at a specific binary.
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

// [roster name, first name box, surname box]
const CASES = [
  // The case this exists for.
  ['M. Willemse',      'M.',   'Willemse'],
  ['J. Willemse',      'J.',   'Willemse'],
  ['M.J. Willemse',    'M.J.', 'Willemse'],
  ['M. J. Willemse',   'M.J.', 'Willemse'],
  ['A.  Ndlovu',       'A.',   'Ndlovu'],      // two spaces
  ['  M. Willemse  ',  'M.',   'Willemse'],    // stray whitespace
  // Everything else keeps its shape. These are the ones that matter: a rule
  // that split any of them would corrupt a correct name.
  ['Willemse',         '',     'Willemse'],
  ['Van Schalkwyk',    '',     'Van Schalkwyk'],
  ['Van Zyl',          '',     'Van Zyl'],
  ['Du Toit',          '',     'Du Toit'],
  ['De Haan',          '',     'De Haan'],
  ['Le Roux',          '',     'Le Roux'],
  ['Gordon-Forbes',    '',     'Gordon-Forbes'],
  ['St. John',         '',     'St. John'],    // two letters, not an initial
  ['Jnr. Adams',       '',     'Jnr. Adams'],
  ['',                 '',     ''],
];

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(800);

  for (const [input, first, surname] of CASES) {
    const got = await page.evaluate(n => splitRosterName(n), input);
    check(`${JSON.stringify(input)} → ${JSON.stringify(first)} + ${JSON.stringify(surname)}`,
          [got.first, got.surname], [first, surname]);
  }

  // The boxes themselves, through the function that actually fills them. The
  // else-branch is the one Preview uses, and it is where a fresh doctor's
  // details are first written.
  for (const isNew of [true, false]) {
    const got = await page.evaluate(nw => {
      state.selectedDoctor = 'M. Willemse';
      state.savedDetails = { firstName:'', surname:'', persal:'', supervisor:'', sigDate:'' };
      document.getElementById('detailsSection').style.display = '';
      restoreDetailsToForm(nw);
      return [document.getElementById('detailFirstName').value,
              document.getElementById('detailSurname').value];
    }, isNew);
    check(`the boxes, restoreDetailsToForm(${isNew})`, got, ['M.', 'Willemse']);
  }

  // The roster key is untouched — it is what tells the two Willemses apart in
  // the staff list and what the schedule is looked up by.
  check('state.selectedDoctor is not rewritten',
        await page.evaluate(() => state.selectedDoctor), 'M. Willemse');

  // A name the doctor has already typed wins over the roster's guess.
  check('a typed first name is not overwritten', await page.evaluate(() => {
    state.selectedDoctor = 'M. Willemse';
    state.savedDetails = { firstName:'Michael', surname:'Willemse', persal:'', supervisor:'', sigDate:'' };
    restoreDetailsToForm(false);
    return [document.getElementById('detailFirstName').value,
            document.getElementById('detailSurname').value];
  }), ['Michael', 'Willemse']);

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nroster initials split correctly');
  process.exit(failed ? 1 : 0);
})();
