#!/usr/bin/env node
// Drive the leave-only panel the way a doctor would, and read the .docx that
// comes out of it. Complements tests/z1a.js, which tests the generator directly.
//
//   node tests/leave-ui.js        (CHROME=/path/to/chrome to pick a binary)
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path'), os = require('os');

const APP = 'file://' + path.join(__dirname, '..', 'index.html');
let failed = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(`${ok ? 'ok  ' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`        want ${JSON.stringify(want)}\n        got  ${JSON.stringify(got)}`);
};

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  // Offline, so the app falls back to the built-in VHW profile and lands on step 1.
  await page.goto(APP);
  await page.waitForTimeout(1200);

  check('the button is there before anything is uploaded',
        await page.isVisible('#z1LeaveBtn'), true);
  check('and it is not gated on extraction',
        await page.isDisabled('#z1LeaveBtn'), false);

  await page.click('#z1LeaveBtn');
  await page.waitForSelector('#z1LeaveOverlay.open', { timeout: 5000 });
  check('opens with Download disabled', await page.isDisabled('#z1LeaveGenerate'), true);
  check('and says why', (await page.textContent('#z1LeaveWhy .txt')).length > 0, true);
  check('the page behind it is inert',
        await page.evaluate(() => document.querySelector('.shell').inert), true);
  check('Component is filled from the profile',
        (await page.textContent('#z1lComponent')).length > 0, true);

  await page.fill('#z1lFirstName', 'Anna');
  await page.fill('#z1lSurname', 'Bester');
  await page.fill('#z1lPersal', '12345678');
  await page.selectOption('#z1lType', 'Leave - Annual');
  await page.fill('#z1lStart', '28/12/2026');
  await page.fill('#z1lEnd', '05/01/2027');
  await page.waitForTimeout(150);

  // 28 Dec – 5 Jan: weekdays are 28,29,30,31 Dec and 1,4,5 Jan. 1 Jan is a
  // public holiday, so six working days.
  check('the count is worked out from the dates', await page.inputValue('#z1lDays'), '6');
  check('the label names the unit',
        (await page.textContent('#z1lDaysLabel')).includes('working days'), true);

  await page.fill('#z1lDays', '7');
  await page.fill('#z1lEnd', '06/01/2027');
  await page.waitForTimeout(150);
  check('an override survives a change of dates', await page.inputValue('#z1lDays'), '7');
  await page.fill('#z1lDays', '');
  await page.waitForTimeout(150);
  check('clearing it hands the field back', await page.inputValue('#z1lDays'), '7');

  await page.selectOption('#z1lType', 'Leave - Unpaid');
  await page.waitForTimeout(150);
  check('changing the type switches the unit',
        (await page.textContent('#z1lDaysLabel')).includes('calendar days'), true);
  check('and recounts in it', await page.inputValue('#z1lDays'), '10');
  await page.selectOption('#z1lType', 'Leave - Annual');
  await page.fill('#z1lEnd', '05/01/2027');
  await page.waitForTimeout(150);

  await page.fill('#z1lAddress', '12 Test Road, Wynberg');
  await page.selectOption('#z1lSupervisorSel', 'Paul Xafis');
  await page.fill('#z1lSigDate', '20/12/2026');
  await page.waitForTimeout(150);
  check('now it unlocks', await page.isDisabled('#z1LeaveGenerate'), false);
  check('and the reason is gone', await page.isHidden('#z1LeaveWhy'), true);

  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#z1LeaveGenerate')]);
  check('filename is taken from the leave period, not #monthSelect',
        dl.suggestedFilename(), 'Z1a_Leave_Anna_Bester_December_2026.docx');

  const tmp = path.join(os.tmpdir(), 'z1a-ui-' + Date.now() + '.docx');
  await dl.saveAs(tmp);
  const bytes = Array.from(fs.readFileSync(tmp));
  const xml = await page.evaluate(async b => {
    const zip = await JSZip.loadAsync(new Uint8Array(b));
    return zip.file('word/document.xml').async('string');
  }, bytes);
  fs.unlinkSync(tmp);
  const cellsOf = x => x.split('<w:tr>').slice(1).map(tr =>
    tr.split('<w:tc>').slice(1).map(tc =>
      (tc.match(/<w:t(?: [^>]*)?>[^<]*/g) || []).map(s => s.replace(/^<w:t(?: [^>]*)?>/, '')).join('')));
  const row = cellsOf(xml).find(c => c[0] === 'Annual Leave') || [];
  check('the downloaded form carries the period', row.slice(0, 4),
        ['Annual Leave', '28/12/2026', '05/01/2027', '6']);
  check('and the typed-in name', xml.includes('Bester'), true);

  await page.waitForTimeout(300);
  check('the panel closed itself', await page.isHidden('#z1LeaveOverlay'), true);
  check('the scroll lock is released',
        await page.evaluate(() => document.documentElement.style.overflow), '');
  check('the page is live again',
        await page.evaluate(() => document.querySelector('.shell').inert), false);
  // The two forms are independent. Section 03 belongs to the doctor whose
  // roster is on screen; nothing typed here may reach it, or the applicant's
  // name and PERSAL end up on a colleague's Annexure C.
  check('section 03 is untouched',
        await page.inputValue('#detailFirstName'), '');
  check('and savedDetails with it',
        await page.evaluate(() => state.savedDetails.firstName), '');
  check('but the panel remembers its own entry',
        await page.evaluate(() => state.leaveDetails.firstName), 'Anna');
  // The reported path: previewing a colleague's roster runs exactly this, and
  // it used to repaint the applicant's name over theirs.
  await page.evaluate(() => {
    state.selectedDoctor = 'Abrahams';
    document.getElementById('detailsSection').style.display = '';
    restoreDetailsToForm(false);
  });
  check('previewing a colleague shows the colleague',
        await page.inputValue('#detailSurname'), 'Abrahams');
  check('and no PERSAL of the applicant\'s',
        await page.inputValue('#detailPersal'), '');
  await page.click('#z1LeaveBtn');
  await page.waitForTimeout(300);
  check('so re-opening prefills the applicant, not the doctor',
        await page.inputValue('#z1lFirstName'), 'Anna');
  check('and the period starts blank again',
        await page.inputValue('#z1lStart'), '');
  await page.click('#z1LeaveCancel');
  await page.waitForTimeout(300);

  // Start over must not leave a name behind for the next person.
  await page.evaluate(() => { if (typeof fullReset === 'function') fullReset(); });
  await page.waitForTimeout(600);
  check('Start over empties savedDetails',
        await page.evaluate(() => state.savedDetails.firstName), '');
  check('and leaveDetails too',
        await page.evaluate(() => state.leaveDetails.firstName), '');

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nthe leave panel works end to end');
  process.exit(failed ? 1 : 0);
})();
