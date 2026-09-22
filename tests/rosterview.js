#!/usr/bin/env node
// The roster viewer is a modeless, movable panel whose find row and scrollbars
// stay at the panel's own edges however far the file is zoomed.
//
//   node tests/rosterview.js
//   ROSTER=/path/to/a/roster.pdf node tests/rosterview.js
//
// The first half needs no file: the panel is opened over stand-in pages, which
// is enough to drive the layout, the drag and the dismissal rules. The rest
// renders a real PDF and runs the find and zoom checks, when ROSTER points at
// one.
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
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**://**', r => r.request().url().startsWith('file:') ? r.continue() : r.abort());
  await page.goto(APP);
  await page.waitForTimeout(1200);
  await page.evaluate(() => { if (typeof showEcSelected === 'function') showEcSelected(); });

  // ── the panel, over stand-in pages ───────────────────────────────────────
  // The viewer's own button is hidden until a file is retained; the panel is
  // the same panel either way, so it is opened through the real handler with
  // the button simply un-hidden, then filled with pages of the right shape.
  const stand = () => page.evaluate(() => {
    document.getElementById('rosterViewBody').innerHTML =
      Array.from({length: 3}, () =>
        '<div class="rv-page" style="height:900px;background:#ddd"></div>').join('');
  });
  await page.evaluate(() => document.getElementById('hdrViewBtn').removeAttribute('hidden'));
  await page.click('#hdrViewBtn');
  await page.waitForTimeout(400);
  await stand();

  check('the viewer opened', await page.isVisible('#rosterViewOverlay'), true);

  // Modeless: the whole point is to read the file against the schedule, which
  // means editing a cell with the panel still up.
  check('the page behind stays live', await page.evaluate(() =>
        [document.querySelector('.shell').inert,
         document.documentElement.style.overflow]), [false, '']);
  check('and the backdrop lets a click through', await page.evaluate(() =>
        getComputedStyle(document.getElementById('rosterViewOverlay')).pointerEvents),
        'none');

  // Clicking away used to dismiss it, which made editing behind it impossible.
  await page.mouse.click(20, 500);
  await page.waitForTimeout(200);
  check('clicking outside does not close it',
        await page.evaluate(() =>
          document.getElementById('rosterViewOverlay').classList.contains('open')), true);

  // ── one scroller, at the panel's edges ───────────────────────────────────
  // The find row used to be sticky inside a .modal-body that scrolled both
  // ways, so the horizontal scrollbar belonged to a box as tall as every page
  // and sat at the bottom of the content — off screen at any zoom.
  const geom = () => page.evaluate(() => {
    const body = document.querySelector('#rosterViewOverlay .modal-body');
    const rv   = document.getElementById('rosterViewBody');
    const tools= document.querySelector('#rosterViewOverlay .rv-tools');
    const panel= document.querySelector('#rosterViewOverlay .modal-panel');
    return {
      bodyScrolls: body.scrollHeight > body.clientHeight + 1,
      rvScrolls:   rv.scrollHeight   > rv.clientHeight + 1,
      // The scroller's bottom edge is the panel's, give or take its padding.
      rvBottomGap: Math.round(panel.getBoundingClientRect().bottom
                            - rv.getBoundingClientRect().bottom),
      toolsTop:    Math.round(tools.getBoundingClientRect().top
                            - panel.getBoundingClientRect().top),
      rvTop:       Math.round(rv.getBoundingClientRect().top),
    };
  });
  const g0 = await geom();
  check('the page area is the scroller, not the panel body',
        [g0.bodyScrolls, g0.rvScrolls], [false, true]);
  check('and its bottom edge is the panel\'s', g0.rvBottomGap <= 26, true);

  // Zoom widens the pages past the panel; the sideways scrollbar has to be on
  // the same box, so it lands at the panel's bottom edge too.
  for (let i = 0; i < 3; i++)
    await page.evaluate(() => document.getElementById('rosterViewIn').click());
  await page.waitForTimeout(250);
  check('zoomed, the sideways scroll is on that same box', await page.evaluate(() => {
    const rv = document.getElementById('rosterViewBody');
    return [rv.scrollWidth > rv.clientWidth + 1,
            document.querySelector('#rosterViewOverlay .modal-body').scrollWidth
              > document.querySelector('#rosterViewOverlay .modal-body').clientWidth + 1];
  }), [true, false]);

  // The find row is out of the scroller entirely, so it cannot scroll away.
  const g1 = await geom();
  await page.evaluate(() => { document.getElementById('rosterViewBody').scrollTop = 1500; });
  await page.waitForTimeout(200);
  const g2 = await geom();
  check('the find row does not move when the pages scroll',
        [g1.toolsTop, g2.toolsTop], [g1.toolsTop, g1.toolsTop]);
  check('the search box and both step buttons are still on screen',
        [await page.isVisible('#rosterViewFind'), await page.isVisible('#rosterViewPrev'),
         await page.isVisible('#rosterViewNext')], [true, true, true]);
  await page.evaluate(() => document.getElementById('rosterViewFit').click());

  // ── dragging it out of the way ───────────────────────────────────────────
  const panelBox = () => page.evaluate(() => {
    const r = document.querySelector('#rosterViewOverlay .modal-panel').getBoundingClientRect();
    return { left: Math.round(r.left), top: Math.round(r.top) };
  });
  const headMid = await page.evaluate(() => {
    const r = document.querySelector('#rosterViewOverlay .modal-head').getBoundingClientRect();
    return { x: Math.round(r.left + 60), y: Math.round(r.top + r.height / 2) };
  });
  const p0 = await panelBox();
  await page.mouse.move(headMid.x, headMid.y);
  await page.mouse.down();
  await page.mouse.move(headMid.x + 220, headMid.y + 160, { steps: 8 });
  await page.mouse.up();
  await page.waitForTimeout(200);
  const p1 = await panelBox();
  check('the head drags the panel', [p1.left - p0.left, p1.top - p0.top], [220, 160]);

  // Dragged down, the panel used to run off the bottom of the window and take
  // the sideways scrollbar with it. It gives up height instead.
  check('and its foot stays inside the window', await page.evaluate(() => {
    const r = document.querySelector('#rosterViewOverlay .modal-panel').getBoundingClientRect();
    const rv = document.getElementById('rosterViewBody').getBoundingClientRect();
    return [r.bottom <= window.innerHeight, rv.bottom <= window.innerHeight];
  }), [true, true]);

  // The head is the only way to bring it back, and the close button rides on
  // it, so it can never be dragged off the edge.
  await page.mouse.move(headMid.x + 220, headMid.y + 160);
  await page.mouse.down();
  await page.mouse.move(4000, 4000, { steps: 10 });
  await page.mouse.up();
  await page.waitForTimeout(200);
  const p2 = await panelBox();
  check('and it cannot be dragged off screen', await page.evaluate(box => {
    const r = document.querySelector('#rosterViewOverlay .modal-panel').getBoundingClientRect();
    return r.left < window.innerWidth - 100 && r.top < window.innerHeight - 40
        && r.right > 100 && r.top >= -1;
  }, p2), true);

  // Parked hard right, it is the close button that has gone over the edge —
  // so the head recentres the panel on a double-click rather than leaving the
  // user to drag it back by the 140px sliver the clamp keeps on screen.
  await page.evaluate(() => {
    const h = document.querySelector('#rosterViewOverlay .modal-head');
    const r = h.getBoundingClientRect();
    h.dispatchEvent(new MouseEvent('dblclick', { bubbles: true,
      clientX: r.left + 60, clientY: r.top + r.height / 2 }));
  });
  await page.waitForTimeout(200);
  check('double-clicking the head recentres it', await panelBox(), p0);

  // The X is now the only dismissal.
  await page.click('#rosterViewCloseBtn');
  await page.waitForTimeout(200);
  check('the close button closes it',
        await page.evaluate(() =>
          document.getElementById('rosterViewOverlay').classList.contains('open')), false);

  if (!ROSTER || !fs.existsSync(ROSTER)) {
    console.log('skip  set ROSTER=/path/to/a/roster.pdf for the find and zoom checks');
    if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
    await browser.close();
    console.log(failed ? `\n${failed} failing` : '\nthe roster viewer stays put and moves when asked');
    process.exit(failed ? 1 : 0);
  }
  await page.setInputFiles('#rosterFile', ROSTER);
  await page.waitForTimeout(400);
  await page.click('#parseBtn');
  await page.waitForTimeout(7000);

  await page.click('#hdrViewBtn');
  await page.waitForTimeout(3000);
  check('the viewer opened on a real file', await page.isVisible('#rosterViewOverlay'), true);

  const rvScrolls = () => page.evaluate(() => {
    const b = document.getElementById('rosterViewBody');
    return b.scrollHeight > b.clientHeight + 200;
  });

  // A one-page consultant roster fits the panel, so there is nothing to scroll
  // until it is zoomed — which is a fair way to reach the state either way,
  // since zooming is exactly when the find row matters most.
  for (let i = 0; i < 4 && !(await rvScrolls()); i++) {
    await page.evaluate(() => document.getElementById('rosterViewIn').click());
    await page.waitForTimeout(250);
  }
  check('there is more than a panel-full of real pages to scroll',
        await rvScrolls(), true);

  await page.evaluate(() => { document.getElementById('rosterViewBody').scrollTop = 1200; });
  await page.waitForTimeout(400);
  check('the find row is still on screen over a scrolled file',
        [await page.isVisible('#rosterViewFind'), await page.isVisible('#rosterViewPrev'),
         await page.isVisible('#rosterViewNext')], [true, true, true]);

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

  // ── Zoom ────────────────────────────────────────────────────────────────
  // A month of roster at the width of a phone is unreadable, so the page box
  // widens past the panel and the body scrolls sideways under it.
  const pageW = () => page.evaluate(() => {
    const pg = document.querySelector('#rosterViewBody .rv-page');
    const bd = document.getElementById('rosterViewBody');
    return { page: Math.round(pg.getBoundingClientRect().width),
             scroll: bd.scrollWidth,
             pct: document.getElementById('rosterViewZoomPct').textContent };
  });

  await page.evaluate(() => document.getElementById('rosterViewFit').click());
  await page.waitForTimeout(250);
  const atFit = await pageW();
  check('Fit is 100%', atFit.pct, '100%');
  check('and the page fits the panel', atFit.page <= atFit.scroll + 1, true);
  check('zoom out is disabled at the bottom',
        await page.evaluate(() => document.getElementById('rosterViewOut').disabled), true);

  await page.evaluate(() => document.getElementById('rosterViewIn').click());
  await page.evaluate(() => document.getElementById('rosterViewIn').click());
  await page.waitForTimeout(300);
  const zoomed = await pageW();
  check('zooming in widens the page', zoomed.page > atFit.page * 1.4, true);
  check('and the body scrolls sideways to reach it', zoomed.scroll > atFit.scroll, true);
  check('the percentage says so', zoomed.pct, '150%');

  // The highlight layer is positioned in percentages, so it has to travel with
  // the page rather than staying where it was painted.
  await page.fill('#rosterViewFind', 'a');
  await page.waitForTimeout(500);
  check('highlights still sit over the page they mark', await page.evaluate(() => {
    const m = document.querySelector('#rosterViewBody .rv-marks i');
    if (!m) return 'no highlight';
    const p = document.querySelector('#rosterViewBody .rv-page').getBoundingClientRect();
    const r = m.getBoundingClientRect();
    return r.left >= p.left - 1 && r.right <= p.right + 1;
  }), true);

  await page.evaluate(() => document.getElementById('rosterViewFit').click());
  await page.waitForTimeout(300);
  check('Fit puts it back', (await pageW()).pct, '100%');

  for (let i = 0; i < 9; i++) await page.evaluate(() => document.getElementById('rosterViewIn').click());
  await page.waitForTimeout(300);
  check('zoom in stops at the top',
        await page.evaluate(() => [document.getElementById('rosterViewZoomPct').textContent,
                                   document.getElementById('rosterViewIn').disabled]),
        ['400%', true]);
  await page.evaluate(() => document.getElementById('rosterViewFit').click());

  // Nothing to say over a page of the file itself.
  check('no note over a PDF page',
        (await page.textContent('#rosterViewNote') || '').trim(), '');
  check('and the empty note takes no space',
        await page.isVisible('#rosterViewNote'), false);

  if (errors.length) { failed++; console.log('\npage errors: ' + errors.join(' | ')); }
  await browser.close();
  console.log(failed ? `\n${failed} failing` : '\nthe find row stays put while the pages scroll');
  process.exit(failed ? 1 : 0);
})();
