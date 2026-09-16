#!/usr/bin/env node
// The phone's step-jump panel, and the leave button at the top of step 1.
//
//   node tests/wizjump.js        (CHROME=/path/to/chrome to pick a binary)
//
// The gating half needs a real parse to reach step 2, which needs a real
// roster PDF — and this repo deliberately holds none (see tests/README.md).
// Point ROSTER at one to run that half; without it those checks are skipped
// and the rest still runs.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const ROSTER = process.env.ROSTER || '';
let failed=0;
const check=(n,g,w)=>{const ok=JSON.stringify(g)===JSON.stringify(w); if(!ok)failed++;
  console.log(`${ok?'ok  ':'FAIL'}  ${n}`); if(!ok)console.log(`        want ${JSON.stringify(w)}\n        got  ${JSON.stringify(g)}`);};
// The number is its own span now — the step's name sits beside it in the
// button, so textContent alone would read "1Upload".
const states = p => p.evaluate(()=>[...document.querySelectorAll('#wizJumpList .btn')]
  .map(b=>({n:b.querySelector('.n').textContent.trim(), off:b.disabled,
            cur:b.getAttribute('aria-current')==='step'})));
(async()=>{
 const b=await chromium.launch(process.env.CHROME?{executablePath:process.env.CHROME}:{});

 // ── phone ────────────────────────────────────────────────────────────────
 const p=await b.newContext({viewport:{width:390,height:844}}).then(c=>c.newPage());
 const errs=[]; p.on('pageerror',e=>errs.push(e.message));
 await p.route('**://**', r=>r.request().url().startsWith('file:')?r.continue():r.abort());
 await p.goto('file:///home/user/VHWRosterTranslator/index.html'); await p.waitForTimeout(1200);

 check('the counter is a button', await p.isVisible('#wizJumpBtn'), true);
 check('and it is in the accent',
   await p.evaluate(()=>getComputedStyle(document.querySelector('#wizStepCount')).color), 'rgb(53, 96, 219)');
 check('leave strip is at the top of step 1', await p.evaluate(()=>{
   const a=document.getElementById('altRoute').getBoundingClientRect();
   const h=document.querySelector('#sec-1 .sec-head').getBoundingClientRect();
   return a.bottom <= h.top+1;}), true);
 check('and it is the blue button', await p.evaluate(()=>
   getComputedStyle(document.getElementById('z1LeaveBtn')).backgroundColor), 'rgb(53, 96, 219)');

 await p.click('#wizJumpBtn'); await p.waitForSelector('#wizJumpOverlay.open');
 check('nothing extracted: only step 1 is open', await states(p),
   [{n:'1',off:false,cur:true},{n:'2',off:true,cur:false},{n:'3',off:true,cur:false},{n:'4',off:true,cur:false}]);
 // The step-1 precondition no longer speaks: Extract data is the button the
 // user is looking at, so the sentence only read it back. The note still
 // carries the reasons that send the user somewhere else, and is hidden when
 // there is nothing to say.
 check('and says nothing about the button already on screen', await p.evaluate(()=>{
   const n=document.getElementById('wizJumpNote');
   return [n.hidden, n.textContent.trim()];}), [true, '']);
 check('the note still speaks for a precondition that is not on screen',
   await p.evaluate(()=>{
     state.consultantData={days:[{date:1,month:4}],doctors:new Set(['Cloete'])};
     wizJumpOpen();
     const n=document.getElementById('wizJumpNote');
     return [n.hidden, n.textContent.includes('Preview schedule')];}), [false, true]);
 await p.evaluate(()=>{ state.consultantData=null; wizJumpOpen(); });
 check('each button names its step', await p.evaluate(()=>
   [...document.querySelectorAll('#wizJumpList .btn .t')].map(t=>t.textContent.trim())),
   ['Upload','Review','Details','Generate']);
 check('and the run-together list underneath is gone', await p.evaluate(()=>{
   const n=document.getElementById('wizJumpNote');
   return /Upload roster files.*Review schedule/.test(n.textContent);}), false);
 await p.keyboard.press('Escape'); await p.waitForTimeout(300);

 // Reaching step 2 the long way needs a real roster; skip if none was given.
 if (!ROSTER || !fs.existsSync(ROSTER)) {
   console.log('skip  the step-2 gating checks — set ROSTER=/path/to/a/roster.pdf to run them');
 } else {
 await p.setInputFiles('#rosterFile', ROSTER);
 await p.click('#parseBtn'); await p.waitForTimeout(6000);
 await p.click('#wizNav1 [data-wiz=next]'); await p.waitForTimeout(600);
 check('we are on step 2', await p.evaluate(()=>wizStep), 2);
 await p.click('#wizJumpBtn'); await p.waitForSelector('#wizJumpOverlay.open');
 const st=await states(p);
 check('step 1 is reachable again, 3 and 4 are not',
   [st[0].off, st[1].cur, st[2].off, st[3].off], [false,true,true,true]);
 await p.click('#wizJumpList .btn[data-go="1"]'); await p.waitForTimeout(700);
 check('clicking 1 goes there', await p.evaluate(()=>wizStep), 1);
 check('and closes the panel', await p.isHidden('#wizJumpOverlay'), true);
 check('scroll lock released', await p.evaluate(()=>document.documentElement.style.overflow), '');
 check('page is live again', await p.evaluate(()=>document.querySelector('.shell').inert), false);
 check('a disabled step cannot be clicked', await p.evaluate(async()=>{
   document.getElementById('wizJumpBtn').click();
   await new Promise(r=>setTimeout(r,200));
   const b=document.querySelector('#wizJumpList .btn[data-go="4"]');
   b.click(); await new Promise(r=>setTimeout(r,200));
   const s=wizStep; document.getElementById('wizJumpCloseBtn').click(); return s;}), 1);
 }
 console.log('phone errors:', errs.length?errs.join(' | '):'none');
 await p.context().close();

 // ── desktop: tabs must still work, counter hidden ────────────────────────
 const d=await b.newContext({viewport:{width:1280,height:900}}).then(c=>c.newPage());
 const e2=[]; d.on('pageerror',e=>e2.push(e.message));
 await d.route('**://**', r=>r.request().url().startsWith('file:')?r.continue():r.abort());
 await d.goto('file:///home/user/VHWRosterTranslator/index.html'); await d.waitForTimeout(1200);
 check('the jump button is hidden on desktop', await d.isHidden('#wizJumpBtn'), true);
 check('the four tabs are shown instead',
   await d.evaluate(()=>document.querySelectorAll('#wizSteps .wizstep').length), 4);
 // The leave button is exactly as wide as the department upload zone beneath
 // it, not the full page. Measured, because the grid that makes it so has to
 // survive every reflow.
 // Guard: a display:none element measures 0x0, and 0-0 would satisfy the
 // comparison below without either box existing.
 check('the leave button is actually on screen', await d.evaluate(()=>
   document.getElementById('z1LeaveBtn').getBoundingClientRect().width > 100), true);
 // The icon sits at the right-hand end, square, like a .dl download card's.
 check('it carries the external-link icon on the right', await d.evaluate(()=>{
   const b=document.getElementById('z1LeaveBtn').getBoundingClientRect();
   const i=document.querySelector('#z1LeaveBtn .alt-route-ico');
   if(!i) return 'no icon';
   const r=i.getBoundingClientRect();
   return [Math.round(r.width), Math.round(r.height),
           // nearer the right edge than the left, and inset by the button's
           // 20px padding plus its 1px border — a .dl card has no border and
           // so sits at a flat 20, the same offset-by-one the nav chevrons hit
           r.left - b.left > b.width/2, Math.round(b.right - r.right)];}),
   [22,22,true,21]);
 check('the leave button matches the left upload zone', await d.evaluate(()=>{
   const b=document.getElementById('z1LeaveBtn').getBoundingClientRect();
   const z=document.querySelector('#sec-1 .two > .upload-zone-col').getBoundingClientRect();
   return [Math.round(b.width-z.width), Math.round(b.left-z.left)];}), [0,0]);
 check('and it is not the full width of the page', await d.evaluate(()=>{
   const b=document.getElementById('z1LeaveBtn').getBoundingClientRect();
   const w=document.querySelector('#sec-1').getBoundingClientRect();
   return b.width < w.width*0.75;}), true);
 // With no consultant zone the upload grid goes to one column, and the button
 // has to follow it rather than staying half-width beside nothing.
 check('with no consultant roster both go full width', await d.evaluate(()=>{
   document.getElementById('altRouteSpacer').style.display='none';
   document.getElementById('consultantZoneWrap').style.display='none';
   const b=document.getElementById('z1LeaveBtn').getBoundingClientRect();
   const z=document.querySelector('#sec-1 .two > .upload-zone-col').getBoundingClientRect();
   return Math.round(b.width-z.width);}), 0);
 console.log('desktop errors:', e2.length?e2.join(' | '):'none');

 await b.close();
 console.log(failed?`\n${failed} failing`:'\nboth changes work');
 process.exit(failed?1:0);
})();
