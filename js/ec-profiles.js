// ═══════════════════════════════════════════════════════════════
// ec-profiles.js — EC profile system: Supabase fetch, profile
//                  application, EC picker UI, and setup wizard.
//
//   fetchProfiles()                 → Promise<profile[]>
//   applyProfile(profile)           — sets activeProfile + UI
//   showEcSelected / showEcPicker / showEcOffline
//   initEcSelector()                — boot entry point
//   openWizard / closeWizard / wizGoto / wizValidate
//   wizRenderPdf / wizSetupYDrag / wizInitColStep
//   setupDivDrag / renderColChips / wizBuildJson
//
// Adding a new EC: insert a row in Supabase ec_profiles table.
// No code changes required here.
//
// Depends on: config.js (SUPA_URL, SUPA_ANON, LS_PROFILE_KEY,
//             LS_PROFILES_KEY, activeProfile)
// ═══════════════════════════════════════════════════════════════

// STEP 0 — EC PROFILE SYSTEM
// ═══════════════════════════════════════════════════════════════


async function fetchProfiles() {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 8000);
  try {
    const res = await fetch(
      `${SUPA_URL}/rest/v1/ec_profiles?status=eq.approved&select=id,ec_name,ec_short,profile`,
      {
        method: 'GET',
        signal: controller.signal,
        headers: {
          'apikey': SUPA_ANON,
          'Authorization': 'Bearer ' + SUPA_ANON,
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        }
      }
    );
    if (!res.ok) throw new Error('Fetch failed: ' + res.status);
    return res.json();
  } finally {
    clearTimeout(timer);
  }
}

function applyProfile(profile) {
  activeProfile = profile;
  // The EC name lives in the header chip (#ecSelectedName) — the brand
  // line stays constant, and the offline badge is cleared here.
  $('headerTitle').textContent = 'Hospital Roster Translator';
  const modeEl = $('ecMode');
  if (modeEl) { modeEl.textContent = ''; modeEl.style.display = 'none'; }
  // Store in localStorage for offline use
  localStorage.setItem(LS_PROFILE_KEY, JSON.stringify(profile));
  // Show/hide consultant zone based on profile type
  // (safe to call even before DOM is fully ready — guarded inside)
  if (typeof updateConsultantZoneVisibility === 'function') updateConsultantZoneVisibility();
  // Department-specific presentation: supervisor names and what this roster
  // calls a duty come from the profile, not from the EC defaults.
  if (typeof applySupervisorMode === 'function') applySupervisorMode();
  if (typeof applyDutyNoun === 'function') applyDutyNoun();
}

function showEcSelected(name) {
  $('ecLoadingRow').style.display  = 'none';
  $('ecPickerRow').style.display   = 'none';
  $('ecSelectedRow').style.display = '';
  $('ecSelectedName').textContent  = name;
  // A department is chosen, so the picker screen gives way to the wizard.
  const _mh=document.querySelector('.masthead'), _wb=document.getElementById('wizBar');
  if(_mh) _mh.hidden=true;
  if(_wb) _wb.hidden=false;
  if(typeof wizGo==='function') wizGo(typeof wizStep==='number'?wizStep:1);
}

function showEcPicker(profiles) {
  $('ecLoadingRow').style.display  = 'none';
  $('ecSelectedRow').style.display = 'none';
  $('ecPickerRow').style.display   = '';
  const _mh=document.querySelector('.masthead'), _wb=document.getElementById('wizBar');
  if(_mh) _mh.hidden=false;
  if(_wb) _wb.hidden=true;
  for(const _s of ['sec-1','sec-2','sec-3','sec-4']){
    const _el=document.getElementById(_s); if(_el) _el.hidden=true;
  }
  const sel = $('ecSelect');
  // Clear existing options except the placeholder
  while (sel.options.length > 1) sel.remove(1);
  profiles.forEach(p => {
    const opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = p.ec_name;
    opt.dataset.profile = JSON.stringify(p.profile);
    opt.dataset.name = p.ec_name;
    sel.appendChild(opt);
  });
}

function showEcOffline(cachedProfile) {
  $('ecOfflineNote').style.display = '';
  if (cachedProfile) {
    applyProfile(cachedProfile);
    showEcSelected(cachedProfile.ec_name || 'Saved EC');
  } else {
    // No cache and no network — hide spinner, show just the warning
    $('ecLoadingRow').style.display = 'none';
    $('ecPickerRow').style.display  = 'none';
  }
}

async function initEcSelector() {
  console.log('[EC] initEcSelector started');
  const cached = localStorage.getItem(LS_PROFILE_KEY);
  const cachedProfile = cached ? JSON.parse(cached) : null;
  console.log('[EC] cached profile:', cachedProfile ? cachedProfile.ec_name : 'none');

  // Try fetching fresh profiles
  let profiles = null;
  try {
    console.log('[EC] fetching profiles...');
    profiles = await fetchProfiles();
    console.log('[EC] fetch succeeded, count:', profiles ? profiles.length : 0);
    // Cache the list for next time
    localStorage.setItem(LS_PROFILES_KEY, JSON.stringify(profiles));
  } catch (e) {
    console.warn('[EC] fetch failed:', e.message || e);
  }

  if (!profiles || profiles.length === 0) {
    console.log('[EC] no profiles — showing offline');
    showEcOffline(cachedProfile);
    return;
  }

  // Online — check if we have a saved selection that still exists
  if (cachedProfile) {
    const still = profiles.find(p => p.ec_name === cachedProfile.ec_name);
    if (still) {
      console.log('[EC] restoring saved profile:', still.ec_name);
      const freshProfile = { ...still.profile, ec_name: still.ec_name, ec_short: still.ec_short || still.profile?.ec_short };
      applyProfile(freshProfile);
      showEcSelected(still.ec_name);
      return;
    }
    // Saved locally but absent from the approved list: an EC set up through
    // the wizard and still awaiting approval, or one loaded by hand. Keep
    // using it rather than dropping the user back to the picker — otherwise
    // a locally activated profile would survive only until the next reload.
    // "Change" still opens the picker when they want a different EC.
    console.log('[EC] restoring local-only profile:', cachedProfile.ec_name);
    applyProfile(cachedProfile);
    showEcSelected(cachedProfile.ec_name || cachedProfile.ec_short || 'Saved EC');
    return;
  }

  // No saved selection — show picker
  console.log('[EC] showing picker with', profiles.length, 'profiles');
  showEcPicker(profiles);
}

// Wire up EC selector events
document.addEventListener('DOMContentLoaded', () => {

  // Confirm button
  $('ecConfirmBtn').addEventListener('click', () => {
    const sel = $('ecSelect');
    const opt = sel.options[sel.selectedIndex];
    if (!opt || !opt.dataset.profile) return;
    const profile = JSON.parse(opt.dataset.profile);
    const fullProfile = { ...profile, ec_name: opt.dataset.name, ec_short: profile.ec_short || opt.dataset.name, roster_type: profile.roster_type };
    applyProfile(fullProfile);
    showEcSelected(opt.dataset.name);
  });

  // Enable confirm button only when an EC is selected
  $('ecSelect').addEventListener('change', () => {
    $('ecConfirmBtn').disabled = !$('ecSelect').value;
  });

  // Change link — go back to picker
  window.reopenEcPicker = async function reopenEcPicker() {
    $('ecSelectedRow').style.display = 'none';
    $('ecLoadingRow').style.display  = '';
    $('headerTitle').textContent = 'Hospital Roster Translator';
    const step1 = $('step1');
    if (step1) step1.style.display = 'none';
    let profiles = null;
    try {
      profiles = await fetchProfiles();
      localStorage.setItem(LS_PROFILES_KEY, JSON.stringify(profiles));
    } catch(e) {
      const cached = localStorage.getItem(LS_PROFILES_KEY);
      profiles = cached ? JSON.parse(cached) : [];
    }
    showEcPicker(profiles);
  };
  $('ecChangeBtn').addEventListener('click', () => window.reopenEcPicker());

  // ── EC Profile Setup Wizard ──────────────────────────────────────────────────
  let wizState = {
    step: 1,
    pdfBuf: null,          // ArrayBuffer of uploaded PDF
    canvasScale: 1,        // PDF→canvas scale factor
    pageW: 0, pageH: 0,    // rendered page dimensions (px)
    divPositions: [],       // array of 5 x-positions (canvas px) for column dividers
    draggingY: false,
    draggingDiv: -1,
    yLinePx: 0,             // red line Y (canvas px)
  };
  const COL_NAMES   = ['slot1','slot2','slot3','meetings','leave','call'];
  // Offered as activity types for a Word-table profile. The EC WD/WE shift
  // types are deliberately absent — those belong to the EC roster alone.
  const WIZ_LEAVE_DEFAULTS = ['Leave - Annual','Leave - Sick','Leave - Family Responsibility',
    'Leave - Study','Leave - Special','Leave - Prenatal','Leave - Maternity','Leave - Paternity',
    'Workshop','Course','Conference'];
  const wizRosterType = () =>
    document.querySelector('input[name="wizRosterType"]:checked')?.value || 'shift';
  // How the department works, which is a separate question from what file the
  // roster arrives in. Only the grid roster asks it; the other two paths carry
  // their pattern in the parser they use.
  const wizPattern = () =>
    document.querySelector('input[name="wizPattern"]:checked')?.value || 'calls';
  const COL_COLOURS = ['#2D6B45','#1A6B3A','#5b9bd5','#9b59b6','#c0392b','#e67e22'];
  const WIZ_STEPS   = 4;

  function openWizard() {
    wizState = { step:1, pdfBuf:null, canvasScale:1, pageW:0, pageH:0, divPositions:[], draggingY:false, draggingDiv:-1, yLinePx:0 };
    $('wizardOverlay').style.display = '';
    wizGoto(1);
  }
  function closeWizard() { $('wizardOverlay').style.display = 'none'; }
  // Opened from the "Set up a new EC" link in the footer.
  window.openWizard = openWizard;

  // "Set up your EC" opens the wizard; the mailto beside it stays as a
  // fallback for anyone who would rather have it done for them.
  const notListed = $('ecNotListedLink');
  if (notListed) notListed.addEventListener('click', e => { e.preventDefault(); openWizard(); });
  $('wizCloseBtn').addEventListener('click', closeWizard);
  $('wizardOverlay').addEventListener('click', e => { if (e.target === $('wizardOverlay')) closeWizard(); });

  function wizGoto(step) {
    wizState.step = step;
    const rt = wizRosterType();
    const isConsultant = rt === 'consultant';
    const isTable      = rt === 'table';
    // Show/hide steps
    for (let i = 1; i <= WIZ_STEPS; i++) {
      const el = $('wizStep' + i);
      if (el) el.style.display = i === step ? '' : 'none';
    }
    // Steps 2 and 3 hold one pane per roster type.
    const pane = (consId, tblId) => {
      const c = $(consId), t = $(tblId);
      if (c) c.style.display = isTable ? 'none' : '';
      if (t) t.style.display = isTable ? '' : 'none';
    };
    pane('wizStep2Cons', 'wizStep2Table');
    pane('wizStep3Cons', 'wizStep3Table');
    if ($('wizTab2')) $('wizTab2').innerHTML = isTable ? '2 &nbsp;Upload Roster File' : '2 &nbsp;Upload PDF';
    if ($('wizTab3')) $('wizTab3').innerHTML = isTable ? '3 &nbsp;Columns &amp; Hours' : '3 &nbsp;Columns &amp; Rules';
    // Update tabs
    for (let i = 1; i <= WIZ_STEPS; i++) {
      const tab = $('wizTab' + i);
      if (!tab) continue;
      tab.className = 'wiz-step-tab' + (i === step ? ' active' : (i < step ? ' done' : ''));
    }
    $('wizStepLabel').textContent = `Step ${step} of ${WIZ_STEPS}`;
    $('wizBackBtn').style.display = step > 1 ? '' : 'none';
    $('wizNextBtn').style.display = step < WIZ_STEPS ? '' : 'none';
    $('wizSubmitBtn').style.display = step === WIZ_STEPS ? '' : 'none';
    $('wizSubmitMsg').style.display = 'none';

    // Step-specific init. A shift-only EC needs no sample file, so it jumps
    // straight to review.
    if ((step === 2 || step === 3) && !isConsultant && !isTable) { wizGoto(4); return; }
    if (step === 3 && isConsultant) wizInitColStep();
    if (step === 3 && isTable) wizInitTableCols();
    if (step === 4) wizBuildJson();
  }

  $('wizBackBtn').addEventListener('click', () => {
    const rt = wizRosterType();
    if (wizState.step === 4 && rt === 'shift') wizGoto(1);
    else if (wizState.step > 1) wizGoto(wizState.step - 1);
  });

  $('wizNextBtn').addEventListener('click', () => {
    if (!wizValidate(wizState.step)) return;
    wizGoto(wizState.step + 1);
  });

  // Step 1: show/hide data_start_y when consultant selected
  document.querySelectorAll('input[name="wizRosterType"]').forEach(r => {
    r.addEventListener('change', () => {
      $('wizDataYWrap').style.display = r.value === 'consultant' ? '' : 'none';
      const isTable = r.value === 'table';
      if ($('wizPatternWrap')) $('wizPatternWrap').style.display = isTable ? '' : 'none';
      if ($('wizTab2')) $('wizTab2').innerHTML = isTable ? '2 &nbsp;Upload Roster File' : '2 &nbsp;Upload PDF';
      if ($('wizTab3')) $('wizTab3').innerHTML = isTable ? '3 &nbsp;Columns &amp; Hours' : '3 &nbsp;Columns &amp; Rules';
    });
  });

  document.querySelectorAll('input[name="wizPattern"]').forEach(r => {
    r.addEventListener('change', () => { if (wizState.step === 3) wizInitTableCols(); });
  });

  function wizValidate(step) {
    if (step === 1) {
      if (!$('wizEcName').value.trim()) { $('wizEcName').focus(); return false; }
      if (!$('wizEcShort').value.trim()) { $('wizEcShort').focus(); return false; }
    }
    if (step === 2 && wizRosterType() === 'consultant') {
      if (!wizState.pdfBuf) { $('wizPdfStatus').textContent = 'Please upload a PDF first.'; return false; }
    }
    if (step === 2 && wizRosterType() === 'table') {
      if (!wizState.tableRows) { $('wizDocStatus').textContent = 'Please upload a sample roster first.'; return false; }
    }
    if (step === 3 && wizRosterType() === 'table') {
      const m = wizReadColMap();
      if (m.dateCol < 0) { alert('Mark which column holds the date.'); return false; }
      if (!Object.keys(m.roleCols).length) { alert('Name at least one duty column.'); return false; }
      if (wizPattern() === 'shifts') {
        const missing = Object.entries(m.roleCols)
          .filter(([, i]) => !(m.map[i] || {}).start || !(m.map[i] || {}).normEnd)
          .map(([n]) => n);
        if (missing.length) {
          alert('Give a start and a normal finish for every shift column: ' + missing.join(', '));
          return false;
        }
      }
    }
    return true;
  }

  // ── Step 2: PDF upload + render ──────────────────────────────────────────
  $('wizPdfFile').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    $('wizPdfStatus').textContent = 'Rendering PDF…';
    try {
      wizState.pdfBuf = await readFile(file);
      await wizRenderPdf(wizState.pdfBuf, 'wizCanvas', 'wizCanvasWrap', 'wizYLine');
      $('wizPdfStatus').textContent = '✓ ' + file.name;
      $('wizCanvasHint').style.display = '';
    } catch(err) {
      $('wizPdfStatus').textContent = 'Error: ' + err.message;
    }
    e.target.value = '';
  });

  async function wizRenderPdf(buf, canvasId, wrapId, yLineId) {
    const pdf  = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
    const page = await pdf.getPage(1);
    const wrap = $(wrapId);
    const maxW = wrap.clientWidth || 700;
    const vp0  = page.getViewport({ scale: 1 });
    const scale = Math.min((maxW - 2) / vp0.width, 2.5);
    const vp   = page.getViewport({ scale });
    const canvas = $(canvasId);
    canvas.width  = vp.width;
    canvas.height = vp.height;
    wizState.canvasScale = scale;
    wizState.pageW = vp.width;
    wizState.pageH = vp.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
    wrap.style.display = '';
    // Position red Y line
    const dataY = parseInt($('wizDataY').value) || 188;
    const yPx   = dataY * scale;
    wizState.yLinePx = yPx;
    if (yLineId) {
      const yLine = $(yLineId);
      yLine.style.top = yPx + 'px';
      $('wizYVal').textContent = dataY;
    }
    // Default col dividers: evenly spaced from 20% to 90% of canvas width
    if (wizState.divPositions.length === 0) {
      wizState.divPositions = [0.25, 0.38, 0.51, 0.64, 0.77].map(f => Math.round(f * vp.width));
    }
    // Setup Y line drag
    if (yLineId) wizSetupYDrag($(yLineId), $(wrapId));
  }

  function wizSetupYDrag(yLine, wrap) {
    const startDrag = (clientY) => {
      wizState.draggingY = true;
      const rect = wrap.getBoundingClientRect();
      const onMove = (cy) => {
        if (!wizState.draggingY) return;
        const raw   = cy - rect.top;
        const clamped = Math.max(10, Math.min(wizState.pageH - 10, raw));
        wizState.yLinePx = clamped;
        yLine.style.top  = clamped + 'px';
        const dataY = Math.round(clamped / wizState.canvasScale);
        $('wizYVal').textContent = dataY;
        $('wizDataY').value      = dataY;
        $('wizDataYDisplay').textContent = dataY;
      };
      const onUp = () => {
        wizState.draggingY = false;
        document.removeEventListener('mousemove', mMove);
        document.removeEventListener('mouseup', mUp);
        document.removeEventListener('touchmove', tMove);
        document.removeEventListener('touchend', mUp);
      };
      const mMove = e => onMove(e.clientY);
      const tMove = e => onMove(e.touches[0].clientY);
      const mUp = onUp;
      document.addEventListener('mousemove', mMove);
      document.addEventListener('mouseup',   mUp);
      document.addEventListener('touchmove', tMove, {passive:false});
      document.addEventListener('touchend',  mUp);
    };
    yLine.addEventListener('mousedown', e  => { e.preventDefault(); startDrag(e.clientY); });
    yLine.addEventListener('touchstart', e => { e.preventDefault(); startDrag(e.touches[0].clientY); }, {passive:false});
  }

  // ── Word-table branch: upload, map columns, build the profile ────────────
  const wizDocFileEl = $('wizDocFile');
  if (wizDocFileEl) wizDocFileEl.addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    $('wizDocStatus').textContent = 'Reading\u2026';
    try {
      const buf = await readFile(file);
      const det = await detectWordTable(buf, file.name);
      wizState.tableRows = det.rows;
      wizState.tableCols = det.columns;
      $('wizDocStatus').textContent =
        `\u2713 ${file.name} \u2014 ${det.columns} columns, ${det.rows.length} rows`;
      wizRenderDocPreview();
      $('wizDocPreviewWrap').style.display = '';
    } catch (err) {
      wizState.tableRows = null;
      $('wizDocStatus').textContent = 'Error: ' + err.message;
      $('wizDocPreviewWrap').style.display = 'none';
    }
    e.target.value = '';
  });

  function wizRenderDocPreview() {
    const rows = wizState.tableRows || [];
    const tbl = $('wizDocPreview');
    if (!tbl) return;
    const head = rows[0] || [];
    const body = rows.slice(1, 7);
    tbl.innerHTML =
      '<thead><tr>' + head.map((c, i) =>
        `<th>${i}: ${c || '<span style="opacity:.5">(blank)</span>'}</th>`).join('') + '</tr></thead>' +
      '<tbody>' + body.map(r =>
        '<tr>' + head.map((_, i) => `<td>${(r[i] || '')}</td>`).join('') + '</tr>').join('') + '</tbody>';
  }

  // Guess each column's meaning so the common case needs no clicking.
  function wizGuessColumns() {
    const rows = wizState.tableRows || [];
    const header = rows[0] || [];
    const body = rows.slice(1);
    const DAYS = /^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$/i;
    const score = (i, test) =>
      body.reduce((n, r) => n + (test(String(r[i] || '').trim()) ? 1 : 0), 0);
    const out = [];
    for (let i = 0; i < (wizState.tableCols || header.length); i++) {
      const label = String(header[i] || '').trim();
      let kind = 'role';
      if (score(i, v => !!parseDateCell(v)) > body.length * 0.5) kind = 'date';
      else if (score(i, v => DAYS.test(v)) > body.length * 0.5) kind = 'day';
      else if (!label && score(i, v => !!v) === 0) kind = 'ignore';
      out.push({ kind, name: label || `Column ${i}`, isCall: kind === 'role' });
    }
    return out;
  }

  function wizInitTableCols() {
    const wrap = $('wizTblCols');
    if (!wrap) return;
    if (!wizState.colMap || wizState.colMap.length !== (wizState.tableCols || 0)) {
      wizState.colMap = wizGuessColumns();
    }
    const rows = wizState.tableRows || [];
    const samples = i => rows.slice(1, 5).map(r => r[i]).filter(Boolean).slice(0, 3).join(', ') || '\u2014';
    const shifts = wizPattern() === 'shifts';
    // A shift department gives each column its own times; a call department
    // shares three band rows and only needs to know which columns are call.
    if ($('wizShiftHours')) $('wizShiftHours').style.display = shifts ? '' : 'none';
    if ($('wizCallHours'))  $('wizCallHours').style.display  = shifts ? 'none' : '';
    const hint = $('wizColsHint');
    if (hint) hint.textContent = shifts
      ? 'Mark the date and weekday columns, then name each shift column \u2014 whatever staff call it. Each one gets its own times below.'
      : 'Mark the date and weekday columns, then name each duty column. Untick \u201Con call\u201D for daytime work such as theatre sessions \u2014 those earn no overnight overtime and no day off afterwards.';
    wrap.innerHTML =
      '<table class="wiz-rules-table"><thead><tr><th>#</th><th>Sample values</th>' +
      '<th>Meaning</th><th>Name on timesheet</th>' + (shifts ? '' : '<th>On call?</th>') +
      '</tr></thead><tbody>' +
      wizState.colMap.map((c, i) => `
        <tr>
          <td style="font-family:var(--font-body);font-variant-numeric:tabular-nums;">${i}</td>
          <td style="font-size:11px;color:var(--text-muted);">${samples(i)}</td>
          <td><select class="wiz-input wizColKind" data-i="${i}" style="min-width:120px;">
            <option value="date"${c.kind==='date'?' selected':''}>Date</option>
            <option value="day"${c.kind==='day'?' selected':''}>Weekday</option>
            <option value="role"${c.kind==='role'?' selected':''}>${shifts?'Shift column':'Duty column'}</option>
            <option value="ignore"${c.kind==='ignore'?' selected':''}>Ignore</option>
          </select></td>
          <td><input class="wiz-input wizColName" data-i="${i}" value="${(c.name||'').replace(/"/g,'&quot;')}"
              ${c.kind==='role'?'':'disabled'} style="min-width:140px;"></td>
          ${shifts ? '' : `<td style="text-align:center;"><input type="checkbox" class="wizColCall" data-i="${i}"
              ${c.isCall?'checked':''} ${c.kind==='role'?'':'disabled'}></td>`}
        </tr>`).join('') + '</tbody></table>';
    if (shifts) wizRenderShiftTimes();

    wrap.querySelectorAll('.wizColKind').forEach(sel =>
      sel.addEventListener('change', () => {
        const i = +sel.dataset.i;
        wizState.colMap[i].kind = sel.value;
        wizInitTableCols();
      }));
    wrap.querySelectorAll('.wizColName').forEach(inp => {
      inp.addEventListener('input', () => { wizState.colMap[+inp.dataset.i].name = inp.value; });
      // On blur, not on every keystroke: the shift-times table below names
      // each column, and re-rendering mid-word would steal the caret.
      inp.addEventListener('change', () => { if (wizPattern() === 'shifts') wizRenderShiftTimes(); });
    });
    wrap.querySelectorAll('.wizColCall').forEach(cb =>
      cb.addEventListener('change', () => { wizState.colMap[+cb.dataset.i].isCall = cb.checked; }));

    const leave = $('wizLeaveTypes');
    if (leave && !leave.children.length) {
      leave.innerHTML = WIZ_LEAVE_DEFAULTS.map(t =>
        `<label class="wiz-radio-lbl"><input type="checkbox" class="wizLeaveType" value="${t}" checked> ${t}</label>`).join('');
    }
  }

  // One row per shift column: when it starts, when normal hours end, and
  // when the overtime tail ends. Everything the timesheet needs.
  function wizRenderShiftTimes() {
    const host = $('wizShiftTimes');
    if (!host) return;
    const cols = (wizState.colMap || [])
      .map((c, i) => ({ c, i })).filter(x => x.c.kind === 'role');
    if (!cols.length) {
      host.innerHTML = '<p class="wiz-hint">Mark at least one shift column above.</p>';
      return;
    }
    host.innerHTML =
      '<table class="wiz-rules-table"><thead><tr><th>Shift column</th>' +
      '<th>Starts</th><th>Normal hours end</th><th>Overtime ends</th></tr></thead><tbody>' +
      cols.map(({ c, i }) => `
        <tr>
          <td>${(c.name || 'Column ' + i).replace(/</g,'&lt;')}</td>
          <td><input class="wiz-time wizShiftStart" data-i="${i}" value="${c.start||''}" placeholder="08:00"></td>
          <td><input class="wiz-time wizShiftNormEnd" data-i="${i}" value="${c.normEnd||''}" placeholder="16:00"></td>
          <td><input class="wiz-time wizShiftOtEnd" data-i="${i}" value="${c.otEnd||''}" placeholder="18:00"></td>
        </tr>`).join('') + '</tbody></table>';
    const bind = (cls, key) => host.querySelectorAll('.' + cls).forEach(inp =>
      inp.addEventListener('input', () => { wizState.colMap[+inp.dataset.i][key] = inp.value.trim(); }));
    bind('wizShiftStart', 'start');
    bind('wizShiftNormEnd', 'normEnd');
    bind('wizShiftOtEnd', 'otEnd');
  }

  function wizReadColMap() {
    const map = wizState.colMap || [];
    const columns = [], roleCols = {};
    let dateCol = -1, dayCol = -1;
    map.forEach((c, i) => {
      columns.push(c.kind === 'role' ? (c.name || `Column ${i}`)
                 : c.kind === 'date' ? 'Date'
                 : c.kind === 'day'  ? 'Day' : (c.name || `Column ${i}`));
      if (c.kind === 'date' && dateCol < 0) dateCol = i;
      if (c.kind === 'day'  && dayCol  < 0) dayCol  = i;
      if (c.kind === 'role') roleCols[c.name || `Column ${i}`] = i;
    });
    return { columns, roleCols, dateCol, dayCol, map };
  }

  function wizBuildTableProfile(ecShort) {
    const t = id => ($(id) ? $(id).value.trim() : '') || null;
    const pair = (a, b) => (t(a) && t(b)) ? [t(a), t(b)] : null;
    const { columns, roleCols, dateCol, dayCol, map } = wizReadColMap();

    // A shift department has no ordinary weekday to fall back on and no
    // post-call day off: you work the shifts you are rostered onto, and each
    // column carries its own times. getTableShifts already does exactly that
    // when default_weekday is null and post_call_off is false.
    if (wizPattern() === 'shifts') {
      const shiftRules = {};
      for (const [name, idx] of Object.entries(roleCols)) {
        const c = map[idx] || {};
        const norm = (c.start && c.normEnd) ? [c.start, c.normEnd] : null;
        const ot   = (c.normEnd && c.otEnd) ? [c.normEnd, c.otEnd] : null;
        const bands = { normal: norm, ot1: ot, ot2: null };
        shiftRules[name] = {
          is_call: false,
          weekday: bands,
          weekend_ph: bands,
          label_weekday: `${name} - Weekday`,
          label_weekend: `${name} - Weekend`,
          label_ph:      `${name} - Public Holiday`,
        };
      }
      return {
        ec_short: ecShort,
        roster_type: 'table',
        work_pattern: 'shifts',
        table: {
          header_row: $('wizDocHeader') ? $('wizDocHeader').checked : true,
          columns, date_col: dateCol, day_col: dayCol,
          role_columns: roleCols, ignore_tokens: [],
        },
        default_weekday: null,
        post_call_off: false,
        duty_noun: 'shifts',
        leave_types: [...document.querySelectorAll('.wizLeaveType')]
          .filter(cb => cb.checked).map(cb => cb.value),
        role_rules: shiftRules,
      };
    }

    const dayRule  = { normal: pair('wt_day_nf','wt_day_nt'), ot1: pair('wt_day_o1f','wt_day_o1t'), ot2: pair('wt_day_o2f','wt_day_o2t') };
    const callWd   = { normal: pair('wt_wd_nf','wt_wd_nt'),   ot1: pair('wt_wd_o1f','wt_wd_o1t'),   ot2: pair('wt_wd_o2f','wt_wd_o2t') };
    const callWe   = { normal: pair('wt_we_nf','wt_we_nt'),   ot1: pair('wt_we_o1f','wt_we_o1t'),   ot2: pair('wt_we_o2f','wt_we_o2t') };
    const DAY_LABEL = 'Normal Hours - Weekday';

    const role_rules = {};
    for (const [name, idx] of Object.entries(roleCols)) {
      const isCall = (map[idx] || {}).isCall !== false;
      role_rules[name] = isCall ? {
        is_call: true,
        weekday: callWd,
        weekend_ph: callWe,
        label_weekday: `${name} On Call - Weekday`,
        label_weekend: `${name} On Call - Weekend`,
        label_ph:      `${name} On Call - Public Holiday`,
      } : {
        is_call: false,
        weekday: dayRule,
        weekend_ph: null,
        label_weekday: DAY_LABEL, label_weekend: DAY_LABEL, label_ph: DAY_LABEL,
      };
    }

    const leaveTypes = [...document.querySelectorAll('.wizLeaveType')]
      .filter(cb => cb.checked).map(cb => cb.value);

    return {
      ec_short: ecShort,
      roster_type: 'table',
      table: {
        header_row: $('wizDocHeader') ? $('wizDocHeader').checked : true,
        columns, date_col: dateCol, day_col: dayCol,
        role_columns: roleCols, ignore_tokens: [],
      },
      work_pattern: 'calls',
      default_weekday: { ...dayRule, label: DAY_LABEL },
      post_call_off: $('wizPostCallOff') ? $('wizPostCallOff').checked : true,
      leave_types: leaveTypes,
      role_rules,
    };
  }

  // ── Step 3: column dividers ───────────────────────────────────────────────
  function wizInitColStep() {
    const wrap3 = $('wizCanvasWrap3');
    const canvas3 = $('wizCanvas3');
    // Copy rendered image from step 2 canvas
    if (wizState.pageW && wizState.pageH) {
      canvas3.width  = wizState.pageW;
      canvas3.height = wizState.pageH;
      const ctx = canvas3.getContext('2d');
      ctx.drawImage($('wizCanvas'), 0, 0);
    }
    // Remove old dividers
    wrap3.querySelectorAll('.wiz-div').forEach(d => d.remove());
    // Build 5 dividers
    COL_NAMES.slice(0, 5).forEach((_, i) => {
      const div = document.createElement('div');
      div.className = 'wiz-div';
      div.dataset.idx = i;
      div.style.cssText = `position:absolute;top:0;bottom:0;width:3px;background:${COL_COLOURS[i]};opacity:0.8;cursor:ew-resize;z-index:10;`;
      div.style.left = wizState.divPositions[i] + 'px';
      // Label pip
      const pip = document.createElement('div');
      pip.style.cssText = `position:absolute;top:4px;left:50%;transform:translateX(-50%);background:${COL_COLOURS[i]};color:#fff;font-size:9px;font-family:var(--font-body);padding:1px 4px;border-radius:2px;white-space:nowrap;`;
      pip.textContent = COL_NAMES[i + 1];
      div.appendChild(pip);
      wrap3.appendChild(div);
      setupDivDrag(div, wrap3);
    });
    // Render colour chips
    renderColChips();
    if (wizState.pageW && wizState.pageH) wrap3.style.display = '';
  }

  function setupDivDrag(divEl, wrap) {
    const startDrag = (clientX) => {
      const idx  = parseInt(divEl.dataset.idx);
      const rect = wrap.getBoundingClientRect();
      const onMove = (cx) => {
        const raw = cx - rect.left;
        const lo  = idx > 0 ? wizState.divPositions[idx - 1] + 8 : 8;
        const hi  = idx < 4 ? wizState.divPositions[idx + 1] - 8 : wizState.pageW - 8;
        const clamped = Math.max(lo, Math.min(hi, raw));
        wizState.divPositions[idx] = clamped;
        divEl.style.left = clamped + 'px';
        renderColChips();
      };
      const onUp = () => {
        document.removeEventListener('mousemove', mMove);
        document.removeEventListener('mouseup',   mUp);
        document.removeEventListener('touchmove', tMove);
        document.removeEventListener('touchend',  mUp);
      };
      const mMove = e => onMove(e.clientX);
      const tMove = e => { e.preventDefault(); onMove(e.touches[0].clientX); };
      const mUp = onUp;
      document.addEventListener('mousemove', mMove);
      document.addEventListener('mouseup',   mUp);
      document.addEventListener('touchmove', tMove, {passive:false});
      document.addEventListener('touchend',  mUp);
    };
    divEl.addEventListener('mousedown',  e => { e.preventDefault(); startDrag(e.clientX); });
    divEl.addEventListener('touchstart', e => { e.preventDefault(); startDrag(e.touches[0].clientX); }, {passive:false});
  }

  function renderColChips() {
    const wrap = $('wizColLabels');
    wrap.innerHTML = '';
    const positions = [0, ...wizState.divPositions, wizState.pageW];
    COL_NAMES.forEach((name, i) => {
      const x_min = Math.round(positions[i] / wizState.canvasScale);
      const x_max = Math.round(positions[i + 1] / wizState.canvasScale);
      const chip = document.createElement('span');
      chip.className = 'wiz-col-chip';
      chip.style.background = COL_COLOURS[i] + '22';
      chip.style.color = COL_COLOURS[i];
      chip.style.border = '1px solid ' + COL_COLOURS[i] + '66';
      chip.textContent = `${name}: ${x_min}–${x_max}`;
      wrap.appendChild(chip);
    });
  }

  // ── Step 4: build profile JSON ───────────────────────────────────────────
  function wizBuildJson() {
    const rt = wizRosterType();
    const isConsultant = rt === 'consultant';
    if (rt === 'table') {
      const built = wizBuildTableProfile($('wizEcShort').value.trim());
      $('wizJsonPreview').textContent = JSON.stringify(built, null, 2);
      return { ecName: $('wizEcName').value.trim(), profile: built };
    }
    const profile = {
      ec_short:    $('wizEcShort').value.trim(),
      roster_type: isConsultant ? 'consultant' : 'shift',
    };
    if (isConsultant) {
      const dataY = parseInt($('wizDataY').value) || 188;
      profile.data_start_y = dataY;
      const positions = [0, ...wizState.divPositions, wizState.pageW];
      const pdf_columns = {};
      COL_NAMES.forEach((name, i) => {
        pdf_columns[name] = {
          x_min: Math.round(positions[i] / wizState.canvasScale),
          x_max: Math.round(positions[i + 1] / wizState.canvasScale),
        };
      });
      profile.pdf_columns = pdf_columns;
      const t = id => $('wr_' + id).value.trim() || null;
      const pair = (a, b) => (t(a) && t(b)) ? [t(a), t(b)] : null;
      profile.time_rules = {
        weekday_slot1_oncall:    { normal: pair('ws1c_nf','ws1c_nt'),  ot1: pair('ws1c_o1f','ws1c_o1t'),  ot2: pair('ws1c_o2f','ws1c_o2t')  },
        weekday_slot1_no_oncall: { normal: pair('ws1n_nf','ws1n_nt'),  ot1: pair('ws1n_o1f','ws1n_o1t'),  ot2: pair('ws1n_o2f','ws1n_o2t')  },
        weekday_slot2_3:         { normal: pair('ws23_nf','ws23_nt'),  ot1: pair('ws23_o1f','ws23_o1t'),  ot2: pair('ws23_o2f','ws23_o2t')  },
        weekday_call_only:       { normal: pair('wco_nf','wco_nt'),    ot1: pair('wco_o1f','wco_o1t'),    ot2: pair('wco_o2f','wco_o2t')    },
        weekend_ph:              { normal: pair('wph_nf','wph_nt'),    ot1: pair('wph_o1f','wph_o1t'),    ot2: pair('wph_o2f','wph_o2t')    },
      };
    }
    $('wizJsonPreview').textContent = JSON.stringify(profile, null, 2);
    return { ecName: $('wizEcName').value.trim(), profile };
  }

  // ── Submit ───────────────────────────────────────────────────────────────
  $('wizSubmitBtn').addEventListener('click', async () => {
    const { ecName, profile } = wizBuildJson();
    const msg = $('wizSubmitMsg');
    $('wizSubmitBtn').disabled = true;
    $('wizSubmitBtn').textContent = 'Submitting…';
    msg.style.display = '';
    msg.style.color = 'var(--text-muted)';
    msg.textContent = 'Submitting to Supabase…';
    // Activate locally FIRST. The database only lets anon read rows whose
    // status is 'approved', so a freshly submitted profile is invisible to the
    // person who just built it — without this they would finish the wizard and
    // still have no working EC.
    const fullProfile = { ...profile, ec_name: ecName, ec_short: profile.ec_short || ecName };
    try {
      applyProfile(fullProfile);
      showEcSelected(ecName);
    } catch (e) {
      console.warn('[Wizard] local activation failed:', e);
    }

    try {
      const res = await fetch(`${SUPA_URL}/rest/v1/ec_profiles`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPA_ANON,
          'Authorization': 'Bearer ' + SUPA_ANON,
          'Prefer': 'return=minimal',
        },
        body: JSON.stringify({
          ec_name: ecName,
          ec_short: profile.ec_short,
          status: 'pending',
          profile: profile,
        }),
      });
      if (res.ok || res.status === 201) {
        msg.style.color = 'var(--success)';
        msg.textContent = '✓ Your department profile is active on this device now. It has also been sent for approval so other staff can pick it from the list.';
        $('wizSubmitBtn').style.display = 'none';
        $('wizNextBtn').style.display = 'none';
      } else {
        const txt = await res.text();
        throw new Error(`Server responded ${res.status}: ${txt}`);
      }
    } catch(err) {
      // Sharing failed, but the profile is already live locally.
      msg.style.color = 'var(--warn)';
      msg.textContent = '✓ Your department profile is active on this device. Could not send it for approval (' +
        err.message + ') — it will stay on this device only.';
      $('wizSubmitBtn').style.display = 'none';
      $('wizNextBtn').style.display = 'none';
    }
    $('wizSubmitBtn').disabled = false;
    $('wizSubmitBtn').textContent = 'Submit for Approval';
  });

  // Run initialisation — called separately below
});

// Initialise EC selector independently so a crash above can't prevent it
console.log('[EC] boot: readyState =', document.readyState);
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => { console.log('[EC] DOMContentLoaded fired'); initEcSelector(); });
} else {
  console.log('[EC] DOM already ready, calling directly');
  initEcSelector();
}


