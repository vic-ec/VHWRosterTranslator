// ═══════════════════════════════════════════════════════════════
// consultant.js — Consultant roster file handling and shift overlay.
//
//   updateConsultantZoneVisibility()
//   renderConsultantList()
//   addConsultantFiles(files)
//   removeConsultantFile(index)
//   setConsultantFile(file)
//   parseAndStoreConsultantRoster()
//   overlayConsultantShifts(doctorName, targetMonth, targetYear)
//
// Depends on: config.js, parser-consultant.js, ec-profiles.js
// ═══════════════════════════════════════════════════════════════

// CONSULTANT ROSTER FILE HANDLING
// ═══════════════════════════════════════════════════════════════

// Show/hide the consultant upload zone based on the active profile type.
// Called whenever a profile is applied.
function updateConsultantZoneVisibility() {
  const wrap = $('consultantZoneWrap');
  if (!wrap) return;
  const isConsultant = activeProfile && activeProfile.roster_type === 'consultant';
  wrap.style.display = isConsultant ? 'block' : 'none';
  // The leave button's grid mirrors this one, and only stays half-width while
  // there is a second cell to hold the track open.
  const spacer = $('altRouteSpacer');
  if (spacer) spacer.style.display = isConsultant ? 'block' : 'none';
  // Hiding the zone has to take the buttons back with it.
  if (typeof syncActionsSide === 'function') syncActionsSide();
}


function renderConsultantList() {
  if (typeof syncActionsSide === 'function') syncActionsSide();
  const list = $('consultantBadgeList');
  if (!list) return;
  const files = state.consultantFiles || [];
  if (!files.length) {
    list.style.display = 'none';
    return;
  }
  list.style.display = 'block';
  list.innerHTML = files.map(f => `
    <div class="roster-item">
      <span class="tag tag-neutral">PDF</span>
      <span class="ri-name">${f.name}</span>
      <span class="ri-days">queued</span>
      <button class="ri-remove" data-name="${f.name}">&times;</button>
    </div>`).join('');
  list.querySelectorAll('.ri-remove').forEach(btn =>
    btn.addEventListener('click', () => removeConsultantFile(btn.dataset.name))
  );
}

function addConsultantFiles(files) {
  if (!state.consultantFiles) state.consultantFiles = [];
  for (const f of files) {
    if (!state.consultantFiles.some(x => x.name === f.name)) state.consultantFiles.push(f);
  }
  state.consultantFile = state.consultantFiles[0] || null; // backward compat
  state.consultantData = null;
  state.consultantFileCount = 0;
  state.consultantFileErrors = [];
  renderConsultantList();
  if (state.consultantFiles.length) {
    $('parseBtn').disabled = false;
    $('clearBtn').style.display = '';
  }
}

function removeConsultantFile(name) {
  state.consultantFiles = (state.consultantFiles || []).filter(f => f.name !== name);
  state.consultantFile = state.consultantFiles[0] || null;
  state.consultantData = null;
  state.consultantFileCount = 0;
  state.consultantFileErrors = [];
  renderConsultantList();
  if (!state.pendingFiles.length && !state.consultantFiles?.length) {
    $('parseBtn').disabled = true;
    $('clearBtn').style.display = 'none';
  }
}

function setConsultantFile(file) {
  // Legacy single-file reset (called from clear button)
  state.consultantFiles = [];
  state.consultantFile = null;
  state.consultantData = null;
  state.consultantFileCount = 0;
  state.consultantFileErrors = [];
  renderConsultantList();
  if (!state.pendingFiles.length) {
    if ($('parseBtn')) $('parseBtn').disabled = true;
    if ($('clearBtn')) $('clearBtn').style.display = 'none';
  }
}

// Parse one or more consultant PDFs and merge results
async function parseAndStoreConsultantRoster() {
  const filesToParse = (state.consultantFiles && state.consultantFiles.length)
    ? state.consultantFiles
    : (state.consultantFile ? [state.consultantFile] : []);
  if (!filesToParse.length || !activeProfile) return;

  const MONTHS_IDX = {january:0,february:1,march:2,april:3,may:4,june:5,
    july:6,august:7,september:8,october:9,november:10,december:11};

  const allDays = [], allDoctors = new Set();
  let lastDetectedMonth = new Date().getMonth();
  let lastDetectedYear  = new Date().getFullYear();
  // How many files actually contributed, and which did not. The status line
  // used to say "across 1 file(s)" however many were queued — it counted
  // state.consultantFile, the single-file field, and the consultant-only
  // branch had the 1 written into the string. A file that throws in here is
  // dropped with only a console message, so the count has to be of files that
  // parsed, and the ones that did not have to be named.
  let okFiles = 0;
  const failedFiles = [];

  for (const cFile of filesToParse) {
    try {
      const buf = await readFile(cFile);
      const result = await parseConsultantRosterPDF(buf, activeProfile);

      // The sheet's own title line wins, the file name is the fallback, and
      // today's date is the last resort. It used to be the file name alone —
      // so "Consultant Duty Roster 2026 Shared.pdf", an April roster with no
      // month in its name, was filed under whatever month it happened to be
      // opened in, and every one of its days was then filtered out of the
      // month the app said it was showing.
      const fnYearMatch  = cFile.name.match(/20\d{2}/);
      const fnMonthMatch = cFile.name.match(
        /January|February|March|April|May|June|July|August|September|October|November|December/i
      );
      const detectedMonth = result.month != null ? result.month
        : (fnMonthMatch ? MONTHS_IDX[fnMonthMatch[0].toLowerCase()]
                        : new Date().getMonth());
      const detectedName = result.month != null ? result.monthName
        : (fnMonthMatch ? fnMonthMatch[0] : '');
      const year = result.year != null ? result.year
        : (fnYearMatch ? parseInt(fnYearMatch[0]) : new Date().getFullYear());

      for (const d of result.days) {
        d.month = detectedMonth;
        d.monthName = detectedName;
      }
      allDays.push(...result.days);
      result.doctors.forEach(d => allDoctors.add(d));
      lastDetectedMonth = detectedMonth;
      lastDetectedYear  = year;
      okFiles++;
    } catch(err) {
      console.error('[Consultant] Parse error for', cFile.name, err);
      failedFiles.push(cFile.name);
    }
  }

  state.consultantData = { days: allDays, doctors: allDoctors };
  state.consultantFileCount  = okFiles;
  state.consultantFileErrors = failedFiles;

  // Merge into main rosterData
  if (state.rosterData && state.parsedFiles.length) {
    allDoctors.forEach(d => state.rosterData.doctors.add(d));
    buildDoctorGrid(state.rosterData.doctors);
  } else {
    // Consultant-only — build rosterData from consultant data
    state.rosterData = { days: allDays, doctors: allDoctors };
    state.availableMonths = new Set(allDays.map(d => d.month));
    buildDoctorGrid(allDoctors);
    $('monthSelect').value = lastDetectedMonth;
    $('yearInput').value = lastDetectedYear;
    step2.style.display = ''; unlock(step2);
    rebuildMonthDropdown();
    checkReady();
  }
}

// Override buildPreview to inject consultant shifts when available
// Consultant shift overlay — called from within buildPreview
function overlayConsultantShifts(doctorName, targetMonth, targetYear) {
  if (!state.consultantData || !activeProfile || activeProfile.roster_type !== 'consultant') return 0;
  const consultantShifts = getConsultantShifts(
    state.consultantData, doctorName, targetMonth, activeProfile, targetYear
  );
  let added = 0;
  for (const [dateStr, shift] of Object.entries(consultantShifts)) {
    const d = parseInt(dateStr);
    if (!state.editedShifts[d]) {
      state.editedShifts[d] = shift;
      state.originalShifts[d] = { ...shift };
      added++;
    }
  }
  return added;
}

document.addEventListener('DOMContentLoaded', () => {

  // Consultant zone visibility is now handled directly inside applyProfile.
  // Run once now in case profile was applied before this DOMContentLoaded fired.
  updateConsultantZoneVisibility();

  // Consultant file input
  const cFile = $('consultantFile');
  if (cFile) {
    cFile.addEventListener('change', e => {
      if (e.target.files.length) addConsultantFiles(Array.from(e.target.files));
      cFile.value = '';
    });
  }

  // Drag and drop on consultant zone
  const cZone = $('consultantZone');
  if (cZone) {
    cZone.addEventListener('dragover', e => { e.preventDefault(); cZone.classList.add('drag-over'); });
    cZone.addEventListener('dragleave', () => cZone.classList.remove('drag-over'));
    cZone.addEventListener('drop', e => {
      e.preventDefault();
      cZone.classList.remove('drag-over');
      if (e.dataTransfer.files.length) addConsultantFiles(Array.from(e.dataTransfer.files));
    });
  }

});
