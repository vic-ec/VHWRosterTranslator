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
}


function renderConsultantList() {
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

  for (const cFile of filesToParse) {
    try {
      const buf = await readFile(cFile);
      const result = await parseConsultantRosterPDF(buf, activeProfile);

      const fnYearMatch = cFile.name.match(/20\d{2}/);
      const year = fnYearMatch ? parseInt(fnYearMatch[0]) : new Date().getFullYear();
      const monthMatch = cFile.name.match(
        /January|February|March|April|May|June|July|August|September|October|November|December/i
      );
      const detectedMonth = monthMatch
        ? MONTHS_IDX[monthMatch[0].toLowerCase()]
        : new Date().getMonth();

      for (const d of result.days) {
        d.month = detectedMonth;
        d.monthName = monthMatch ? monthMatch[0] : '';
      }
      allDays.push(...result.days);
      result.doctors.forEach(d => allDoctors.add(d));
      lastDetectedMonth = detectedMonth;
      lastDetectedYear  = year;
    } catch(err) {
      console.error('[Consultant] Parse error for', cFile.name, err);
    }
  }

  state.consultantData = { days: allDays, doctors: allDoctors };

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
