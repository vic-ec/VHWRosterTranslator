window.addEventListener('error', (event) => { const el = document.getElementById('bootStatus'); if (el) el.textContent = `Runtime error: ${event.message}`; });
window.addEventListener('unhandledrejection', (event) => { const el = document.getElementById('bootStatus'); if (el) el.textContent = `Promise error: ${event.reason?.message || event.reason}`; });
import { DESIGNATIONS, MONTHS, SHIFT_CATEGORIES, SHIFT_DEFINITIONS, SUPERVISORS } from './constants.js';
import { state, resetState } from './state.js';
import { parseExcelRoster } from './parsers/excelParser.js';
import { parsePdfRoster } from './parsers/pdfParser.js';
import { normalizeParsedEntries, filterRoster, rosterContainsLeave } from './roster/normalizer.js';
import { downloadBlob, escapeHtml, getDayNameFromIso, hoursBetween } from './utils.js';
import { generateDutyRosterWorkbook } from './export/dutyRoster.js';
import { generateAnnexureCDocx } from './export/annexureC.js';
import { generateZ1aDocx } from './export/z1a.js';

const $ = (id) => document.getElementById(id);
const setBoot = (msg, isError = false) => {
  const el = $('bootStatus');
  if (!el) return;
  el.textContent = msg;
  el.classList.toggle('message', isError);
  el.classList.toggle('error', isError);
};

try {
  const els = {
    bootStatus: $('bootStatus'),
    dropZone: $('dropZone'),
    dropZoneInner: document.querySelector('#dropZone .drop-zone-inner'),
    rosterFiles: $('rosterFiles'),
    fileStatus: $('fileStatus'),
    extractBtn: $('extractBtn'),
    clearFilesBtn: $('clearFilesBtn'),
    uploadedFilesPanel: $('uploadedFilesPanel'),
    uploadedFilesList: $('uploadedFilesList'),
    parseMessages: $('parseMessages'),
    doctorSelect: $('doctorSelect'),
    timesheetName: $('timesheetName'),
    clearDoctorBtn: $('clearDoctorBtn'),
    monthSelect: $('monthSelect'),
    yearSelect: $('yearSelect'),
    previewBtn: $('previewBtn'),
    previewEmptyState: $('previewEmptyState'),
    previewPanel: $('previewPanel'),
    previewTitle: $('previewTitle'),
    previewSummary: $('previewSummary'),
    addShiftBtn: $('addShiftBtn'),
    sortShiftsBtn: $('sortShiftsBtn'),
    shiftTableBody: $('shiftTableBody'),
    firstName: $('firstName'),
    surname: $('surname'),
    persalNumber: $('persalNumber'),
    designationSelect: $('designationSelect'),
    designationOther: $('designationOther'),
    supervisorSelect: $('supervisorSelect'),
    supervisorOther: $('supervisorOther'),
    signatureDate: $('signatureDate'),
    leaveDetailsCard: $('leaveDetailsCard'),
    shiftWorkerSelect: $('shiftWorkerSelect'),
    casualEmployeeSelect: $('casualEmployeeSelect'),
    leaveAddress: $('leaveAddress'),
    validationPanel: $('validationPanel'),
    downloadDutyRosterBtn: $('downloadDutyRosterBtn'),
    downloadAnnexureBtn: $('downloadAnnexureBtn'),
    downloadLeaveBtn: $('downloadLeaveBtn'),
    resetAppBtn: $('resetAppBtn'),
    shiftRowTemplate: $('shiftRowTemplate')
  };

  function showMessage(text, type = 'info') {
    const div = document.createElement('div');
    div.className = `message ${type}`;
    div.textContent = text;
    els.parseMessages.appendChild(div);
  }

  function clearMessages() {
    els.parseMessages.innerHTML = '';
  }

  function filesFromEventList(fileList) {
    return Array.from(fileList || []).filter((file) => /\.(pdf|xlsx|xls|csv)$/i.test(file.name));
  }

  function populateStaticControls() {
    els.designationSelect.innerHTML = '<option value="">— select —</option>';
    DESIGNATIONS.forEach((item) => {
      const option = document.createElement('option');
      option.value = item;
      option.textContent = item;
      els.designationSelect.appendChild(option);
    });
    els.supervisorSelect.innerHTML = '<option value="">— select —</option>';
    SUPERVISORS.forEach((item) => {
      const option = document.createElement('option');
      option.value = item;
      option.textContent = item;
      els.supervisorSelect.appendChild(option);
    });
    els.monthSelect.innerHTML = '<option value="">— select —</option>';
    MONTHS.forEach((month) => {
      const option = document.createElement('option');
      option.value = month;
      option.textContent = month;
      els.monthSelect.appendChild(option);
    });
    els.yearSelect.innerHTML = '<option value="">— select —</option>';
    [2024, 2025, 2026, 2027].forEach((year) => {
      const option = document.createElement('option');
      option.value = year;
      option.textContent = year;
      els.yearSelect.appendChild(option);
    });
  }

  function renderFiles() {
    els.uploadedFilesList.innerHTML = '';
    els.uploadedFilesPanel.classList.toggle('hidden', !state.files.length);
    state.files.forEach((file) => {
      const li = document.createElement('li');
      li.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
      els.uploadedFilesList.appendChild(li);
    });
    els.fileStatus.textContent = state.files.length ? `${state.files.length} file(s) selected` : 'Waiting for roster file…';
  }

  function updateDoctorOptions() {
    els.doctorSelect.innerHTML = '<option value="">— select —</option>';
    [...state.rostersByDoctor.keys()].sort().forEach((doctor) => {
      const option = document.createElement('option');
      option.value = doctor;
      option.textContent = doctor;
      els.doctorSelect.appendChild(option);
    });
  }

  function syncConditionalFields() {
    els.designationOther.classList.toggle('hidden', els.designationSelect.value !== 'Other…');
    els.supervisorOther.classList.toggle('hidden', els.supervisorSelect.value !== 'Other…');
    const hasLeave = rosterContainsLeave(state.currentShifts);
    els.leaveDetailsCard.classList.toggle('hidden', !hasLeave);
    els.downloadLeaveBtn.classList.toggle('hidden', !hasLeave);
  }

  function renderShiftRows() {
    els.shiftTableBody.innerHTML = '';
    state.currentShifts.forEach((shift) => {
      const fragment = els.shiftRowTemplate.content.cloneNode(true);
      const row = fragment.querySelector('tr');
      row.dataset.id = shift.id;
      const fields = Object.fromEntries([...row.querySelectorAll('[data-field]')].map((el) => [el.dataset.field, el]));
      fields.date.value = shift.dateIso;
      fields.dayName.value = shift.dayName;
      fields.startTime.value = shift.startTime || '';
      fields.endTime.value = shift.endTime || '';
      fields.hours.value = shift.hours ?? 0;
      fields.comment.value = shift.comment || '';
      SHIFT_DEFINITIONS.forEach((def) => {
        const opt = document.createElement('option');
        opt.value = def.key;
        opt.textContent = def.label;
        fields.shiftType.appendChild(opt);
      });
      fields.shiftType.value = shift.shiftType;
      SHIFT_CATEGORIES.forEach((cat) => {
        const opt = document.createElement('option');
        opt.value = cat.key;
        opt.textContent = cat.label;
        fields.category.appendChild(opt);
      });
      fields.category.value = shift.category;
      row.addEventListener('input', (event) => {
        const target = event.target;
        const item = state.currentShifts.find((s) => s.id === row.dataset.id);
        if (!item) return;
        const field = target.dataset.field;
        if (!field) return;
        item[field] = target.value;
        if (field === 'date') item.dayName = fields.dayName.value = getDayNameFromIso(target.value);
        if (field === 'startTime' || field === 'endTime') item.hours = fields.hours.value = hoursBetween(fields.startTime.value, fields.endTime.value);
        validateForDownloads();
        syncConditionalFields();
      });
      row.addEventListener('click', (event) => {
        const action = event.target?.dataset?.action;
        if (!action) return;
        const index = state.currentShifts.findIndex((s) => s.id === row.dataset.id);
        if (index < 0) return;
        if (action === 'delete') state.currentShifts.splice(index, 1);
        if (action === 'duplicate') state.currentShifts.splice(index + 1, 0, { ...structuredClone(state.currentShifts[index]), id: crypto.randomUUID() });
        renderShiftRows();
        validateForDownloads();
        syncConditionalFields();
      });
      els.shiftTableBody.appendChild(fragment);
    });
  }

  function renderPreview() {
    if (!state.selectedDoctor || !state.selectedMonth || !state.selectedYear) {
      els.previewEmptyState.classList.remove('hidden');
      els.previewPanel.classList.add('hidden');
      return;
    }
    els.previewEmptyState.classList.add('hidden');
    els.previewPanel.classList.remove('hidden');
    els.previewTitle.textContent = `${state.selectedDoctor} — ${state.selectedMonth} ${state.selectedYear}`;
    els.previewSummary.textContent = `${state.currentShifts.length} shift item(s)`;
    renderShiftRows();
  }

  function validateForDownloads() {
    const issues = [];
    if (!state.currentShifts.length) issues.push('No shifts selected for preview/export.');
    if (!state.employee.firstName) issues.push('First name is required.');
    if (!state.employee.surname) issues.push('Surname is required.');
    if (!state.employee.persalNumber) issues.push('PERSAL number is required.');
    if (!state.employee.designation) issues.push('Designation is required.');
    if (state.employee.designation === 'Other…' && !state.employee.designationOther) issues.push('Other designation is required.');
    if (!state.employee.supervisor) issues.push('Supervisor is required.');
    if (state.employee.supervisor === 'Other…' && !state.employee.supervisorOther) issues.push('Other supervisor is required.');
    if (!state.employee.signatureDate) issues.push('Date of signature is required.');
    if (rosterContainsLeave(state.currentShifts)) {
      if (!state.employee.shiftWorker) issues.push('Shift worker selection is required for leave form.');
      if (!state.employee.casualEmployee) issues.push('Casual employee selection is required for leave form.');
      if (!state.employee.leaveAddress) issues.push('Address during leave is required.');
    }
    els.validationPanel.classList.toggle('hidden', issues.length === 0);
    els.validationPanel.innerHTML = issues.length ? `<strong>Required to enable downloads:</strong><ul>${issues.map(i => `<li>${escapeHtml(i)}</li>`).join('')}</ul>` : '';
    const ok = issues.length === 0;
    els.downloadDutyRosterBtn.disabled = !ok;
    els.downloadAnnexureBtn.disabled = !ok;
    els.downloadLeaveBtn.disabled = !ok || !rosterContainsLeave(state.currentShifts);
  }

  async function handleExtract() {
    clearMessages();
    if (!state.files.length) {
      showMessage('Select at least one roster file first.', 'error');
      return;
    }
    state.parseResults = [];
    for (const file of state.files) {
      try {
        if (file.name.toLowerCase().endsWith('.pdf')) {
          showMessage(`PDF parsing is not implemented yet for ${file.name}. Please use Excel roster files for now.`, 'error');
        } else {
          state.parseResults.push(await parseExcelRoster(file));
          showMessage(`Parsed ${file.name}`, 'success');
        }
      } catch (error) {
        showMessage(`Failed to parse ${file.name}: ${error.message}`, 'error');
      }
    }
    state.rostersByDoctor = normalizeParsedEntries(state.parseResults);
    updateDoctorOptions();
  }

  async function downloadDutyRoster() {
    const response = await fetch('./assets/templates/Duty-Rosters-2026.xlsx');
    if (!response.ok) throw new Error('Could not load duty roster template from assets/templates');
    const blob = await response.blob();
    const templateFile = new File([blob], 'Duty-Rosters-2026.xlsx', { type: blob.type });
    const outBlob = await generateDutyRosterWorkbook({ templateFile, employee: state.employee, shifts: state.currentShifts, month: state.selectedMonth, year: state.selectedYear });
    downloadBlob(outBlob, `Duty-Roster-${state.selectedMonth}-${state.selectedYear}-${state.employee.surname || 'employee'}.xlsx`);
  }

  async function downloadAnnexure() {
    const blob = await generateAnnexureCDocx({ employee: state.employee, shifts: state.currentShifts, month: state.selectedMonth, year: state.selectedYear });
    downloadBlob(blob, `Annexure-C-${state.selectedMonth}-${state.selectedYear}-${state.employee.surname || 'employee'}.docx`);
  }

  async function downloadLeave() {
    const blob = await generateZ1aDocx({ employee: state.employee, shifts: state.currentShifts });
    downloadBlob(blob, `Z1a-Leave-${state.selectedMonth}-${state.selectedYear}-${state.employee.surname || 'employee'}.docx`);
  }

  function wireEmployeeFields() {
    [
      ['firstName', els.firstName], ['surname', els.surname], ['persalNumber', els.persalNumber],
      ['designation', els.designationSelect], ['designationOther', els.designationOther],
      ['supervisor', els.supervisorSelect], ['supervisorOther', els.supervisorOther],
      ['signatureDate', els.signatureDate], ['timesheetName', els.timesheetName],
      ['shiftWorker', els.shiftWorkerSelect], ['casualEmployee', els.casualEmployeeSelect], ['leaveAddress', els.leaveAddress]
    ].forEach(([key, element]) => {
      ['input', 'change'].forEach((evt) => element.addEventListener(evt, () => {
        state.employee[key] = element.value;
        syncConditionalFields();
        validateForDownloads();
      }));
    });
  }

  function wireEvents() {
    const openPicker = () => {
      setBoot('Opening file picker…');
      els.rosterFiles.click();
    };
    els.dropZone.addEventListener('click', (e) => { e.preventDefault(); openPicker(); });
    els.dropZoneInner?.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); openPicker(); });
    els.dropZone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        openPicker();
      }
    });
    ['dragenter', 'dragover'].forEach((evt) => els.dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      els.dropZone.classList.add('dragover');
      setBoot('Drop files to upload');
    }));
    ['dragleave', 'drop'].forEach((evt) => els.dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      els.dropZone.classList.remove('dragover');
    }));
    els.dropZone.addEventListener('drop', (e) => {
      state.files = filesFromEventList(e.dataTransfer?.files);
      renderFiles();
      setBoot(state.files.length ? `Loaded ${state.files.length} file(s). Click Extract data.` : 'No supported files dropped.');
    });
    els.rosterFiles.addEventListener('change', () => {
      state.files = filesFromEventList(els.rosterFiles.files);
      renderFiles();
      setBoot(state.files.length ? `Loaded ${state.files.length} file(s). Click Extract data.` : 'No supported files selected.');
    });
    els.clearFilesBtn.addEventListener('click', () => { state.files = []; els.rosterFiles.value = ''; renderFiles(); clearMessages(); setBoot('Files cleared.'); });
    els.extractBtn.addEventListener('click', handleExtract);
    els.doctorSelect.addEventListener('change', () => { state.selectedDoctor = els.doctorSelect.value; els.timesheetName.value = state.employee.timesheetName = els.doctorSelect.value; });
    els.clearDoctorBtn.addEventListener('click', () => { state.selectedDoctor = ''; els.doctorSelect.value = ''; state.currentShifts = []; renderPreview(); validateForDownloads(); });
    els.monthSelect.addEventListener('change', () => state.selectedMonth = els.monthSelect.value);
    els.yearSelect.addEventListener('change', () => state.selectedYear = els.yearSelect.value);
    els.previewBtn.addEventListener('click', () => {
      state.selectedDoctor = els.doctorSelect.value;
      state.selectedMonth = els.monthSelect.value;
      state.selectedYear = els.yearSelect.value;
      state.currentShifts = filterRoster(state.rostersByDoctor, state.selectedDoctor, state.selectedMonth, state.selectedYear);
      renderPreview();
      syncConditionalFields();
      validateForDownloads();
    });
    els.addShiftBtn.addEventListener('click', () => {
      const monthIndex = MONTHS.indexOf(state.selectedMonth) + 1 || 1;
      const baseDate = `${state.selectedYear || 2026}-${String(monthIndex).padStart(2, '0')}-01`;
      state.currentShifts.push({ id: crypto.randomUUID(), doctorName: state.selectedDoctor, month: state.selectedMonth, year: Number(state.selectedYear || 2026), dateIso: baseDate, dayName: getDayNameFromIso(baseDate), shiftType: 'OTHER', shiftLabel: 'Other', startTime: '', endTime: '', hours: 0, category: 'notOvertime', comment: '' });
      renderShiftRows();
      validateForDownloads();
    });
    els.sortShiftsBtn.addEventListener('click', () => { state.currentShifts.sort((a, b) => a.dateIso.localeCompare(b.dateIso) || (a.startTime || '').localeCompare(b.startTime || '')); renderShiftRows(); });
    els.resetAppBtn.addEventListener('click', () => {
      resetState();
      els.rosterFiles.value = '';
      document.querySelectorAll('input, select, textarea').forEach((el) => {
        if (el.type === 'file') return;
        if (el.tagName === 'SELECT') el.selectedIndex = 0;
        else el.value = '';
      });
      renderFiles();
      clearMessages();
      updateDoctorOptions();
      renderPreview();
      syncConditionalFields();
      validateForDownloads();
      setBoot('App reset.');
    });
    els.downloadDutyRosterBtn.addEventListener('click', downloadDutyRoster);
    els.downloadAnnexureBtn.addEventListener('click', downloadAnnexure);
    els.downloadLeaveBtn.addEventListener('click', downloadLeave);
  }

  populateStaticControls();
  renderFiles();
  updateDoctorOptions();
  wireEmployeeFields();
  wireEvents();
  syncConditionalFields();
  validateForDownloads();
  setBoot('App loaded. Click the upload area to choose files.');
} catch (error) {
  console.error(error);
  setBoot(`App failed to initialize: ${error.message}`, true);
}
