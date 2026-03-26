(function () {
  'use strict';

  // ─── Boot status helpers ───────────────────────────────────────────────────
  const boot = document.getElementById('bootStatus');
  const setBoot = (msg, isError = false) => {
    if (!boot) return;
    boot.textContent = msg;
    boot.className = isError ? 'panel message error' : 'panel';
  };

  // ─── Global error traps ────────────────────────────────────────────────────
  window.addEventListener('error', (e) => setBoot(`Runtime error: ${e.message}`, true));
  window.addEventListener('unhandledrejection', (e) => setBoot(`Promise error: ${e.reason?.message || e.reason}`, true));

  // ─── Constants ─────────────────────────────────────────────────────────────
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const DESIGNATIONS = [
    'Intern','Community Service Medical Officer','Medical Officer Grade 1','Medical Officer Grade 2',
    'Medical Officer Grade 3','Registrar','Medical Specialist Grade 1','Medical Specialist Grade 2',
    'Medical Specialist Grade 3','Other…'
  ];
  const SUPERVISORS = ['Philip Cloete', 'Sebastian De Haan', 'Paul Xafis', 'Other…'];
  const SHIFT_DEFINITIONS = [
    { key: 'WD_0800_1800', label: 'WD shift 08h00-18h00',  category: 'onCall1',     startTime: '08:00', endTime: '18:00', hours: 10 },
    { key: 'WD_1200_2200', label: 'WD shift 12h00-22h00',  category: 'onCall1',     startTime: '12:00', endTime: '22:00', hours: 10 },
    { key: 'WD_1500_2300', label: 'WD shift 15h00-23h00',  category: 'onCall1',     startTime: '15:00', endTime: '23:00', hours:  8 },
    { key: 'WD_2200_1000', label: 'WD shift 22h00-10h00',  category: 'onCall1',     startTime: '22:00', endTime: '10:00', hours: 12 },
    { key: 'WE_0800_2000', label: 'WE shift 08h00-20h00',  category: 'onCall1',     startTime: '08:00', endTime: '20:00', hours: 12 },
    { key: 'WE_1300_2300', label: 'WE shift 13h00-23h00',  category: 'onCall1',     startTime: '13:00', endTime: '23:00', hours: 10 },
    { key: 'WE_2000_1000', label: 'WE shift 20h00-10h00',  category: 'onCall1',     startTime: '20:00', endTime: '10:00', hours: 14 },
    { key: 'ANNUAL_LEAVE',        label: 'Annual leave',                  category: 'leave',       startTime: '', endTime: '', hours: 0 },
    { key: 'SICK_LEAVE',          label: 'Sick leave',                    category: 'leave',       startTime: '', endTime: '', hours: 0 },
    { key: 'FAMILY_RESPONSIBILITY',label:'Family responsibility leave',   category: 'leave',       startTime: '', endTime: '', hours: 0 },
    { key: 'SPECIAL_LEAVE',       label: 'Special leave',                 category: 'leave',       startTime: '', endTime: '', hours: 0 },
    { key: 'PRE_NATAL',           label: 'Pre-natal leave',               category: 'leave',       startTime: '', endTime: '', hours: 0 },
    { key: 'UNPAID_LEAVE',        label: 'Unpaid leave',                  category: 'leave',       startTime: '', endTime: '', hours: 0 },
    { key: 'WORKSHOP',            label: 'Workshop',                      category: 'notOvertime', startTime: '', endTime: '', hours: 0 },
    { key: 'COURSE',              label: 'Course',                        category: 'notOvertime', startTime: '', endTime: '', hours: 0 },
    { key: 'OFF_DAY',             label: 'Off-day',                       category: 'notOvertime', startTime: '', endTime: '', hours: 0 },
    { key: 'OTHER',               label: 'Other',                         category: 'notOvertime', startTime: '', endTime: '', hours: 0 }
  ];
  const SHIFT_DEF_MAP = new Map(SHIFT_DEFINITIONS.map(d => [d.key, d]));
  const SHIFT_CATEGORIES = [
    { key: 'normal',      label: 'Normal working hours'  },
    { key: 'onCall1',     label: '1st on call - on site' },
    { key: 'onCall2',     label: '2nd on call - off site'},
    { key: 'notOvertime', label: 'Not overtime'          },
    { key: 'leave',       label: 'Leave'                 }
  ];
  const HOLIDAYS = {
    '2026-01-01': 'Public Holiday',      '2026-03-21': 'Human Rights Day',
    '2026-04-03': 'Good Friday',          '2026-04-06': 'Family Day',
    '2026-05-01': "Workers' Day",         '2026-06-16': 'Youth Day',
    '2026-08-09': "National Women's Day", '2026-09-23': 'Heritage Day',
    '2026-12-16': 'Day of Reconciliation','2026-12-25': 'Christmas Day',
    '2026-12-26': 'Day of Goodwill'
  };
  const INSTITUTION     = 'VICTORIA HOSPITAL';
  const DEPARTMENT      = 'HEALTH';
  const COMPONENT       = 'Emergency Centre';
  const ANNEXURE_GROUP  = '3.2';
  const TEMPLATE_PATH   = './assets/templates/Duty-Rosters-2026.xlsx';

  // ─── State ──────────────────────────────────────────────────────────────────
  const state = {
    files: [],
    parseResults: [],
    rostersByDoctor: new Map(),
    selectedDoctor: '',
    selectedMonth: '',
    selectedYear: '',
    currentShifts: [],
    employee: {
      firstName: '', surname: '', persalNumber: '', designation: '',
      designationOther: '', supervisor: '', supervisorOther: '',
      signatureDate: '', timesheetName: '', shiftWorker: '',
      casualEmployee: '', leaveAddress: ''
    }
  };

  // ─── Utilities ──────────────────────────────────────────────────────────────
  const $ = (id) => document.getElementById(id);
  const pad = (n) => String(n).padStart(2, '0');
  const getDayName = (iso) => new Date(`${iso}T00:00:00`).toLocaleDateString('en-ZA', { weekday: 'long' });
  const hoursBetween = (s, e) => {
    if (!s || !e) return 0;
    const [sh, sm] = s.split(':').map(Number);
    const [eh, em] = e.split(':').map(Number);
    let start = sh * 60 + sm, end = eh * 60 + em;
    if (end < start) end += 1440;
    return Number(((end - start) / 60).toFixed(2));
  };
  const escapeHtml = (v = '') => String(v)
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
  const xmlEscape = (v = '') => String(v)
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;');
  const fullName = (emp) => [emp.firstName, emp.surname].filter(Boolean).join(' ').trim();
  const initials = (emp) => [emp.firstName, emp.surname].filter(Boolean).map(p => p.trim()[0].toUpperCase()).join('');
  const downloadBlob = (blob, name) => {
    const url = URL.createObjectURL(blob);
    const a = Object.assign(document.createElement('a'), { href: url, download: name });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  const monthToNumber = (name) => {
    const idx = MONTHS.findIndex(m => m.toLowerCase() === String(name).toLowerCase());
    return idx >= 0 ? idx + 1 : null;
  };

  // ─── DOM element refs ───────────────────────────────────────────────────────
  const els = {
    bootStatus:           $('bootStatus'),
    dropZone:             $('dropZone'),
    dropZoneInner:        document.querySelector('#dropZone .drop-zone-inner'),
    rosterFiles:          $('rosterFiles'),
    fileStatus:           $('fileStatus'),
    extractBtn:           $('extractBtn'),
    clearFilesBtn:        $('clearFilesBtn'),
    uploadedFilesPanel:   $('uploadedFilesPanel'),
    uploadedFilesList:    $('uploadedFilesList'),
    parseMessages:        $('parseMessages'),
    doctorSelect:         $('doctorSelect'),
    timesheetName:        $('timesheetName'),
    clearDoctorBtn:       $('clearDoctorBtn'),
    monthSelect:          $('monthSelect'),
    yearSelect:           $('yearSelect'),
    previewBtn:           $('previewBtn'),
    previewEmptyState:    $('previewEmptyState'),
    previewPanel:         $('previewPanel'),
    previewTitle:         $('previewTitle'),
    previewSummary:       $('previewSummary'),
    addShiftBtn:          $('addShiftBtn'),
    sortShiftsBtn:        $('sortShiftsBtn'),
    shiftTableBody:       $('shiftTableBody'),
    firstName:            $('firstName'),
    surname:              $('surname'),
    persalNumber:         $('persalNumber'),
    designationSelect:    $('designationSelect'),
    designationOther:     $('designationOther'),
    supervisorSelect:     $('supervisorSelect'),
    supervisorOther:      $('supervisorOther'),
    signatureDate:        $('signatureDate'),
    leaveDetailsCard:     $('leaveDetailsCard'),
    shiftWorkerSelect:    $('shiftWorkerSelect'),
    casualEmployeeSelect: $('casualEmployeeSelect'),
    leaveAddress:         $('leaveAddress'),
    validationPanel:      $('validationPanel'),
    downloadDutyRosterBtn:$('downloadDutyRosterBtn'),
    downloadAnnexureBtn:  $('downloadAnnexureBtn'),
    downloadLeaveBtn:     $('downloadLeaveBtn'),
    resetAppBtn:          $('resetAppBtn'),
    shiftRowTemplate:     $('shiftRowTemplate')
  };

  // ─── Message helpers ────────────────────────────────────────────────────────
  const showMessage = (text, type = 'info') => {
    const d = document.createElement('div');
    d.className = `message ${type}`;
    d.textContent = text;
    els.parseMessages.appendChild(d);
  };
  const clearMessages = () => { els.parseMessages.innerHTML = ''; };

  // ─── Excel time normaliser ──────────────────────────────────────────────────
  // ExcelJS sometimes returns time as a decimal fraction (e.g. 0.333 = 08:00)
  // or as a string like "08:00", "08h00", "8.00"
  function cleanTime(raw) {
    if (raw === null || raw === undefined || raw === '') return '';
    // Numeric fraction (Excel date serial for time-only values)
    if (typeof raw === 'number') {
      const totalMins = Math.round(raw * 1440);
      return `${pad(Math.floor(totalMins / 60) % 24)}:${pad(totalMins % 60)}`;
    }
    // Date object (ExcelJS can return these)
    if (raw instanceof Date) {
      return `${pad(raw.getUTCHours())}:${pad(raw.getUTCMinutes())}`;
    }
    const text = String(raw).trim().replace(/h/ig, ':').replace(/\./g, ':');
    const m = text.match(/(\d{1,2})[:](\d{2})/);
    if (!m) return '';
    return `${pad(Number(m[1]))}:${m[2]}`;
  }

  // ─── Excel parser ───────────────────────────────────────────────────────────
  function inferShiftType(comment, startTime, endTime, dayName) {
    const t = String(comment || '').toLowerCase();
    if (t.includes('annual leave'))      return 'ANNUAL_LEAVE';
    if (t.includes('sick'))              return 'SICK_LEAVE';
    if (t.includes('family'))            return 'FAMILY_RESPONSIBILITY';
    if (t.includes('special leave'))     return 'SPECIAL_LEAVE';
    if (t.includes('pre-natal'))         return 'PRE_NATAL';
    if (t.includes('unpaid'))            return 'UNPAID_LEAVE';
    if (t.includes('course'))            return 'COURSE';
    if (t.includes('workshop'))          return 'WORKSHOP';
    if (t.includes('off'))               return 'OFF_DAY';
    const weekend = ['Saturday', 'Sunday'].includes(dayName);
    const key = `${startTime}-${endTime}`;
    if (weekend) {
      if (key === '08:00-20:00') return 'WE_0800_2000';
      if (key === '13:00-23:00') return 'WE_1300_2300';
      if (key === '20:00-10:00') return 'WE_2000_1000';
    } else {
      if (key === '08:00-18:00') return 'WD_0800_1800';
      if (key === '12:00-22:00') return 'WD_1200_2200';
      if (key === '15:00-23:00') return 'WD_1500_2300';
      if (key === '22:00-10:00') return 'WD_2200_1000';
    }
    return 'OTHER';
  }

  async function parseExcelRoster(file) {
    if (typeof ExcelJS === 'undefined') throw new Error('ExcelJS library failed to load. Check CDN connectivity.');
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const entries = [];

    wb.worksheets.forEach((ws) => {
      const sheetMatch = (ws.name || '').match(/([A-Za-z]+)\s+(\d{4})/);
      if (!sheetMatch) return;
      const month = sheetMatch[1];
      const year  = Number(sheetMatch[2]);
      const monthNum = monthToNumber(month);
      if (!monthNum) return;

      // Try to detect doctor name from common header cells (rows 1–7, various columns)
      let detectedDoctor = '';
      const headerCells = ['C5','C4','C3','B5','B4','B3','C6','D5'];
      for (const addr of headerCells) {
        const v = String(ws.getCell(addr).text || ws.getCell(addr).value || '').trim();
        // Heuristic: looks like a name (≥2 words, mostly letters)
        if (v && v.split(/\s+/).length >= 2 && /^[A-Za-z\s.'-]+$/.test(v)) {
          detectedDoctor = v;
          break;
        }
      }
      // Fallback: strip year from sheet name
      if (!detectedDoctor) detectedDoctor = ws.name.replace(/\s+\d{4}$/, '').trim() || 'Unknown Doctor';

      // Scan data rows (row 9 to 39 matches standard template; scan wider to be safe)
      for (let rowNum = 7; rowNum <= 42; rowNum++) {
        const row = ws.getRow(rowNum);
        const rawDay  = row.getCell(1).value;
        const rawDayName = String(row.getCell(2).text || row.getCell(2).value || '').trim();

        const day = typeof rawDay === 'number' ? rawDay : parseInt(rawDay, 10);
        if (!day || isNaN(day) || day < 1 || day > 31) continue;
        if (!rawDayName || !/^[A-Za-z]/.test(rawDayName)) continue;

        const dateIso = `${year}-${pad(monthNum)}-${pad(day)}`;

        const normal   = String(row.getCell(3).text || row.getCell(3).value || '').trim();
        const oc1From  = cleanTime(row.getCell(5).value ?? row.getCell(5).text);
        const oc1To    = cleanTime(row.getCell(6).value ?? row.getCell(6).text);
        const oc2From  = cleanTime(row.getCell(7).value ?? row.getCell(7).text);
        const oc2To    = cleanTime(row.getCell(8).value ?? row.getCell(8).text);
        const notFrom  = cleanTime(row.getCell(9).value ?? row.getCell(9).text);
        const notTo    = cleanTime(row.getCell(10).value ?? row.getCell(10).text);
        const comments = String(row.getCell(11).text || row.getCell(11).value || '').trim();
        const dayName  = rawDayName || getDayName(dateIso);

        const push = (startTime, endTime, category, comment = '') => {
          if (!startTime && !endTime && !comment) return;
          const shiftType = inferShiftType(comment || normal, startTime, endTime, dayName);
          entries.push({
            doctorName: detectedDoctor,
            month, year, dateIso, dayName, shiftType,
            shiftLabel: SHIFT_DEF_MAP.get(shiftType)?.label || comment || 'Other',
            startTime, endTime,
            hours: hoursBetween(startTime, endTime),
            category,
            comment: comment || '',
            sourceFile: file.name
          });
        };

        if (normal) {
          const isLeave   = /leave/i.test(normal);
          const isNotOt   = /course|workshop|off.?day/i.test(normal);
          const shiftType = inferShiftType(normal, '', '', dayName);
          entries.push({
            doctorName: detectedDoctor,
            month, year, dateIso, dayName, shiftType,
            shiftLabel: SHIFT_DEF_MAP.get(shiftType)?.label || normal,
            startTime: '', endTime: '', hours: 0,
            category: isLeave ? 'leave' : isNotOt ? 'notOvertime' : 'normal',
            comment: normal,
            sourceFile: file.name
          });
        }

        push(oc1From, oc1To, 'onCall1', comments);
        push(oc2From, oc2To, 'onCall2', comments);
        push(notFrom, notTo,  'notOvertime', comments);
      }
    });

    return { source: file.name, entries };
  }

  // ─── Normaliser ─────────────────────────────────────────────────────────────
  function normalizeParsedEntries(results) {
    const byDoctor = new Map();
    results.flatMap(r => r.entries || []).forEach((entry) => {
      const doctor   = String(entry.doctorName || 'Unknown Doctor').trim();
      const shiftDef = SHIFT_DEF_MAP.get(entry.shiftType) || SHIFT_DEF_MAP.get('OTHER');
      const holiday  = HOLIDAYS[entry.dateIso] || '';
      let comment    = entry.comment || '';
      if (holiday && !comment.includes('PUBLIC HOLIDAY')) {
        comment = [comment, holiday].filter(Boolean).join(' | ');
      }
      const item = {
        id:          crypto.randomUUID(),
        doctorName:  doctor,
        month:       entry.month,
        year:        Number(entry.year),
        dateIso:     entry.dateIso,
        dayName:     entry.dayName || getDayName(entry.dateIso),
        shiftType:   shiftDef.key,
        shiftLabel:  entry.shiftLabel || shiftDef.label,
        startTime:   entry.startTime || shiftDef.startTime || '',
        endTime:     entry.endTime   || shiftDef.endTime   || '',
        hours:       entry.hours     ?? hoursBetween(entry.startTime || '', entry.endTime || ''),
        category:    entry.category  || shiftDef.category,
        comment,
        holidayName: holiday,
        sourceFile:  entry.sourceFile || ''
      };
      if (!byDoctor.has(doctor)) byDoctor.set(doctor, []);
      byDoctor.get(doctor).push(item);
    });
    for (const [doctor, entries] of byDoctor) {
      entries.sort((a, b) => a.dateIso.localeCompare(b.dateIso) || (a.startTime || '').localeCompare(b.startTime || ''));
      byDoctor.set(doctor, entries);
    }
    return byDoctor;
  }

  const filterRoster = (byDoctor, doctor, month, year) =>
    (byDoctor.get(doctor) || []).filter(i => i.month === month && Number(i.year) === Number(year));

  const rosterContainsLeave = (shifts) => shifts.some(s => s.category === 'leave');

  // ─── DOCX builder ───────────────────────────────────────────────────────────
  function docxParagraph(text, opts = {}) {
    const sz   = opts.size || 22;
    const bold = opts.bold      ? '<w:b/>'                    : '';
    const ul   = opts.underline ? '<w:u w:val="single"/>'     : '';
    const jc   = opts.align     ? `<w:jc w:val="${opts.align}"/>` : '';
    return `<w:p><w:pPr>${jc}</w:pPr><w:r><w:rPr>${bold}${ul}<w:sz w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
  }

  function docxTable(rows, widths) {
    const cols = widths.map(w => `<w:gridCol w:w="${w}"/>`).join('');
    const trs  = rows.map(row =>
      `<w:tr>${row.map((cell, i) =>
        `<w:tc><w:tcPr><w:tcW w:w="${widths[i]}" w:type="dxa"/></w:tcPr>${docxParagraph(cell.text || '', cell)}</w:tc>`
      ).join('')}</w:tr>`
    ).join('');
    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid>${cols}</w:tblGrid>${trs}</w:tbl>`;
  }

  async function buildDocx({ paragraphs = [], tables = [] }) {
    if (typeof JSZip === 'undefined') throw new Error('JSZip library failed to load.');
    const body = [
      ...paragraphs.map(p => docxParagraph(p.text || '', p)),
      ...tables.map(t => docxTable(t.rows, t.widths))
    ].join('');
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wne="http://schemas.microsoft.com/office/2006/wordml" mc:Ignorable="w14"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
    zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    zip.folder('word').file('document.xml', docXml);
    zip.folder('docProps').file('core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Roster Translator</dc:title></cp:coreProperties>`);
    zip.folder('docProps').file('app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Roster Translator</Application></Properties>`);
    return zip.generateAsync({ type: 'blob' });
  }

  // ─── Annexure C generator ────────────────────────────────────────────────────
  async function generateAnnexureCDocx({ employee, shifts, month, year }) {
    const name = fullName(employee).toUpperCase();
    const rank = (employee.designation === 'Other…' ? employee.designationOther : employee.designation || '').toUpperCase();
    const counts = new Map();
    shifts.filter(s => s.category === 'leave').forEach(s => counts.set(s.shiftLabel, (counts.get(s.shiftLabel) || 0) + 1));
    const leaveRows = [
      ['Type of leave taken', 'No reduction', 'Reduce overtime'],
      ['Vacation leave',               `${counts.get('Annual leave')                  || 0} day(s)`, ''],
      ['Sick leave',                   `${counts.get('Sick leave')                     || 0} day(s)`, ''],
      ['Family responsibility leave',  `${counts.get('Family responsibility leave')    || 0} day(s)`, ''],
      ['Special leave',                `${counts.get('Special leave')                  || 0} day(s)`, ''],
      ['Pre-natal leave',              `${counts.get('Pre-natal leave')                || 0} day(s)`, ''],
      ['Official courses/workshops',   `${(counts.get('Course') || 0) + (counts.get('Workshop') || 0)} day(s)`, '']
    ].map((r, i) => r.map(text => ({ text, bold: i === 0, size: i === 0 ? 20 : 18 })));
    return buildDocx({
      paragraphs: [
        { text: 'Annexure C',                      bold: true, underline: true,  align: 'center', size: 24 },
        { text: 'VERIFICATION OF COMMUTED OVERTIME WORKED FOR THE PERIOD', bold: true, underline: true, align: 'center', size: 22 },
        { text: `Date: ${String(month).toUpperCase()} ${year}`,             bold: true, size: 20 },
        { text: 'PART A',                          bold: true, underline: true,  size: 20 },
        { text: `Name: ${name}    Persal No: ${employee.persalNumber || ''}`, size: 20 },
        { text: `Rank: ${rank}    Department: ${DEPARTMENT}`,                size: 20 },
        { text: `Institution: ${INSTITUTION}`,                               size: 20 },
        { text: 'PART B',                          bold: true, underline: true,  size: 20 },
        { text: `1. I hereby certify that the above-named employee has performed the number of hours overtime (Group ${ANNEXURE_GROUP}) as agreed to his/her commuted overtime contract as well as that reflected in the duty roster for this particular month.`, size: 18 },
        { text: '2. I hereby certify that the above-named employee has performed the required number of working hours in this particular month (i.e. 40 hours per week).', size: 18 },
        { text: '3. During this particular period the following leave/no leave has been utilized by the employee in question:', size: 18 },
        { text: shifts.some(s => s.category === 'leave') ? 'Leave has been utilized' : 'No leave has been utilized', bold: true, size: 18 },
        { text: `${employee.signatureDate || ''}    HEAD OF CLINICAL DEPARTMENT / INSTITUTIONAL HEAD`, bold: true, size: 18 },
        { text: 'This document must be signed by the Head of Clinical Department or Institutional Head and not the Employee', bold: true, size: 18 }
      ],
      tables: [{ widths: [3600, 2200, 2200], rows: leaveRows }]
    });
  }

  // ─── Z1(a) leave form generator ─────────────────────────────────────────────
  async function generateZ1aDocx({ employee, shifts }) {
    const leave = shifts.filter(s => s.category === 'leave').sort((a, b) => a.dateIso.localeCompare(b.dateIso));
    const leaveRange = leave.length ? { start: leave[0].dateIso, end: leave[leave.length - 1].dateIso, days: leave.length, type: leave[0].shiftLabel } : null;
    const rowsA = [
      ['Type of Leave Taken as Working Days', 'Start Date', 'End Date', 'Number of Working Days'],
      [leaveRange?.type || '', leaveRange?.start || '', leaveRange?.end || '', String(leaveRange?.days || '')]
    ].map((r, i) => r.map(text => ({ text, bold: i === 0, size: i === 0 ? 20 : 18 })));
    return buildDocx({
      paragraphs: [
        { text: 'Z1 (a)',                                   bold: true, size: 22 },
        { text: 'APPLICATION FOR LEAVE OF ABSENCE',         bold: true, size: 22 },
        { text: `Surname: ${employee.surname || ''}`,                   size: 18 },
        { text: `Initials: ${initials(employee)}`,                      size: 18 },
        { text: `PERSAL Number: ${employee.persalNumber || ''}`,        size: 18 },
        { text: `Shift Worker: ${employee.shiftWorker || ''}`,          size: 18 },
        { text: `Casual Employee: ${employee.casualEmployee || ''}`,    size: 18 },
        { text: `Department: ${DEPARTMENT}`,                            size: 18 },
        { text: `Component: ${COMPONENT}`,                              size: 18 },
        { text: 'SECTION A: For periods covering a full day',  bold: true, size: 20 },
        { text: `Address during leave: ${employee.leaveAddress || ''}`, size: 18 },
        { text: `Employee signature date: ${employee.signatureDate || ''}`, size: 18 }
      ],
      tables: [{ widths: [3600, 1800, 1800, 1800], rows: rowsA }]
    });
  }

  // ─── Duty Roster Excel generator ────────────────────────────────────────────
  async function generateDutyRosterWorkbook({ employee, shifts, month, year }) {
    if (typeof ExcelJS === 'undefined') throw new Error('ExcelJS library failed to load.');
    setBoot('Loading roster template…');
    const response = await fetch(TEMPLATE_PATH);
    if (!response.ok) throw new Error(`Could not load template from ${TEMPLATE_PATH} (${response.status}). Make sure Duty-Rosters-2026.xlsx is in assets/templates/.`);
    const buffer = await response.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const sheetName = `${String(month).toUpperCase()} ${year}`;
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) throw new Error(`Worksheet "${sheetName}" not found in template. Available sheets: ${workbook.worksheets.map(w => w.name).join(', ')}`);

    // Clear editable area (rows 9–39, columns 3–11) except PUBLIC HOLIDAY labels in col 3
    for (let row = 9; row <= 39; row++) {
      for (let col = 3; col <= 11; col++) {
        const cell = sheet.getRow(row).getCell(col);
        const existing = String(cell.text || cell.value || '');
        if (col === 3 && existing.includes('PUBLIC HOLIDAY')) continue;
        cell.value = null;
      }
    }

    // Header fields
    sheet.getCell('C5').value = fullName(employee);
    sheet.getCell('H5').value = employee.persalNumber || '';
    sheet.getCell('F46').value = fullName(employee);
    sheet.getCell('J46').value = employee.signatureDate || '';
    const supervisorName = employee.supervisor === 'Other…' ? employee.supervisorOther : employee.supervisor;
    sheet.getCell('F50').value = supervisorName || '';
    sheet.getCell('J50').value = employee.signatureDate || '';

    // Write shift rows
    shifts.forEach((shift) => {
      const day = Number(String(shift.dateIso).slice(-2));
      if (!day || day < 1 || day > 31) return;
      const row = sheet.getRow(day + 8); // row 9 = day 1
      const holiday = HOLIDAYS[shift.dateIso] || '';

      if (shift.category === 'normal' || shift.category === 'leave') {
        row.getCell(3).value = shift.shiftLabel || shift.comment || '';
      }
      if (shift.category === 'onCall1') {
        row.getCell(5).value = shift.startTime || '';
        row.getCell(6).value = shift.endTime   || '';
      }
      if (shift.category === 'onCall2') {
        row.getCell(7).value = shift.startTime || '';
        row.getCell(8).value = shift.endTime   || '';
      }
      if (shift.category === 'notOvertime') {
        if (shift.startTime || shift.endTime) {
          row.getCell(9).value  = shift.startTime || '';
          row.getCell(10).value = shift.endTime   || '';
        } else {
          row.getCell(9).value = shift.shiftLabel || shift.comment || '';
        }
      }
      const parts = [holiday ? 'PUBLIC HOLIDAY' : '', shift.comment || ''].filter(Boolean).join(' | ');
      if (parts) row.getCell(11).value = parts;
    });

    const out = await workbook.xlsx.writeBuffer();
    return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // ─── UI helpers ─────────────────────────────────────────────────────────────
  function renderFiles() {
    els.uploadedFilesList.innerHTML = '';
    els.uploadedFilesPanel.classList.toggle('hidden', !state.files.length);
    state.files.forEach((file) => {
      const li = document.createElement('li');
      li.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
      els.uploadedFilesList.appendChild(li);
    });
    els.fileStatus.textContent = state.files.length
      ? `${state.files.length} file(s) selected`
      : 'Waiting for roster file…';
  }

  function updateDoctorOptions() {
    els.doctorSelect.innerHTML = '<option value="">— select —</option>';
    [...state.rostersByDoctor.keys()].sort().forEach((doctor) => {
      const o = document.createElement('option');
      o.value = doctor; o.textContent = doctor;
      els.doctorSelect.appendChild(o);
    });
  }

  function syncConditionalFields() {
    els.designationOther.classList.toggle('hidden', els.designationSelect.value !== 'Other…');
    els.supervisorOther.classList.toggle('hidden', els.supervisorSelect.value !== 'Other…');
    const hasLeave = rosterContainsLeave(state.currentShifts);
    els.leaveDetailsCard.classList.toggle('hidden', !hasLeave);
    els.downloadLeaveBtn.classList.toggle('hidden', !hasLeave);
  }

  function validateForDownloads() {
    const issues = [];
    if (!state.currentShifts.length)                        issues.push('No shifts selected for preview/export.');
    if (!state.employee.firstName)                          issues.push('First name is required.');
    if (!state.employee.surname)                            issues.push('Surname is required.');
    if (!state.employee.persalNumber)                       issues.push('PERSAL number is required.');
    if (!state.employee.designation)                        issues.push('Designation is required.');
    if (state.employee.designation === 'Other…' && !state.employee.designationOther) issues.push('Other designation is required.');
    if (!state.employee.supervisor)                         issues.push('Supervisor is required.');
    if (state.employee.supervisor === 'Other…' && !state.employee.supervisorOther) issues.push('Other supervisor is required.');
    if (!state.employee.signatureDate)                      issues.push('Date of signature is required.');
    if (rosterContainsLeave(state.currentShifts)) {
      if (!state.employee.shiftWorker)    issues.push('Shift worker selection is required (leave form).');
      if (!state.employee.casualEmployee) issues.push('Casual employee selection is required (leave form).');
      if (!state.employee.leaveAddress)   issues.push('Address during leave is required.');
    }
    els.validationPanel.classList.toggle('hidden', issues.length === 0);
    els.validationPanel.innerHTML = issues.length
      ? `<strong>Required to enable downloads:</strong><ul>${issues.map(i => `<li>${escapeHtml(i)}</li>`).join('')}</ul>`
      : '';
    const ok = issues.length === 0;
    els.downloadDutyRosterBtn.disabled = !ok;
    els.downloadAnnexureBtn.disabled   = !ok;
    els.downloadLeaveBtn.disabled      = !ok || !rosterContainsLeave(state.currentShifts);
  }

  function renderShiftRows() {
    els.shiftTableBody.innerHTML = '';
    state.currentShifts.forEach((shift) => {
      const fragment = els.shiftRowTemplate.content.cloneNode(true);
      const row = fragment.querySelector('tr');
      row.dataset.id = shift.id;
      const fields = Object.fromEntries(
        [...row.querySelectorAll('[data-field]')].map(el => [el.dataset.field, el])
      );
      fields.date.value      = shift.dateIso;
      fields.dayName.value   = shift.dayName;
      fields.startTime.value = shift.startTime || '';
      fields.endTime.value   = shift.endTime   || '';
      fields.hours.value     = shift.hours     ?? 0;
      fields.comment.value   = shift.comment   || '';

      SHIFT_DEFINITIONS.forEach((def) => {
        const o = document.createElement('option');
        o.value = def.key; o.textContent = def.label;
        fields.shiftType.appendChild(o);
      });
      fields.shiftType.value = shift.shiftType;

      SHIFT_CATEGORIES.forEach((cat) => {
        const o = document.createElement('option');
        o.value = cat.key; o.textContent = cat.label;
        fields.category.appendChild(o);
      });
      fields.category.value = shift.category;

      row.addEventListener('input', (event) => {
        const target = event.target;
        const item   = state.currentShifts.find(s => s.id === row.dataset.id);
        if (!item) return;
        const field = target.dataset.field;
        if (!field) return;
        item[field] = target.value;
        if (field === 'date')                            item.dayName = fields.dayName.value = getDayName(target.value);
        if (field === 'startTime' || field === 'endTime') item.hours = fields.hours.value = hoursBetween(fields.startTime.value, fields.endTime.value);
        if (field === 'shiftType') {
          const def = SHIFT_DEF_MAP.get(target.value);
          if (def) {
            item.shiftLabel = def.label;
            if (def.startTime) { item.startTime = fields.startTime.value = def.startTime; }
            if (def.endTime)   { item.endTime   = fields.endTime.value   = def.endTime;   }
            if (def.hours)     { item.hours     = fields.hours.value     = def.hours;     }
            if (def.category)  { item.category  = fields.category.value  = def.category;  }
          }
        }
        validateForDownloads();
        syncConditionalFields();
      });

      row.addEventListener('click', (event) => {
        const action = event.target?.dataset?.action;
        if (!action) return;
        const index = state.currentShifts.findIndex(s => s.id === row.dataset.id);
        if (index < 0) return;
        if (action === 'delete')    state.currentShifts.splice(index, 1);
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
    els.previewTitle.textContent   = `${state.selectedDoctor} — ${state.selectedMonth} ${state.selectedYear}`;
    els.previewSummary.textContent = `${state.currentShifts.length} shift item(s)`;
    renderShiftRows();
  }

  function populateStaticControls() {
    const makeOpts = (select, items) => {
      select.innerHTML = '<option value="">— select —</option>';
      items.forEach(v => { const o = document.createElement('option'); o.value = v; o.textContent = v; select.appendChild(o); });
    };
    makeOpts(els.designationSelect, DESIGNATIONS);
    makeOpts(els.supervisorSelect,   SUPERVISORS);
    makeOpts(els.monthSelect,        MONTHS);
    [2024, 2025, 2026, 2027].forEach(v => {
      const o = document.createElement('option'); o.value = v; o.textContent = v; els.yearSelect.appendChild(o);
    });
  }

  // ─── Extract handler ─────────────────────────────────────────────────────────
  async function handleExtract() {
    clearMessages();
    if (!state.files.length) {
      showMessage('Select at least one roster file first.', 'error');
      return;
    }
    state.parseResults = [];
    for (const file of state.files) {
      try {
        if (/\.pdf$/i.test(file.name)) {
          showMessage(`PDF parsing is not yet implemented for "${file.name}". Please use the Excel roster format.`, 'error');
        } else {
          const result = await parseExcelRoster(file);
          state.parseResults.push(result);
          const n = result.entries.length;
          showMessage(`Parsed "${file.name}" — ${n} shift row${n !== 1 ? 's' : ''} found.`, n > 0 ? 'success' : 'info');
        }
      } catch (error) {
        showMessage(`Failed to parse "${file.name}": ${error.message}`, 'error');
        console.error(error);
      }
    }
    state.rostersByDoctor = normalizeParsedEntries(state.parseResults);
    updateDoctorOptions();
    if (!state.parseResults.length) {
      showMessage('No files were parsed successfully.', 'error');
    } else if (!state.rostersByDoctor.size) {
      showMessage('Files parsed, but no roster rows were detected. Check that the Excel format matches the expected template (day numbers in col A, day names in col B, data from row 9).', 'error');
    } else {
      showMessage(`Extraction complete: ${state.rostersByDoctor.size} doctor roster(s) available. Select a doctor to preview.`, 'success');
    }
  }

  // ─── Download handlers ───────────────────────────────────────────────────────
  async function withDownloadGuard(label, fn) {
    setBoot(`Generating ${label}…`);
    try {
      await fn();
      setBoot(`${label} downloaded.`);
    } catch (error) {
      showMessage(`Export failed: ${error.message}`, 'error');
      setBoot(`Export error: ${error.message}`, true);
      console.error(error);
    }
  }

  // ─── Wire up events ──────────────────────────────────────────────────────────
  function wireEvents() {
    // File deduplication helper
    const mergeFiles = (incoming) => {
      const map = new Map(state.files.map(f => [`${f.name}|${f.size}|${f.lastModified}`, f]));
      filesFromList(incoming).forEach(f => map.set(`${f.name}|${f.size}|${f.lastModified}`, f));
      state.files = [...map.values()];
    };
    const filesFromList = (list) => Array.from(list || []).filter(f => /\.(pdf|xlsx|xls|csv)$/i.test(f.name));

    // Prevent browser default drag-open on body
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt =>
      document.addEventListener(evt, e => e.preventDefault(), false)
    );

    // Drop zone drag visuals
    ['dragenter', 'dragover'].forEach(evt =>
      els.dropZone.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); els.dropZone.classList.add('dragover'); })
    );
    ['dragleave', 'drop'].forEach(evt =>
      els.dropZone.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); els.dropZone.classList.remove('dragover'); })
    );

    // Drop
    els.dropZone.addEventListener('drop', (e) => {
      mergeFiles(e.dataTransfer?.files);
      renderFiles();
      setBoot(state.files.length ? `${state.files.length} file(s) loaded. Click "Extract data".` : 'No supported files dropped.');
    });

    // File input change — reset value after reading so re-selecting same file works
    els.rosterFiles.addEventListener('change', () => {
      mergeFiles(els.rosterFiles.files);
      els.rosterFiles.value = '';
      renderFiles();
      setBoot(state.files.length ? `${state.files.length} file(s) loaded. Click "Extract data".` : 'No supported files selected.');
    });

    els.clearFilesBtn.addEventListener('click', () => {
      state.files = [];
      els.rosterFiles.value = '';
      renderFiles();
      clearMessages();
      setBoot('Files cleared.');
    });

    els.extractBtn.addEventListener('click', async () => {
      setBoot('Extracting data…');
      await handleExtract();
      setBoot(state.rostersByDoctor.size
        ? `Data extracted — ${state.rostersByDoctor.size} roster(s). Select doctor, month, and year then click Preview.`
        : 'No roster rows extracted. Check file format or the browser console for details.'
      );
    });

    els.doctorSelect.addEventListener('change', () => {
      state.selectedDoctor = els.doctorSelect.value;
      els.timesheetName.value = state.employee.timesheetName = els.doctorSelect.value;
    });

    els.clearDoctorBtn.addEventListener('click', () => {
      state.selectedDoctor = '';
      els.doctorSelect.value = '';
      state.currentShifts = [];
      renderPreview();
      validateForDownloads();
    });

    els.monthSelect.addEventListener('change', () => state.selectedMonth = els.monthSelect.value);
    els.yearSelect.addEventListener('change', ()  => state.selectedYear  = els.yearSelect.value);

    els.previewBtn.addEventListener('click', () => {
      state.selectedDoctor = els.doctorSelect.value;
      state.selectedMonth  = els.monthSelect.value;
      state.selectedYear   = els.yearSelect.value;
      if (!state.selectedDoctor || !state.selectedMonth || !state.selectedYear) {
        showMessage('Please select a doctor, month, and year before previewing.', 'info');
        return;
      }
      state.currentShifts = filterRoster(state.rostersByDoctor, state.selectedDoctor, state.selectedMonth, state.selectedYear);
      if (!state.currentShifts.length) {
        showMessage(`No shifts found for ${state.selectedDoctor} in ${state.selectedMonth} ${state.selectedYear}. You can add shifts manually using "Add shift".`, 'info');
      }
      renderPreview();
      syncConditionalFields();
      validateForDownloads();
    });

    els.addShiftBtn.addEventListener('click', () => {
      const monthIdx  = MONTHS.indexOf(state.selectedMonth) + 1 || 1;
      const baseDate  = `${state.selectedYear || 2026}-${pad(monthIdx)}-01`;
      state.currentShifts.push({
        id: crypto.randomUUID(), doctorName: state.selectedDoctor,
        month: state.selectedMonth, year: Number(state.selectedYear || 2026),
        dateIso: baseDate, dayName: getDayName(baseDate),
        shiftType: 'OTHER', shiftLabel: 'Other',
        startTime: '', endTime: '', hours: 0, category: 'notOvertime', comment: ''
      });
      renderShiftRows();
      validateForDownloads();
    });

    els.sortShiftsBtn.addEventListener('click', () => {
      state.currentShifts.sort((a, b) => a.dateIso.localeCompare(b.dateIso) || (a.startTime || '').localeCompare(b.startTime || ''));
      renderShiftRows();
    });

    els.resetAppBtn.addEventListener('click', () => {
      state.files = []; state.parseResults = []; state.rostersByDoctor = new Map();
      state.selectedDoctor = ''; state.selectedMonth = ''; state.selectedYear = '';
      state.currentShifts = [];
      Object.keys(state.employee).forEach(k => state.employee[k] = '');
      els.rosterFiles.value = '';
      document.querySelectorAll('input, select, textarea').forEach((el) => {
        if (el.type === 'file') return;
        el.tagName === 'SELECT' ? (el.selectedIndex = 0) : (el.value = '');
      });
      renderFiles(); clearMessages(); updateDoctorOptions(); renderPreview(); syncConditionalFields(); validateForDownloads();
      setBoot('App reset.');
    });

    // Employee fields
    [
      ['firstName',      els.firstName],        ['surname',          els.surname],
      ['persalNumber',   els.persalNumber],      ['designation',      els.designationSelect],
      ['designationOther',els.designationOther], ['supervisor',       els.supervisorSelect],
      ['supervisorOther',els.supervisorOther],   ['signatureDate',    els.signatureDate],
      ['timesheetName',  els.timesheetName],     ['shiftWorker',      els.shiftWorkerSelect],
      ['casualEmployee', els.casualEmployeeSelect],['leaveAddress',   els.leaveAddress]
    ].forEach(([key, el]) => {
      ['input', 'change'].forEach(evt => el.addEventListener(evt, () => {
        state.employee[key] = el.value;
        syncConditionalFields();
        validateForDownloads();
      }));
    });

    // Downloads
    els.downloadDutyRosterBtn.addEventListener('click', () =>
      withDownloadGuard('duty roster', async () => {
        const blob = await generateDutyRosterWorkbook({ employee: state.employee, shifts: state.currentShifts, month: state.selectedMonth, year: state.selectedYear });
        downloadBlob(blob, `Duty-Roster-${state.selectedMonth}-${state.selectedYear}-${state.employee.surname || 'employee'}.xlsx`);
      })
    );

    els.downloadAnnexureBtn.addEventListener('click', () =>
      withDownloadGuard('Annexure C', async () => {
        const blob = await generateAnnexureCDocx({ employee: state.employee, shifts: state.currentShifts, month: state.selectedMonth, year: state.selectedYear });
        downloadBlob(blob, `Annexure-C-${state.selectedMonth}-${state.selectedYear}-${state.employee.surname || 'employee'}.docx`);
      })
    );

    els.downloadLeaveBtn.addEventListener('click', () =>
      withDownloadGuard('Z1(a) leave form', async () => {
        const blob = await generateZ1aDocx({ employee: state.employee, shifts: state.currentShifts });
        downloadBlob(blob, `Z1a-Leave-${state.selectedMonth}-${state.selectedYear}-${state.employee.surname || 'employee'}.docx`);
      })
    );
  }

  // ─── Init ────────────────────────────────────────────────────────────────────
  try {
    populateStaticControls();
    wireEvents();
    renderFiles();
    updateDoctorOptions();
    syncConditionalFields();
    validateForDownloads();
    setBoot('App loaded. Click the upload area to choose your roster file(s).');
    boot.classList.remove('hidden');
  } catch (error) {
    console.error(error);
    setBoot(`App failed to initialize: ${error.message}`, true);
  }

})();
