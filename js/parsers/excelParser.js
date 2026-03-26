import { monthNameToNumber, toIsoDate, getDayNameFromIso } from '../utils.js';

function inferShiftType(comment, startTime, endTime, dayName) {
  const text = String(comment || '').toLowerCase();
  if (text.includes('annual leave')) return 'ANNUAL_LEAVE';
  if (text.includes('sick')) return 'SICK_LEAVE';
  if (text.includes('family')) return 'FAMILY_RESPONSIBILITY';
  if (text.includes('special leave')) return 'SPECIAL_LEAVE';
  if (text.includes('pre-natal')) return 'PRE_NATAL';
  if (text.includes('course')) return 'COURSE';
  if (text.includes('workshop')) return 'WORKSHOP';
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
  return text.includes('off') ? 'OFF_DAY' : 'OTHER';
}

function cleanTime(value) {
  if (!value) return '';
  const text = String(value).trim().replace(/h/ig, ':').replace(/[.]/g, ':');
  const m = text.match(/(\d{1,2})[:](\d{2})/);
  if (!m) return '';
  return `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`;
}

export async function parseExcelRoster(file) {
  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const results = [];

  wb.worksheets.forEach((ws) => {
    const sheetName = ws.name || '';
    const monthMatch = sheetName.match(/([A-Za-z]+)\s+(\d{4})/);
    const monthName = monthMatch?.[1] || '';
    const year = Number(monthMatch?.[2] || 0);
    const month = monthNameToNumber(monthName);
    if (!month || !year) return;

    const headerText = [];
    for (let r = 1; r <= 15; r += 1) {
      const row = ws.getRow(r);
      row.eachCell((cell) => headerText.push(String(cell.text || cell.value || '').trim()));
    }
    const employeeHeader = headerText.find((t) => /name of employee/i.test(t));
    const doctorName = employeeHeader ? '' : '';

    for (let rowNum = 9; rowNum <= 39; rowNum += 1) {
      const row = ws.getRow(rowNum);
      const dateValue = row.getCell(1).value;
      const dayName = String(row.getCell(2).text || row.getCell(2).value || '').trim();
      const day = Number(dateValue);
      if (!day || !dayName) continue;

      const normal = String(row.getCell(3).text || row.getCell(3).value || '').trim();
      const oc1From = cleanTime(row.getCell(5).text || row.getCell(5).value || '');
      const oc1To = cleanTime(row.getCell(6).text || row.getCell(6).value || '');
      const oc2From = cleanTime(row.getCell(7).text || row.getCell(7).value || '');
      const oc2To = cleanTime(row.getCell(8).text || row.getCell(8).value || '');
      const notOtFrom = cleanTime(row.getCell(9).text || row.getCell(9).value || '');
      const notOtTo = cleanTime(row.getCell(10).text || row.getCell(10).value || '');
      const comments = String(row.getCell(11).text || row.getCell(11).value || '').trim();
      const dateIso = toIsoDate(year, month, day);

      const pushEntry = (startTime, endTime, category, comment = '') => {
        if (!startTime && !endTime && !comment) return;
        results.push({
          doctorName: doctorName || 'Unknown Doctor',
          month: monthName,
          year,
          dateIso,
          dayName: dayName || getDayNameFromIso(dateIso),
          shiftType: inferShiftType(comment || normal, startTime, endTime, dayName),
          startTime,
          endTime,
          category,
          comment: comment || normal,
          sourceFile: file.name
        });
      };

      if (normal) {
        results.push({
          doctorName: doctorName || 'Unknown Doctor',
          month: monthName,
          year,
          dateIso,
          dayName: dayName || getDayNameFromIso(dateIso),
          shiftType: inferShiftType(normal, '', '', dayName),
          startTime: '',
          endTime: '',
          category: /leave|course|workshop|off/i.test(normal) ? (/leave/i.test(normal) ? 'leave' : 'notOvertime') : 'normal',
          comment: normal,
          sourceFile: file.name
        });
      }

      pushEntry(oc1From, oc1To, 'onCall1', comments);
      pushEntry(oc2From, oc2To, 'onCall2', comments);
      pushEntry(notOtFrom, notOtTo, 'notOvertime', comments);
    }
  });

  return { source: file.name, entries: results };
}
