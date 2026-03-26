import { fullEmployeeName } from '../utils.js';
import { getHolidayName } from '../holidays.js';

function pickSheet(workbook, month, year) {
  const name = `${String(month).toUpperCase()} ${year}`;
  return workbook.getWorksheet(name);
}

function clearEditableArea(sheet) {
  for (let row = 9; row <= 39; row += 1) {
    for (let col = 3; col <= 11; col += 1) {
      const cell = sheet.getRow(row).getCell(col);
      if (col === 3 && String(sheet.getRow(row).getCell(3).text || '').includes('PUBLIC HOLIDAY')) continue;
      cell.value = null;
    }
  }
}

export async function generateDutyRosterWorkbook({ templateFile, employee, shifts, month, year }) {
  const buffer = await templateFile.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = pickSheet(workbook, month, year);
  if (!sheet) throw new Error(`Worksheet not found for ${month} ${year}`);

  clearEditableArea(sheet);
  sheet.getCell('C5').value = fullEmployeeName(employee);
  sheet.getCell('H5').value = employee.persalNumber || '';
  sheet.getCell('F46').value = fullEmployeeName(employee);
  sheet.getCell('J46').value = employee.signatureDate || '';
  sheet.getCell('F50').value = employee.supervisor === 'Other…' ? employee.supervisorOther : employee.supervisor;
  sheet.getCell('J50').value = employee.signatureDate || '';

  shifts.forEach((shift) => {
    const day = Number(String(shift.dateIso).slice(-2));
    const row = day + 8;
    const holiday = getHolidayName(shift.dateIso);
    if (shift.category === 'normal' || shift.category === 'leave') {
      sheet.getRow(row).getCell(3).value = shift.shiftLabel || shift.comment || '';
    }
    if (shift.category === 'onCall1') {
      sheet.getRow(row).getCell(5).value = shift.startTime || '';
      sheet.getRow(row).getCell(6).value = shift.endTime || '';
    }
    if (shift.category === 'onCall2') {
      sheet.getRow(row).getCell(7).value = shift.startTime || '';
      sheet.getRow(row).getCell(8).value = shift.endTime || '';
    }
    if (shift.category === 'notOvertime') {
      if (shift.startTime || shift.endTime) {
        sheet.getRow(row).getCell(9).value = shift.startTime || '';
        sheet.getRow(row).getCell(10).value = shift.endTime || '';
      } else {
        sheet.getRow(row).getCell(9).value = shift.shiftLabel || shift.comment || '';
      }
    }
    const comments = [holiday ? 'PUBLIC HOLIDAY' : '', shift.comment || ''].filter(Boolean).join(' | ');
    if (comments) sheet.getRow(row).getCell(11).value = comments;
  });

  const out = await workbook.xlsx.writeBuffer();
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
