import { COMPONENT, DEPARTMENT } from '../constants.js';
import { buildSimpleDocx } from './docxHelpers.js';
import { initialsFromName } from '../utils.js';

function firstLeaveRange(shifts) {
  const leave = shifts.filter(s => s.category === 'leave').sort((a, b) => a.dateIso.localeCompare(b.dateIso));
  if (!leave.length) return null;
  return { start: leave[0].dateIso, end: leave[leave.length - 1].dateIso, days: leave.length, type: leave[0].shiftLabel };
}

export async function generateZ1aDocx({ employee, shifts }) {
  const leave = firstLeaveRange(shifts);
  const rowsA = [
    ['Type of Leave Taken as Working Days', 'Start Date', 'End Date', 'Number of Working Days'],
    [leave?.type || '', leave?.start || '', leave?.end || '', String(leave?.days || '')]
  ].map((r, i) => r.map((text) => ({ text, bold: i === 0, size: i === 0 ? 20 : 18 })));

  return buildSimpleDocx({
    paragraphs: [
      { text: 'Z1 (a)', bold: true, size: 22 },
      { text: 'APPLICATION FOR LEAVE OF ABSENCE', bold: true, size: 22 },
      { text: `Surname: ${employee.surname || ''}`, size: 18 },
      { text: `Initials: ${initialsFromName(employee)}`, size: 18 },
      { text: `PERSAL Number: ${employee.persalNumber || ''}`, size: 18 },
      { text: `Shift Worker: ${employee.shiftWorker || ''}`, size: 18 },
      { text: `Casual Employee: ${employee.casualEmployee || ''}`, size: 18 },
      { text: `Department: ${DEPARTMENT}`, size: 18 },
      { text: `Component: ${COMPONENT}`, size: 18 },
      { text: 'SECTION A: For periods covering a full day', bold: true, size: 20 },
      { text: `Address during leave: ${employee.leaveAddress || ''}`, size: 18 },
      { text: `Employee signature date: ${employee.signatureDate || ''}`, size: 18 }
    ],
    tables: [{ widths: [3600, 1800, 1800, 1800], rows: rowsA }]
  });
}
