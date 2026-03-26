import { ANNEXURE_GROUP, DEPARTMENT, INSTITUTION } from '../constants.js';
import { buildSimpleDocx } from './docxHelpers.js';
import { fullEmployeeName } from '../utils.js';

function leaveSummary(shifts) {
  const counts = new Map();
  shifts.filter(s => s.category === 'leave').forEach((s) => {
    counts.set(s.shiftLabel, (counts.get(s.shiftLabel) || 0) + 1);
  });
  return counts;
}

export async function generateAnnexureCDocx({ employee, shifts, month, year }) {
  const name = fullEmployeeName(employee).toUpperCase();
  const rank = (employee.designation === 'Other…' ? employee.designationOther : employee.designation || '').toUpperCase();
  const summary = leaveSummary(shifts);
  const leaveRows = [
    ['Type of leave taken', 'No reduction', 'Reduce overtime'],
    ['Vacation leave', `${summary.get('Annual leave') || 0} days`, ''],
    ['Sick leave', `${summary.get('Sick leave') || 0} days`, ''],
    ['Family responsibility leave', `${summary.get('Family responsibility leave') || 0} days`, ''],
    ['Special leave', `${summary.get('Special leave') || 0} days`, ''],
    ['Pre-natal leave', `${summary.get('Pre-natal leave') || 0} days`, ''],
    ['Official courses/workshops', `${summary.get('Course') || 0 + (summary.get('Workshop') || 0)} days`, '']
  ].map((r, i) => r.map((text) => ({ text, bold: i === 0, size: i === 0 ? 20 : 18 })));

  return buildSimpleDocx({
    paragraphs: [
      { text: 'Annexure C', bold: true, underline: true, align: 'center', size: 24 },
      { text: 'VERIFICATION OF COMMUTED OVERTIME WORKED FOR THE PERIOD', bold: true, underline: true, align: 'center', size: 22 },
      { text: `Date: ${String(month).toUpperCase()} ${year}`, bold: true, size: 20 },
      { text: 'PART A', bold: true, underline: true, size: 20 },
      { text: `Name: ${name}    Persal No: ${employee.persalNumber || ''}`, size: 20 },
      { text: `Rank: ${rank}    Department: ${DEPARTMENT}`, size: 20 },
      { text: `Institution: ${INSTITUTION}`, size: 20 },
      { text: 'PART B', bold: true, underline: true, size: 20 },
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
