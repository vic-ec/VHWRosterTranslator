import { SHIFT_DEFINITIONS } from '../constants.js';
import { getHolidayName } from '../holidays.js';
import { getDayNameFromIso, hoursBetween } from '../utils.js';

const defsByKey = new Map(SHIFT_DEFINITIONS.map((item) => [item.key, item]));

export function normalizeParsedEntries(results) {
  const byDoctor = new Map();
  results.flatMap(result => result.entries || []).forEach((entry) => {
    const doctorName = String(entry.doctorName || 'Unknown Doctor').trim();
    const shiftDef = defsByKey.get(entry.shiftType) || defsByKey.get('OTHER');
    const normalized = {
      id: crypto.randomUUID(),
      doctorName,
      month: entry.month,
      year: Number(entry.year),
      dateIso: entry.dateIso,
      dayName: entry.dayName || getDayNameFromIso(entry.dateIso),
      shiftType: shiftDef.key,
      shiftLabel: shiftDef.label,
      startTime: entry.startTime || shiftDef.startTime || '',
      endTime: entry.endTime || shiftDef.endTime || '',
      hours: entry.hours ?? (entry.startTime && entry.endTime ? hoursBetween(entry.startTime, entry.endTime) : shiftDef.hours || 0),
      category: entry.category || shiftDef.category,
      comment: entry.comment || '',
      holidayName: getHolidayName(entry.dateIso),
      sourceFile: entry.sourceFile || ''
    };
    if (normalized.holidayName && !normalized.comment.includes('PUBLIC HOLIDAY')) {
      normalized.comment = [normalized.comment, normalized.holidayName].filter(Boolean).join(' | ');
    }
    if (!byDoctor.has(doctorName)) byDoctor.set(doctorName, []);
    byDoctor.get(doctorName).push(normalized);
  });
  for (const [doctor, entries] of byDoctor.entries()) {
    entries.sort((a, b) => a.dateIso.localeCompare(b.dateIso) || a.startTime.localeCompare(b.startTime));
    byDoctor.set(doctor, entries);
  }
  return byDoctor;
}

export function filterRoster(byDoctor, doctorName, month, year) {
  const entries = byDoctor.get(doctorName) || [];
  return entries.filter(item => item.month === month && Number(item.year) === Number(year));
}

export function rosterContainsLeave(shifts) {
  return shifts.some((shift) => shift.category === 'leave');
}
