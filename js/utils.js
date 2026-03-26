import { MONTHS } from './constants.js';

export function escapeHtml(value = '') {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

export function pad(number) {
  return String(number).padStart(2, '0');
}

export function monthNameToNumber(monthName) {
  const index = MONTHS.findIndex((m) => m.toLowerCase() === String(monthName).toLowerCase());
  return index >= 0 ? index + 1 : null;
}

export function toIsoDate(year, month, day) {
  return `${year}-${pad(month)}-${pad(day)}`;
}

export function getDayNameFromIso(dateIso) {
  const d = new Date(`${dateIso}T00:00:00`);
  return d.toLocaleDateString('en-ZA', { weekday: 'long' });
}

export function hoursBetween(startTime, endTime) {
  if (!startTime || !endTime) return 0;
  const [sh, sm] = startTime.split(':').map(Number);
  const [eh, em] = endTime.split(':').map(Number);
  let start = sh * 60 + sm;
  let end = eh * 60 + em;
  if (end < start) end += 24 * 60;
  return Number(((end - start) / 60).toFixed(2));
}

export function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function fullEmployeeName(employee) {
  return [employee.firstName, employee.surname].filter(Boolean).join(' ').trim();
}

export function initialsFromName(employee) {
  return [employee.firstName, employee.surname].filter(Boolean).map(part => part.trim().charAt(0).toUpperCase()).join('');
}
