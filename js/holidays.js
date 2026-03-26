const HOLIDAYS_BY_YEAR = {
  2026: {
    '2026-01-01': 'Public Holiday',
    '2026-03-21': 'Human Rights Day',
    '2026-04-03': 'Good Friday',
    '2026-04-06': 'Family Day',
    '2026-05-01': 'Workers' Day',
    '2026-06-16': 'Youth Day',
    '2026-08-09': 'National Women's Day',
    '2026-09-23': 'Heritage Day',
    '2026-12-16': 'Day of Reconciliation',
    '2026-12-25': 'Christmas Day',
    '2026-12-26': 'Day of Goodwill'
  }
};

export function getHolidayName(dateIso) {
  const year = Number(String(dateIso).slice(0, 4));
  return HOLIDAYS_BY_YEAR[year]?.[dateIso] || '';
}
