export const MONTHS = [
  'January','February','March','April','May','June','July','August','September','October','November','December'
];

export const DESIGNATIONS = [
  'Intern',
  'Community Service Medical Officer',
  'Medical Officer Grade 1',
  'Medical Officer Grade 2',
  'Medical Officer Grade 3',
  'Registrar',
  'Medical Specialist Grade 1',
  'Medical Specialist Grade 2',
  'Medical Specialist Grade 3',
  'Other…'
];

export const SUPERVISORS = ['Philip Cloete', 'Sebastian De Haan', 'Paul Xafis', 'Other…'];

export const SHIFT_DEFINITIONS = [
  { key: 'WD_0800_1800', label: 'WD shift 08h00-18h00', category: 'onCall1', startTime: '08:00', endTime: '18:00', hours: 10 },
  { key: 'WD_1200_2200', label: 'WD shift 12h00-22h00', category: 'onCall1', startTime: '12:00', endTime: '22:00', hours: 10 },
  { key: 'WD_1500_2300', label: 'WD shift 15h00-23h00', category: 'onCall1', startTime: '15:00', endTime: '23:00', hours: 8 },
  { key: 'WD_2200_1000', label: 'WD shift 22h00-10h00', category: 'onCall1', startTime: '22:00', endTime: '10:00', hours: 12 },
  { key: 'WE_0800_2000', label: 'WE shift 08h00-20h00', category: 'onCall1', startTime: '08:00', endTime: '20:00', hours: 12 },
  { key: 'WE_1300_2300', label: 'WE shift 13h00-23h00', category: 'onCall1', startTime: '13:00', endTime: '23:00', hours: 10 },
  { key: 'WE_2000_1000', label: 'WE shift 20h00-10h00', category: 'onCall1', startTime: '20:00', endTime: '10:00', hours: 14 },
  { key: 'ANNUAL_LEAVE', label: 'Annual leave', category: 'leave', startTime: '', endTime: '', hours: 0 },
  { key: 'SICK_LEAVE', label: 'Sick leave', category: 'leave', startTime: '', endTime: '', hours: 0 },
  { key: 'FAMILY_RESPONSIBILITY', label: 'Family responsibility leave', category: 'leave', startTime: '', endTime: '', hours: 0 },
  { key: 'SPECIAL_LEAVE', label: 'Special leave', category: 'leave', startTime: '', endTime: '', hours: 0 },
  { key: 'PRE_NATAL', label: 'Pre-natal leave', category: 'leave', startTime: '', endTime: '', hours: 0 },
  { key: 'UNPAID_LEAVE', label: 'Unpaid leave', category: 'leave', startTime: '', endTime: '', hours: 0 },
  { key: 'WORKSHOP', label: 'Workshop', category: 'notOvertime', startTime: '', endTime: '', hours: 0 },
  { key: 'COURSE', label: 'Course', category: 'notOvertime', startTime: '', endTime: '', hours: 0 },
  { key: 'OFF_DAY', label: 'Off-day', category: 'notOvertime', startTime: '', endTime: '', hours: 0 },
  { key: 'OTHER', label: 'Other', category: 'notOvertime', startTime: '', endTime: '', hours: 0 }
];

export const SHIFT_CATEGORIES = [
  { key: 'normal', label: 'Normal working hours' },
  { key: 'onCall1', label: '1st on call - on site' },
  { key: 'onCall2', label: '2nd on call - off site' },
  { key: 'notOvertime', label: 'Not overtime' },
  { key: 'leave', label: 'Leave' }
];

export const LEAVE_SHIFT_KEYS = new Set([
  'ANNUAL_LEAVE','SICK_LEAVE','FAMILY_RESPONSIBILITY','SPECIAL_LEAVE','PRE_NATAL','UNPAID_LEAVE'
]);

export const INSTITUTION = 'VICTORIA HOSPITAL';
export const DEPARTMENT = 'HEALTH';
export const COMPONENT = 'Emergency Centre';
export const ANNEXURE_GROUP = '3.2';
