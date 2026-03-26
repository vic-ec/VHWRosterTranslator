export const state = {
  files: [],
  parseResults: [],
  rostersByDoctor: new Map(),
  selectedDoctor: '',
  selectedMonth: '',
  selectedYear: '',
  currentShifts: [],
  employee: {
    firstName: '',
    surname: '',
    persalNumber: '',
    designation: '',
    designationOther: '',
    supervisor: '',
    supervisorOther: '',
    signatureDate: '',
    timesheetName: '',
    shiftWorker: '',
    casualEmployee: '',
    leaveAddress: ''
  }
};

export function resetState() {
  state.files = [];
  state.parseResults = [];
  state.rostersByDoctor = new Map();
  state.selectedDoctor = '';
  state.selectedMonth = '';
  state.selectedYear = '';
  state.currentShifts = [];
  Object.keys(state.employee).forEach((key) => state.employee[key] = '');
}
