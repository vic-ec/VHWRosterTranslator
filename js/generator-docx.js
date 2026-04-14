// ═══════════════════════════════════════════════════════════════
// generator-docx.js — Word document generators.
//
//   generateAnnexureCDocx(details, shifts, month, year, profile)
//     → Promise<Blob>  — WCG Annexure C overtime claim form
//
//   generateZ1ADocx(details, shifts, month, year, leaveType)
//     → Promise<Blob>  — Z1(a) leave application form
//
// Depends on: config.js, holidays.js  (docx library loaded globally)
// ═══════════════════════════════════════════════════════════════

// === FORM GENERATORS ===
// === Form Generators (Annexure C + Z1a) ===
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign,
  convertInchesToTwip, UnderlineType, TabStopType
} = docx;
const FS = 18; const FSS = 16;
const S0 = { before:0, after:0, line:240, lineRule:'exact' };
const NONE_B = { style:BorderStyle.NONE, size:0, color:'FFFFFF' };
const SNG    = { style:BorderStyle.SINGLE, size:4, color:'000000' };
const allSng = { top:SNG, bottom:SNG, left:SNG, right:SNG };
const pageWidth = 10800;

function t(txt, opts={}) { return new TextRun({text:String(txt||''),font:'Calibri',size:FS,...opts}); }
function b(txt, opts={}) { return t(txt, {bold:true,...opts}); }
function u(txt, opts={}) { return t(txt, {underline:{type:UnderlineType.SINGLE},...opts}); }
function it(txt, opts={}) { return t(txt, {italics:true,...opts}); }
function p(ch, opts={}) { return new Paragraph({children:Array.isArray(ch)?ch:[ch],spacing:S0,...opts}); }
function blank(bef=50) { return new Paragraph({children:[t('')],spacing:{before:bef,after:0,line:240,lineRule:'exact'}}); }
function botLine() { return new Paragraph({children:[t('')],border:{bottom:SNG},spacing:{before:0,after:4,line:240,lineRule:'exact'}}); }
function nb(w,ch,opts={}) {
  return new TableCell({children:Array.isArray(ch)?ch:[p(ch)],width:{size:w,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},...opts});
}

// ── Leave table ───────────────────────────────────────────────────────────────
const COL_W = [4200, 3000, 3600];
const tblTop = { style:BorderStyle.SINGLE,size:4,color:'000000' };
const tblBot = { style:BorderStyle.SINGLE,size:4,color:'000000' };
const tblL   = { style:BorderStyle.SINGLE,size:4,color:'000000' };
const tblR   = { style:BorderStyle.SINGLE,size:4,color:'000000' };
const colD   = { style:BorderStyle.SINGLE,size:4,color:'000000' };
const hBL = {top:tblTop,bottom:SNG,left:tblL,right:colD};
const hBM = {top:tblTop,bottom:SNG,left:colD,right:colD};
const hBR = {top:tblTop,bottom:SNG,left:colD,right:tblR};
const dBL = {top:NONE_B,bottom:SNG,left:tblL,right:colD};
const dBM = {top:NONE_B,bottom:SNG,left:colD,right:colD};
const dBR = {top:NONE_B,bottom:SNG,left:colD,right:tblR};
const lBL = {top:NONE_B,bottom:tblBot,left:tblL,right:colD};
const lBM = {top:NONE_B,bottom:tblBot,left:colD,right:colD};
const lBR = {top:NONE_B,bottom:tblBot,left:colD,right:tblR};

function leaveHdrRow() {
  const m = {top:25,bottom:25,left:80,right:80};
  return new TableRow({children:[
    new TableCell({children:[p([b('Type of leave taken:')],S0)],width:{size:COL_W[0],type:WidthType.DXA},borders:hBL,margins:m}),
    new TableCell({children:[p([b('No reduction'),t(' in commuted overtime, individual did meet overtime commitment')],S0)],width:{size:COL_W[1],type:WidthType.DXA},borders:hBM,margins:m}),
    new TableCell({children:[p([b('Reduce'),t(' commuted overtime with number of days indicated as individual was not able to meet overtime commitment')],S0)],width:{size:COL_W[2],type:WidthType.DXA},borders:hBR,margins:m}),
  ]});
}

function leaveDataRow(label, col1Val, col2Val, isLast) {
  // col1Val = days for "No reduction" column; col2Val = days for "Reduce" column (0 = show grey "days")
  const v1 = col1Val > 0 ? String(col1Val) : '';
  const v2 = col2Val > 0 ? String(col2Val) : '';
  const BL = isLast?lBL:dBL, BM = isLast?lBM:dBM, BR = isLast?lBR:dBR;
  const m = {top:16,bottom:16,left:80,right:80};
  function daysCell(val, bdr, w) {
    const dm = {top:16,bottom:16,left:80,right:80};
    return new TableCell({
      children:[p(val ? [b(val+' '),t('days')] : [t('days',{color:'AAAAAA'})],
        {...S0, alignment:AlignmentType.CENTER})],
      width:{size:w,type:WidthType.DXA}, borders:bdr, margins:dm,
    });
  }
  return new TableRow({children:[
    new TableCell({children:[p([t(label)],S0)],width:{size:COL_W[0],type:WidthType.DXA},borders:BL,margins:m}),
    daysCell(v1, BM, COL_W[1]),
    daysCell(v2, BR, COL_W[2]),
  ]});
}

function reduceRow(label, days, labelChildren) {
  // 3-column table row: col1 = leave label, col2 = bottom-border line, col3 = "days (reduce...)"
  const dayStr = days > 0 ? String(days) : '';
  const botBorder = { style:BorderStyle.SINGLE, size:4, color:'000000' }; // 0.5pt
  const noBorder  = { style:BorderStyle.NONE, size:0, color:'FFFFFF' };
  const cb = { top:noBorder, bottom:noBorder, left:noBorder, right:noBorder };
  const rowS = {before:20,after:20,line:240,lineRule:'exact'};

  const col1 = new TableCell({
    children: labelChildren || [p([t(label)], {spacing:rowS})],
    width: { size:4800, type:WidthType.DXA },
    borders: cb,
    margins: { top:0, bottom:0, left:0, right:80 },
  });

  const col2 = new TableCell({
    children: [new Paragraph({
      children: [t(dayStr)],
      border: { bottom: botBorder },
      spacing: rowS,
      alignment: AlignmentType.CENTER,
    })],
    width: { size:900, type:WidthType.DXA },
    borders: cb,
    margins: { top:0, bottom:0, left:40, right:40 },
    verticalAlign: VerticalAlign.BOTTOM,
  });

  const col3 = new TableCell({
    children: [p([t('days ('), b('reduce'), t(' commuted overtime with number of days indicated)')], {spacing:rowS})],
    width: { size:5100, type:WidthType.DXA },
    borders: cb,
    margins: { top:0, bottom:0, left:40, right:0 },
  });

  return new TableRow({ children: [col1, col2, col3] });
}

// ── Document ─────────────────────────────────────────────────────────────────

// ── Annexure C Generator ─────────────────────────────────────────────────────
async function generateAnnexureCDocx(d) {
  const fullName = ((d.firstName||'') + ' ' + (d.surname||'')).trim();
  const persal = d.persal || '';
  const rank = d.designation || '';
  const month = d.month;
  const year = d.year;
  const sigDate = d.signatureDate || '';
  const lc = (function() {
    const LEAVE_MAP_ANN = {
      'Leave - Annual':'vacation','Leave - Sick':'sick',
      'Leave - Family Responsibility':'family','Leave - Study':'study',
      'Leave - Prenatal':'prenatal','Leave - Paternity':'paternity',
      'Leave - Special':'special','Leave - Maternity':'maternity',
      'Workshop':'official','Course':'official','Conference':'official',
    };
    const c = {};
    if (d.editedShifts) {
      for (const es of Object.values(d.editedShifts)) {
        const t = LEAVE_MAP_ANN[es.typeLabel];
        if (t) c[t] = (c[t]||0) + 1;
      }
    }
    return c;
  })();
  const hasLeave = Object.keys(lc).length > 0;

const annexureDoc = new Document({sections:[{properties:{
  page:{
    size:{width:convertInchesToTwip(8.27),height:convertInchesToTwip(11.69)},
    margin:{top:320,bottom:320,left:720,right:720},
  },
},children:[

  // Small top spacer so "Annexure C" header clears the margin
  new Paragraph({children:[t('')], spacing:{before:0,after:60,line:240,lineRule:'exact'}}),

  // Title + "Annexure C" on same line
  new Table({
    width:{size:pageWidth,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B},
    rows:[new TableRow({children:[
      new TableCell({
        children:[p([u('VERIFICATION OF COMMUTED OVERTIME WORKED FOR THE PERIOD',{bold:true,size:FS})],{alignment:AlignmentType.CENTER,spacing:S0})],
        width:{size:8000,type:WidthType.DXA},borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},
        margins:{top:0,bottom:0,left:0,right:0},verticalAlign:VerticalAlign.CENTER,
      }),
      new TableCell({
        children:[p([b('Annexure C',{size:32})],{alignment:AlignmentType.RIGHT,spacing:{...S0,before:80}})],
        width:{size:2800,type:WidthType.DXA},borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},
        margins:{top:0,bottom:0,left:0,right:0},verticalAlign:VerticalAlign.CENTER,
      }),
    ]})]
  }),

  blank(30),
  p([b('Date:  '),u(MONTH_NAMES[month].toUpperCase()+' '+String(year||''))],{alignment:AlignmentType.CENTER,spacing:{before:0,after:50,line:240,lineRule:'exact'}}),

  // PART A
  p([b('PART A',{underline:{type:UnderlineType.SINGLE}})],{spacing:{before:40,after:20,line:240,lineRule:'exact'}}),
  p([it('Particulars of participant in the commuted overtime system.')],{spacing:{before:0,after:60,line:240,lineRule:'exact'}}),

  new Table({
    width:{size:pageWidth,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B},
    rows:[
      new TableRow({children:[
        new TableCell({children:[new Paragraph({children:[b('Name: '),t(fullName||'')],border:{bottom:SNG},spacing:{before:0,after:30,line:240,lineRule:'exact'}})],
          width:{size:6800,type:WidthType.DXA},borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},margins:{top:0,bottom:0,left:0,right:200}}),
        new TableCell({children:[new Paragraph({children:[b('Persal No:  '),t(persal||'')],border:{bottom:SNG},spacing:{before:0,after:30,line:240,lineRule:'exact'}})],
          width:{size:4000,type:WidthType.DXA},borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},margins:{top:0,bottom:0,left:0,right:0}}),
      ]}),
      new TableRow({children:[
        new TableCell({children:[new Paragraph({children:[b('Rank: '),t(rank||'')],border:{bottom:SNG},spacing:{before:0,after:30,line:240,lineRule:'exact'}})],
          width:{size:6800,type:WidthType.DXA},borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},margins:{top:0,bottom:0,left:0,right:200}}),
        new TableCell({children:[p([b('Department: '),u('HEALTH',{bold:true})])],
          width:{size:4000,type:WidthType.DXA},borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},margins:{top:0,bottom:0,left:0,right:0}}),
      ]}),
      new TableRow({children:[
        new TableCell({children:[p([b('Institution: '),u('VICTORIA  HOSPITAL',{bold:true})])],
          width:{size:pageWidth,type:WidthType.DXA},columnSpan:2,
          borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},margins:{top:0,bottom:0,left:0,right:0}}),
      ]}),
    ]
  }),

  // PART B
  blank(50),
  p([b('PART B',{underline:{type:UnderlineType.SINGLE}})],{spacing:{before:0,after:30,line:240,lineRule:'exact'}}),
  p([b('For completion by the Head of Clinical Department/Institutional Head: ',{italics:true})],{spacing:{before:0,after:25,line:240,lineRule:'exact'}}),

  p([t('1.\t'),t('I hereby certify that '),b('the above-named employee has performed the number of hours overtime ('),
     b('Group 3.2',{size:FS+2}),b(') as agreed to his/her commuted'),
     t(' overtime contract as well as that reflected in the '),b('duty roster'),t(' for this particular month.')],
    {spacing:{before:0,after:18,line:240,lineRule:'exact'},tabStops:[{type:TabStopType.LEFT,position:400}]}),

  p([t('2.\t'),b('the above-named employee has performed the required number of working hours in this particular month (i.e.: 40 hours per week).')],
    {spacing:{before:0,after:18,line:240,lineRule:'exact'},tabStops:[{type:TabStopType.LEFT,position:400}]}),

  p([t('3.\tDuring this particular period the following leave/no leave has been utilized by the employee in question:')],
    {spacing:{before:0,after:20,line:240,lineRule:'exact'},tabStops:[{type:TabStopType.LEFT,position:400}]}),

  // FIX 1: "No leave has been utilized" + checkbox — inline left-aligned
  // Use a single paragraph with the text, then a small inline table for the box
  new Table({
    width:{size:pageWidth,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B},
    rows:[new TableRow({children:[
      new TableCell({
        children:[p([b('No leave has been utilized')],S0)],
        width:{size:3400,type:WidthType.DXA},
        borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},
        margins:{top:0,bottom:0,left:0,right:100},
        verticalAlign:VerticalAlign.CENTER,
      }),
      new TableCell({
        children:[p([t(hasLeave?'':'✓',{bold:true})],{...S0,alignment:AlignmentType.CENTER})],
        width:{size:360,type:WidthType.DXA},
        borders:allSng,
        margins:{top:30,bottom:30,left:0,right:0},
        verticalAlign:VerticalAlign.CENTER,
      }),
      new TableCell({
        children:[p([t('')],S0)],
        width:{size:7040,type:WidthType.DXA},
        borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},
      }),
    ]})]
  }),

  blank(30),

  // Leave table
  new Table({
    width:{size:pageWidth,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B},
    rows:[
      leaveHdrRow(),
      leaveDataRow('Vacation leave',lc.vacation||0,0),
      leaveDataRow('Sick leave',lc.sick||0,0),
      leaveDataRow('Family responsibility leave',lc.family||0,0),
      leaveDataRow('Special leave for study purposes (Prep & exams only)',lc.study||0,0),
      leaveDataRow('Pre-natal leave',lc.prenatal||0,0,true),
    ]
  }),

  blank(30),

  // Reducible leave — 3-column table (no borders, no headers)
  new Table({
    width: { size: pageWidth, type: WidthType.DXA },
    columnWidths: [4800, 900, 5100],
    borders: { top:NONE_B, bottom:NONE_B, left:NONE_B, right:NONE_B, insideH:NONE_B, insideV:NONE_B },
    rows: [
      reduceRow('Special leave (other)', lc.special||0),
      reduceRow('Sabbatical leave', 0),
      reduceRow('Maternity leave', lc.maternity||0),
      reduceRow('Leave without pay', 0),
      reduceRow('Suspended from duty', 0),
      reduceRow('', lc.official||0,
        [p([t('Official courses/symposia/congresses (as well as act as examiners) '),
            b('in excess of 10 days per annum')],
           {spacing:{before:20,after:20,line:240,lineRule:'exact'}})]),
    ],
  }),

  blank(60),

  // HOD Signature
  new Table({
    width:{size:pageWidth,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B},
    rows:[new TableRow({children:[
      new TableCell({
        children:[
          botLine(),
          p([b('HEAD OF CLINICAL DEPARTMENT/'),new TextRun({text:'INSTITUTIONAL HEAD',font:'Calibri',size:FS,bold:true,break:1})])
        ],
        width:{size:6200,type:WidthType.DXA},borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},
        margins:{top:0,bottom:0,left:0,right:400},
      }),
      new TableCell({
        children:[botLine(),p([b('DATE')])],
        width:{size:4600,type:WidthType.DXA},borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},
        margins:{top:0,bottom:0,left:0,right:0},
      }),
    ]})]
  }),

  blank(25),
  p([t('This document must be signed by the Head of Clinical Department or Institutional Head and not the Employee',{bold:true,italics:true})],
    {alignment:AlignmentType.CENTER,spacing:{before:0,after:50,line:240,lineRule:'exact'}}),

  p(['='.repeat(98)],{spacing:{before:30,after:20,line:240,lineRule:'exact'}}),
  p([b('For Personnel Office Only',{underline:{type:UnderlineType.SINGLE}})],{spacing:{before:0,after:25,line:240,lineRule:'exact'}}),

  // FIX 3: PERSAL verification — "(Name and Signature)" and "Date" aligned under their lines
  // Layout: [prefix text][__Name+Sig line__][__Date line__]
  //         [            ](Name and Signature)    (Date)
  new Table({
    width:{size:pageWidth,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B},
    rows:[
      new TableRow({children:[
        nb(5200,[p([t('1.  Verification form checked and verified with PERSAL records : ',{size:FS})],S0)],{margins:{top:0,bottom:0,left:0,right:0}}),
        nb(3400,[botLine()],{margins:{top:0,bottom:0,left:0,right:160}}),
        nb(2200,[botLine()],{margins:{top:0,bottom:0,left:0,right:0}}),
      ]}),
      new TableRow({children:[
        nb(5200,[p([t('')],S0)]),
        nb(3400,[p([t('(Name and Signature)',{size:FSS})],{...S0,alignment:AlignmentType.CENTER})],{margins:{top:0,bottom:0,left:0,right:160}}),
        nb(2200,[p([t('Date',{size:FSS})],{...S0,alignment:AlignmentType.CENTER})],{margins:{top:0,bottom:0,left:0,right:0}}),
      ]}),
    ]
  }),

  blank(12),

  // FIX 4: Must COT be reduced — left-aligned YES/NO boxes
  new Table({
    width:{size:pageWidth,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B},
    rows:[new TableRow({children:[
      nb(3200,[p([t('Must COT be reduced:',{size:FS})],S0)],{margins:{top:0,bottom:0,left:0,right:80},verticalAlign:VerticalAlign.CENTER}),
      new TableCell({children:[p([b('YES')],{...S0,alignment:AlignmentType.CENTER})],
        width:{size:700,type:WidthType.DXA},borders:allSng,margins:{top:25,bottom:25,left:0,right:0},verticalAlign:VerticalAlign.CENTER}),
      nb(160,[p([t('')],S0)]),
      new TableCell({children:[p([b('NO')],{...S0,alignment:AlignmentType.CENTER})],
        width:{size:700,type:WidthType.DXA},borders:allSng,margins:{top:25,bottom:25,left:0,right:0},verticalAlign:VerticalAlign.CENTER}),
      nb(5040,[p([t('')],S0)]),
    ]})]
  }),

  blank(30),
  p([t('2.  Calculation of commuted overtime for the above leave to be deducted from salary:')],
    {spacing:{before:0,after:25,line:240,lineRule:'exact'}}),
  p([t('\u2666Formula \u2013 ((Notch x 7) \u00f7 (365 x 40) x 4 \u00f7 3) x hours x 52) \u00f7 12 x (* days worked \u00f7 Calendar days in month) '),
     b('* fraction is very important as it impacts on the calculation',{size:FSS}),
     t(' e.g. 18 days in March should be 18/31 = .58.',{size:FSS})],
    {spacing:{before:0,after:30,line:240,lineRule:'exact'}}),

  // FIX 5: Calculations table — Date: labels immediately next to their lines, all on page
  // Row 1: Received | ♦Should have received | Deduct from salary — each as underlined label+value
  new Table({
    width:{size:pageWidth,type:WidthType.DXA},
    borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B},
    rows:[
      // Top row: Received | Should have received | Deduct from salary (merged across Date+value cols)
      new TableRow({children:[
        nb(2200,[new Paragraph({children:[t('Received: R',{size:FS})],border:{bottom:SNG},spacing:{before:0,after:16,line:240,lineRule:'exact'}})],{margins:{top:0,bottom:0,left:0,right:120}}),
        nb(4800,[new Paragraph({children:[t('\u2666Should have received: R',{size:FS})],border:{bottom:SNG},spacing:{before:0,after:16,line:240,lineRule:'exact'}})],{margins:{top:0,bottom:0,left:0,right:120}}),
        new TableCell({
          children:[new Paragraph({children:[t('Deduct from salary: R',{size:FS})],border:{bottom:SNG},spacing:{before:0,after:16,line:240,lineRule:'exact'}})],
          width:{size:3800,type:WidthType.DXA},
          columnSpan:2,
          borders:{top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B},
          margins:{top:0,bottom:0,left:0,right:0},
        }),
      ]}),
      // label | ___line___ | Date: | __line_with_bottom_border__
      ...['Calculated by:','Calculation verified by:','Captured on Persal by:','Approved on Persal by:'].map(label =>
        new TableRow({children:[
          nb(2200,[p([t(label,{size:FS})],S0)],{margins:{top:0,bottom:0,left:0,right:80}}),
          nb(4800,[new Paragraph({children:[t('')],border:{bottom:SNG},spacing:{before:0,after:16,line:240,lineRule:'exact'}})],{margins:{top:0,bottom:0,left:0,right:40}}),
          nb(400,[p([t('Date:',{size:FS})],{...S0,alignment:AlignmentType.LEFT})],{margins:{top:0,bottom:0,left:0,right:4}}),
          new TableCell({
            children:[new Paragraph({children:[t('')],border:{bottom:SNG},spacing:{before:0,after:16,line:240,lineRule:'exact'}})],
            width:{size:3400,type:WidthType.DXA},
            borders:{top:NONE_B,bottom:{style:BorderStyle.SINGLE,size:4,color:'000000'},left:{style:BorderStyle.SINGLE,size:4,color:'000000'},right:NONE_B},
            margins:{top:0,bottom:0,left:0,right:0},
          }),
        ]})
      ),
    ]
  }),

  blank(30),
  p([t("Make sure that all leave, as indicated above, has been captured on PERSAL and Z1 (a)'s are on file.",{bold:true,italics:true,size:FSS})],
    {spacing:{before:0,after:0,line:240,lineRule:'exact'}}),

]}]});

  return await Packer.toBlob(annexureDoc);
}


// ── Z1(a) Leave Form Generator ───────────────────────────────────────────────
async function generateZ1ADocx(d) {
  // Z1a font/spacing — independent of Annexure C
  const FS = 16, FSS = 16;  // 8pt for all table text (half-points)
  const S0 = { before:0, after:0, line:240, lineRule:'exact' };  // 1.0 line spacing
  const firstName = d.firstName || '';
  const surname = d.surname || '';
  const persal = d.persal || '';
  const sigDate = d.signatureDate || '';
  const month = d.month;
  const year = d.year;
  const shiftWorker = d.shiftWorker || 'yes';
  const casualEmployee = d.casualEmployee || 'no';
  const addressDuringLeave = d.addressDuringLeave || '';
  const supervisorName = d.supervisorName || '';
  const leaveData = (function() {
    const LEAVE_MAP_Z1 = {
      'Leave - Annual':'Annual Leave','Leave - Sick':'Normal Sick Leave',
      'Leave - Family Responsibility':'Family Responsibility Leave',
      'Leave - Study':'Special Leave (Study purposes)',
      'Leave - Prenatal':'Pre-natal Leave','Leave - Paternity':'Paternity Leave',
      'Leave - Special':'Special Leave','Leave - Maternity':'Maternity Leave',
    };
    if (!d.editedShifts) return [];
    const days = Object.keys(d.editedShifts).map(Number).sort((a,b)=>a-b);
    const groups = {};
    for (const day of days) {
      const es = d.editedShifts[day];
      if (!es.typeLabel || es.typeLabel.startsWith('WD') || es.typeLabel.startsWith('WE') || es.typeLabel.startsWith('On Call') || es.typeLabel === 'Normal Hours - Weekday') continue;
      const lbl = LEAVE_MAP_Z1[es.typeLabel] || es.typeLabel;
      if (!groups[lbl]) groups[lbl] = { type: es.typeLabel, startDay: day, endDay: day, count: 0 };
      groups[lbl].endDay = day;
      groups[lbl].count++;
    }
    return Object.values(groups).map(g => ({
      type: g.type,
      startDate: String(g.startDay).padStart(2,'0') + '/' + String(month+1).padStart(2,'0') + '/' + year,
      endDate: String(g.endDay).padStart(2,'0') + '/' + String(month+1).padStart(2,'0') + '/' + year,
      count: g.count,
    }));
  })();

        const leaveMap = {};
  const LEAVE_LABELS_Z1 = {
    'Leave - Annual': 'Annual Leave',
    'Leave - Sick': 'Normal Sick Leave (Provide supporting evidence when applicable)',
    'Leave - Family Responsibility': 'Family Responsibility Leave (Provide supporting evidence)',
    'Leave - Study': 'Special Leave for study purposes (Prep &amp; exams only)',
    'Leave - Prenatal': 'Pre-natal Leave (Provide supporting evidence)',
    'Leave - Paternity': 'Paternity Leave (Provide supporting evidence)',
    'Leave - Special': 'Special Leave ((Provide supporting evidence)',
    'Leave - Maternity': 'Maternity Leave (Provide supporting evidence))',
  };
  for (const ld of leaveData) {
    const label = LEAVE_LABELS_Z1[ld.type];
    if (label) leaveMap[label] = ld;
  }

const fullSurname = surname || '';
const initials = firstName ? firstName.trim().split(/\s+/).map(w=>w[0]+'.').join('') : '';
const addressText = addressDuringLeave || '';

// Local overrides: Z1(a) text at 8pt Calibri, line spacing 1.15
function t(txt, opts={}) { return new TextRun({text:String(txt||''),font:'Calibri',size:FS,...opts}); }
function b(txt, opts={}) { return t(txt, {bold:true,...opts}); }
function p(ch, opts={}) { return new Paragraph({children:Array.isArray(ch)?ch:[ch],spacing:S0,...opts}); }
function cell(children, width, opts={}) {
  return new TableCell({
    children: Array.isArray(children) ? children : [p(children)],
    width: { size:width, type:WidthType.DXA },
    borders: opts.borders || allSng,
    ...(opts.gridSpan ? { columnSpan:opts.gridSpan } : {}),
    ...(opts.shading ? { shading:opts.shading } : {}),
    verticalAlign: opts.valign || VerticalAlign.CENTER,
    margins: opts.margins || { top:4, bottom:4, left:50, right:50 },
  });
}

function hdrCell(text, width, gridSpan) {
  return cell([p([b(text,{size:20})],S0)], width, {
    gridSpan, borders:allSng,
    shading:{ type:ShadingType.CLEAR, fill:'D9D9D9' },
    margins:{ top:3, bottom:3, left:40, right:40 },
  });
}

function sectionRow(text) {
  return new TableRow({ children:[
    cell([p([b(text.toUpperCase(),{size:20})],S0)], 10512, {
      gridSpan:27, borders:allSng,
      shading:{ type:ShadingType.CLEAR, fill:'D9D9D9' },
      margins:{ top:18, bottom:18, left:60, right:60 },
    })
  ]});
}

function leaveRow4(label, leaveMap) {
  const m = {top:3,bottom:3,left:40,right:40};
  const ld = leaveMap[label];
  const shade = ld ? { type:ShadingType.CLEAR, fill:'DEEAF1' } : undefined;
  return new TableRow({ children:[
    cell([p([b(label)],S0)], 4785, { gridSpan:12, borders:allSng, shading:shade, margins:m }),
    cell([p([t(ld?ld.startDate:'')],{...S0,alignment:AlignmentType.CENTER})], 1408, { gridSpan:4, borders:allSng, shading:shade }),
    cell([p([t(ld?ld.endDate:'')],{...S0,alignment:AlignmentType.CENTER})], 1610, { gridSpan:4, borders:allSng, shading:shade }),
    cell([p([t(ld?String(ld.count):'')],{...S0,alignment:AlignmentType.CENTER})], 2709, { gridSpan:7, borders:allSng, shading:shade }),
  ]});
}

function smallRow(label, fullWidth) {
  const m = {top:3,bottom:3,left:40,right:40};
  return new TableRow({ children:[
    cell([p([b(label)],S0)], fullWidth||4785, { gridSpan:fullWidth?27:12, borders:allSng, margins:m }),
    ...(fullWidth ? [] : [cell([p([t('')],S0)], 5727, { gridSpan:15, borders:allSng })]),
  ]});
}

function sigLine(value='') {
  return new Paragraph({
    children:[t(value,{size:FSS})],
    border:{ bottom:{ style:BorderStyle.SINGLE, size:4, color:'000000' } },
    spacing:{ before:120, after:60, line:276, lineRule:'exact' },
  });
}

// Helper: 2-column sig+date layout  
function sigRow2Col(sigValue, sigLabel, dateValue) {
  const FULL = 10512;
  const SIG_W = Math.round(FULL * 0.68);
  const DAT_W = FULL - SIG_W;
  const bBot = { bottom:{ style:BorderStyle.SINGLE, size:4, color:'000000' } };
  const bNone = { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} };
  const mL = {top:0,bottom:0,left:0,right:80};
  const mR = {top:0,bottom:0,left:80,right:0};
  return [
    new Table({
      width:{size:FULL,type:WidthType.DXA},
      borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE}},
      rows:[
        new TableRow({ children:[
          new TableCell({
            children:[new Paragraph({children:[t(sigValue,{size:FSS})],border:bBot,spacing:{before:80,after:0,line:240,lineRule:'exact'}})],
            width:{size:SIG_W,type:WidthType.DXA},borders:bNone,margins:mL,
          }),
          new TableCell({
            children:[new Paragraph({children:[t(dateValue,{size:FSS})],border:bBot,spacing:{before:80,after:0,line:240,lineRule:'exact'}})],
            width:{size:DAT_W,type:WidthType.DXA},borders:bNone,margins:mR,
          }),
        ]}),
        new TableRow({ children:[
          new TableCell({
            children:[p([b(sigLabel,{size:FSS})],{spacing:{before:0,after:20,line:240,lineRule:'exact'}})],
            width:{size:SIG_W,type:WidthType.DXA},borders:bNone,margins:mL,
          }),
          new TableCell({
            children:[p([b('DATE',{size:FSS})],{spacing:{before:0,after:20,line:240,lineRule:'exact'}})],
            width:{size:DAT_W,type:WidthType.DXA},borders:bNone,margins:mR,
          }),
        ]}),
      ],
    }),
  ];
}


function calRow(label, note) {
  const m = {top:3,bottom:3,left:40,right:40};
  return new TableRow({ children:[
    cell([p([b(label)],S0)], 4785, { gridSpan:12, borders:allSng, margins:m }),
    cell([p([t('')],S0)], 1408, { gridSpan:4, borders:allSng }),
    cell([p([t('')],S0)], 1610, { gridSpan:4, borders:allSng }),
    cell([p([b(note)],{...S0,alignment:AlignmentType.CENTER})], 1958, { gridSpan:6, borders:allSng, margins:m }),
    cell([p([t('')],S0)], 751, { borders:allSng }),
  ]});
}

function secBRow(label, isFirst) {
  const m = {top:3,bottom:3,left:40,right:40};
  return new TableRow({ children:[
    cell([p([b(label)],S0)], 4378, { gridSpan:10, borders:allSng, margins:m }),
    cell([p([t('')],S0)], 1239, { gridSpan:5, borders:allSng }),
    cell([p([t('')],S0)], 1109, { gridSpan:2, borders:allSng }),
    cell([p([t('')],S0)], 1077, { gridSpan:3, borders:allSng }),
    cell([p([t(isFirst?'H':'h')],{...S0,alignment:AlignmentType.CENTER})], 1124, { gridSpan:3, borders:allSng }),
    cell([p([t(isFirst?'M':'m')],{...S0,alignment:AlignmentType.CENTER})], 1585, { gridSpan:4, borders:allSng }),
  ]});
}

// ── Right-side block: Shift Worker row with matching column layout as Casual Employee ──
// Both rows use: [label:gs6][Yes:gs2][box:gs3][No][box]
// Total right = 5090 twips: label=2381(gs6) Yes=851(gs2) box=585(gs3) No=522 box=751
// FIX 3: make Shift Worker right side identical to Casual Employee right side
// Left side: PERSAL row = 5422 twips (same as Casual Employee left)

// The PERSAL row left part needs to split differently from Surname row (which uses 1092+4330=5422)
// We use: [PERSAL label:1092][PERSAL value:4330=gs13] then right side
// The right side for both rows must be IDENTICAL:
//   [label:2381=gs6][Yes:851=gs2][box:585=gs3][No:522][box:751]
// Total right: 2381+851+585+522+751 = 5090 ✓
// Total left: 1092+4330 = 5422 ✓
// Grand total: 5422+5090 = 10512 ✓

const RIGHT_YN = (label, value) => {
  const isYes = (value||'').toLowerCase() === 'yes';
  return [
    cell([p([b(label)],S0)], 2381, {gridSpan:6,borders:allSng,margins:{top:4,bottom:4,left:50,right:50}}),
    cell([p([b('Yes')],S0)], 851, {gridSpan:2,borders:allSng,margins:{top:4,bottom:4,left:50,right:50}}),
    cell([p([t(isYes?'✓':'')],{...S0,alignment:AlignmentType.CENTER})], 585, {gridSpan:3,borders:allSng}),
    cell([p([b('No')],S0)], 522, {borders:allSng,margins:{top:4,bottom:4,left:50,right:50}}),
    cell([p([t(isYes?'':'✓')],{...S0,alignment:AlignmentType.CENTER})], 751, {borders:allSng}),
  ];
};

const doc = new Document({ sections:[{ properties:{
  page:{
    size:{ width:convertInchesToTwip(8.27), height:convertInchesToTwip(11.69) },
    margin:{ top:240, bottom:240, left:360, right:360 },
  }
}, children:[

  new Paragraph({ children:[b('Z1 (a)',{size:20})], alignment:AlignmentType.RIGHT, spacing:{before:0,after:14,line:240,lineRule:'exact'} }),
  new Paragraph({ children:[b('APPLICATION FOR LEAVE OF ABSENCE',{size:20})], alignment:AlignmentType.CENTER, spacing:{before:0,after:0,line:240,lineRule:'exact'} }),
  new Paragraph({ children:[t('',{size:16})], spacing:{before:0,after:0,line:240,lineRule:'exact'} }),

  new Table({
    width:{ size:10512, type:WidthType.DXA },
    alignment: AlignmentType.CENTER,
    borders:{ top:NONE_B,bottom:NONE_B,left:NONE_B,right:NONE_B,insideH:NONE_B,insideV:NONE_B },
    rows:[

      // Surname | Initials
      new TableRow({ children:[
        cell([p([b('Surname')])], 1092, {borders:allSng}),
        cell([p([t(fullSurname)])], 4330, {gridSpan:13,borders:allSng}),
        cell([p([b('Initials:')])], 1411, {gridSpan:4,borders:allSng}),
        cell([p([t(initials)])], 3679, {gridSpan:9,borders:allSng}),
      ]}),

      // PERSAL + Shift Worker — FIX 3: right columns match Casual Employee exactly
      new TableRow({ children:[
        // PERSAL label + value in one merged cell (avoids label wrapping)
        cell([p([b('PERSAL Number:  '),t(persal||'')],S0)], 5422, {gridSpan:14,borders:allSng}),
        ...RIGHT_YN('Shift Worker', shiftWorker),
      ]}),

      // Casual Employee — same right layout as Shift Worker above
      new TableRow({ children:[
        // Left: address during leave occupies the blank block
        cell([
          p([b('Address during leave: ',{size:FSS}),t(addressText||'',{size:FSS})],{...S0,lineRule:'exact'}),
        ], 5422, {
          gridSpan:14,
          borders:{top:NONE_B,bottom:NONE_B,left:SNG,right:NONE_B},
          margins:{top:2,bottom:2,left:40,right:40},
        }),
        ...RIGHT_YN('Casual Employee', casualEmployee),
      ]}),

      // Department
      new TableRow({ children:[
        cell([p([t('')],S0)], 5422, {gridSpan:14,borders:{top:NONE_B,bottom:NONE_B,left:SNG,right:NONE_B}}),
        cell([p([b('Department')])], 5090, {gridSpan:13,borders:allSng}),
      ]}),
      new TableRow({ children:[
        cell([p([t('')],S0)], 5422, {gridSpan:14,borders:{top:NONE_B,bottom:NONE_B,left:SNG,right:NONE_B}}),
        cell([p([t('Western Cape Department of Health and Wellness',{size:FSS})])], 5090, {gridSpan:13,borders:allSng}),
      ]}),

      // Component
      new TableRow({ children:[
        cell([p([t('')],S0)], 5422, {gridSpan:14,borders:{top:NONE_B,bottom:NONE_B,left:SNG,right:NONE_B}}),
        cell([p([b('Component')])], 5090, {gridSpan:13,borders:allSng}),
      ]}),
      new TableRow({ children:[
        cell([p([t('')],S0)], 5422, {gridSpan:14,borders:{top:NONE_B,bottom:{style:BorderStyle.SINGLE,size:4,color:'000000'},left:SNG,right:NONE_B}}),
        cell([p([t('Emergency Medicine \u2014 Victoria Hospital',{size:FSS})])], 5090, {gridSpan:13,borders:allSng}),
      ]}),

      sectionRow('SECTION A: For Periods covering a full day'),

      new TableRow({ children:[
        hdrCell('Type of Leave Taken as Working Days',4785,12),
        hdrCell('Start Date',1408,4),
        hdrCell('End Date',1610,4),
        hdrCell('Number of Working Days',2709,7),
      ]}),

      leaveRow4('Annual Leave', leaveMap),
      leaveRow4('Normal Sick Leave (Provide supporting evidence when applicable)', leaveMap),

      new TableRow({ children:[
        cell([p([b('Temporary Incapacity Leave')],S0)], 4785, {gridSpan:12,borders:allSng,margins:{top:3,bottom:3,left:40,right:40}}),
        cell([p([t('Temporary incapacity leave must be applied for on the application form prescribed in terms of the Policy and Procedure on Incapacity Leave and Ill-health Retirement for Public Service Employees.',{size:14,italics:true})],{spacing:{before:0,after:0,line:240,lineRule:'exact'}})], 5727, {gridSpan:15,borders:allSng,margins:{top:3,bottom:3,left:40,right:40}}),
      ]}),

      leaveRow4('Leave for Occupational Injuries and Diseases', leaveMap),
      leaveRow4('Adoption Leave (Provide supporting evidence)', leaveMap),
      leaveRow4('Family Responsibility Leave (Provide supporting evidence)', leaveMap),
      leaveRow4('Pre-natal Leave (Provide supporting evidence)', leaveMap),
      leaveRow4('Paternity Leave (Provide supporting evidence)', leaveMap),
      leaveRow4('Special Leave ((Provide supporting evidence)', leaveMap),
      smallRow('Specify Type of Special Leave'),
      leaveRow4('Leave for Union Office Bearers (Provide supporting evidence)', leaveMap),
      leaveRow4('Leave for Union Shop Stewards (Provide supporting evidence)', leaveMap),
      smallRow('Specify Union Affiliation'),

      new TableRow({ children:[
        hdrCell('Type of Leave Taken as Calendar Days/Months/Weeks',4785,12),
        hdrCell('Start Date',1408,4),
        hdrCell('End Date',1610,4),
        hdrCell('Number of Calendar Days',2709,7),
      ]}),
      leaveRow4('Unpaid Leave (Provide motivation)', leaveMap),
      calRow('Maternity Leave (Provide supporting evidence))','No. of Calendar Months'),
      calRow('Surrogacy Leave: Committing Parent (Provide supporting evidence)','No. of Calendar Months'),
      calRow('Surrogacy Leave: Surrogate mother (Provide supporting evidence)','No of weeks'),

      sectionRow('SECTION B: For periods covering parts of a day or fractions'),

      new TableRow({ children:[
        hdrCell('Type of Leave Taken as Working Days',4378,10),
        hdrCell('Date',1239,5),
        hdrCell('Start Time',1109,2),
        hdrCell('End Time',1077,3),
        hdrCell('Number of Hours/ Minutes',2709,7),
      ]}),

      secBRow('Annual Leave',true),
      secBRow('Normal Sick Leave',false),
      secBRow('Family Responsibility Leave (Provide supporting evidence)',false),
      secBRow('Pre-natal Leave (Provide supporting evidence)',false),
      secBRow('Paternity Leave (Provide supporting evidence)',false),
      secBRow('Special Leave',false),
      new TableRow({ children:[
        cell([p([b('Specify Type of Special Leave')],S0)], 4378, {gridSpan:10,borders:allSng,margins:{top:3,bottom:3,left:40,right:40}}),
        cell([p([t('')],S0)], 6134, {gridSpan:17,borders:allSng}),
      ]}),
      secBRow('Leave for Union Office Bearers (Provide supporting evidence)',false),
      secBRow('Leave for Union Shop Stewards (Provide supporting evidence)',false),
      new TableRow({ children:[
        cell([p([b('Specify Union Affiliation')],S0)], 4378, {gridSpan:10,borders:allSng,margins:{top:3,bottom:3,left:40,right:40}}),
        cell([p([t('')],S0)], 6134, {gridSpan:17,borders:allSng}),
      ]}),

      // Certification + Employee Sig
      new TableRow({ children:[
        cell([
          p([t('I hereby certify that I have acquainted myself of my available leave credits and with the rules governing the leave I have applied for. Further, I am certifying that the information provided is correct. Any falsification of information in this regard may form ground for disciplinary action. Furthermore, I fully understand that if I do not have sufficient leave credits from my previous or current leave cycle to cover for my application, my capped leave as at 30 June 2000 will be automatically utilised.',{size:14,italics:true})],{spacing:{before:0,after:0,line:240,lineRule:'exact'}}),
          ...sigRow2Col((firstName||'')+' '+(surname||''), 'EMPLOYEE SIGNATURE', sigDate||''),
        ], 10512, {gridSpan:27,borders:allSng,margins:{top:3,bottom:3,left:55,right:55}}),
      ]}),

      // Recommendation
      new TableRow({ children:[
        cell([p([b('Recommendation by Supervisor/Manager (Mark with X)')],{...S0,alignment:AlignmentType.CENTER})], 10512, {gridSpan:27,shading:{type:ShadingType.CLEAR,fill:'D9D9D9'},borders:allSng,margins:{top:4,bottom:4,left:40,right:40}}),
      ]}),
      new TableRow({ children:[
        cell([p([b('Recommended')])],1742,{gridSpan:2,borders:allSng}),
        cell([p([t('')])],1277,{gridSpan:4,borders:allSng}),
        cell([p([b('Not Recommended')])],3174,{gridSpan:10,borders:allSng}),
        cell([p([t('')])],1287,{gridSpan:3,borders:allSng}),
        cell([p([b('Rescheduled')])],1562,{gridSpan:5,borders:allSng}),
        cell([p([t('')])],1470,{gridSpan:3,borders:allSng}),
      ]}),
      new TableRow({ children:[
        cell([
          p([t('REMARKS (If not recommended please state the reasons & the dates in the case of rescheduling):',{size:14,italics:true})],{spacing:{before:0,after:0,line:240,lineRule:'exact'}}),
          ...sigRow2Col(supervisorName, "MANAGER'S/SUPERVISOR'S SIGNATURE", ''),
        ], 10512, {gridSpan:27,borders:allSng,margins:{top:3,bottom:3,left:55,right:55}}),
      ]}),

      // HOD Approval
      new TableRow({ children:[
        cell([p([b('Approval by Head of Department (Mark with X)')],{...S0,alignment:AlignmentType.CENTER})], 10512, {gridSpan:27,shading:{type:ShadingType.CLEAR,fill:'D9D9D9'},borders:allSng,margins:{top:4,bottom:4,left:40,right:40}}),
      ]}),
      new TableRow({ children:[
        cell([p([b('Approved With Full Pay')])],4500,{gridSpan:11,borders:allSng}),
        cell([p([t('')])],450,{gridSpan:2,borders:allSng}),
        cell([p([b('Approved Without Pay')])],2853,{gridSpan:7,borders:allSng}),
        cell([p([t('')])],556,{borders:allSng}),
        cell([p([b('Not Approved')])],1402,{gridSpan:5,borders:allSng}),
        cell([p([t('')])],751,{borders:allSng}),
      ]}),
      new TableRow({ children:[
        cell([
          p([t('REMARKS (If approved with a change in condition of payment or not approved, please provide motivation):',{size:14,italics:true})],{spacing:{before:0,after:0,line:240,lineRule:'exact'}}),
          ...sigRow2Col(supervisorName, 'SIGNATURE OF HOD OR DESIGNEE', ''),
        ], 10512, {gridSpan:27,borders:allSng,margins:{top:3,bottom:3,left:55,right:55}}),
      ]}),

      // Data Capturing
      new TableRow({ children:[
        cell([p([b('Data Capturing')],{...S0,alignment:AlignmentType.CENTER})], 10512, {gridSpan:27,shading:{type:ShadingType.CLEAR,fill:'D9D9D9'},borders:allSng,margins:{top:4,bottom:4,left:40,right:40}}),
      ]}),
      new TableRow({ children:[
        cell([
          p([b('Captured By:'),t('_________________________   '),b('Captured On '),t('______________   '),b('Signature'),t('_____________________')],{spacing:{before:0,after:0,line:240,lineRule:'exact'}}),
          p([b('Checked By:'),t('__________________________   '),b('Checked On:'),t('________________   '),b('Signature'),t('_____________________')],{spacing:{before:20,after:0,line:240,lineRule:'exact'}}),
        ], 10512, {gridSpan:27,borders:allSng,margins:{top:3,bottom:3,left:55,right:55}}),
      ]}),

    ],
  }),

]}]});

  return await Packer.toBlob(doc);
}


// === UI ===


function patchCell(xml,ref,val) {
  const re=new RegExp(`<c r="${ref.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}"[^>]*>.*?<\\/c>`,'s');
  return re.test(xml)?xml.replace(re,`<c r="${ref}" t="inlineStr"><is><t>${escXml(val)}</t></is></c>`):xml;
}
function patchNumCell(xml,ref,num) {
  const re=new RegExp(`<c r="${ref.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}"[^>]*>.*?<\\/c>`,'s');
  return re.test(xml)?xml.replace(re,`<c r="${ref}"><v>${num}</v></c>`):xml;
}
function patchCellKeepStyle(xml,ref,val,bold=false) {
  const esc=ref.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  const re=new RegExp('<c r="'+esc+'"([^>/]*?)\\s*(?:/>|>[\\s\\S]*?<\\/c>)');
  if(!re.test(xml)) return xml;
  return xml.replace(re,function(match,attrs){
    const cleanAttrs=attrs.replace(/\s*t="[^"]*"/g,'');
    const inner=bold
      ? '<r><rPr><b/></rPr><t>'+escXml(val)+'</t></r>'
      : '<t>'+escXml(val)+'</t>';
    return '<c r="'+ref+'"'+cleanAttrs+' t="inlineStr"><is>'+inner+'</is></c>';
  });
}
// Safe cell replacement - does NOT cross cell boundaries (no 's' flag cross-row bug)
function replaceCell(xml,ref,newCellXml) {
  const esc=ref.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  const re=new RegExp('<c r="'+esc+'"([^>/]*?)\\s*(?:/>|>[\\s\\S]*?<\\/c>)');
  return re.test(xml) ? xml.replace(re,newCellXml) : xml;
}

