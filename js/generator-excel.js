// ═══════════════════════════════════════════════════════════════
// generator-excel.js — Excel timesheet generator.
//
//   generateExcel(details, shifts, month, year, profile)
//     → Promise<Blob>  — WCG-formatted monthly timesheet workbook
//
// Helper functions (internal):
//   b64ToUint8, escXml, parseTimeMins, minsToStr, splitShift,
//   patchStyleOnly, patchCell, patchNumCell, patchCellKeepStyle,
//   replaceCell, dateKeyLocal
//
// typeOptsFor(isWE, isPH, selectedType) — activity type dropdown
//   helper used by the preview UI row editor.
//
// Depends on: config.js, holidays.js
// ═══════════════════════════════════════════════════════════════

function b64ToUint8(b64) {
  const bin=atob(b64),arr=new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i); return arr;
}
function escXml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}

// === SHIFT SPLIT ===
function parseTimeMins(t) { const [h,m]=t.split(':').map(Number); return h*60+m; }
function minsToStr(mins) {
  mins=((mins%1440)+1440)%1440;
  return String(Math.floor(mins/60)).padStart(2,'0')+'H'+String(mins%60).padStart(2,'0');
}
function splitShift(startStr,endStr) {
  let s=parseTimeMins(startStr),e=parseTimeMins(endStr);
  if(e<=s)e+=1440; const ne=s+480;
  return e<=ne ? {nf:minsToStr(s),nt:minsToStr(e),of:null,ot:null}
               : {nf:minsToStr(s),nt:minsToStr(ne),of:minsToStr(ne),ot:minsToStr(e)};
}

function patchStyleOnly(xml,ref,styleIdx) {
  const re=new RegExp(`<c r="${ref.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}"([^>]*)>.*?<\/c>`,'s');
  if(!re.test(xml)) return xml;
  return xml.replace(re,`<c r="${ref}" s="${styleIdx}"></c>`);
}

// ── Generic template approach — works for any year/month ────────────────────
const TEMPLATE_GENERIC_SIG_START = 46; // always: 9 + 31 data rows + 6 = 46

// Style indices from template — verified against xl/styles.xml xf list (0-based)
// Template has 32 styles (s0-s31). We inject 4 more: s32, s33 (missing WE), s34 (leave-WD), s35 (leave-WE)
// Fills: fl3=#4F81BD(blue), fl4=#FFFFFF(white)
// WD = weekday (white fill), WE = weekend/PH (blue fill)
// s20: f6(bold),fl3(BLUE),b15  | s21: f6(bold),fl4(WHITE),b16
// s22: f7,fl4(WHITE),b17       | s23: f7,fl4(WHITE),b18
// s24: f7,fl4(WHITE),b15       | s25: f7,fl3(BLUE),b18
// s26: f7,fl3(BLUE),b15        | s27: f6(bold),fl3(BLUE),b18
// Injected: s32 = f6(bold),fl3(BLUE),b16 (WE col B)
//           s33 = f7,fl3(BLUE),b17 (WE col C)
// Style indices from template (0-31 native) + injected (32-36):
//   Template native styles used:
//     s20: bold, BLUE fill, border15(med-L) → WE col A
//     s21: bold, WHITE fill, border16(med-LR) → WD col B
//     s22: normal, WHITE fill, border17(thin-all) → WD inner cells (C)
//     s23: normal, WHITE fill, border18(thin-L,med-R) → WD group-end (D,F,H,J,K)
//     s24: normal, WHITE fill, border15(med-L) → WD group-start (E,G,I)
//     s25: normal, BLUE fill, border18(med-R) → WE group-end (D,F,H,J,K)
//     s26: normal, BLUE fill, border15(med-L) → WE group-start (E,G,I)
//     s27: bold, BLUE fill, border18, align=left → PH label
//   Injected styles:
//     s32: bold, WHITE fill, border15 → WD col A
//     s33: bold, BLUE fill, border16 → WE col B
//     s34: normal, BLUE fill, border17 → WE inner cells (C)
//     s35: normal, WHITE fill, border17, wrapText → WD leave label
//     s36: normal, BLUE fill, border17, wrapText → WE leave label
const COL_STYLES_WD = {A:'32',B:'21',C:'22',D:'23',E:'24',F:'23',G:'24',H:'23',I:'24',J:'23',K:'23'};
const COL_STYLES_WE = {A:'20',B:'33',C:'34',D:'25',E:'26',F:'25',G:'26',H:'25',I:'26',J:'25',K:'25'};

function sigRow(monthIdx) { return TEMPLATE_GENERIC_SIG_START; }

async function generateExcel(monthIdx, year, details) {
  const b64 = TEMPLATE_GENERIC;
  const zip = await JSZip.loadAsync(b64ToUint8(b64));
  const sheetFile = zip.file('xl/worksheets/sheet1.xml');
  if(!sheetFile) throw new Error('Sheet not found');
  let xml = await sheetFile.async('string');

  // Patch styles.xml: inject 5 new cellXf entries (s32-s36)
  const stylesFile = zip.file('xl/styles.xml');
  if(stylesFile) {
    let sxml = await stylesFile.async('string');
    if(!sxml.includes('PATCHED_STYLES')) {
      const xfCountM = sxml.match(/<cellXfs count="(\d+)"/);
      if(xfCountM) {
        const newCount = parseInt(xfCountM[1]) + 5;
        sxml = sxml.replace(`<cellXfs count="${xfCountM[1]}"`, `<cellXfs count="${newCount}"`);
        // s32: WD col A — bold, WHITE fill, med-L border
        const wdA  = `<xf numFmtId="0" fontId="6" fillId="4" borderId="15" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="bottom"/></xf>`;
        // s33: WE col B — bold, BLUE fill, med-LR border
        const weB  = `<xf numFmtId="0" fontId="6" fillId="3" borderId="16" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="bottom"/></xf>`;
        // s34: WE inner cells (C) — normal, BLUE fill, thin-all border
        const weC  = `<xf numFmtId="0" fontId="7" fillId="3" borderId="17" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="bottom"/></xf>`;
        // s35: WD leave — WHITE fill, thin-all, LEFT-aligned, no wrap
        const wdLv = `<xf numFmtId="0" fontId="7" fillId="4" borderId="17" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="bottom"/></xf>`;
        // s36: WE leave — BLUE fill, thin-all, LEFT-aligned, no wrap
        const weLv = `<xf numFmtId="0" fontId="7" fillId="3" borderId="17" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="bottom"/></xf>`;
        sxml = sxml.replace('</cellXfs>', `${wdA}${weB}${weC}${wdLv}${weLv}<!-- PATCHED_STYLES --></cellXfs>`);
      }
      zip.file('xl/styles.xml', sxml);
    }
  }

  // Patch header: month name, year, employee name, PERSAL
  xml = patchCell(xml, 'E4', MONTH_NAMES[monthIdx].toUpperCase());
  xml = patchNumCell(xml, 'G4', year);
  xml = patchCell(xml, 'C5', (details.firstName + ' ' + details.surname).trim());
  xml = patchCell(xml, 'H5', details.persal || '');

  // Compute days in month and public holidays for the actual year
  const daysInMonth = new Date(year, monthIdx + 1, 0).getDate();
  const phMap = getSAPublicHolidays(year);  // Map of dateKey → phName
  const DAY_NAMES_SHORT = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

  // ── Patch date rows 9-39 (dates 1-31) ───────────────────────────────────
  for(let d = 1; d <= 31; d++) {
    const row = 8 + d;
    if(d > daysInMonth) {
      // Hide this row: set height to 0, keep customHeight, and add hidden
      // Template rows have: ht="19.5" customHeight="1"
      // Use word-boundary-safe match to only target the 'ht' attribute (not 'customHeight')
      xml = xml.replace(
        new RegExp('(<row r="'+row+'"[^>]*\\bht=")[^"]*(")', 'g'),
        '$10$2'
      );
      // Add hidden="1" only if not already present
      if (!new RegExp('<row r="'+row+'"[^>]*hidden=').test(xml)) {
        xml = xml.replace(
          new RegExp('(<row r="'+row+'"[^>]*)>'),
          '$1 hidden="1">'
        );
      }
      continue;
    }
    const dt = new Date(year, monthIdx, d);
    const dow = dt.getDay(); // 0=Sun, 6=Sat
    const isWE = dow === 0 || dow === 6;
    const dk = `${year}-${String(monthIdx+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const phName = phMap ? phMap.get(dk) : null;
    const isSpecial = isWE || !!phName;
    const styles = isSpecial ? COL_STYLES_WE : COL_STYLES_WD;
    const dayName = DAY_NAMES_SHORT[dow];

    // Patch A: date number with correct style
    xml = replaceCell(xml, 'A'+row, '<c r="A'+row+'" s="'+styles.A+'" t="n"><v>'+d+'</v></c>');
    // Patch B: day name with correct style
    xml = replaceCell(xml, 'B'+row, '<c r="B'+row+'" s="'+styles.B+'" t="inlineStr"><is><t>'+dayName+'</t></is></c>');
    // Apply correct styles to C-K (blank fill, no value)
    for(const col of ['C','D','E','F','G','H','I','J','K']) {
      xml = replaceCell(xml, col+row, '<c r="'+col+row+'" s="'+styles[col]+'"/>');
    }
    // PH name in K
    if(phName) {
      xml = replaceCell(xml, 'K'+row, '<c r="K'+row+'" s="'+styles.K+'" t="inlineStr"><is><r><rPr><b/></rPr><t>'+escXml(phName)+'</t></r></is></c>');
    }
  }

  // ── Patch shift/leave data from editedShifts ───────────────────────────
  const es_map = details.editedShifts || {};
  for(const [dayStr, es] of Object.entries(es_map)) {
    const d = parseInt(dayStr);
    if(!d || d < 1 || d > daysInMonth) continue;
    const row = 8 + d;
    const dt = new Date(year, monthIdx, d);
    const dow = dt.getDay();
    const isWE = dow === 0 || dow === 6;
    const dk = `${year}-${String(monthIdx+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const phName = phMap ? phMap.get(dk) : null;
    const isSpecial = isWE || !!phName;
    const isShift = es.typeLabel && (es.typeLabel.startsWith('WD Shift') || es.typeLabel.startsWith('WE Shift'));
    const isConsultantType = es.typeLabel && (
      es.typeLabel.startsWith('On Call -') || es.typeLabel === 'Normal Hours - Weekday'
    );

    if(isConsultantType) {
      // Consultant mode: 6 time columns
      // C = Normal From, D = Normal To
      // E = 1st On Call From (ot1f), F = 1st On Call To (ot1t)
      // G = 2nd On Call From (ot2f), H = 2nd On Call To (ot2t)
      // Use stored times; fall back to CONSULTANT_SHIFT_TIMES defaults
      const cTimes = CONSULTANT_SHIFT_TIMES[es.typeLabel] || {};
      const nf  = es.nf  !== undefined ? es.nf  : (cTimes.nf  || '');
      const nt  = es.nt  !== undefined ? es.nt  : (cTimes.nt  || '');
      const ot1f = es.ot1f !== undefined ? es.ot1f : (cTimes.ot1f || '');
      const ot1t = es.ot1t !== undefined ? es.ot1t : (cTimes.ot1t || '');
      const ot2f = es.ot2f !== undefined ? es.ot2f : (cTimes.ot2f || '');
      const ot2t = es.ot2t !== undefined ? es.ot2t : (cTimes.ot2t || '');
      if(nf)   xml = patchCellKeepStyle(xml, `C${row}`, nf);
      if(nt)   xml = patchCellKeepStyle(xml, `D${row}`, nt);
      if(ot1f) xml = patchCellKeepStyle(xml, `E${row}`, ot1f);
      if(ot1t) xml = patchCellKeepStyle(xml, `F${row}`, ot1t);
      if(ot2f) xml = patchCellKeepStyle(xml, `G${row}`, ot2f);
      if(ot2t) xml = patchCellKeepStyle(xml, `H${row}`, ot2t);
    } else if(isShift) {
      const hasNormalTimes = es.nf && es.nf !== '00H00' && es.nt && es.nt !== '00H00';
      const hasOTTimes = es.of && es.of !== '00H00' && es.ot && es.ot !== '00H00';
      if(hasNormalTimes) {
        xml = patchCellKeepStyle(xml, `C${row}`, es.nf);
        xml = patchCellKeepStyle(xml, `D${row}`, es.nt);
      }
      if(hasOTTimes) {
        xml = patchCellKeepStyle(xml, `E${row}`, es.of);
        xml = patchCellKeepStyle(xml, `F${row}`, es.ot);
      }
    } else if(es.typeLabel) {
      // Non-shift (leave/workshop): merge C+D, bold label, white fill + wrapText
      const LEAVE_LABEL_MAP = {
        'Leave - Annual':'Annual Leave','Leave - Sick':'Sick Leave',
        'Leave - Family Responsibility':'Family Responsibility Leave',
        'Leave - Study':'Study Leave','Leave - Prenatal':'Pre-natal Leave',
        'Leave - Paternity':'Paternity Leave','Leave - Special':'Special Leave',
        'Leave - Maternity':'Maternity Leave','Workshop':'Workshop',
        'Course':'Course','Conference':'Conference / Symposium',
      };
      const leaveDisplayLabel = LEAVE_LABEL_MAP[es.typeLabel] || es.typeLabel;
      // Write label left-aligned in C only — style s35 (WD) or s36 (WE), both left-aligned
      const leaveStyleIdx = isSpecial ? '36' : '35';
      xml = replaceCell(xml, 'C'+row,
        '<c r="C'+row+'" s="'+leaveStyleIdx+'" t="inlineStr"><is><t>'+escXml(leaveDisplayLabel)+'</t></is></c>');
    }
  }

  // ── Sig row patching ────────────────────────────────────────────────────
  const sr = TEMPLATE_GENERIC_SIG_START;
  xml = patchCellKeepStyle(xml, `F${sr}`, (details.firstName+' '+details.surname).trim());
  xml = patchCellKeepStyle(xml, `J${sr}`, details.signatureDate || '');
  xml = patchCellKeepStyle(xml, `F${sr+4}`, details.supervisorName || '');
  xml = patchCellKeepStyle(xml, `J${sr+4}`, details.signatureDate || '');

  zip.file('xl/worksheets/sheet1.xml', xml);
  return await zip.generateAsync({type:'arraybuffer', compression:'DEFLATE'});
}


// === ACTIVITY TYPE DROPDOWN HELPER ===
function dateKeyLocal(d) {
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}

// Return filtered ACTIVITY_TYPES for a given row (WD/WE/PH-aware)
// isWE = Saturday or Sunday; isPH = public holiday (can be any day)
function typeOptsFor(isWE, isPH, selectedType) {
  const isSpecial = isWE || isPH;
  const isConsultantMode = activeProfile && activeProfile.roster_type === 'consultant' && state.consultantData;
  let filtered;
  if (isConsultantMode) {
    // Normalise selectedType: Consultant Day - HHhMM → Normal Hours - Weekday
    if (selectedType && selectedType.startsWith('Consultant Day')) selectedType = 'Normal Hours - Weekday';
    filtered = CONSULTANT_ACTIVITY_TYPES.filter(t => {
      if (t === 'Normal Hours - Weekday') return !isSpecial;
      if (t === 'On Call - Weekday')      return !isWE && !isPH;
      if (t === 'On Call - Weekend')      return isWE && !isPH;
      if (t === 'On Call - Public Holiday') return isPH;
      return true;
    });
  } else {
    filtered = ACTIVITY_TYPES.filter(t => {
      if(t.startsWith('WD Shift')) return !isSpecial;
      if(t.startsWith('WE Shift')) return isSpecial;
      return true;
    });
  }
  // If selectedType is filtered out, keep it anyway
  const opts = filtered.includes(selectedType) ? filtered : [selectedType, ...filtered];
  const blankOpt=`<option value=""${!selectedType?' selected':''}>— clear row —</option>`;
  return blankOpt+opts.map(t=>`<option value="${t}"${t===selectedType?' selected':''}>${t}</option>`).join('');
}
