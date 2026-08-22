// ═══════════════════════════════════════════════════════════════
// parser-xlsx.js — Minimal .xlsx reader built on JSZip.
//
//   readXlsxSheets(arrayBuffer) → Promise<[{name, rows: string[][]}]>
//
// An .xlsx is a zip of XML, so the same JSZip bundle the app already
// loads can read it — no SheetJS needed. This deliberately implements
// only what parseRosterExcel consumes: sheets in workbook order, each
// as rows of cell text (the equivalent of SheetJS's
// sheet_to_json({header:1, raw:false})).
//
// Not supported: legacy binary .xls (a completely different format),
// formulas (the cached result is used), charts, styles beyond the
// number format needed to recognise dates.
//
// Depends on: JSZip (vendor bundle, already loaded)
// ═══════════════════════════════════════════════════════════════

const X_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const X_REL  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const X_PKG  = 'http://schemas.openxmlformats.org/package/2006/relationships';

// Built-in number formats that denote a date/time (ECMA-376 §18.8.30).
const BUILTIN_DATE_FMTS = new Set([14,15,16,17,18,19,20,21,22,45,46,47]);

function xlsxParseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Malformed XML in .xlsx');
  return doc;
}

// "BC12" → 54 (0-based column index)
function colFromRef(ref) {
  let n = 0;
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

// Excel serial → "D Month YYYY" (or with time appended when there is one).
const XLSX_MONTHS = ['January','February','March','April','May','June',
  'July','August','September','October','November','December'];

function serialToDateText(serial, date1904) {
  // Excel's day 60 is a non-existent 29 Feb 1900; shift to stay aligned.
  let days = Math.floor(serial);
  const frac = serial - days;
  if (!date1904 && days > 59) days -= 1;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  const d = new Date(epoch + days * 86400000);
  if (isNaN(d.getTime())) return String(serial);
  let out = `${d.getUTCDate()} ${XLSX_MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
  if (frac > 1e-9) {
    const mins = Math.round(frac * 1440);
    out += ` ${String(Math.floor(mins / 60)).padStart(2,'0')}:${String(mins % 60).padStart(2,'0')}`;
  }
  return out;
}

// Concatenate the <t> runs under a node (handles <si> with rich-text <r> runs).
function xlsxText(node) {
  if (!node) return '';
  const ts = node.getElementsByTagNameNS(X_MAIN, 't');
  let s = '';
  for (let i = 0; i < ts.length; i++) s += ts[i].textContent || '';
  return s;
}

async function readXlsxSheets(arrayBuffer) {
  let zip;
  try {
    zip = await JSZip.loadAsync(arrayBuffer);
  } catch (e) {
    // An .xls saved with an .xlsx extension lands here, as does any non-zip.
    throw new Error('This file is not a readable .xlsx workbook. If it is an older .xls, open it in Excel and Save As .xlsx.');
  }
  const get = async p => { const f = zip.file(p); return f ? f.async('string') : null; };

  const wbXml = await get('xl/workbook.xml');
  if (!wbXml) throw new Error('Not an .xlsx workbook (missing xl/workbook.xml)');
  const wb = xlsxParseXml(wbXml);

  const pr = wb.getElementsByTagNameNS(X_MAIN, 'workbookPr')[0];
  const date1904 = !!pr && /^(1|true)$/i.test(pr.getAttribute('date1904') || pr.getAttribute('backupFile') || '');

  // rId → worksheet path
  const relsXml = await get('xl/_rels/workbook.xml.rels');
  const relMap = {};
  if (relsXml) {
    const rels = xlsxParseXml(relsXml).getElementsByTagNameNS(X_PKG, 'Relationship');
    for (let i = 0; i < rels.length; i++) {
      let t = rels[i].getAttribute('Target') || '';
      if (t.startsWith('/xl/')) t = t.slice(1);
      else if (!t.startsWith('xl/')) t = 'xl/' + t.replace(/^\.\//, '');
      relMap[rels[i].getAttribute('Id')] = t;
    }
  }

  // Shared strings
  const sstXml = await get('xl/sharedStrings.xml');
  const sst = [];
  if (sstXml) {
    const sis = xlsxParseXml(sstXml).getElementsByTagNameNS(X_MAIN, 'si');
    for (let i = 0; i < sis.length; i++) sst.push(xlsxText(sis[i]));
  }

  // Style index → is-a-date
  const stylesXml = await get('xl/styles.xml');
  const styleIsDate = [];
  if (stylesXml) {
    const st = xlsxParseXml(stylesXml);
    const custom = {};
    const nfs = st.getElementsByTagNameNS(X_MAIN, 'numFmt');
    for (let i = 0; i < nfs.length; i++) {
      const code = nfs[i].getAttribute('formatCode') || '';
      // Strip quoted literals and escapes before looking for date tokens.
      const bare = code.replace(/"[^"]*"/g, '').replace(/\\./g, '');
      custom[nfs[i].getAttribute('numFmtId')] = /[ymdYMD]/.test(bare) && !/^[#0.,%\s]*$/.test(bare);
    }
    const xfsWrap = st.getElementsByTagNameNS(X_MAIN, 'cellXfs')[0];
    if (xfsWrap) {
      const xfs = xfsWrap.getElementsByTagNameNS(X_MAIN, 'xf');
      for (let i = 0; i < xfs.length; i++) {
        const id = xfs[i].getAttribute('numFmtId') || '0';
        styleIsDate.push(BUILTIN_DATE_FMTS.has(parseInt(id, 10)) || custom[id] === true);
      }
    }
  }

  // Sheets, in workbook order
  const out = [];
  const sheetEls = wb.getElementsByTagNameNS(X_MAIN, 'sheet');
  for (let s = 0; s < sheetEls.length; s++) {
    const el = sheetEls[s];
    const name = el.getAttribute('name') || `Sheet${s + 1}`;
    const rid = el.getAttributeNS(X_REL, 'id') || el.getAttribute('r:id');
    const path = relMap[rid] || `xl/worksheets/sheet${s + 1}.xml`;
    const xml = await get(path);
    if (!xml) { out.push({ name, rows: [] }); continue; }

    const rows = [];
    const rowEls = xlsxParseXml(xml).getElementsByTagNameNS(X_MAIN, 'row');
    for (let r = 0; r < rowEls.length; r++) {
      const cells = [];
      const cs = rowEls[r].getElementsByTagNameNS(X_MAIN, 'c');
      for (let i = 0; i < cs.length; i++) {
        const c = cs[i];
        const ref = c.getAttribute('r') || '';
        const idx = ref ? colFromRef(ref) : cells.length;
        const type = c.getAttribute('t') || 'n';
        const vEl = c.getElementsByTagNameNS(X_MAIN, 'v')[0];
        let val = '';

        if (type === 's') {
          const k = parseInt(vEl ? vEl.textContent : '', 10);
          val = Number.isFinite(k) ? (sst[k] || '') : '';
        } else if (type === 'inlineStr') {
          val = xlsxText(c.getElementsByTagNameNS(X_MAIN, 'is')[0]);
        } else if (type === 'b') {
          val = vEl && vEl.textContent === '1' ? 'TRUE' : 'FALSE';
        } else if (type === 'e') {
          val = vEl ? vEl.textContent : '';
        } else if (vEl) {
          const raw = vEl.textContent;
          const num = parseFloat(raw);
          const si = parseInt(c.getAttribute('s') || '', 10);
          if (type === 'n' && Number.isFinite(num) && styleIsDate[si]) val = serialToDateText(num, date1904);
          else val = raw;
        }

        while (cells.length < idx) cells.push('');
        cells[idx] = val;
      }
      const rNum = parseInt(rowEls[r].getAttribute('r') || '', 10);
      const target = Number.isFinite(rNum) ? rNum - 1 : rows.length;
      while (rows.length < target) rows.push([]);
      rows[target] = cells;
    }
    out.push({ name, rows });
  }
  return out;
}
