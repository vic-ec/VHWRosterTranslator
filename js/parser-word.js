// ═══════════════════════════════════════════════════════════════
// parser-word.js — Word roster table parsers (.docx and legacy .doc).
//
//   readOleStreams(arrayBuffer)             → Map<name, Uint8Array>
//   extractDocText(arrayBuffer)             → string  (Word 97 text stream)
//   extractDocxTables(arrayBuffer)          → Promise<string[][][]>
//   extractDocTables(arrayBuffer, nCols)    → string[][][]
//   extractWordTables(buf, fileName, nCols) → Promise<string[][][]>
//   parseWordRosterTable(buf, profile, fn)  → Promise<{days, doctors, warnings}>
//   getTableShifts(data, name, month, profile, year) → {[date]: shiftObj}
//
// Why this exists: a Word roster carries its grid explicitly — rows and
// cells are part of the file format. Unlike the PDF parsers, there is no
// coordinate inference and nothing to calibrate. A profile only has to
// say what the columns MEAN, not where they are.
//
// .docx is the reliable path (OOXML gives exact row/cell structure).
// .doc is supported because that is what departments actually circulate,
// but the legacy format marks cell-end and row-end with the same byte
// (0x07) and only distinguishes them via paragraph properties, so rows
// are recovered using the column count declared in the profile.
//
// Depends on: config.js, holidays.js, JSZip (vendor bundle, already loaded)
// ═══════════════════════════════════════════════════════════════

// ── CP1252 high range (0x80–0x9F differs from Latin-1) ──────────────────────
const CP1252_HIGH = {
  0x80:0x20AC, 0x82:0x201A, 0x83:0x0192, 0x84:0x201E, 0x85:0x2026, 0x86:0x2020,
  0x87:0x2021, 0x88:0x02C6, 0x89:0x2030, 0x8A:0x0160, 0x8B:0x2039, 0x8C:0x0152,
  0x8E:0x017D, 0x91:0x2018, 0x92:0x2019, 0x93:0x201C, 0x94:0x201D, 0x95:0x2022,
  0x96:0x2013, 0x97:0x2014, 0x98:0x02DC, 0x99:0x2122, 0x9A:0x0161, 0x9B:0x203A,
  0x9C:0x0153, 0x9E:0x017E, 0x9F:0x0178,
};

// ── OLE / CFB compound-document reader ──────────────────────────────────────
// Legacy .doc is a mini filesystem. We only need two streams out of it
// ("WordDocument" and the table stream), but reaching them means walking
// the FAT the same way the format's own reader would.
function readOleStreams(arrayBuffer) {
  const dv = new DataView(arrayBuffer);
  const u8 = new Uint8Array(arrayBuffer);
  const SIG = [0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1];
  for (let i = 0; i < 8; i++) {
    if (u8[i] !== SIG[i]) throw new Error('Not a Word 97-2003 document (bad OLE signature)');
  }

  const secSize      = 1 << dv.getUint16(30, true);
  const miniSize     = 1 << dv.getUint16(32, true);
  const dirStart     = dv.getUint32(48, true);
  const miniCutoff   = dv.getUint32(56, true);
  const miniFatStart = dv.getUint32(60, true);
  const difatStart   = dv.getUint32(68, true);

  const FREE = 0xFFFFFFFF, ENDOFCHAIN = 0xFFFFFFFE;
  const secOff = s => (s + 1) * secSize;

  // DIFAT → list of FAT sectors
  const fatSectors = [];
  for (let i = 0; i < 109; i++) {
    const v = dv.getUint32(76 + i * 4, true);
    if (v === FREE || v === ENDOFCHAIN) break;
    fatSectors.push(v);
  }
  let ds = difatStart, guard = 0;
  const perSec = secSize / 4;
  while (ds !== ENDOFCHAIN && ds !== FREE && guard++ < 100000) {
    const base = secOff(ds);
    for (let i = 0; i < perSec - 1; i++) {
      const v = dv.getUint32(base + i * 4, true);
      if (v !== FREE) fatSectors.push(v);
    }
    ds = dv.getUint32(base + (perSec - 1) * 4, true);
  }

  // FAT
  const fat = new Uint32Array(fatSectors.length * perSec);
  let k = 0;
  for (const fs of fatSectors) {
    const base = secOff(fs);
    for (let i = 0; i < perSec; i++) fat[k++] = dv.getUint32(base + i * 4, true);
  }

  const chain = (start, table) => {
    const out = [];
    let s = start, g = 0;
    while (s !== ENDOFCHAIN && s !== FREE && s < table.length && g++ < 1000000) {
      out.push(s);
      s = table[s];
    }
    return out;
  };

  const readSectors = (sectors, size, sSize, source) => {
    const out = new Uint8Array(sectors.length * sSize);
    sectors.forEach((s, i) => {
      const off = source ? s * sSize : secOff(s);
      out.set(source ? source.subarray(off, off + sSize) : u8.subarray(off, off + sSize), i * sSize);
    });
    return out.subarray(0, size);
  };

  // Mini FAT
  const miniFatSectors = chain(miniFatStart, fat);
  const miniFat = new Uint32Array(miniFatSectors.length * perSec);
  k = 0;
  for (const ms of miniFatSectors) {
    const base = secOff(ms);
    for (let i = 0; i < perSec; i++) miniFat[k++] = dv.getUint32(base + i * 4, true);
  }

  // Directory entries
  const dirSectors = chain(dirStart, fat);
  const dirBytes = readSectors(dirSectors, dirSectors.length * secSize, secSize, null);
  const ddv = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
  const entries = [];
  for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
    const nameLen = ddv.getUint16(off + 64, true);
    if (nameLen < 2) { entries.push(null); continue; }
    let name = '';
    for (let i = 0; i < nameLen - 2; i += 2) name += String.fromCharCode(ddv.getUint16(off + i, true));
    entries.push({
      name,
      type:  ddv.getUint8(off + 66),
      left:  ddv.getUint32(off + 68, true),
      right: ddv.getUint32(off + 72, true),
      child: ddv.getUint32(off + 76, true),
      start: ddv.getUint32(off + 116, true),
      size:  ddv.getUint32(off + 120, true),
    });
  }

  // Root entry holds the mini stream
  const root = entries.find(e => e && e.type === 5);
  let miniStream = null;
  if (root && root.size > 0) {
    miniStream = readSectors(chain(root.start, fat), root.size, secSize, null);
  }

  // The directory is a red-black tree, not a flat list. Collect only the
  // root's direct children — an embedded object (a pasted Word picture, say)
  // has its own storage containing a stream also called "WordDocument", and
  // taking entries by name alone would let it shadow the real one.
  const topLevel = [];
  const visit = (id, depth) => {
    if (id === FREE || id >= entries.length || depth > 10000) return;
    const e = entries[id];
    if (!e) return;
    visit(e.left, depth + 1);
    topLevel.push(e);
    visit(e.right, depth + 1);
  };
  if (root) visit(root.child, 0);

  const streams = new Map();
  for (const e of topLevel) {
    if (e.type !== 2 || e.size === 0 || streams.has(e.name)) continue;
    const data = (e.size < miniCutoff && miniStream)
      ? readSectors(chain(e.start, miniFat), e.size, miniSize, miniStream)
      : readSectors(chain(e.start, fat), e.size, secSize, null);
    streams.set(e.name, data);
  }
  return streams;
}

// ── Word 97 text-stream reconstruction (FIB + piece table) ──────────────────
function extractDocText(arrayBuffer) {
  const streams = readOleStreams(arrayBuffer);
  const wd = streams.get('WordDocument');
  if (!wd) throw new Error('No WordDocument stream — not a Word document');
  const wdv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);

  const flags     = wdv.getUint16(10, true);
  const tableName = (flags & 0x0200) ? '1Table' : '0Table';
  const fcClx     = wdv.getUint32(418, true);
  const lcbClx    = wdv.getUint32(422, true);

  const tbl = streams.get(tableName);
  if (!tbl) throw new Error('Missing ' + tableName + ' stream');
  const clx = tbl.subarray(fcClx, fcClx + lcbClx);
  const cdv = new DataView(clx.buffer, clx.byteOffset, clx.byteLength);

  // Skip Prc blocks (0x01), land on the Pcdt (0x02)
  let i = 0;
  while (i < clx.length && clx[i] === 0x01) i += 3 + cdv.getUint16(i + 1, true);
  if (clx[i] !== 0x02) throw new Error('Malformed piece table');

  const lcbPlc = cdv.getUint32(i + 1, true);
  const plcOff = i + 5;
  const n = Math.floor((lcbPlc - 4) / 12);
  const cps = [];
  for (let j = 0; j <= n; j++) cps.push(cdv.getUint32(plcOff + j * 4, true));

  let text = '';
  for (let p = 0; p < n; p++) {
    const pcd = plcOff + 4 * (n + 1) + p * 8;
    let fc = cdv.getUint32(pcd + 2, true);
    const compressed = (fc & 0x40000000) !== 0;
    fc &= 0x3FFFFFFF;
    const len = cps[p + 1] - cps[p];
    if (compressed) {
      const base = fc >> 1;
      for (let c = 0; c < len; c++) {
        const b = wd[base + c];
        text += String.fromCharCode(b >= 0x80 && b <= 0x9F ? (CP1252_HIGH[b] || b) : b);
      }
    } else {
      for (let c = 0; c < len; c++) text += String.fromCharCode(wdv.getUint16(fc + c * 2, true));
    }
  }
  return text;
}

// ── Table recovery from a .doc text stream ──────────────────────────────────
// 0x07 marks BOTH end-of-cell and end-of-row; the two are told apart only by
// paragraph properties. Rather than parse those, we use the column count the
// profile already declares: every row is exactly nCols cells + 1 row mark.
function extractDocTables(arrayBuffer, nCols) {
  const text = extractDocText(arrayBuffer);
  const last = text.lastIndexOf('\u0007');
  if (last < 0 || !nCols) return [];
  const tokens = text.slice(0, last + 1).split('\u0007');
  tokens.pop(); // trailing empty piece after the final mark

  const rows = [];
  for (let i = 0; i + nCols <= tokens.length; i += nCols + 1) {
    rows.push(tokens.slice(i, i + nCols).map(cleanCellText));
  }
  return rows.length ? [rows] : [];
}

function cleanCellText(s) {
  return String(s == null ? '' : s)
    .replace(/\u00a0/g, ' ')                 // nbsp placeholders in empty cells
    .replace(/[\r\n]/g, ' ')           // paragraph breaks inside a cell
    .replace(/[\u0000-\u0006\u0008\u000e-\u001f]/g, '') // field/picture markers
    .replace(/\s+/g, ' ')
    .trim();
}

// ── .docx table extraction (OOXML — exact structure, nothing inferred) ──────
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function extractDocxTables(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('Not a Word .docx (no word/document.xml)');
  const xml = await entry.async('string');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Corrupt document.xml');

  const kids = (el, tag) =>
    Array.from(el.childNodes).filter(c => c.nodeType === 1 && c.localName === tag);

  const out = [];
  const tbls = doc.getElementsByTagNameNS(W_NS, 'tbl');
  for (let t = 0; t < tbls.length; t++) {
    const rows = [];
    for (const tr of kids(tbls[t], 'tr')) {
      const cells = [];
      for (const tc of kids(tr, 'tc')) {
        const ts = tc.getElementsByTagNameNS(W_NS, 't');
        let txt = '';
        for (let x = 0; x < ts.length; x++) txt += ts[x].textContent || '';
        cells.push(cleanCellText(txt));
        // A horizontally merged cell occupies several grid columns; pad so
        // column indexes in the profile still line up.
        const pr = kids(tc, 'tcPr')[0];
        const gs = pr && kids(pr, 'gridSpan')[0];
        const span = gs ? parseInt(gs.getAttributeNS(W_NS, 'val') || gs.getAttribute('w:val') || '1', 10) : 1;
        for (let s = 1; s < span; s++) cells.push('');
      }
      rows.push(cells);
    }
    if (rows.length) out.push(rows);
  }
  return out;
}

async function extractWordTables(arrayBuffer, fileName, nCols) {
  const isDocx = /\.docx$/i.test(fileName || '') ||
    (new Uint8Array(arrayBuffer, 0, 2)[0] === 0x50 && new Uint8Array(arrayBuffer, 0, 2)[1] === 0x4B);
  return isDocx ? extractDocxTables(arrayBuffer) : extractDocTables(arrayBuffer, nCols);
}

// ── Date-cell parsing ───────────────────────────────────────────────────────
const MONTH_ABBR = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};

function parseDateCell(raw) {
  const s = cleanCellText(raw);
  if (!s) return null;
  // "01-Aug", "1 Aug", "01/Aug", "1 August"
  let m = s.match(/^(\d{1,2})\s*[-/ ]\s*([A-Za-z]{3,})\.?$/);
  if (m) {
    const mi = MONTH_ABBR[m[2].slice(0, 3).toLowerCase()];
    if (mi !== undefined) return { day: parseInt(m[1], 10), month: mi };
  }
  // "Aug-01", "August 1"
  m = s.match(/^([A-Za-z]{3,})\.?\s*[-/ ]\s*(\d{1,2})$/);
  if (m) {
    const mi = MONTH_ABBR[m[1].slice(0, 3).toLowerCase()];
    if (mi !== undefined) return { day: parseInt(m[2], 10), month: mi };
  }
  // "01/08" (day/month)
  m = s.match(/^(\d{1,2})\s*\/\s*(\d{1,2})$/);
  if (m) return { day: parseInt(m[1], 10), month: parseInt(m[2], 10) - 1 };
  // bare day number
  m = s.match(/^(\d{1,2})$/);
  if (m) {
    const d = parseInt(m[1], 10);
    if (d >= 1 && d <= 31) return { day: d, month: null };
  }
  return null;
}

const DAY_NAME_RE = /^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$/i;
const WEEKEND_NAMES = new Set(['saturday', 'sunday']);

// Split a roster cell into individual people.
// "Moonieya/Ndebele" → two names; "Woermann AM" → name + note.
function splitNames(cell, opts) {
  const ignore = new Set(((opts && opts.ignore_tokens) || []).map(s => s.toLowerCase()));
  const out = [];
  for (const part of cleanCellText(cell).split(/[\/,&+]|\band\b/i)) {
    let name = part.trim();
    if (!name) continue;
    let note = '';
    const q = name.match(/^(.*?)[\s(]*\b(AM|PM|am|pm|½|1\/2)\b\)?$/);
    if (q && q[1].trim()) { name = q[1].trim(); note = q[2].toUpperCase(); }
    name = name.replace(/[.’']+$/, '').trim();
    if (!name || ignore.has(name.toLowerCase())) continue;
    out.push({ name, note });
  }
  return out;
}

// ── Main entry: Word table → roster days ────────────────────────────────────
async function parseWordRosterTable(arrayBuffer, profile, fileName) {
  const spec = (profile && profile.table) || {};
  const columns = spec.columns || [];
  if (!columns.length) throw new Error('Profile has no table.columns — cannot map this roster');

  const nCols     = columns.length;
  const dateCol   = spec.date_col == null ? 0 : spec.date_col;
  const dayCol    = spec.day_col  == null ? 1 : spec.day_col;
  const roleCols  = spec.role_columns || {};
  const warnings  = [];

  const tables = await extractWordTables(arrayBuffer, fileName, nCols);
  if (!tables.length) throw new Error('No tables found in this document');

  // Pick the table that actually looks like the roster: prefer a header row
  // matching the profile, otherwise the one with the most parseable dates.
  const norm = s => cleanCellText(s).toLowerCase();
  const score = rows => rows.reduce((n, r) => n + (parseDateCell(r[dateCol]) ? 1 : 0), 0);
  let table = null, best = -1;
  for (const rows of tables) {
    const hdr = rows[0] || [];
    const headerMatch = columns.every((c, i) => !c || norm(hdr[i] || '') === norm(c));
    const s = score(rows) + (headerMatch ? 1000 : 0);
    if (s > best) { best = s; table = rows; }
  }
  if (!table || score(table) === 0) {
    throw new Error('No table in this document has a recognisable date column');
  }

  // Drop the header row if present
  let rows = table;
  if (spec.header_row !== false && rows.length && !parseDateCell(rows[0][dateCol])) {
    const hdr = rows[0];
    const mismatch = columns
      .map((c, i) => (c && norm(hdr[i] || '') !== norm(c)) ? `col ${i}: expected "${c}", found "${hdr[i] || ''}"` : null)
      .filter(Boolean);
    if (mismatch.length) warnings.push('Header differs from profile — ' + mismatch.join('; '));
    rows = rows.slice(1);
  }

  const fnMonth = monthFromFileName(fileName);
  const days = [], doctors = new Set();

  for (const row of rows) {
    if (!row || !row.length) continue;
    const parsed = parseDateCell(row[dateCol]);
    if (!parsed) {
      if (row.some(c => c)) warnings.push('Skipped row with unreadable date: ' + row.filter(Boolean).join(' | '));
      continue;
    }
    const month = parsed.month != null ? parsed.month : fnMonth;

    let dayName = cleanCellText(row[dayCol] || '');
    if (!DAY_NAME_RE.test(dayName)) dayName = '';

    const roles = {};
    const allNames = [];
    for (const [role, idx] of Object.entries(roleCols)) {
      const people = splitNames(row[idx], spec);
      roles[role] = people;
      for (const p of people) {
        doctors.add(p.name);
        allNames.push(p.name);
        if (p.note) warnings.push(`${cleanCellText(row[dateCol])} ${role}: "${p.name}" marked ${p.note} — partial day, check hours`);
        if (p.name.length <= 2) warnings.push(`${cleanCellText(row[dateCol])} ${role}: "${p.name}" is very short — a code rather than a name?`);
      }
    }

    days.push({
      date: parsed.day,
      month,
      monthName: month != null ? MONTH_NAMES[month] : '',
      dayName,
      isWeekend: WEEKEND_NAMES.has(dayName.toLowerCase()),
      shiftType: WEEKEND_NAMES.has(dayName.toLowerCase()) ? 'weekend' : 'weekday',
      roles,
      allNames,
      shifts: [[], [], [], []],   // kept for shape-compatibility with the PDF parsers
      consultant: null,
      rosterType: 'table',
    });
  }

  return { days, doctors, warnings };
}

function monthFromFileName(fileName) {
  const m = String(fileName || '').match(
    /January|February|March|April|May|June|July|August|September|October|November|December/i
  );
  return m ? MONTH_NAMES.findIndex(n => n.toLowerCase() === m[0].toLowerCase()) : null;
}

// ── Table roster → timesheet rows ───────────────────────────────────────────
// Mirrors getConsultantShifts' output contract so the preview table, the
// Excel generator and the docx generators need no changes.
function getTableShifts(tableData, doctorName, targetMonth, profile, targetYear) {
  if (!tableData || !profile) return {};
  const rules     = profile.role_rules || {};
  const roleOrder = Object.keys((profile.table && profile.table.role_columns) || {});
  const result    = {};
  const nl   = String(doctorName || '').toLowerCase();
  const year = targetYear || new Date().getFullYear();

  // Public holidays take the weekend bands even midweek.
  const phDays = new Set();
  if (typeof getSAPublicHolidays === 'function') {
    const phMap = getSAPublicHolidays(year);
    for (let d = 1; d <= 31; d++) {
      const dt = new Date(year, targetMonth, d);
      if (dt.getMonth() !== targetMonth) continue;
      const key = year + '-' + String(dt.getMonth() + 1).padStart(2, '0') + '-' +
        String(dt.getDate()).padStart(2, '0');
      if (phMap.has(key)) phDays.add(d);
    }
  }

  const byDate = new Map();
  for (const day of tableData.days) {
    if (day.month === targetMonth) byDate.set(day.date, day);
  }

  // The role this doctor is rostered for on a given date, or null. First
  // match wins, in the profile's declared column order.
  const roleOn = (d) => {
    const day = byDate.get(d);
    if (!day) return null;
    for (const key of roleOrder) {
      const hit = (day.roles[key] || []).find(p => p.name.toLowerCase() === nl);
      if (hit) return { role: key, note: hit.note };
    }
    return null;
  };

  // A role is on call unless the profile says otherwise.
  const isCallRole = (r) => !!r && (rules[r.role] || {}).is_call !== false;

  const t = v => (v || '').replace(':', 'H');
  const band = (rule, label, note) => ({
    nf: t(rule.normal ? rule.normal[0] : ''),
    nt: t(rule.normal ? rule.normal[1] : ''),
    of: t(rule.ot2 ? rule.ot2[0] : (rule.ot1 ? rule.ot1[0] : '')),
    ot: t(rule.ot2 ? rule.ot2[1] : (rule.ot1 ? rule.ot1[1] : '')),
    typeLabel: note ? `${label} (${note})` : label,
    ot1f: t(rule.ot1 ? rule.ot1[0] : ''),
    ot1t: t(rule.ot1 ? rule.ot1[1] : ''),
    ot2f: t(rule.ot2 ? rule.ot2[0] : ''),
    ot2t: t(rule.ot2 ? rule.ot2[1] : ''),
  });

  // A call roster records overtime only — the ordinary working week is
  // implied, so it is filled in here rather than read from the file.
  const dflt        = profile.default_weekday || null;
  const postCallOff = profile.post_call_off !== false;
  const daysInMonth = new Date(year, targetMonth + 1, 0).getDate();

  for (let d = 1; d <= daysInMonth; d++) {
    const dow       = new Date(year, targetMonth, d).getDay();
    const isWeekend = dow === 0 || dow === 6;
    const isPH      = phDays.has(d);
    const today     = roleOn(d);

    // Rostered today: the role's own bands win, whatever day it is. A role
    // with no rule for this kind of day (daytime work on a weekend, say)
    // falls through to the ordinary-day logic below.
    if (today) {
      const rr   = rules[today.role];
      const rule = rr && ((isWeekend || isPH) ? rr.weekend_ph : rr.weekday);
      if (rule) {
        const label = isPH      ? (rr.label_ph      || `${today.role} - Public Holiday`)
                    : isWeekend ? (rr.label_weekend || `${today.role} - Weekend`)
                                : (rr.label_weekday || `${today.role} - Weekday`);
        result[d] = band(rule, label, today.note);
        continue;
      }
    }

    // Post-call: yesterday's call runs into this morning, and the rest of
    // the day is off. Only actual call counts — a daytime role such as a
    // theatre session earns no day off. Day 1 cannot be judged, since the
    // previous month is not loaded, so it is treated as an ordinary day.
    if (postCallOff && d > 1 && isCallRole(roleOn(d - 1))) continue;

    // Otherwise: an ordinary working weekday, or nothing at all.
    if (isWeekend || isPH || !dflt) continue;
    result[d] = band(dflt, dflt.label || 'Normal Hours - Weekday', '');
  }
  return result;
}

// ── Overlay: table roster → editedShifts ────────────────────────────────────
// Mirrors overlayConsultantShifts so buildPreview can call both blindly.
function overlayTableShifts(doctorName, targetMonth, targetYear) {
  if (typeof isTableRosterMode !== 'function' || !isTableRosterMode()) return 0;
  const shifts = getTableShifts(state.tableData, doctorName, targetMonth, activeProfile, targetYear);
  let added = 0;
  for (const [dateStr, shift] of Object.entries(shifts)) {
    const d = parseInt(dateStr, 10);
    if (!state.editedShifts[d]) {
      state.editedShifts[d] = shift;
      state.originalShifts[d] = { ...shift };
      added++;
    }
  }
  return added;
}

// ── Setup-time table detection (no profile yet) ──────────────────────────────
// The EC wizard needs the grid before a profile exists to describe it. For
// .docx the structure is explicit. For .doc the row length is unknown — cell
// and row marks are the same byte — so candidate widths are scored by how
// many rows begin with something date-shaped, and the best one wins.
function inferDocColumnCount(text) {
  const last = text.lastIndexOf('\u0007');
  if (last < 0) return 0;
  const tokens = text.slice(0, last + 1).split('\u0007');
  tokens.pop();
  let bestN = 0, bestRatio = 0;
  for (let n = 2; n <= 15; n++) {
    let dated = 0, rows = 0;
    for (let i = 0; i + n <= tokens.length; i += n + 1) {
      rows++;
      if (parseDateCell(tokens[i])) dated++;
    }
    if (rows < 3 || dated < 3) continue;
    // Score by the PROPORTION of rows that start with a date, not the count.
    // A divisor of the true width lands on a row boundary every so often and
    // so finds just as many dates, but across far more rows — the true width
    // is the one where nearly every row begins with one. Ties go to the wider
    // grid, since a divisor can never beat it outright.
    const ratio = dated / rows;
    if (ratio > bestRatio + 1e-9 || (Math.abs(ratio - bestRatio) < 1e-9 && n > bestN)) {
      bestRatio = ratio; bestN = n;
    }
  }
  return bestN;
}

async function detectWordTable(arrayBuffer, fileName) {
  const isDocx = /\.docx$/i.test(fileName || '') ||
    (new Uint8Array(arrayBuffer, 0, 2)[0] === 0x50 && new Uint8Array(arrayBuffer, 0, 2)[1] === 0x4B);

  let rows;
  if (isDocx) {
    const tables = await extractDocxTables(arrayBuffer);
    if (!tables.length) throw new Error('No tables found in this document');
    // The roster is the table with the most date-shaped first cells.
    let best = null, bestScore = -1;
    for (const t of tables) {
      const s = t.reduce((n, r) => n + (parseDateCell(r[0]) ? 1 : 0), 0);
      if (s > bestScore) { bestScore = s; best = t; }
    }
    rows = best;
  } else {
    const text = extractDocText(arrayBuffer);
    const n = inferDocColumnCount(text);
    if (!n) throw new Error('Could not find a table in this .doc — try saving it as .docx');
    const tables = extractDocTables(arrayBuffer, n);
    rows = tables.length ? tables[0] : [];
  }

  if (!rows || !rows.length) throw new Error('No table rows found');
  const width = Math.max(...rows.map(r => r.length));
  const norm = rows.map(r => { const c = r.slice(); while (c.length < width) c.push(''); return c; });
  return { rows: norm, columns: width, isDocx };
}
