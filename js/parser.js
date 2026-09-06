// ═══════════════════════════════════════════════════════════════
// parser.js — Shift roster PDF and Excel parsers.
//
//   inferYear(pages, parsedFiles)
//   findAnchors(words)
//   nearestCol(x, anchors)
//   extractNamesWithAnchors(words, anchors, pageH)
//   parseRosterPDF(arrayBuffer, existingParsed)  → {days, doctors, month, year}
//   parseRosterExcel(arrayBuffer)  async           → {days, doctors, month, year}
//   getDoctorShifts(rosterData, doctorName)      → Map<day, shiftObj>
//
// Depends on: config.js, holidays.js
// ═══════════════════════════════════════════════════════════════

// Infer roster year from page text or from already-detected dates.
function inferYear(pageText, detectedDates) {
  // 1. Look for 4-digit year in page text
  const m = pageText.match(/\b(202\d|203\d)\b/);
  if (m) return parseInt(m[1]);
  // 2. Cross-reference: find a year where a detected date's day-of-week matches
  const DAY_IDX = {Sunday:0,Monday:1,Tuesday:2,Wednesday:3,Thursday:4,Friday:5,Saturday:6};
  const ref = detectedDates.find(d => d.dayName && DAY_IDX[d.dayName] !== undefined);
  if (ref) {
    const cur = new Date().getFullYear();
    for (const y of [cur-1, cur, cur+1, cur-2, cur+2]) {
      if (new Date(y, ref.month, ref.date).getDay() === DAY_IDX[ref.dayName]) return y;
    }
  }
  return new Date().getFullYear();
}

// ── Column anchor detection ─────────────────────────────────────────────────
// Matches standalone "08:00" or "18h00" tokens
// The separator is ';' as well as ':' and 'h' on purpose. The EC roster
// template has carried a typo since at least 2023 — its first band reads
// "08;00 - 18:00" — and because that token then failed to look like a time,
// findAnchors() located only three of the four columns and every name in the
// 08:00 column fell outside maxDist and was dropped. Accepting ';' cannot
// change a correctly typed roster: those tokens already matched.
const TIME_TOK_SINGLE = /^\d{2}[h:;]\d{2}$/i;
// Matches combined range tokens like "08:00 - 18h00", "13:00 - 23:00", etc.
const TIME_TOK_RANGE = /^(\d{2}[h:;]\d{2})\s*[-–]\s*(\d{2}[h:;]\d{2})$/i;

function findAnchors(words, minCount) {
  // Collect time tokens: both single ("08:00") and combined ("08:00 - 18h00")
  // For combined tokens, use the word's x position (= the start time column)
  const timeToks = [];
  for (const w of words) {
    if (TIME_TOK_SINGLE.test(w.text)) {
      timeToks.push(w);
    } else if (TIME_TOK_RANGE.test(w.text)) {
      // Combined range token — treat it as a single column anchor at its x position
      timeToks.push({ text: w.text.match(TIME_TOK_RANGE)[1], x: w.x, y: w.y });
    }
  }
  if (timeToks.length < minCount) return null;
  const rowGroups = [];
  for (const t of timeToks) {
    const g = rowGroups.find(g => Math.abs(g.y - t.y) <= 15);
    if (g) { g.toks.push(t); g.ySum += t.y; g.yCount++; }
    else rowGroups.push({ y:t.y, ySum:t.y, yCount:1, toks:[t] });
  }
  const candidates = rowGroups
    .filter(g => g.toks.length >= minCount)
    .map(g => {
      const xs = g.toks.map(t => t.x).sort((a,b) => a-b);
      return { y:g.ySum/g.yCount, toks:g.toks, spread:xs[xs.length-1]-xs[0] };
    })
    .filter(g => g.spread >= 50)
    .sort((a,b) => b.toks.length - a.toks.length || b.y - a.y);
  if (!candidates.length) return null;
  const best = candidates[0];
  const sorted = best.toks.sort((a,b) => a.x - b.x);
  const deduped = [];
  for (const t of sorted)
    if (!deduped.length || t.x - deduped[deduped.length-1].x > 12) deduped.push(t);
  let anchors = deduped;
  if (anchors.length >= 6) {
    const cols = [];
    let i = 0;
    while (i < anchors.length) {
      cols.push(anchors[i]);
      if (i+1 < anchors.length && anchors[i+1].x - anchors[i].x < 50) i += 2;
      else i++;
    }
    anchors = cols;
  }
  if (anchors.length < minCount) return null;
  const anchorXs = anchors.slice(0,4).map(t => t.x);
  return { xs:anchorXs, y:best.y };
}

function nearestCol(x, anchorXs, maxDist) {
  let best=-1, bestDist=maxDist;
  for (let i=0; i<anchorXs.length; i++) {
    const d = Math.abs(x - anchorXs[i]);
    if (d < bestDist) { bestDist=d; best=i; }
  }
  return best;
}

// ── Name utilities ──────────────────────────────────────────────────────────
// ── Built-in VHW fallback profile ──────────────────────────────────────────
// The EC Shift Roster parser (parseRosterPDF) needs no profile at all — its
// column positions are detected dynamically from each PDF's own time-band
// headers. The Consultant On-Call parser, however, needs real column
// coordinates and time rules. This hardcoded default lets BOTH upload boxes
// work the very first time the app is ever opened, with no internet
// connection and no prior cached profile — matching the app's original
// purpose as a VHW-specific tool. If a real profile is later fetched from
// Supabase (or restored from cache), it takes priority over this fallback.
const VHW_FALLBACK_PROFILE = {
  ec_name: 'VHW Emergency Medicine',
  // Kept genuinely short: it is what the offline badge and note print, and a
  // full department name there wraps the header at narrow widths.
  ec_short: 'VHW EM',
  roster_type: 'consultant',
  // Named here rather than relying on supervisorOptionsFor()'s legacy branch,
  // which only fires for roster_type 'shift' — VHW's own profile is
  // 'consultant', so it was falling through to a free-text box. Spelled out
  // rather than referencing LEGACY_EC_SUPERVISORS: that const is inlined from
  // config.js further down the bundle, so naming it here is a dead-zone throw
  // at load.
  supervisors: ['Philip Cloete', 'Sebastian De Haan', 'Paul Xafis'],
  data_start_y: 188,
  time_rules: {
    weekend_ph: {
      ot1: ['07:30', '11:30'],
      ot2: ['11:30', '07:30'],
    },
    weekday_slot2_3: {
      ot1: ['15:30', '16:30'],
      normal: ['07:30', '15:30'],
    },
    weekday_call_only: {
      ot2: ['16:30', '07:30'],
    },
    weekday_slot1_oncall: {
      ot1: ['15:30', '16:30'],
      ot2: ['16:30', '07:30'],
      normal: ['07:30', '15:30'],
    },
    weekday_slot1_no_oncall: {
      ot1: ['15:30', '16:30'],
      normal: ['07:30', '15:30'],
    },
  },
  known_names: ['Cloete', 'Xafis', 'De Haan', 'Els', 'Zaayman'],
  pdf_columns: {
    call:     { x_max: 792, x_min: 618 },
    leave:    { x_max: 618, x_min: 535 },
    slot1:    { x_max: 295, x_min: 220 },
    slot2:    { x_max: 360, x_min: 295 },
    slot3:    { x_max: 445, x_min: 360 },
    meetings: { x_max: 535, x_min: 445 },
  },
};

const NAME_PREFIXES = new Set(['Van','De','Du','Von','Le']);
const NOISE_RE = /^(WEEK|CONSULTANT|REGISTRAR|COSMO|INTERN|PERSAL|MDHS|Shifts|Hours|TOTAL|NUMBER|Leave|worked|Nights|PH|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|January|February|March|April|May|June|July|August|September|October|November|December|Good|Family|Freedom|Worker|Heritage|Youth|Reconciliation|Christmas|Goodwill|Day|Week|No|TOT|Psych)/i;

const isNoise = s => !s || s.length<2 || /^\d+$/.test(s) || /\d{2}[h:]\d{2}/i.test(s) || NOISE_RE.test(s);
const isNameTok = s => /^[A-Z][a-zA-Z\-]{1,14}$/.test(s);

function extractNamesWithAnchors(rowWords, anchorXs, maxDist) {
  const out = [];
  const sorted = [...rowWords].sort((a,b) => a.x - b.x);
  // Left boundary: names must be at or right of (first anchor - half the inter-column gap)
  // This prevents consultant names (x≈325-335) from being assigned to col 0 (x≈386)
  const leftBoundary = anchorXs.length >= 2
    ? anchorXs[0] - (anchorXs[1] - anchorXs[0]) * 0.4
    : anchorXs[0] - 25;
  let i = 0;
  while (i < sorted.length) {
    const w = sorted[i];
    const raw = w.text.replace(/\(T\)/gi,'').replace(/\(Psy\)/gi,'').replace(/\[|\]/g,'').trim();
    if (!raw) { i++; continue; }

    // Skip words too far left of the shift columns (date area, consultant area)
    if (w.x < leftBoundary) { i++; continue; }

    // Handle compound names returned as a single token: "Van Schalkwyk", "De Haan", etc.
    const COMPOUND_RE = /^(Van|De|Du|Von|Le)\s+([A-Z][a-zA-Z\-]{1,20})$/;
    const compoundMatch = raw.match(COMPOUND_RE);
    if (compoundMatch) {
      const name = compoundMatch[1] + ' ' + compoundMatch[2];
      const col = nearestCol(w.x, anchorXs, maxDist);
      if (col >= 0) out.push({ name, col });
      i++; continue;
    }

    if (isNoise(raw) || !isNameTok(raw)) { i++; continue; }
    let name = raw;
    if (NAME_PREFIXES.has(raw) && i+1 < sorted.length) {
      const nw = sorted[i+1];
      const nraw = nw.text.replace(/\(T\)/gi,'').replace(/\(Psy\)/gi,'').trim();
      // Increased gap tolerance from 80 to 120px for compound names
      if (isNameTok(nraw) && nw.x - w.x < 120) { name = raw+' '+nraw; i += 2; }
      else i++;
    } else i++;
    const col = nearestCol(w.x, anchorXs, maxDist);
    if (col >= 0) out.push({ name, col });
  }
  return out;
}

// ── Main PDF parser ─────────────────────────────────────────────────────────
async function parseRosterPDF(arrayBuffer) {
  // A PDF with no text at all is a scan or a photo, not an export. Saying so
  // is the difference between a useful message and a green tick over nothing.
  let sawText = false;
  const pdfjsLib = window.pdfjsLib;
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const doctors = new Set(), days = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const words = content.items
      .map(it => ({ text:it.str.trim(), x:Math.round(it.transform[4]), y:Math.round(it.transform[5]) }))
      .filter(w => w.text.length > 0);

    const pageText = words.map(w => w.text).join(' ');
    sawText = sawText || words.length > 0;
    if (!HAS_DATE.test(pageText)) continue;

    // ── Column anchors ───────────────────────────────────────────────────
    const wdAnchors = findAnchors(words, 3);
    const weToken = words.find(w => /^13[h:]00$/i.test(w.text));
    const weY = weToken ? weToken.y : -1;
    let weAnchors = null;
    if (weY > 0) weAnchors = findAnchors(words.filter(w => w.y <= weY+20), 3);

    const wdAnchorXs = wdAnchors?.xs ?? (weAnchors?.xs ?? [285,360,437,515]);
    const weAnchorXs = weAnchors?.xs ?? (wdAnchors?.xs.slice(0,3) ?? [285,360,437]);
    const wdMaxDist = wdAnchors ? Math.round((wdAnchorXs[1]-wdAnchorXs[0])*0.75) : 60;
    const weMaxDist = weAnchors ? Math.round((weAnchorXs[1]-weAnchorXs[0])*0.75) : 60;

    // ── Group words into y-rows ──────────────────────────────────────────
    const rowMap = new Map();
    for (const w of words) {
      let hit = false;
      for (const [ky] of rowMap)
        if (Math.abs(ky - w.y) <= 6) { rowMap.get(ky).push(w); hit=true; break; }
      if (!hit) rowMap.set(w.y, [w]);
    }
    const sortedYs = [...rowMap.keys()].sort((a,b) => b-a);

    // ── PASS 1: detect date lines ────────────────────────────────────────
    const dateLinesOnPage = [];
    const seenDays = new Set();

    // Year and PH calendar are inferred lazily after first normal date is found
    let phCal = null;   // Map<"YYYY-MM-DD", dayName>
    let rosterYear = -1;

    for (let yi = 0; yi < sortedYs.length; yi++) {
      const y = sortedYs[yi];
      const row = rowMap.get(y).sort((a,b) => a.x - b.x);
      const lineText = row.map(w => w.text).join(' ').trim();
      let dm = lineText.match(DATE_RE);
      let consultantOverride = null;

      if (!dm) {
        const pm = lineText.match(PARTIAL_DATE_RE);
        if (pm) {
          // Build or refresh the PH calendar once we can infer the year
          if (rosterYear < 0) {
            rosterYear = inferYear(pageText, dateLinesOnPage);
            phCal = buildPHCalendar(rosterYear);
          }
          const dateNum = parseInt(pm[1]);
          const monthIdx = MONTHS_IDX[pm[2].toLowerCase()];

          // --- NEW: scan adjacent rows (within 15px below) for weekday + consultant ---
          // PH dates often split: "27 April" at y=506, "Monday Xafis" at y=498
          let foundWeekday = null;
          const WEEKDAY_RE = /\b(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b/i;
          for (let adj = yi + 1; adj < sortedYs.length && sortedYs[adj] >= y - 15; adj++) {
            const adjRow = rowMap.get(sortedYs[adj]).sort((a,b) => a.x - b.x);
            const adjText = adjRow.map(w => w.text).join(' ').trim();
            const wdMatch = adjText.match(WEEKDAY_RE);
            if (wdMatch) {
              foundWeekday = wdMatch[1];
              // Extract consultant: text in the left half (x < 380) after the weekday
              const consultantWords = adjRow.filter(w =>
                w.x > 280 && w.x < 400 && !WEEKDAY_RE.test(w.text) && !/^(PH|Good|Family|Freedom|Worker|Heritage|Youth|Day|Reconciliation|Christmas|Goodwill)$/i.test(w.text)
              );
              if (consultantWords.length > 0) {
                consultantOverride = consultantWords.map(w => w.text).join(' ').trim();
              }
              break;
            }
          }

          // Look up weekday from SA PH calendar
          const phKey = rosterYear+'-'+String(monthIdx+1).padStart(2,'0')+'-'+String(dateNum).padStart(2,'0');
          let computedDay = foundWeekday || (phCal ? phCal.get(phKey) : null);

          // If not in PH calendar, compute from date arithmetic anyway
          if (!computedDay) {
            computedDay = DAY_NAMES_FULL[new Date(rosterYear, monthIdx, dateNum).getDay()];
          }

          const synthetic = pm[1]+' '+pm[2]+' '+computedDay;
          dm = synthetic.match(DATE_RE);
        }

        if (!dm) continue;
      }

      if (!dm) continue;
      const monthIdx = MONTHS_IDX[dm[2].toLowerCase()];
      const dayKey = `${monthIdx}-${dm[1]}`;
      if (seenDays.has(dayKey)) continue;
      seenDays.add(dayKey);

      const dn = dm[3].charAt(0).toUpperCase() + dm[3].slice(1).toLowerCase();
      const isWE = ['Saturday','Sunday'].includes(dn) || (weY > 0 && y <= weY);
      dateLinesOnPage.push({
        y, date:parseInt(dm[1]), month:monthIdx, monthName:dm[2],
        dayName:dn, isWE, consultant: consultantOverride || (dm[4]||'').trim()||null,
      });
    }

    if (dateLinesOnPage.length === 0) continue;
    dateLinesOnPage.sort((a,b) => b.y - a.y);

    // ── PASS 2: zone boundaries with generous overlap buffer ──────────────
    // Use a fixed buffer (% of inter-zone gap) so names near boundaries
    // are always captured without zones stealing from each other.
    const pageTopY = sortedYs[0] + 50;
    const pageBottomY = sortedYs[sortedYs.length-1] - 50;
    const ZONE_BUFFER = 0.25; // each zone claims 25% extra into adjacent zone

    const zones = dateLinesOnPage.map((dl, i) => {
      const prevY = i === 0 ? pageTopY : dateLinesOnPage[i-1].y;
      const nextY = i === dateLinesOnPage.length-1 ? pageBottomY : dateLinesOnPage[i+1].y;
      const mid_up = (dl.y + prevY) / 2;
      const mid_dn = (dl.y + nextY) / 2;
      const buf_up = (prevY - dl.y) * ZONE_BUFFER;
      const buf_dn = (dl.y - nextY) * ZONE_BUFFER;
      return {
        ...dl,
        upperY: mid_up + buf_up,   // extend upward into previous zone
        lowerY: mid_dn - buf_dn,   // extend downward into next zone
        shifts: [[], [], [], []],
        allNames: [],
      };
    });

    // Priority: when zones overlap, the zone whose date y is CLOSEST to the name y wins.
    for (const y of sortedYs) {
      const row = rowMap.get(y).sort((a,b) => a.x - b.x);
      const lineText = row.map(w => w.text).join(' ').trim();
      if (DATE_RE.test(lineText)) continue;

      const candidates = zones.filter(z => y <= z.upperY && y >= z.lowerY);
      if (!candidates.length) continue;
      const zone = candidates.reduce((best, z) =>
        Math.abs(z.y - y) < Math.abs(best.y - y) ? z : best
      );

      const isWERow = weY > 0 && y <= weY;
      const anchorXs = isWERow ? weAnchorXs : wdAnchorXs;
      const maxDist  = isWERow ? weMaxDist  : wdMaxDist;

      for (const { name, col } of extractNamesWithAnchors(row, anchorXs, maxDist)) {
        zone.shifts[col].push(name);
        zone.allNames.push(name);
        if (!/psy/i.test(name)) doctors.add(name);
      }
    }

    for (const zone of zones) {
      // ── Validation: check shift counts per day ───────────────────────────
      const expectedCols = zone.isWE ? 3 : 4;
      const filledCols = zone.shifts.slice(0, expectedCols).filter(s => s.length > 0).length;
      const totalNames = zone.allNames.length;
      if (filledCols < expectedCols && totalNames > 0) {
        console.warn(`[Parser] ⚠ ${zone.date} ${zone.monthName} (${zone.dayName}): only ${filledCols}/${expectedCols} shift columns have names (${totalNames} names total)`);
      }
      if (totalNames === 0) {
        console.warn(`[Parser] ⚠ ${zone.date} ${zone.monthName} (${zone.dayName}): NO names detected in any shift column`);
      }
      days.push({
        date:zone.date, month:zone.month, monthName:zone.monthName,
        dayName:zone.dayName, isWeekend:zone.isWE, shiftType:zone.isWE?'weekend':'weekday',
        shifts:zone.shifts, consultant:zone.consultant, allNames:zone.allNames,
      });
    }
  }

  if (!sawText) throw new Error(
    'this PDF has no text in it \u2014 it is a scan or a photo. Only a PDF ' +
    'exported from the roster file itself can be read.');
  return { days, doctors };
}

// ── Excel parser (.xlsx via parser-xlsx.js / JSZip) ─────────────────────────
async function parseRosterExcel(arrayBuffer) {
  // Reads the workbook with JSZip (see parser-xlsx.js) rather than SheetJS —
  // an .xlsx is a zip of XML, so no extra library is needed. Legacy binary
  // .xls is a different format and is not supported.
  const sheets = await readXlsxSheets(arrayBuffer);
  const doctors = new Set(), days = [];
  const MONTHS = {January:0,February:1,March:2,April:3,May:4,June:5,July:6,August:7,September:8,October:9,November:10,December:11};
  const norm = s => s.replace(/\[|\]/g,'').replace(/\(T\)/gi,'').replace(/\(Psy\)/gi,'').trim();

  for (const sheet of sheets) {
    const lines = sheet.rows
      .flat().filter(Boolean).map(String).map(s => s.trim()).filter(Boolean);
    let cur = null, st = 'weekday';
    const seenDays = new Set();
    const flush = () => { if (cur) { days.push(cur); cur = null; } };
    for (let i = 0; i < lines.length; i++) {
      const l = lines[i];
      if (/08:00\s*[-–]\s*18/i.test(l)) { st='weekday'; continue; }
      if (/08:00\s*[-–]\s*20/i.test(l)) { st='weekend'; continue; }
      const dm = l.match(DATE_RE);
      if (dm) {
        const monthIdx = MONTHS[dm[2]];
        const dayKey = `${monthIdx}-${dm[1]}`;
        if (seenDays.has(dayKey)) { flush(); continue; }
        seenDays.add(dayKey);
        flush();
        const dn = dm[3].charAt(0).toUpperCase() + dm[3].slice(1).toLowerCase();
        cur = { date:parseInt(dm[1]), month:monthIdx, monthName:dm[2], dayName:dn,
          isWeekend:['Saturday','Sunday'].includes(dn), shiftType:st,
          shifts:[[],[],[],[]], consultant:null, allNames:[] };
        if (i+1 < lines.length) { cur.consultant = norm(lines[i+1]); i++; }
        continue;
      }
      if (cur) {
        const raw = norm(l);
        if (raw && raw.length > 1 && raw.length < 30 && /^[A-Z]/.test(raw) && !/^\d/.test(raw)) {
          cur.allNames.push(raw);
          if (!/psy/i.test(raw)) doctors.add(raw);
        }
      }
    }
    flush();
  }
  return { days, doctors };
}

// ── getDoctorShifts ─────────────────────────────────────────────────────────
function getDoctorShifts(rosterData, doctorName, targetMonth) {
  if (!rosterData) return {};
  const result = {}, nl = doctorName.toLowerCase();
  for (const day of rosterData.days) {
    if (day.month !== targetMonth) continue;
    const allNames = [...(day.allNames||[]), ...(day.shifts?.flat()||[])];
    if (!allNames.some(n => n.toLowerCase() === nl)) continue;
    let shiftDef = null;
    for (let col = 0; col < 4; col++) {
      if ((day.shifts[col]||[]).some(n => n.toLowerCase() === nl)) {
        const defs = day.isWeekend ? WEEKEND_SHIFTS : WEEKDAY_SHIFTS;
        shiftDef = defs[col] || defs[defs.length-1];
        break;
      }
    }
    if (!shiftDef) shiftDef = (day.isWeekend ? WEEKEND_SHIFTS : WEEKDAY_SHIFTS)[0];
    result[day.date] = { start:shiftDef.start, end:shiftDef.end, label:shiftDef.label,
      dayName:day.dayName, isWeekend:day.isWeekend };
  }
  return result;
}


// ── Consultant Roster PDF Parser ────────────────────────────────────────────
// Profile-driven: reads column boundaries and time rules from activeProfile.
// Returns { days, doctors } in the same format as parseRosterPDF / parseRosterExcel.

