// ═══════════════════════════════════════════════════════════════
// parser-consultant.js — Consultant roster PDF parser.
//
//   parseConsultantRosterPDF(arrayBuffer, profile) → {days, doctors, month, year}
//   colFor(x, divPositions)
//   joinTokens(tokens)
//   getConsultantShifts(consultantData, doctorName, targetMonth,
//                       profile, targetYear) → {[dateStr]: shiftObj}
//
// Profile-driven: column boundaries and time rules come from
// activeProfile (fetched from Supabase ec_profiles table).
// Adding a new EC never requires changes to this file — only a
// new Supabase row with the correct profile JSON.
//
// Depends on: config.js, holidays.js
// ═══════════════════════════════════════════════════════════════

// Two consultants in one duty cell — "Cloete & Els", "Cloete / Els" — are two
// people on duty together, not one name: the junior is first on call and the
// next consultant supervises, and HR needs each of them to have claimed the
// day. Both separators appear in the real exports. This never applies to the
// Meetings column, which is not a duty column at all.
const PAIR_SEP = /\s*[/&]\s*/;

// The six column names every consultant export prints above its grid. Declared
// out here because findHeaderRow runs near the top of the parse, before a const
// inside the function body would have been initialised.
const HEADER_WANT = [['slot1',/^1$/], ['slot2',/^2$/], ['slot3',/^3$/],
                     ['meetings',/^meetings/i], ['leave',/^leave$/i], ['call',/^call$/i]];

async function parseConsultantRosterPDF(arrayBuffer, profile) {
  let   cols  = profile.pdf_columns;   // {slot1,slot2,slot3,meetings,leave,call}
  const rules = profile.time_rules;    // weekday/weekend time bands
  const knownNames = new Set((profile.known_names || []).map(n => n.toLowerCase()));

  const pdf   = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
  const page  = await pdf.getPage(1);
  const tc    = await page.getTextContent();
  const vp    = page.getViewport({ scale: 1 });
  const H     = vp.height;

  // ── Collect words with x/y (PDF y is bottom-up → flip to top-down) ──
  const words = [];
  for (const item of tc.items) {
    if (!item.str.trim()) continue;
    // Split fused tokens like "1Wednesday" into date+weekday
    // "1Wednesday" and "1 Wednesday" are both used. Without the optional
    // space the whole cell fails both the date test and the weekday test, so
    // the row has no date, currentDate never advances and every row of the
    // file is skipped — the April export parsed as zero days because of it.
    const fused = item.str.match(/^(\d{1,2})\s*([A-Z][a-z]+)$/);
    if (fused) {
      const x = item.transform[4];
      const y = H - item.transform[5];
      words.push({ text: fused[1], x, y });
      words.push({ text: fused[2], x: x + 8, y });
    } else {
      words.push({ text: item.str.trim(), x: item.transform[4], y: H - item.transform[5] });
    }
  }

  // ── Group every word into rows first, so the header can be found ──
  const byRow = ws => {
    const m = new Map();
    for (const w of ws) {
      const ry = Math.round(w.y / 4) * 4;
      if (!m.has(ry)) m.set(ry, []);
      m.get(ry).push(w);
    }
    return [...m.entries()].sort((a, b) => a[0] - b[0])
      .map(([y, g]) => ({ y, ws: g.sort((a, b) => a.x - b.x) }));
  };
  const allRows = byRow(words);

  // ── Where the data starts ──
  // profile.data_start_y is a fixed 188, and none of the real exports agree
  // with it: July's first row sits at y=132 and May's at y=124, so the first
  // six and eight days of those months were dropped before anything looked at
  // them — a timesheet quietly missing its first week. The header row is the
  // real boundary, and it is already being located for the columns.
  const header = findHeaderLabels(allRows);
  const DATA_Y = header ? header.y + 4 : (profile.data_start_y || 188);
  const dataWords = words.filter(w => w.y >= DATA_Y);

  // ── Group words into rows by y (±4px tolerance) ──
  const rowMap = new Map();
  for (const w of dataWords) {
    const ry = Math.round(w.y / 4) * 4;
    if (!rowMap.has(ry)) rowMap.set(ry, []);
    rowMap.get(ry).push(w);
  }
  const rows = [...rowMap.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, ws]) => ws.sort((a, b) => a.x - b.x));

  // ── Column assignment ──
  const WEEKDAYS = new Set(['Monday','Tuesday','Wednesday','Thursday','Friday']);
  const WEEKENDS = new Set(['Saturday','Sunday']);
  const MONTHS_IDX = {January:0,February:1,March:2,April:3,May:4,June:5,
    July:6,August:7,September:8,October:9,November:10,December:11};

  function colFor(x) {
    for (const [name, {x_min, x_max}] of Object.entries(cols)) {
      if (x >= x_min && x < x_max) return name;
    }
    return null;
  }

  // ── Finding the columns ─────────────────────────────────────────────────
  // Seven real exports, and nothing about the geometry is constant. A column
  // label may sit over its data (May), a whole column to the right of it
  // (January's spreadsheet export), or on a different row from the other
  // labels (February). The profile's fixed pdf_columns fit two of the seven.
  //
  // Two things do hold: the columns are always in the order slot 1, 2, 3,
  // Meetings, Leave, Call from left to right, and the data itself forms
  // columns. So the labels are located, the data is clustered into columns of
  // its own, and the two are matched in that order — which needs no assumption
  // about how far a label sits from what it heads, and lets a label with
  // nothing beneath it match nothing at all.

  // Meetings, Leave and Call share a row in every export seen; the slot
  // numbers may be on that row or the one above. "1" is far too common a token
  // to hunt for on its own — dates and printed spreadsheet row numbers are the
  // same string — so they are read right to left from a band around that row,
  // each one left of the column that follows it.
  function findHeaderLabels(allRows) {
    for (const row of allRows) {
      const m = row.ws.find(w => /^meetings/i.test(w.text));
      const l = row.ws.find(w => /^leave$/i.test(w.text));
      const c = row.ws.find(w => /^call$/i.test(w.text));
      if (!m || !l || !c || !(m.x < l.x && l.x < c.x)) continue;
      const band = allRows.filter(r => Math.abs(r.y - row.y) <= 20);
      const rightmost = (label, limit) => {
        let best = null;
        for (const r of band) for (const w of r.ws)
          if (w.text.trim() === label && w.x < limit && (!best || w.x > best.x)) best = w;
        return best;
      };
      const s3 = rightmost('3', m.x);
      const s2 = s3 && rightmost('2', s3.x);
      const s1 = s2 && rightmost('1', s2.x);
      if (!s1) continue;
      return { at: [s1.x, s2.x, s3.x, m.x, l.x, c.x],
               y: Math.max(row.y, s1.y, s2.y, s3.y) };
    }
    return null;
  }

  // Where the data actually sits. A column is a run of x values within tol of
  // where the run started; a single stray word does not make one.
  function clusterColumns(xs, tol) {
    const sorted = [...xs].sort((a, b) => a - b);
    const runs = [];
    for (const x of sorted) {
      const last = runs[runs.length - 1];
      if (last && x - last.start <= tol) last.xs.push(x);
      else runs.push({ start: x, xs: [x] });
    }
    // Every distinct position counts, even one used once. July writes a single
    // "Retreat" in its Meetings column and May a single "PH" left of slot 1;
    // dropping those as noise left no column there, and the neighbouring
    // column's range — which runs to the midpoint of the next column along —
    // reached over and swallowed them into the duty slots.
    return runs.map(r => r.xs.reduce((a, b) => a + b, 0) / r.xs.length);
  }

  // Match labels to columns in order, left to right, allowing a label to match
  // nothing. Plain edit-distance shape: the cost of a match is how far apart
  // the two are, an unmatched label costs skipCost, and a column no label
  // wants is free — the date, the weekday, January's row numbers and the
  // trailing totals are all columns nothing should be read out of.
  function alignLabelsToColumns(labels, columns, skipCost) {
    const L = labels.length, C = columns.length, INF = Infinity;
    const best = [], back = [];
    for (let i = 0; i <= L; i++) { best.push(new Array(C + 1).fill(INF)); back.push(new Array(C + 1).fill(null)); }
    for (let j = 0; j <= C; j++) best[0][j] = 0;
    for (let i = 1; i <= L; i++) {
      for (let j = 0; j <= C; j++) {
        let b = best[i-1][j] + skipCost, f = 'skipLabel';
        if (j > 0) {
          const drop = best[i][j-1];
          if (drop < b) { b = drop; f = 'dropColumn'; }
          const take = best[i-1][j-1] + Math.abs(labels[i-1] - columns[j-1]);
          if (take < b) { b = take; f = 'match'; }
        }
        best[i][j] = b; back[i][j] = f;
      }
    }
    const pick = new Array(L).fill(-1);
    let i = L, j = C;
    while (i > 0) {
      const f = back[i][j];
      if (f === 'dropColumn') j--;
      else if (f === 'skipLabel') i--;
      else { pick[i-1] = j-1; i--; j--; }
    }
    return pick;
  }

  // Each matched column is bounded by the midpoints to its neighbouring
  // columns — its own neighbours, not the labels' — so a column nothing was
  // matched to cannot bleed into the one beside it.
  function columnsFromLabels(head, dataRows) {
    const columns = clusterColumns(dataRows.flatMap(r => r.ws.map(w => w.x)), 10);
    if (columns.length < 3) return null;
    const gaps = columns.slice(1).map((x, k) => x - columns[k]).sort((a, b) => a - b);
    const skipCost = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 70;
    const pick = alignLabelsToColumns(head.at, columns, skipCost);
    const out = {};
    HEADER_WANT.forEach(([key], n) => {
      const k = pick[n];
      if (k < 0) { out[key] = { x_min: 0, x_max: 0 }; return; }   // no data under this label
      out[key] = { x_min: k > 0 ? (columns[k-1] + columns[k]) / 2 : columns[k] - 30,
                   x_max: k < columns.length - 1 ? (columns[k] + columns[k+1]) / 2 : Infinity };
    });
    out.__first = Math.min(...HEADER_WANT.map(([key]) => out[key].x_max > 0 ? out[key].x_min : Infinity));
    return out;
  }

  // Join compound surnames: "De" + "Haan" → "De Haan"
  // Also handles slash pairs: "Cloete" + "/" + "Els" → "Cloete/Els"
  function joinTokens(tokens) {
    const result = [];
    let i = 0;
    while (i < tokens.length) {
      const t = tokens[i];
      // Check if next token is "/" or "&" then a name (a pair)
      if (i + 2 < tokens.length && (tokens[i+1] === '/' || tokens[i+1] === '&')) {
        result.push(t + '/' + tokens[i+2]);
        i += 3;
        continue;
      }
      // Check if this looks like first part of compound surname (De, Van, etc.)
      if (i + 1 < tokens.length && NAME_PREFIXES.has(t) &&
          /^[A-Z][a-z]+$/.test(tokens[i+1])) {
        result.push(t + ' ' + tokens[i+1]);
        i += 2;
        continue;
      }
      result.push(t);
      i++;
    }
    return result;
  }

  if (header) {
    const fromLabels = columnsFromLabels(header, byRow(dataWords));
    if (fromLabels) cols = fromLabels;
  }
  // Everything left of the first real column is the date and weekday.
  const leftBound = (cols.__first != null && isFinite(cols.__first))
    ? cols.__first : cols.slot1.x_min;

  // ── Parse each row ──
  const days = [];
  const doctors = new Set();
  let currentDate = null, currentDayName = null, currentMonth = null;

  for (const rowWords of rows) {
    // Bucket words by column
    const buckets = { slot1:[], slot2:[], slot3:[], meetings:[], leave:[], call:[] };
    let dateNum = null, weekdayName = null;

    for (const w of rowWords) {
      if (w.x < leftBound) continue;               // date/weekday, read below
      const c = colFor(w.x);
      if (c && c in buckets) buckets[c].push(w.text);
    }

    // The date is the number immediately to the left of the weekday name.
    // January's export prints the spreadsheet's own row numbers in a column
    // further left again, and taking the first number on the row made those
    // the dates — which is how that file reported a 34th of the month.
    const leftWords = rowWords.filter(w => w.x < leftBound);
    const wd = leftWords.find(w => WEEKDAYS.has(w.text) || WEEKENDS.has(w.text));
    if (wd) {
      weekdayName = wd.text;
      // The last cell of January's grid reads "31-Jan" rather than "31".
      // Without the month suffix it is not a date, the nearest number left of
      // the weekday becomes the printed spreadsheet row number, and that file
      // reported a 34th of the month with the 31st missing.
      const DATE_CELL = /^(\d{1,2})(?:[-\s][A-Za-z]{3,9})?$/;
      let nearest = null, nearestNum = null;
      for (const w of leftWords) {
        const m = DATE_CELL.exec(w.text);
        if (m && w.x < wd.x && (!nearest || w.x > nearest.x)) { nearest = w; nearestNum = m[1]; }
      }
      if (nearest) dateNum = parseInt(nearestNum);
    }

    const startsDay = dateNum !== null && !!weekdayName;
    if (startsDay) {
      currentDate    = dateNum;
      currentDayName = weekdayName;
    }
    if (currentDate === null) continue;
    // A row with no date and no weekday of its own is a continuation of the
    // day above — or, at the foot of January's grid, a stray left by the
    // spreadsheet. Either way its names belong to that day. Pushing a second
    // day object for the same date instead was double-counting the last day
    // of the month in every staff-list tally.
    if (!startsDay && days.length) {
      const prev = days[days.length - 1];
      const add = (arr, more) => { for (const n of more) if (!arr.includes(n)) arr.push(n); };
      add(prev.slot1, joinTokens(buckets.slot1));
      add(prev.slot2, joinTokens(buckets.slot2));
      add(prev.slot3, joinTokens(buckets.slot3));
      add(prev.callNames, joinTokens(buckets.call.filter(t => !/^\d+$/.test(t))));
      add(prev.leaveNames, joinTokens(buckets.leave.filter(t => t !== 'Leave')));
      for (const n of [...prev.slot1, ...prev.slot2, ...prev.slot3, ...prev.callNames])
        for (const one of n.split(PAIR_SEP)) if (one.trim().length > 1) doctors.add(one.trim());
      prev.allNames = [...new Set([...prev.slot1, ...prev.slot2, ...prev.slot3, ...prev.callNames]
        .flatMap(n => n.split(PAIR_SEP)).map(n => n.trim()))];
      continue;
    }

    // Parse slot names
    const slot1Names = joinTokens(buckets.slot1);
    const slot2Names = joinTokens(buckets.slot2);
    const slot3Names = joinTokens(buckets.slot3);

    // Parse leave: "De Haan Leave" → name is "De Haan", "Leave" is literal word
    const leaveRaw = joinTokens(buckets.leave.filter(t => t !== 'Leave'));
    const isLeave  = buckets.leave.includes('Leave') || leaveRaw.length > 0;
    const leaveNames = leaveRaw;

    // Parse call: names + optional trailing number
    const callRaw   = buckets.call.filter(t => !/^\d+$/.test(t));
    const callNames = joinTokens(callRaw);

    // Collect all names on this day for the doctors set
    const allSlotNames = [...slot1Names, ...slot2Names, ...slot3Names];
    for (const name of allSlotNames) {
      // Expand pairs. Two names in one duty cell are two consultants on duty
      // together — the junior is first on call and the next supervises, and
      // HR needs both to have claimed the day — so both are credited.
      for (const n of name.split(PAIR_SEP)) {
        const trimmed = n.trim();
        if (trimmed.length > 1) doctors.add(trimmed);
      }
    }
    for (const name of callNames) {
      for (const n of name.split(PAIR_SEP)) {
        const trimmed = n.trim();
        if (trimmed.length > 1) doctors.add(trimmed);
      }
    }

    const isWeekend = WEEKENDS.has(currentDayName);
    const isPublicHoliday = false; // PH detection can be added later

    days.push({
      date: currentDate,
      dayName: currentDayName,
      isWeekend,
      // Consultant-specific fields
      slot1: slot1Names,
      slot2: slot2Names,
      slot3: slot3Names,
      callNames,
      leaveNames,
      isLeave: isLeave && leaveNames.length > 0,
      // Keep month from filename/year detection — set post-parse
      month: null,
      monthName: null,
      // Standard fields for compatibility
      allNames: [...new Set([...allSlotNames.flatMap(n=>n.split(PAIR_SEP)).map(n=>n.trim()),
                             ...callNames.flatMap(n=>n.split(PAIR_SEP)).map(n=>n.trim())])],
      shifts: [slot1Names, slot2Names, slot3Names, callNames],
      shiftType: isWeekend ? 'weekend' : 'weekday',
      consultant: null,
      rosterType: 'consultant',
    });
  }

  return { days, doctors };
}

// ── getConsultantShifts ──────────────────────────────────────────────────────
// Converts a consultant roster day into editedShifts entries for a named doctor.
// Returns { [date]: { nf, nt, of, ot, typeLabel } }

function getConsultantShifts(consultantData, doctorName, targetMonth, profile, targetYear) {
  if (!consultantData || !profile) return {};
  const result = {};
  const rules = profile.time_rules || {};
  const nl = doctorName.toLowerCase();

  // Helper: find if a name (possibly slash-pair) matches the target doctor
  // Also handles cases where call column has extra text like "(2nd)" appended
  const nameMatches = (name) => {
    if (!name) return false;
    return name.split(PAIR_SEP).some(n => {
      const clean = n.trim().toLowerCase().replace(/[^a-z\s]/g,'').trim();
      // Exact match, or compound-surname containment (e.g. 'de haan' within 'de haan')
      if (clean === nl) return true;
      // Allow prefix match only if nl is at least 4 chars (avoids 'els' matching 'el')
      if (nl.length >= 4 && clean.startsWith(nl)) return true;
      if (nl.length >= 4 && clean === nl) return true;
      // For short names (3 chars), require exact match
      return false;
    });
  };

  // Build SA public holiday map for the target month's year
  // (we need to know if a day is a PH so we apply weekend_ph rule)
  const phDays = new Set();
  if (typeof getSAPublicHolidays === 'function') {
    // Use targetYear for correct PH detection (2026 != current year potentially)
    const phYear = targetYear || new Date().getFullYear();
    const phMap = getSAPublicHolidays(phYear);
    for (let d = 1; d <= 31; d++) {
      const dateObj = new Date(phYear, targetMonth, d);
      const key = phYear + '-' +
        String(dateObj.getMonth()+1).padStart(2,'0') + '-' +
        String(dateObj.getDate()).padStart(2,'0');
      if (phMap.has(key)) phDays.add(d);
    }
  }

  for (const day of consultantData.days) {
    if (day.month !== targetMonth) continue;

    const inSlot1 = day.slot1.some(nameMatches);
    const inSlot2 = day.slot2.some(nameMatches);
    const inSlot3 = day.slot3.some(nameMatches);
    const inCall  = day.callNames.some(nameMatches);
    const onLeave = day.leaveNames.some(nameMatches);

    if (onLeave) {
      result[day.date] = { nf:'', nt:'', of:'', ot:'', typeLabel:'Leave - Annual' };
      continue;
    }

    if (!inSlot1 && !inSlot2 && !inSlot3 && !inCall) continue;

    const isWE = day.isWeekend;
    const isPH = phDays.has(day.date);
    const isSpecial = isWE || isPH;

    // Determine time rule key
    let rule = null;
    if (isSpecial) {
      // Weekends and public holidays: OT on-site 07h30-11h30, OT off-site 11h30-07h30
      // No normal hours — this is not a standard shift
      rule = rules.weekend_ph;
    } else if (inSlot1 && inCall) {
      // On-call day: normal day shift + OT1 (15h30-16h30) + OT2 off-site (16h30-07h30)
      rule = rules.weekday_slot1_oncall;
    } else if (inSlot1) {
      rule = rules.weekday_slot1_no_oncall;
    } else if (inSlot2 || inSlot3) {
      rule = rules.weekday_slot2_3;
    } else if (inCall) {
      rule = rules.weekday_call_only;
    }

    if (!rule) continue;

    // Build shift entry from rule
    // For weekend_ph: no normal hours, OT1 = on-site, OT2 = off-site
    // For weekday rules: normal hours present, OT1/OT2 as applicable
    const t = (v) => (v||'').replace(':','H');

    if (isSpecial) {
      // Weekend/PH: no normal hours
      // OT From = 07H30 (start of on-site), OT To = 07H30 next day (end of off-site)
      // OT1 = on-site 07H30-11H30, OT2 = off-site 11H30-07H30
      result[day.date] = {
        nf: '',
        nt: '',
        of: t(rule.ot1 ? rule.ot1[0] : (rule.ot2 ? rule.ot2[0] : '')),
        ot: t(rule.ot2 ? rule.ot2[1] : (rule.ot1 ? rule.ot1[1] : '')),
        typeLabel: isPH ? 'On Call - Public Holiday' : 'On Call - Weekend',
        ot1f: t(rule.ot1 ? rule.ot1[0] : ''),
        ot1t: t(rule.ot1 ? rule.ot1[1] : ''),
        ot2f: t(rule.ot2 ? rule.ot2[0] : ''),
        ot2t: t(rule.ot2 ? rule.ot2[1] : ''),
      };
    } else {
      // Weekday: Normal = 07H30-15H30
      // OT1 = handover 15H30-16H30 (shown as OT From start)
      // OT2 = off-site overnight 16H30-07H30 (shown as OT From/To when on-call)
      const nf  = t(rule.normal ? rule.normal[0] : '');
      const nt  = t(rule.normal ? rule.normal[1] : '');
      // If there's an OT2 (on-call overnight), show OT2 in the OT columns
      // OT1 (handover, 1hr) is informational — represented by OT From = OT2 start
      const of_ = rule.ot2 ? t(rule.ot2[0]) : t(rule.ot1 ? rule.ot1[0] : '');
      const ot_ = rule.ot2 ? t(rule.ot2[1]) : t(rule.ot1 ? rule.ot1[1] : '');
      const label = inCall
        ? 'On Call - Weekday'
        : (nf ? `Consultant Day - ${nf}` : 'Normal Hours - Weekday');
      result[day.date] = {
        nf, nt, of: of_, ot: ot_, typeLabel: label,
        // Keep OT1 for Excel generator
        ot1f: t(rule.ot1 ? rule.ot1[0] : ''),
        ot1t: t(rule.ot1 ? rule.ot1[1] : ''),
        ot2f: t(rule.ot2 ? rule.ot2[0] : ''),
        ot2t: t(rule.ot2 ? rule.ot2[1] : ''),
      };
    }
  }
  return result;
}

