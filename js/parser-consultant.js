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

async function parseConsultantRosterPDF(arrayBuffer, profile) {
  const cols  = profile.pdf_columns;   // {slot1,slot2,slot3,meetings,leave,call}
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
    const fused = item.str.match(/^(\d{1,2})([A-Z][a-z]+)$/);
    if (fused) {
      const x = item.transform[4];
      const y = H - item.transform[5];
      words.push({ text: fused[1], x, y });
      words.push({ text: fused[2], x: x + 8, y });
    } else {
      words.push({ text: item.str.trim(), x: item.transform[4], y: H - item.transform[5] });
    }
  }

  // ── Find data start row: first row whose y >= profile.data_start_y ──
  const DATA_Y = profile.data_start_y || 188;
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

  // Join compound surnames: "De" + "Haan" → "De Haan"
  // Also handles slash pairs: "Cloete" + "/" + "Els" → "Cloete/Els"
  function joinTokens(tokens) {
    const result = [];
    let i = 0;
    while (i < tokens.length) {
      const t = tokens[i];
      // Check if next token is "/" then a name (slash pair)
      if (i + 2 < tokens.length && tokens[i+1] === '/') {
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

  // ── Parse each row ──
  const days = [];
  const doctors = new Set();
  let currentDate = null, currentDayName = null, currentMonth = null;

  for (const rowWords of rows) {
    // Bucket words by column
    const buckets = { slot1:[], slot2:[], slot3:[], meetings:[], leave:[], call:[] };
    let dateNum = null, weekdayName = null;

    for (const w of rowWords) {
      const c = colFor(w.x);
      // Date column (x < slot1.x_min)
      if (w.x < cols.slot1.x_min) {
        if (/^\d{1,2}$/.test(w.text)) dateNum = parseInt(w.text);
        else if (WEEKDAYS.has(w.text) || WEEKENDS.has(w.text)) weekdayName = w.text;
      } else if (c && c in buckets) {
        buckets[c].push(w.text);
      }
    }

    if (dateNum !== null && weekdayName) {
      currentDate    = dateNum;
      currentDayName = weekdayName;
    }
    if (currentDate === null) continue;

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
      // Expand slash pairs
      for (const n of name.split('/')) {
        const trimmed = n.trim();
        if (trimmed.length > 1) doctors.add(trimmed);
      }
    }
    for (const name of callNames) {
      for (const n of name.split('/')) {
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
      allNames: [...new Set([...allSlotNames.flatMap(n=>n.split('/')).map(n=>n.trim()),
                             ...callNames.flatMap(n=>n.split('/')).map(n=>n.trim())])],
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
    return name.split('/').some(n => {
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

