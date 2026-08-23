# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A browser-based, offline tool that parses monthly shift roster PDFs and generates WCG-formatted duty roster timesheets, overtime verification forms, and leave application documents. Built for Victoria Hospital Emergency Centre, Cape Town, with the ability to add other Emergency Centres (ECs) via a Supabase-backed profile database.

## Running / testing

There is no build step, package manager, or test suite — this is a static site.

- Open `index.html` directly in a browser (Chrome/Edge recommended for PDF parsing; Firefox also supported), or serve the folder with any static file server.
- Deployed via GitHub Pages directly from `index.html`.
- Verify changes manually in-browser: upload a roster PDF, walk through doctor/month/year selection, edit the preview table, and download each output document (Excel, Annexure C, Z1(a)).

## Architecture — the critical thing to understand

**`index.html` is a single-file bundle.** Everything the app needs — HTML, CSS, and nearly all JS — lives inline in that one file. The `js/` folder (other than the two vendor libraries) is a set of **modular source files that are manually concatenated into `index.html`'s inline `<script>` block**, with their header comments stripped:

- `js/config.js` → app constants, Excel template (base64), Supabase credentials, shift/activity maps, `state` object
- `js/ec-profiles.js` → EC profile definitions
- `js/holidays.js` → SA public holiday calendar
- `js/parser.js` → roster PDF parsing (coordinate-based extraction via PDF.js)
- `js/parser-consultant.js` → consultant-roster PDF parsing
- `js/parser-word.js` → Word roster table parsing (`.docx` via JSZip/OOXML, legacy `.doc` via an OLE + piece-table reader)
- `js/parser-xlsx.js` → minimal `.xlsx` reader on JSZip, feeding `parseRosterExcel`
- `js/generator-excel.js` → duty roster `.xlsx` generation (JSZip over the base64 template)
- `js/generator-docx.js` → Annexure C / Z1(a) `.docx` generation (docx.js)
- `js/ui.js` → wizard/step UI, doctor grid, preview table, event wiring
- `js/consultant.js` → consultant file upload/merge logic

Only two vendor bundles are loaded via `<script src>` in `index.html`: `js/pdf.min.js` and `js/jszip.min.js` (plus `js/pdf.worker.min.js`, loaded by PDF.js itself). Everything else runs from the inline script.

**Because of this, the `js/*.js` module files are not actually loaded at runtime by `index.html`.** When changing app logic (not just markup/CSS), the corresponding code exists in two places that must be kept in sync:
1. the module file under `js/` (the readable, commented source), and
2. the matching block inside `index.html`'s inline `<script>` (uncommented, what actually ships).

If you only edit one, the app's real behavior (driven by `index.html`) won't change, and/or the module source will drift out of sync. When asked to fix or add functionality, check whether the same logic block exists in both places and update both — don't assume editing `js/` alone is sufficient.

## Design system — "Modernist"

The interface is built on the Modernist design system: flat and architectural, headings and label-as-object in Inter Tight 800, prose in IBM Plex Sans 400, near-mono ink (`#201e1d`) on bone (`#f3f2f2`) with a single royal-blue accent (`--color-accent-500`, `#3560db`), **zero corner radius**, and 2px rules between major sections instead of cards or shadows. Error state is the one thing outside the accent: a three-step `--color-danger-*` ramp, used for nothing else.

- The token sheet and component layer live at the top of `index.html`'s inline `<style>` block: `:root` custom properties (`--color-*`, `--font-*`, `--space-*`, `--radius-*`, `--shadow-*`), then the component classes (`.btn`, `.input`, `.field`, `.tag`, `.table`), then the app layout.
- **Take every colour, font and spacing value from the tokens** — never hard-code a hex, a font name or a radius. Ramp steps (`--color-neutral-100…900`, `--color-accent-100…900`) exist for tints, hovers and pressed states.
- A block of legacy aliases (`--bg`, `--surface`, `--border`, `--text-muted`, `--accent-mid`, `--warn`, `--success`, `--sans`, `--mono`, `--radius`) maps the old variable names — still emitted by some template strings in the app script — onto Modernist tokens. Don't add new uses; prefer the `--color-*` tokens.
- Rules that must hold: no rounded corners, no centred hero copy, hovers and pressed states come from the accent ramp, and keyboard focus is the 2px accent `:focus-visible` ring.
- Page structure: sticky header (brand + EC chip + Start over) → masthead (hero + process list, and the EC picker before one is chosen) → numbered sections `#sec-1`…`#sec-4` → footer. The progressive-disclosure gates (`#step1`, `#step2`, `#detailsSection`) are still toggled by the app script; a small `MutationObserver` at the end of `ui.js` mirrors those toggles onto the placeholder blocks (`#step1Empty`, `#step2Empty`, `#sec3Empty`).

## Output documents

| Document | Format | Purpose |
|---|---|---|
| Duty Roster | `.xlsx` | Monthly timesheet submitted to payroll |
| Annexure C | `.docx` | Overtime hours verification for HOD sign-off |
| Z1(a) Leave Form | `.docx` | Official WCG leave application |

## Roster types

`activeProfile.roster_type` selects the parsing path:

| `roster_type` | Input | Parser | Shift lookup |
|---|---|---|---|
| `shift` | PDF / Excel | `parser.js` | `getDoctorShifts` (hardcoded shift bands) |
| `consultant` | PDF | `parser-consultant.js` | `getConsultantShifts` (profile `time_rules`) |
| `table` | Excel `.xlsx`, Word `.docx` / `.doc` | `parser-word.js` | `getTableShifts` (profile `role_rules`) |

`consultant` and `table` both emit normal + OT1 + OT2 bands, so they share the
wider preview layout via `isExtendedRosterMode()`. A Word table carries its
grid explicitly, so `table` profiles declare what columns *mean*, not where
they are — see `profiles/README.md`.

## Notes

- **A `table` roster is a grid, whatever file carries it.** `gridKind()` picks
  the reader by extension first (an `.xlsx` and a `.docx` are both zips, so the
  PK magic cannot tell them apart), and `extractWordTables` returns rows for
  all three of `.xlsx`, `.docx` and `.doc`. Both the wizard's sample reader
  (`detectWordTable`) and the runtime parser (`parseWordRosterTable`) go
  through it, and `ui.js` routes an uploaded `.xlsx` to the table parser when
  the active profile is `table` rather than to `parseRosterExcel`. A grid
  profile therefore needs no coordinate calibration — only PDF does.
- **Shifts on the grid path.** `getTableShifts` already supports a shift
  department with no extra code: set `default_weekday: null` and
  `post_call_off: false`, and give every duty column its own hours. The wizard
  captures this as `work_pattern: 'shifts'`. Note this is *not* the legacy
  `roster_type: 'shift'`, which is the coordinate-parsed Victoria Hospital EC
  export with shift bands hardcoded in `holidays.js` / `config.js`; that path
  is unchanged and still fits only a roster with VHW EC's layout and times.
- All PDF parsing is coordinate-based against PDF.js text positions, since roster layouts vary — logic in `parser.js`/`parser-consultant.js` (and their inlined counterparts) is layout-sensitive. Word parsing is not: `parser-word.js` reads real cell boundaries and needs no calibration.
- **The app no longer bundles SheetJS.** `js/xlsx.full.min.js` had been a GitHub error page rather than a library, so `window.XLSX` was never defined and Excel upload always failed. It is replaced by `js/parser-xlsx.js`, a small reader over JSZip — an `.xlsx` is a zip of XML, so no extra dependency is needed. SheetJS was not restored because the newest version obtainable from npm is 0.18.5 (March 2022), which carries two unfixed high-severity parsing advisories (GHSA-4r6h-8v6p-xvw6, GHSA-5pgg-2g8v-p4x9); the patched builds are only on `cdn.sheetjs.com`. Consequence: legacy binary `.xls` is no longer accepted — the upload zone takes `.pdf,.xlsx,.docx,.doc`, and an `.xls` gets a message telling the user to Save As `.xlsx`.
- `parseRosterExcel` is **async** (it awaits `readXlsxSheets`); call sites must `await` it.
- Date cells render as `D Month YYYY` rather than SheetJS's locale-dependent `M/D/YY`. Text cells — which is what the roster parsers actually match on — are byte-identical to SheetJS's `sheet_to_json({header:1, raw:false})`.
- The EC setup wizard's overlay must stay a direct child of `<body>`. It previously sat inside `#detailsSection`, which is `display:none` at the EC-picker step, so the modal could never render.
- The SA public holiday calendar (`holidays.js`) is a deterministic, hardcoded calendar — no network calls, works fully offline.
- EC profile submission (the in-app wizard for adding a new EC) writes to Supabase; credentials for this live in `config.js`/the inlined config block.
- **What persists, and what the privacy note in section 01 promises.** Roster
  files, parsed shifts, staff names and the personal details typed into
  section 03 live only in the `state` object — they are never written to
  storage and never sent anywhere; all three documents are generated in the
  page (the wording lives in the `#privacyOverlay` panel, opened by the
  header chip `#privacyBtn`). `localStorage` holds exactly two keys, both of
  them configuration:
  `ec_roster_profile` (the chosen department profile) and
  `ec_roster_profiles_list` (a cache of the public catalogue). Neither is ever
  removed by "Change", but "Start over" (`fullReset`) removes
  `ec_roster_profile` and reopens the picker; the catalogue cache is left
  alone, since it is the public department list and holds nothing about the
  user. Outbound requests are a GET of the
  approved-profile catalogue on load, the wizard's POST of a new profile
  (structure and hours only, no roster content), and the Google Fonts
  stylesheet. Keep it that way, or change the panel copy to match.
