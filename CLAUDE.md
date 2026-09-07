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
- Page structure: sticky header (brand + EC chip + Your data is safe) → `.wizbar` (step list + period/department context) → masthead, shown only before an EC is chosen → numbered sections `#sec-1`…`#sec-4` → footer. The old progressive-disclosure gates (`#step1`, `#step2`, `#detailsSection`) are still toggled by the app script; a small `MutationObserver` at the end of `ui.js` mirrors those toggles onto the placeholder blocks (`#step1Empty`, `#step2Empty`, `#sec3Empty`).

## The four-step wizard

`#sec-1`…`#sec-4` are one wizard, one step on screen at a time — the others
carry the `hidden` attribute. The shell is the `WIZARD SHELL` block at the end
of `ui.js` (and its inline twin). See `docs/wizard.md` for the whole picture;
the parts worth knowing before editing:

- **`wizBlockedReason(step)` is the single source of truth** for whether the
  step's Continue button is disabled *and* for the sentence saying why. Add a
  new precondition there, not in two places.
- **Only two variables belong to the wizard**, `wizStep` and `wizReviewed`,
  both session-only. Nothing about position, roster or details is persisted;
  `localStorage` still holds only `ec_roster_profile` and
  `ec_roster_profiles_list`. Do not add a third.
- **The review acknowledgement (`#reviewAck`) may never be ticked in code**,
  and any schedule edit must clear it — `markDirty()` calls
  `wizInvalidateReview()` for exactly this. The attention panel is a reading
  aid: it flags what looks odd, marks nothing reviewed, and never replaces
  looking at the whole month.
- **Anything that changes what the user is looking at calls `wizRefresh()`.**
  Extraction and preview already do.
- **`--control-h` is the one height** for every row and control in a list or
  form on a phone. Do not reintroduce a per-component `min-height` — a
  component that sets its own `padding` (as `.actions-row` did) outranks a
  plain `.btn` rule and silently opts out of it.
- **The date-of-signature icon is a real `<input type="date">`** sitting under
  the glyph at full size, not a button calling `showPicker()` on a hidden one
  — a browser will not open a picker for an input that is 0x0 and
  `pointer-events: none`.
- **`.topbar` is the one sticky element.** It wraps the header and the wizard
  bar; neither is sticky on its own. Two stacked sticky boxes — the second
  offset by a measured header height — came apart under an iOS over-scroll and
  both slid away.
- **The EC setup wizard uses the app's own chrome.** Its close button carries
  `.modal-close` like every other panel — only the ink is overridden, since its
  head is an accent band — and its Back/Next carry the step nav's chevrons at
  the same 19px inset and one `--control-h`. `#wizBackBtn`/`#wizNextBtn` need a
  `min-width` for that: the step nav keeps a short label clear of its chevron
  via its 264px, which will not fit two-up inside a modal on a phone.
- **Extract data and Clear all are a cell of the upload grid.** The row used
  to sit outside `.two`, so it was as wide as both upload zones; as a grid item
  with `grid-column: 1` it tracks the department zone through every reflow,
  including the case where the consultant zone is hidden and `auto-fit`
  collapses to one full-width column. Its `margin-top` is zeroed there — the
  grid's own `gap` provides the space.
- **Hairlines between cells are borders, not a 1px grid gap.** The printing
  steps drew their rules as `gap: 1px` over a divider-coloured container. With
  `auto-fit` the tracks take fractional widths, and a gap whose two edges round
  onto the same device pixel paints nothing — so rules went missing and
  flickered while the window resized. Fixed column counts per breakpoint plus a
  `border-left` per cell fixes both: a border always paints, and a known column
  count is the only way to identify the first cell of a row (`:nth-child(3n+1)`)
  so its rule can be dropped. `auto-fit` never allows that.
- **The setup wizard's tabs carry two labels.** `.t-full` and `.t-short` sit
  side by side and CSS picks by width — four full names were being cut off
  mid-word on a phone. Tabs 2 and 3 are rewritten in JS per roster type, so
  those writes go through `wizTabLabel(n, full, short)` and emit both forms;
  writing a bare string there silently loses the short label.
- **Destructive controls go through `askConfirm()`.** A capture-phase listener
  matches `CONFIRM_ACTIONS` and asks before the real handler runs, then
  re-sends the click. It ignores untrusted events on purpose: the app presses
  these buttons itself (`fullReset()` clicks Clear all), and asking there left
  the page inert waiting on an answer nobody could give.
- **Start over is in the header on a desktop** (`#hdrResetBtn`) and in the
  Edit panel on a phone (`#wizStartOver`); each width shows exactly one of
  them. The masthead button all three share a handler with is hidden from the
  moment a department is chosen.
- **The scroll lock goes on `<html>`, not `<body>`.** `scrollbar-gutter:
  stable` is declared on `html`; body's overflow only reaches the viewport by
  propagation, and once it does, the gutter reserved on `html` is released and
  the whole layout shifts by the scrollbar's width when a modal opens.
- **`.modal-choice[hidden]` needs its own rule.** The class sets
  `display: flex`, which outranks the browser's rule for `[hidden]`, so
  setting the attribute alone left both choices on screen.
- **The period and department are header controls on a desktop**
  (`#hdrPeriodBtn`, `#hdrDeptBtn`), each opening `#wizEditOverlay` filtered to
  its own action by `showEditChoices()`. The bar's `.wizctx` carries them on a
  phone instead and is hidden above 860px.
- **A wizard-built profile can carry `supervisors`.** The setup wizard asks for
  them (`#wizSupervisors`, one name per line) and both profile builders attach
  the list when there is one. Without it every wizard profile fell through to
  the free-text box, which read as the dropdown being broken.
- **`showEcPicker()` hides `#hdrCtx` as well as the bar.** The header context
  mirrors the wizard bar, and everywhere else that pairing is maintained by
  `wizRefresh()`. This is the one place that hides the bar without it, and
  `reopenEcPicker()` is async — so a caller cannot sync the header itself, its
  line would run first. Before this, Start over left the header showing the
  period and department of the roster just cleared.
- **`.shell` is the flex column that makes the footer sit on the bottom edge.**
  Not `body` — making the body the flex container changes how its children
  resolve their width. `.shell > .wrap` needs an explicit `width: 100%`
  because its auto side margins otherwise stop it stretching.
- **The phone schedule (≤860px) is a card per day**, and its cells are
  selected by the field they hold (`td:has([data-field=nf])`), not by column
  number — so the standard and extended column sets share one set of rules.
  Adding a column needs no new CSS; adding a *band* does.

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
- **The offline VHW fallback and the Supabase row are the same profile.**
  `VHW_FALLBACK_PROFILE` (in `js/parser.js` and its inlined twin) is a
  field-for-field copy of row 1's `profile` jsonb in `ec_profiles`, which is
  what `fetchProfiles()` selects. Verified equal end-to-end: with the network
  blocked, the `activeProfile` the app applies matches the row on all eight
  fields. If that row is ever edited, edit the fallback to match, in both
  places. Note the table also has a `pdf_columns_v2` column, holding different
  coordinates and an extra `leave_weekend` key: it belongs to a dual-template
  experiment added on 2026-06-26 (`f0cb8eb`) and reverted the next day
  (`3d88465`, "revert to working consultant roster parser"). Nothing reads it
  — neither the select list nor any parser — so it is stale data, not part of
  the profile, and the fallback deliberately does not carry it.
- **Viewing the uploaded file back.** `View roster file` (step 1, beside
  Preview schedule) reopens whatever was uploaded — a PDF through PDF.js onto
  canvases, a grid file through `extractWordTables` — so a wrong-looking parse
  can be inspected in the app. It relies on `state.parsedFiles[].file` keeping
  the `File` handle after extraction; nothing is written to storage. Note that
  `extractWordTables` returns an array of *tables*, not rows: `gridRowsFor`
  picks the largest and passes the profile's column count, which a legacy
  `.doc` needs to find row boundaries at all. The viewer has a **find box**
  (`#rosterViewFind`, prefilled with `state.selectedDoctor`) that highlights
  every occurrence of a name and steps through them. Three things about it:
  a PDF page is indexed *once*, at render time, into a string plus a map back
  to the text item and character each position came from — so searching never
  re-reads the file or re-paints a page, and a name split across two text
  items still matches; highlight boxes are percentages of the viewport in a
  layer over the canvas, so they hold as it scales to the panel width; and
  spaces in the query are loosened to `\s*`, because a PDF may set "De Haan"
  as `De` + `Haan` with no space character between them. Nothing is "current"
  until the user steps, so the panel does not jump while they type. It opens
  from three places that share one preparer (`rosterViewOpen`): the button
  beside Preview schedule, `#hdrViewBtn` in the header (desktop) and
  `#wizViewBtn` in the wizard bar (phone) — both inside the sticky `.topbar`,
  so the file stays reachable while scrolling the schedule. The two icons are
  *hidden* until a file is retained rather than disabled, which needs explicit
  `.hdr-icon[hidden]` / `.wizctx-edit[hidden]` rules: both classes set
  `display: inline-flex`, which outranks the UA rule for `[hidden]`.
- **The EC roster template has a typo the parser has to tolerate.** Its first
  time band reads `08;00 - 18:00` — a semicolon — in at least the 2023 and 2024
  exports. `TIME_TOK_SINGLE`/`TIME_TOK_RANGE` therefore accept `;` alongside
  `:` and `h`. Without that the token did not look like a time, `findAnchors()`
  located three of the four columns instead of four, and every name in the
  08:00 column fell outside `maxDist` and was silently dropped — a ~30% under-
  count of shifts. Widening the class cannot change a correctly typed roster,
  since those tokens already matched: A/B on real files showed the 2019 export
  (proper colons) byte-identical, while 2023 went from 0/16 to 16/16 matching
  the roster's own printed shift tally. **This is the pattern to follow for
  layout drift**: accept an extra spelling of something the parser already
  understands, never loosen what a match means.
- **A PDF with no text throws rather than returning nothing.** A scan or photo
  has no text layer, and the old code reported `✓ 0 days · 0 staff` in green
  over it. `parseRosterPDF` tracks whether any page had words and throws if
  none did, and the parse loop keeps the first error's message so the status
  line says what was wrong instead of `1 error(s)`.
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
