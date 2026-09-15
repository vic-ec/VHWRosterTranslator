# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A browser-based, offline tool that parses monthly shift roster PDFs and generates WCG-formatted duty roster timesheets, overtime verification forms, and leave application documents. Built for Victoria Hospital Emergency Centre, Cape Town, with the ability to add other Emergency Centres (ECs) via a Supabase-backed profile database.

## Running / testing

There is no build step, package manager, or test suite — this is a static site.

- Open `index.html` directly in a browser (Chrome/Edge recommended for PDF parsing; Firefox also supported), or serve the folder with any static file server.
- Deployed via GitHub Pages directly from `index.html`.
- **Every change to the app bumps the version string to that day's date**, as
  `vDD.MM.YYYY` — so an edit made on 14 September 2026 ships as `v14.09.2026`.
  It appears in exactly two places in `index.html`, the header `.brand-ver` and
  the footer `.copy`, and they must always agree; `grep -c 'v[0-9][0-9]\.' index.html`
  should return 2. Bump it for anything that changes what the app does or looks
  like — `index.html` or a `js/` module — and leave it alone for
  documentation, `CLAUDE.md`, `tests/` and `profiles/`, which ship nothing to
  the user. Treat it as part of the change, not a follow-up: the version is how
  the user tells, on their phone, whether the page they are looking at is the
  one that was just deployed.
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

**Never edit `index.html` in GitHub's web editor.** The bundle is over 1 MB, and
a web edit on 2026-09-14 (`2ee6814`, changing only the two version strings)
saved it back at exactly 1,048,576 bytes — 1 MiB precisely — cut mid-statement
inside a `$('detailSigDate')` listener, taking 2,460 lines with it: the whole
wizard shell, every overlay's handlers and both generators. It committed
cleanly and the diff read as an ordinary two-line change, so nothing announced
it; the app simply loaded as a much older version. A one-character change to
this file still has to go through a clone, or through a `js/` module plus the
matching inline block. If the file is ever a round power of two in size, it has
been truncated — check the last line is `</html>`.

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
- **The viewer's find row is sticky, and the panel's top padding is zeroed for
  it.** Stepping through matches is exactly when the next-match button is
  wanted, and it used to scroll out of the panel with the first page. Sticky
  offsets pin the **margin** box, so a negative top margin does not pull the
  bar up to the scrollport — it pushes the visible box *down* by that much and
  leaves a strip of `.modal-body`'s padding above it for pages to show
  through. `#rosterViewOverlay .modal-body { padding-top: 0 }` and `.rv-tools`
  supplies its own instead. The negative *side* margins are still wanted, or a
  page slides through the 22px gap beside it, and the bar needs an opaque
  background because pages pass behind it. `rvStep` already scrolls a match
  with `block: 'center'`, so nothing lands underneath the bar.
- **Extract data and Clear all follow the files.** `syncActionsSide()` puts the
  row under whichever upload zone holds something — the department roster wins
  when both do, a consultant-only upload moves it under the right-hand zone —
  and it is called from both list renderers and from
  `updateConsultantZoneVisibility()`, since hiding the zone has to bring the
  buttons back. The move is gated by a **container query** on `.two` at 624px
  (`2 x 300 + 24`, the width at which `auto-fit` grants a second track): a
  fixed breakpoint would only approximate it because the threshold is the
  grid's own width, and `grid-column: 2` in a one-column grid invents a track
  rather than failing, which would wreck the phone layout. Both tracks are
  equal, so the buttons keep their width exactly.
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
  writing a bare string there silently loses the short label. `Columns` is the
  longest short form and sets the floor: below 380px it needs the tracking
  gone as well to fit a 320px screen.
- **On a phone the step counter is the only way to move between steps.** The
  four `.wizstep` tabs are `display: none` below 861px, so `#wizStepMobile`
  carries `#wizJumpBtn`, which opens `#wizJumpOverlay` — a row of 1–4. Its
  buttons carry **`data-go`**, the same attribute the desktop tabs use, so the
  delegated listener in `ui.js` performs the navigation and the two routes
  cannot disagree. Gating is `wizCanEnter(n)` recomputed on open, the same rule
  as the tabs and the Continue buttons; no "visited" state exists or should be
  added. The panel also prints the first unmet precondition, which the desktop
  tabs still do not.
- **Each jump button carries its own name**, `WIZ_STEPS_DEF[].short` in a
  `.t` span under the number's `.n` span — Upload, Review, Details, Generate,
  with the full `title` still on the button's `aria-label`. The names used to
  run underneath as one paragraph, which wrapped wherever the panel ended and
  so lined up with no button in particular. `#wizJumpNote` is now *only* the
  blocked reason and is `hidden` when there is none; it needs its own
  `[hidden]` rule, and `tests/wizjump.js` reads the number from `.n` rather
  than the button's `textContent`, which now says "1Upload".
- **`#wizJumpBtn` is static markup, not rendered.** `wizRenderSteps()` runs on
  every `wizRefresh()`; it writes `textContent` into `#wizStepCount`. If the
  button were part of that string it would be destroyed and re-created, and the
  listener `dialog()` bound to it would be lost. (The old
  `<span class="t">` in that paragraph was already dead for the same reason —
  the first render removed it and nothing rebuilt it.)
- **The leave button is laid out like a `.dl` download card** — the app's other
  big block button with a right-hand icon: `justify-content: space-between`,
  a 16px gap, the two text lines in their own `.alt-route-text` column, and a
  22px Lucide external-link `.alt-route-ico` with `flex: none` so a wrapping
  note cannot squeeze it out of square. Its right inset measures 21px, not the
  20px of a `.dl`: `.btn` carries a 1px border and `.dl` does not, the same
  offset-by-one the nav chevrons' 19px already accounts for.
- **The leave button is a cell of a grid that copies `.two`.** `.alt-route`
  declares the same `repeat(auto-fit, minmax(300px, 1fr))` and 24px gap as the
  upload row below it, with the button in `grid-column: 1` and an empty
  `.alt-route-spacer` holding the second track open, so the button is exactly
  as wide as the department upload zone at every width. A fixed width or a
  breakpoint gets two cases wrong that this one handles for free: a wide
  window, where `auto-fit` collapses its empty tracks so two cells stay half
  each however much room there is (`auto-fill` would have kept splitting and
  left the button a quarter wide), and a profile with no consultant roster,
  where `updateConsultantZoneVisibility()` hides the spacer alongside
  `#consultantZoneWrap` and both grids fall to one full-width column. The
  spacer has to be an element: `auto-fit` collapses a track with nothing in it,
  which would stretch the button back across the page.
- **The leave button is the first thing in `#sec-1`**, above the section head,
  and `#altRoute` is toggled alongside `#step1` in `showEcSelected` /
  `showEcPicker` / `reopenEcPicker` — a Z1(a) needs a department for its
  Component line and supervisor list, and every section renders briefly at boot
  before the picker hides them. Its rule sets `padding` and so must outrank the
  phone-wide `.btn.btn { padding-block: 0 }` floor: it is written
  `.alt-route .btn.btn`, the same specificity trap `.actions-row` hit.
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

## The leave-only route

A doctor can produce a Z1(a) without a roster: **Leave form only** at the foot
of step 1 opens `#z1LeaveOverlay`, which collects one leave period and
downloads the form. A department must be chosen first, because that is what
makes `z1ComponentFor()` and the supervisor list right — but nothing is
uploaded, parsed or reviewed.

- **The generator takes the rows directly.** `generateZ1ADocx` accepts an
  optional `d.leaveRows` — `{type, startDate, endDate, count, specify}` with
  dates already `DD/MM/YYYY` — in preference to deriving them from
  `editedShifts`. Do not try to synthesise `editedShifts` instead: the roster
  path builds both dates from one `d.month`, so a period crossing a month
  boundary cannot be expressed that way at all.
- **The unit belongs to the row, not the leave.** The form has two Section A
  blocks: working days for most types, **calendar days** for Unpaid, and
  **calendar months** for Maternity (the one type rendered by `calRow`). The
  panel's count field switches label and calculation with the type; get it
  wrong and a working-day count lands under "Number of Calendar Days".
- **Study is a Special Leave** on the printed form, with the kind written on
  the "Specify Type of Special Leave" line. Study and Special therefore share a
  label, and their periods interleave by date under it rather than one silently
  replacing the other.
- **A leave type prints one row per period, not one row per type.** On the
  roster path `leaveData` splits each type into contiguous blocks: a block
  continues across a gap only when every day in it is one the doctor would not
  have worked anyway — a weekend or a public holiday from `buildPHCalendar`.
  So a fortnight stays one row while two separate weeks become two, and annual
  leave on the 3rd and again on the 27th no longer prints as a single 25-day
  span. `leaveMap[label]` is therefore a **list**, and `leaveRows4` /
  `calRows` emit one row each (falling back to the form's single blank row when
  a type is unused), which is why every call site spreads them.
- **The count auto-fills but is editable.** `data-auto` on `#z1lDays` tracks
  whether the doctor has taken it over; blanking the box hands it back. The
  automatic figure is often wrong for an EC, where a weekend day is a working
  day, which is the whole reason it is editable.
- **`z1LeaveWire()` is called from `wire()`, not at load.** The overlay markup
  sits ~300 lines *after* the inline script, so at script-execution time none
  of its elements exist — and the `?.` in those listeners would attach nothing,
  silently. Section 03's listeners can be top-level only because `#sec-3`
  precedes the script.
- **No designation field.** `generateZ1ADocx` never reads `d.designation`
  (Annexure C does). Asking for it would be a required field that changes
  nothing in the output.
- `supervisorEls`/`readSupervisor`/`showSupervisorBox`/`applySupervisorMode`/
  `setSupervisorValue` all take an optional element prefix, defaulting to
  `detail`, so the panel reuses section 03's dropdown-vs-free-text rule instead
  of copying it.
- **`state.savedDetails` now declares `address`,** the key that was always
  written and read; the initialiser used to declare a dead `addressDuringLeave`
  and so dropped the real one on reset. `shiftWorker` and `casualEmployee` are
  gone from it entirely — see the note below on why the Z1(a) always says No. `fullReset()` empties the object too —
  both forms prefill from state directly, so clearing the DOM boxes is no
  longer enough for Start over to mean it. (`d.addressDuringLeave` remains the
  *generator's* parameter name.)
- **The panel stores itself in `state.leaveDetails`, never in
  `state.savedDetails`.** `z1LeaveOpen()` reads the one and `z1LeaveSaveOwn()`
  writes it; nothing in the panel touches section 03's boxes. Section 03
  belongs to the doctor whose roster is on screen and this belongs to whoever
  is applying — the panel used to copy its fields across "so the roster route
  needs no retyping", and the result was that filling in a leave form and then
  previewing a colleague put the applicant's name, PERSAL, supervisor and
  address on *their* Annexure C, with only `detailSurname` even hinting at it
  (it falls back to `state.selectedDoctor` only when the saved surname is
  empty). `fullReset()` clears both objects. `tests/leave-ui.js` asserts the
  separation in both directions.

## Names on the roster vs names on the form

A roster tells two doctors of the same surname apart by prefixing a first
initial — `M. Willemse` beside `J. Willemse`. `splitRosterName()` sends that
initial to the first-name box and the rest to the surname box, so section 03
reads as a name rather than putting `M. Willemse` under Surname.

- **Only a single letter followed by a full stop is an initial**
  (`/^((?:[A-Z]\.[ \t]*){1,3})([A-Za-z].*)$/`, so up to three of them). That is
  what leaves every other shape alone: `Van Schalkwyk`, `Du Toit`, `Le Roux`
  and `Gordon-Forbes` carry no full stop, and `St.` is two letters. A looser
  rule — splitting on the first space, say — would corrupt every compound
  surname in the department.
- **`state.selectedDoctor` is never rewritten.** It is the key the schedule is
  looked up by and the label that tells the two Willemses apart in the staff
  list; only the details boxes get the split. Dropping the initial instead
  would have made the two indistinguishable on their own paperwork.
- **The `isNewDoctor` branch of `restoreDetailsToForm` is not the one that
  seeds a fresh doctor.** Clicking a chip restores only while `#detailsSection`
  is already visible, which it is not the first time, so Preview's
  `restoreDetailsToForm(false)` is what fills the boxes — both branches
  therefore fall back to the split, and a name the doctor has typed still wins.
- `tests/names.js` pins all of this, the compound surnames especially.

**The parser reads an initial in either of the two layouts a PDF can produce.**
`M. Willemse` may arrive as one text item or as `M.` followed by `Willemse`,
depending on the export. `INITIAL_RE` has always read the single-item
spelling; `INITIAL_TOK_RE` reads the other, joining a lone initial to the name
token on its right within the same 120px the compound-name join uses. Both
name paths do it — the shift columns in `extractNamesWithAnchors`, and the
leave column, which had to be sorted by x first so a lookahead means anything.

Before that join the lone `M.` failed `isNameTok` (a full stop is outside
`[A-Z][a-zA-Z\-]{1,14}`), was skipped, and the two doctors **merged into one
`Willemse` carrying both their shifts** — one person's hours silently doubled
and the other gone from the roster entirely. This follows the rule the
semicolon fix set: accept one more spelling of something the parser already
understands, never loosen what a match means. `isNameTok` is unchanged, and
all three real rosters parse byte-identically before and after, so a roster
without initials cannot have been affected. `tests/initials.js` replays a real
fixture in both layouts; it fails on the pre-fix bundle for the split layout
only, which is exactly the shape of the bug.

## A consultant roster on its own

A consultant on-call PDF is a complete roster, not only a supplement to a
department one. `parseAndStoreConsultantRoster()` has always had the branch
for it: with no department file it builds `rosterData` from the consultant
days, fills the staff list, sets the month and year, unlocks step 2 and
rebuilds the month dropdown — and `buildPreview` reads consultant days rather
than `rosterData` whenever `isExtendedRosterMode()` is true. Everything
downstream worked.

- **Two gates were all that blocked it.** `wizExtracted()` counted only
  `state.parsedFiles` and `state.tableData`, so step 1 never completed; and
  the `wizRefresh()` after extraction ran *before* the consultant parse, so
  even once the first was fixed Continue stayed disabled over a staff list the
  app had already drawn. `wizRefresh()` now runs after both branches of the
  parse. Verified end to end on three real consultant rosters.
- **The month and year come from the file name**, in
  `parseAndStoreConsultantRoster()` — the grid itself carries no month, and
  `getConsultantShifts` skips every day whose `month` does not equal the target
  (`day.month !== targetMonth`). A file renamed without its month name falls
  back to *today's* month and silently matches nothing. Worth remembering when
  testing: `c-jul2026.pdf` parses as September.
- **Cells naming an event and two consultants become one junk name.** The May
  2026 roster yields `Cloete & Els`, `SAPA Cloete & Els`, `Clin Gov Cloete &
  Els` and `Retreat Cloete & Els` beside the four real consultants; each
  matches no day, so picking one gives an empty month. Splitting on `&` would
  credit both consultants with those days, which changes hours and is a policy
  question, not a parsing one — so nothing splits them yet.

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

## Tests

`tests/` holds replay fixtures for the shift parser and a runner —
`node tests/run.js`, Playwright's Chromium the only requirement. Beside it:
`tests/z1a.js` (the leave form's rows), `tests/leave-ui.js` (the leave-only
panel end to end), `tests/names.js` (splitting a roster initial off a surname),
`tests/initials.js` (the parser reading one in either PDF layout) and
`tests/wizjump.js` (the phone step jump), `tests/actionsrow.js` (which zone
Extract data sits under), `tests/rosterview.js` (the viewer's sticky find
row — needs `ROSTER=/path/to/a/roster.pdf`, and skips without it) and
`tests/consultant.js` (a consultant roster standing alone — its first half
drives the gate with no file at all, the rest runs a real upload when
`CONSULTANT_ROSTER` points at one). It is a
dev-only tool: the app still has no build step and no dependencies.

A fixture is not a PDF. The parser reads nothing from one but `numPages` and,
per page, a list of `{str, x, y}`, so the fixture *is* that list with every
name replaced — no glyphs, fonts, metadata or incremental-save history to leak
a real roster into a public repository. Coordinates are the real ones, since
that is what the parser is sensitive to.

A substitute has to be indistinguishable *to the parser*: `make-fixture.js`
preserves length, capitalisation and hyphen positions, and verifies both that
every item classifies identically (`isNameTok`, `isNoise`, compound match) and
that every row keeps its date-grammar verdict (`DATE_RE`, `PARTIAL_DATE_RE`,
`HAS_DATE`). The row check is the one that catches the subtle cases — a noise-
shaped stand-in for "Wednesday" passed every item-level check and still broke
`DATE_RE`. Never commit the real→substitute map: it is the re-identification
key. `tests/README.md` has the whole picture, including what these tests do
*not* prove.

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
- **The Z1(a) declares No for Shift Worker and Casual Employee, always.** Both
  rows are ticked in the No box by `RIGHT_YN` in `generator-docx.js`, on every
  route. This is policy, not a missing feature — it was once a parameter fed
  from `#detailShiftWorker` / `#detailCasualEmployee`, two elements that never
  existed in the DOM, so the value was always its default anyway. The leave
  panel used to offer dropdowns for both; they were removed because the
  generator dropped the answers on the floor, and two controls that change
  nothing are worse than none. `tests/z1a.js` asserts the ticked row on both
  routes, so the question cannot go unanswered again.
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
