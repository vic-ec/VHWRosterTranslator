# The four-step wizard

The app used to be one long page: upload, preview, details and downloads all
on screen at once, gated by `display:none` toggles. It is now four steps, one
visible at a time, with a persistent bar saying where you are and what you are
working on.

Nothing about parsing, hour calculation, editing, the public-holiday calendar
or document generation changed. The wizard is a shell around them.

## What is on each step

| Step | Section | What it does |
|---|---|---|
| 1 · Upload roster files | `#sec-1` | File upload and Extract data — unchanged |
| 2 · Review schedule | `#sec-2` | Staff list, month, the editable day-by-day schedule, hours summary, attention panel, review acknowledgement |
| 3 · Your details | `#sec-3` | The personal and sign-off fields — unchanged |
| 4 · Generate documents | `#sec-4` | The three document cards and the printing guide |

Each step shows a `.wiznav` footer with Back and Continue. When Continue is
disabled, the sentence beside it says why; the button points at that sentence
with `aria-describedby`, and the sentence is a `role="status"` region, so a
screen reader hears the reason rather than only a dead button.

## Components added

All of it lives in the two places every change in this repo lives — the module
file under `js/` and the matching block inside `index.html`'s inline `<script>`
(see the top of `CLAUDE.md`).

**Markup** (`index.html`)
- `.wizbar` — sticky bar under the header: the step list (`#wizSteps`, a
  `<nav aria-label="Progress">`), a mobile counter (`#wizStepMobile`,
  `aria-live="polite"`) that prints only "Step n of 4" because the section
  heading below already names the step, the context row (`#wizPeriod`,
  `#wizDept`) and an icon-only `#wizEditBtn`. It sticks at
  `top: var(--header-h)`, a value JS measures from the header and keeps
  current with a `ResizeObserver` — the header wraps to two and three lines on
  narrow screens, and at `top: 0` the bar covered it on the first scroll.
- `#totalsOverlay` — the hours breakdown, a bottom sheet
  (`.modal-overlay.is-sheet`). One row per figure, value right: days on duty,
  days on leave, normal, overtime, weekend, public holiday. A public holiday
  falling on a weekend counts once, as a holiday. On a phone this is the only
  place the month total appears — `.totals-bar` keeps its figures on a desktop
  but shows nothing except the button that opens this, and is not sticky:
  pinned to the bottom it cost a fifth of a landscape screen.
  `updatePreviewTotals()` therefore has to survive its own readouts being off
  screen, and guards every write.
- `#rosterViewOverlay` — View roster file, a wide dialog beside Preview
  schedule and Hours breakdown on step 1. It shows the file the user actually
  uploaded, so a suspect parse can be checked without leaving the page: a PDF
  is rendered page by page onto canvases at scale 2 via PDF.js, and a grid file
  (`.xlsx` / `.docx` / `.doc`) is drawn as the rows the reader recovered.
  `extractWordTables` returns *every* table it finds, so `gridRowsFor` takes
  the largest one and passes the profile's column count — a legacy `.doc`
  needs it to recover row boundaries at all. The note under the picker says
  which of the two the user is looking at, because they mean different things:
  the PDF is the source, the grid is what the parser saw. The file picker
  (`.rv-pick`) is hidden unless two or more files are loaded.
  This needs the `File` handle to survive extraction, so the objects pushed
  onto `state.parsedFiles` now carry `file`; it stays in memory only, like the
  rest of `state`. `wizRefresh()` disables the button until at least one
  retained file exists.
- `#wizEditOverlay` — the Edit panel: Change period, Change department and
  Start over. Spelling all three out in the bar cost two lines on a phone.
  Start over lives here because the wizard hides the masthead that used to
  hold it, so this is now the only way to reach it once a roster is loaded.
- `.attention` (`#attentionPanel`) — above the schedule on step 2.
- `.reviewack` (`#reviewAckWrap`) — the acknowledgement below the schedule.
- `.wiznav` (`#wizNav1`…`#wizNav4`) — the per-step footers.
- `.totals-toggle` / `.totals-detail` — the hours breakdown in the summary bar.

**Controller** (`js/ui.js`, section `WIZARD SHELL`)
- `wizExtracted()` / `wizPreviewed()` / `wizDetailsDone()` — read the existing
  state; they add no state of their own.
- `wizBlockedReason(step)` — the single source of both the disabled Continue
  button and the sentence explaining it.
- `wizGo(step)` / `wizRefresh()` / `wizRenderSteps()` / `wizRenderContext()` /
  `wizRenderNav()` — navigation and rendering.
- `buildAttentionItems()` — the reading aid described below.
- `updateTotalsDetail()` — re-splits the hours the summary bar already shows.

**One control height.** `--control-h` (38px) is the height of every row and
control that sits in a list or a form on a phone: a staff row, an activity
select, a time box, a step button, the Edit button. It is above the 24px WCAG
2.5.8 minimum target size and replaces a spread of 38-52px that made the
schedule scroll far further than it needed to. A textarea is the exception —
it is sized by its rows and has to be able to grow.

**Confirmation.** `#confirmOverlay` sits in front of Clear all, Clear
selection, removing a queued file, removing a day, and Start over. One panel,
Yes/No, closable by the x, the backdrop or Escape — all of which answer no.
`askConfirm()` closes any panel already open first, so there is never a
question stacked on top of a question.

**The staff list collapses.** Once a name is chosen the other rows are hidden
(`.doctor-grid.is-collapsed`) and a Show all / Show fewer button appears in the
head. Presentation only — every chip stays in the DOM and stays clickable once
shown.

**Phone layout.** At ≤860px the schedule stops being a table and becomes one
card per day: date and weekday as a heading, then the activity, then one line
per time band with the start and end side by side under a shared label. A day
with no duty stays a single line. The cells are found by the field they hold
(`td:has([data-field=nf])`), not by column number, so the standard and extended
column sets need no separate rules.

## Session-only state

The wizard adds exactly two variables, both module-level and neither persisted:

```js
let wizStep = 1;        // which step is on screen
let wizReviewed = false; // has the acknowledgement been ticked
```

Reloading the page puts you back on step 1 with an empty schedule, because
everything the wizard reads — `state.parsedFiles`, `state.editedShifts`,
`state.savedDetails` — has always lived in memory only.

`localStorage` still holds exactly the two configuration keys it held before:
`ec_roster_profile` and `ec_roster_profiles_list`. No roster, no schedule, no
personal detail, and no wizard position is written to any storage; there is no
`sessionStorage`, no IndexedDB and no cookie. The only outbound requests remain
the approved-profile catalogue, the setup wizard's profile submission
(structure and hours only) and the Google Fonts stylesheet — none of which
carries roster content or personal data.

## The review acknowledgement

The whole month is on screen and the user must go through it. The
acknowledgement is a plain checkbox:

> I have reviewed all days in this month, including shifts, swaps, leave and
> additional hours.

Until it is ticked, Continue on step 2 is disabled and says
"Confirm you have reviewed every day in the month."

It is never ticked programmatically, and it does not survive an edit: any
change to the schedule calls `wizInvalidateReview()`, which unticks it and
re-disables Continue — the user confirmed the month as it was, not as it now
is.

**The attention panel is a reading aid, not a substitute.** It lists what looks
odd (a band with only one of its two times, a day over 24 hours, a parser
warning) and each entry scrolls to that day. With nothing to flag it still
says the month has not been checked for accuracy. It marks nothing reviewed,
it hides nothing, and it does not shorten what the acknowledgement asserts.

## Wording that stays honest

Step 4 says the documents are built in the browser, that nothing is uploaded,
and — explicitly — that downloading a document does not submit it: the user
still sends the files to payroll or their HOD themselves. The app has never
sent anything to HR, and nothing in the redesign implies that it does.

## Preserved unchanged

- All roster parsing: `parser.js`, `parser-consultant.js`, `parser-word.js`,
  `parser-xlsx.js`.
- Hour calculation and the band logic, including the `of`/`ot` mirror of the
  OT2 band.
- Manual schedule editing: every time field, the activity dropdown, add a day
  (`+`), remove a day (`×`) and undo — all still there, on both layouts.
- Public-holiday categorisation from `holidays.js`, its tint, its footnote
  letters and its effect on which band applies.
- Document generation: `generator-excel.js` and `generator-docx.js`, and the
  Z1(a) Component derivation.
- The department profile: `ec_name` is `VHW Emergency Medicine` and that exact
  string is what the header chip and the wizard context row display.
