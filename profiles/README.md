# EC profile examples

Profiles live in the Supabase `ec_profiles` table. A `table` profile can be
built in the app — "Not listed? Set up your hospital department profile" —
which reads a sample roster (`.xlsx`, `.docx` or `.doc`), guesses what each
column means, and asks for the hours.
These files are reference copies of the same shape, for editing by hand.

| File | roster_type | Source format |
|---|---|---|
| `vhw-anaesthetics.example.json` | `table` | Excel `.xlsx`, Word `.docx` / `.doc` |

## `roster_type: "table"`

For rosters that are a grid — an Excel sheet or a Word table. The grid is
explicit in the file, so the profile only declares what the columns *mean*.
Which format carries it changes nothing but how the rows are recovered.

### `work_pattern`

Two kinds of department share this shape.

- **`"calls"`** (the default) — everyone works the ordinary weekday, and the
  roster records only who is on call. `default_weekday` supplies the implied
  week and `post_call_off` gives the day after a call off.
- **`"shifts"`** — nobody works an ordinary weekday; staff work only the
  shifts they are rostered onto. Such a profile sets `default_weekday: null`
  and `post_call_off: false`, and each duty column carries its own hours in
  `role_rules`. That is all `getTableShifts` needs: with no default weekday to
  fall back on and no post-call rule, only rostered days produce hours.

`work_pattern` is recorded for readability; the behaviour follows from
`default_weekday` and `post_call_off`.

The column declarations:

- `table.columns` — header labels, in order. Its **length is significant**:
  legacy `.doc` marks end-of-cell and end-of-row with the same byte, so the
  column count is what recovers the row boundaries.
- `table.date_col` / `table.day_col` — indexes of the date and weekday columns.
- `table.role_columns` — maps a role name to its column index. Role names are
  the join to `role_rules`.
- `table.ignore_tokens` — cell values to treat as "nobody" rather than a name.
- `role_rules[role].weekday` / `.weekend_ph` — `normal` / `ot1` / `ot2` time
  bands as `["HH:MM", "HH:MM"]`, or `null` where a band does not apply.
  Public holidays use the SA calendar in `holidays.js` and take `weekend_ph`.

### Hours in `vhw-anaesthetics.example.json`

**The roster records call only.** The ordinary working week never appears in
the file, so the parser fills it in. Each day of the month resolves in this
order:

1. **On call that day** (named in any role column) → that role's bands.
2. **Post-call** (on call the day before) → *off*. Last night's call already
   runs to 07:30 that morning, and the rest of the day is leave.
   Set `post_call_off: false` to disable.
3. **An ordinary weekday** → `default_weekday`.
4. Otherwise (weekend or public holiday, not on call) → nothing.

On call today always wins, so a consultant covering a whole week is on duty
throughout rather than post-call every second day.

| | Normal | OT1 (handover / on-site) | OT2 (call) |
|---|---|---|---|
| `default_weekday` | 07:30–15:30 | 15:30–16:00 | — |
| Call, weekday | 07:30–15:30 | 15:30–16:00 | 16:00–07:30 |
| Call, weekend / PH | — | 07:30–11:30 | 11:30–07:30 |

Weekends and public holidays follow the EC consultant split: an on-site
morning band, then off-site until the next morning. Public holidays come
from `holidays.js`, so they take the weekend bands even midweek.

Two limits worth knowing. The first of the month cannot be judged post-call,
because the previous month's roster is not loaded — it is treated as an
ordinary day. And every role carries identical bands; if a column turns out
to be daytime work rather than call — `Sessions` is the likely candidate,
since the roster marks half-days there ("Woermann AM") — change just that
role's entry.

## Presentation and form fields

Optional keys any profile can set:

| Key | Effect | Default |
|---|---|---|
| `duty_noun` | What section 2 calls a duty — "Preview & edit *calls*" | `shifts` |
| `z1_component` | Component line on the Z1(a) leave form | the profile's own `ec_short` / `ec_name`, with the hospital stripped; `Emergency Medicine` if nothing is left |
| `supervisors` | Names for the supervisor dropdown. Empty or absent gives a free-text box (except the original EC `shift` profile, which keeps its built-in list) | — |
| `leave_types` | Activity types offered alongside the roster's own duty labels | the standard leave list |
| `work_pattern` | `calls` or `shifts` — see above | `calls` |

The Z1(a) files the **department alone** — the hospital is already established
by the rest of the form — so a leading or trailing `VHW`, `VH` or
`Victoria Hospital` is stripped from whichever name is used. `VHW Anaesthetics`
and `Anaesthetics — Victoria Hospital` both come out as `Anaesthetics`. Setting
`z1_component` explicitly is still the clearest option; the fallback to the
profile's own name exists so a profile cached before this key was added still
names the right department.

Leave is classified **positively**: an activity counts as leave only if it is a
known leave type or appears in `leave_types`. Anything else is duty. Deciding
by exclusion used to misread role-prefixed labels such as
`COSMO/SN On Call - Weekday` as leave, which put call days on the leave form
and wrote the label into the timesheet instead of the hours. Only `Leave - *`
types unlock the Z1(a); a workshop or course is official duty and belongs on
Annexure C.
