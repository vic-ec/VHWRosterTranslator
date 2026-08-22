# EC profile examples

Profiles live in the Supabase `ec_profiles` table — these files are reference
copies showing the shape of each `roster_type`, so a new EC can be set up by
copying and editing rather than starting from a blank JSON.

| File | roster_type | Source format |
|---|---|---|
| `vhw-anaesthetics.example.json` | `table` | Word `.docx` / `.doc` |

## `roster_type: "table"`

For rosters that are a Word table. The grid is explicit in the file, so the
profile only declares what the columns *mean*:

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

Confirmed with the department: it is an on-call roster, so appearing in any
role column means being on call that day.

| | Normal | OT1 (handover / on-site) | OT2 (on call) |
|---|---|---|---|
| Weekday | 07:30–15:30 | 15:30–16:00 | 16:00–07:30 |
| Weekend / PH | — | 07:30–11:30 | 11:30–07:30 |

Weekend and public holidays follow the EC consultant split: an on-site
morning band, then off-site until the next morning. Public holidays are
detected from `holidays.js`, so they take the weekend bands even midweek.

Every role currently carries identical bands. If a column turns out to be
daytime-only rather than call — `Sessions` is the likely candidate, since
the roster marks half-days there ("Woermann AM") — change just that role's
`weekday`/`weekend_ph` entry; nothing else needs touching.
