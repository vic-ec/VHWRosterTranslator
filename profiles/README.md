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

> **The hours in `vhw-anaesthetics.example.json` are placeholders.** They were
> copied from the VHW EC consultant pattern so the pipeline could be tested
> end to end, and have not been confirmed with the Anaesthetics department.
> Confirm them before this profile is used for a real payroll submission.
