# Parser fixtures

Replay tests for the shift-roster parser, built from real Victoria Hospital EC
rosters with every name replaced.

```
node tests/run.js            # check the fixtures against index.html
node tests/run.js --update   # re-record expectations from current behaviour
```

Needs Playwright's Chromium (`CHROME=/path/to/chrome` to point at a specific
binary). Nothing else — no package.json, no build step, and the app itself
still has no dependencies.

## Why there is no PDF here

The parser reads nothing from a PDF except `numPages` and, per page, a list of
`{str, x, y}`. So a fixture does not need to be a PDF — it is that list:

```json
[["WEEK 1", 89, 547], ["CONSULTANT", 240, 547], ["08:00 - 18:00", 316, 547]]
```

That matters for more than tidiness. Real rosters carry staff names, and this
repository is public. Editing a PDF to remove them is the wrong tool twice
over: **blurring the page leaves the text layer intact**, so the names are
still extractable and nothing is anonymised, while **rasterising destroys the
text layer**, leaving a file the parser cannot read at all. PDFs also keep
incremental-save history and document metadata, so a "redacted" PDF can still
carry the original text — the standard way redactions fail.

Recording the text items sidesteps all of it: no glyphs, no fonts, no
metadata, no revision history. Coordinates are the real ones, because that is
what the parser is sensitive to; only the words change.

## What a substitute has to preserve

A replacement must be indistinguishable **to the parser**, not merely
plausible. Every one of these was a real bug found while building the
generator:

| Substitute | What went wrong |
|---|---|
| `Northcote` | silently vanishes — `NOISE_RE` matches the prefix `No` |
| `Ziegler2`  | stops being a name — `isNameTok` forbids digits |
| any tidy name for `Dayar` | `Dayar` is *already* noise (`Day` prefix), so a well-formed substitute **adds** a doctor the real roster never had |
| a noise-shaped word for `Wednesday` | kept every per-item property and still broke `DATE_RE`, rerouting four days down the public-holiday path |

So `make-fixture.js` preserves **length, capitalisation and hyphen positions**
exactly, and verifies on two levels against the parser's own predicates:

- every item must classify identically (`isNameTok`, `isNoise`, compound-name
  match, `NAME_PREFIXES`), and
- every **row** must keep the same date-grammar verdict (`DATE_RE`,
  `PARTIAL_DATE_RE`, `HAS_DATE`).

The row check is the one that matters most: `Wednesday` passed every item-level
check and still corrupted the fixture. Without it the fixture would have looked
fine and tested the wrong thing.

Each fixture reproduces its source PDF exactly — same day count, same staff
count, identical shift-count multiset.

## What these tests do and do not prove

They pin **current behaviour**. They would have caught the 2023/2024 shift
undercount the moment it appeared: reverting that one fix turns
`ec-roster-2023` red with the exact shifts it loses, while 2019 — which never
had the typo — stays green.

They are a change detector, not a correctness oracle. A fixture recorded today
also enshrines whatever is still wrong. The real oracle is the tally printed on
the roster's own summary page; where that is known it is worth recording as
the expected answer rather than whatever the parser currently says.

## Adding a fixture

```
node tests/make-fixture.js some-roster.pdf     # writes some-roster.fixture.json
node tests/run.js --update                     # record its expectations
```

Check the generator's report before committing: `items classifying
differently` and `pages whose row grammar moved` must both be zero, and
`original words still present` must be `none`. **Never commit the
real→substitute mapping** — it is the re-identification key, and a fixture
plus its key is no better than the original file.

Note the generator's word list is tuned to this roster layout. A different
department's roster may use vocabulary it does not recognise as structural,
which the verification will report rather than hide.
