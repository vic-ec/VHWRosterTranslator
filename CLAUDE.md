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

**`index.html` is a single-file bundle.** Everything the app needs — HTML, CSS, and nearly all JS — lives inline in that one file. The `js/` folder (other than the three vendor libraries) is a set of **modular source files that are manually concatenated into `index.html`'s inline `<script>` block**, with their header comments stripped:

- `js/config.js` → app constants, Excel template (base64), Supabase credentials, shift/activity maps, `state` object
- `js/ec-profiles.js` → EC profile definitions
- `js/holidays.js` → SA public holiday calendar
- `js/parser.js` → roster PDF parsing (coordinate-based extraction via PDF.js)
- `js/parser-consultant.js` → consultant-roster PDF parsing
- `js/generator-excel.js` → duty roster `.xlsx` generation (SheetJS)
- `js/generator-docx.js` → Annexure C / Z1(a) `.docx` generation (docx.js)
- `js/ui.js` → wizard/step UI, doctor grid, preview table, event wiring
- `js/consultant.js` → consultant file upload/merge logic

Only the three vendor bundles are loaded via `<script src>` in `index.html`: `js/pdf.min.js`, `js/xlsx.full.min.js`, `js/jszip.min.js` (plus `js/pdf.worker.min.js`, loaded by PDF.js itself). Everything else runs from the inline script starting around `index.html:644`.

**Because of this, the `js/*.js` module files are not actually loaded at runtime by `index.html`.** When changing app logic (not just markup/CSS), the corresponding code exists in two places that must be kept in sync:
1. the module file under `js/` (the readable, commented source), and
2. the matching block inside `index.html`'s inline `<script>` (uncommented, what actually ships).

If you only edit one, the app's real behavior (driven by `index.html`) won't change, and/or the module source will drift out of sync. When asked to fix or add functionality, check whether the same logic block exists in both places and update both — don't assume editing `js/` alone is sufficient.

## Output documents

| Document | Format | Purpose |
|---|---|---|
| Duty Roster | `.xlsx` | Monthly timesheet submitted to payroll |
| Annexure C | `.docx` | Overtime hours verification for HOD sign-off |
| Z1(a) Leave Form | `.docx` | Official WCG leave application |

## Notes

- All parsing is coordinate-based against PDF.js text positions, since roster layouts vary — logic in `parser.js`/`parser-consultant.js` (and their inlined counterparts) is layout-sensitive.
- The SA public holiday calendar (`holidays.js`) is a deterministic, hardcoded calendar — no network calls, works fully offline.
- EC profile submission (the in-app wizard for adding a new EC) writes to Supabase; credentials for this live in `config.js`/the inlined config block.
