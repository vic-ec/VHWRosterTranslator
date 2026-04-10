# EC Roster Translator

A browser-based tool that parses monthly shift rosters and generates WCG-formatted duty roster timesheets, overtime verification forms, and leave application documents — entirely offline, no installation required. This tool has been built for Victoria Hospital Emergency Centre, Cape Town.

---

## What it does

1. **Upload and extract** one or more monthly roster PDFs
2. **Select** a doctor, month, and year
3. **Preview and edit** the parsed shift data in an interactive table
4. **Download** the completed documents:
   - **Duty Roster Excel** — WCG-formatted timesheet (`.xlsx`)
   - **Annexure C (Overtime) Form** — overtime verification document (`.docx`)
   - **Z1(a) Leave Form** — leave application document (`.docx`)

---

## Features

- **Automatic shift detection** — parses shift times, day types (weekday/weekend), and public holidays directly from PDF layout using coordinate-based extraction
- **SA public holiday calendar** — deterministic calendar (no internet required) covering all South African public holidays
- **Editable preview** — add, remove, or correct any shift before downloading; changes can be undone row by row
- **WCG-compliant output** — Excel timesheet and Word forms formatted to Western Cape Government standards
- **Fully offline** — all libraries bundled; works from a USB drive, local folder, or GitHub Pages with no internet connection

---

## Usage

### GitHub Pages
Open the hosted URL directly in any modern browser.

### Local (other)
Serve the folder with any static file server, or open `index.html` directly in a browser (Chrome/Edge recommended for PDF parsing).

---

## Output documents

| Document | Format | Purpose |
|---|---|---|
| Duty Roster | `.xlsx` | Monthly timesheet submitted to payroll |
| Annexure C | `.docx` | Overtime hours verification for HOD sign-off |
| Z1(a) Leave Form | `.docx` | Official WCG leave application |

---

## Technical notes

- **PDF parsing** uses PDF.js with positional coordinate analysis to identify shift zones, time values, and staff names across varying roster layouts
- **Excel generation** uses SheetJS; **Word generation** uses docx.js
- All libraries are bundled in `js/` — no CDN dependencies
- Tested on Chrome and Edge (desktop); Firefox supported

---

## Folder structure

```
index.html          Main application
js/
  pdf.min.js        PDF.js (Mozilla)
  pdf.worker.min.js PDF.js worker
  xlsx.full.min.js  SheetJS
  jszip.min.js      JSZip
README.md
```

---

## Developed for

Victoria Hospital Emergency Centre, Cape Town


