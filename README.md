# Roster Translator

## Hierarchy

```text
roster-translator/
├── index.html
├── styles.css
├── README.md
├── assets/
│   └── templates/
│       └── Duty-Rosters-2026.xlsx
└── js/
    ├── app.js
    ├── constants.js
    ├── holidays.js
    ├── state.js
    ├── utils.js
    ├── parsers/
    │   ├── excelParser.js
    │   └── pdfParser.js
    ├── roster/
    │   └── normalizer.js
    └── export/
        ├── annexureC.js
        ├── docxHelpers.js
        ├── dutyRoster.js
        └── z1a.js
```

## Notes

- Put the supplied `Duty-Rosters-2026.xlsx` into `assets/templates/`.
- This version is modular and GitHub Pages compatible.
- PDF parsing is still a placeholder and should be tailored to the exact EC roster PDF text layout.
- Excel export uses the real workbook template structure.
- Annexure C and Z1(a) export as `.docx`.
