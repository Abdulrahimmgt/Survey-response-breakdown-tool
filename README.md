# Browser Data Analyzer

A simple static web app for analyzing Excel and CSV survey-style data directly in the browser.

## What it does

- Upload `.xlsx`, `.xls`, or `.csv` files.
- Choose a sheet when an Excel file has multiple sheets.
- Preview the first 20 rows.
- Create multiple chart cards from the same file.
- Analyze one column by response count and percentage.
- Compare one column against another.
- Add one or more checklist filters.
- Hide selected responses or manually combine similar response labels.
- Export chart images, summary tables, and filtered data as CSV.

Uploaded files are processed only in your browser. The app does not use a backend, database, sign-in, API key, or paid service.

## Files

- `index.html`
- `style.css`
- `script.js`
- `README.md`

The app uses CDN links for:

- SheetJS
- Chart.js
- Chart.js Data Labels plugin

## Run locally

Open `index.html` in your browser.

## Deploy with GitHub Pages

1. Create a GitHub repository.
2. Upload the project files.
3. Open the repository settings.
4. Open the Pages section.
5. Select deployment from the main branch.
6. Open the generated GitHub Pages URL.

## Notes

- The first row is treated as the column header row.
- Completely empty rows are ignored.
- Blank cells are grouped as `No Response`.
- Zero and `false` values are preserved as valid responses.
- The original uploaded file is not changed.
- The app is designed for normal Excel and CSV files up to about 50,000 rows.
