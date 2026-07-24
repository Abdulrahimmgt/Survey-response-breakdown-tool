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
- Generate sheet-style distribution reports from an uploaded workbook or public Google Sheet.
- Use an input/config sheet to choose the report name and breakdown columns.
- Merge optional site-level columns into the report.
- View the generated distribution table in the app with percentage heat maps.
- Export chart images, summary tables, and filtered data as CSV.
- Export generated distribution reports as CSV or Excel.

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

## Distribution Report mode

The Distribution Report section appears below the chart maker. It can use:

- The uploaded workbook.
- A public Google Sheet link that can be opened without signing in.

For best results, include sheets like:

- A raw data sheet with one header row.
- A question list sheet with a question-column field and optional display-name field.
- An optional input/config sheet with columns such as `Survey`, `Generate`, `New Sheet Name`, `Breakdown 1`, `Breakdown 2`, `Breakdown 3`, and site-level breakdown columns.
- An optional site lookup sheet with a `Site` column and site-level characteristics.

Private Google Sheets are not connected in this version because that would require Google sign-in/API setup.

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
