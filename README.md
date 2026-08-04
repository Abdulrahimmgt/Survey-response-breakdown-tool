# Survey Response Analyzer

A static web app for analyzing Excel, CSV, and public Google Sheets survey-style data directly in the browser.

## What it does

- Upload `.xlsx`, `.xls`, or `.csv` files.
- Load a public Google Sheet that does not require sign-in.
- Choose a sheet when an Excel file has multiple sheets.
- Create multiple chart cards from the same file.
- Hide open-ended-style columns from chart choices when they have more than 15 unique responses.
- Analyze one column by response count and percentage.
- Compare one column against another.
- Add one or more checklist filters.
- Hide selected responses or manually combine similar response labels.
- Navigate focused Charts, Question Breakdown Report, and Data Preview workspaces.
- Generate sheet-style question breakdown reports from the active workbook or public Google Sheet.
- Preview the first 100 rows, search the dataset, inspect column quality, and hide columns from chart analysis without changing the source.
- Choose response columns to include with checked lists.
- Optionally break those response columns down by one selected column.
- Filter breakdown reports by one selected column with fewer than 500 unique values.
- View the generated breakdown table in a fixed scrollable panel with percentage heat maps.
- Hide open-ended-style columns from report choices when they have more than 15 unique responses.
- Export chart images, summary tables, and filtered data as CSV.
- Export generated breakdown reports as CSV or Excel.

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

## Question Breakdown Report

The Question Breakdown Report uses the currently active dataset. It can use:

- The uploaded workbook.
- A public Google Sheet link that can be opened without signing in. Loaded public sheets can also be charted and used for breakdown reports.

For best results, include sheets like:

- A raw data sheet with one header row.

Response-column choices appear checked by default so you can uncheck anything you do not want in the output. Columns with no responses or more than 15 unique responses are hidden from report choices so empty and open-ended questions do not create unusable breakdowns. The breakdown dropdown starts with no breakdown selected; when you choose one, that choice appears above the on-screen report. The report filter lets you select a column with fewer than 500 unique values, then uncheck values you want to exclude.

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
