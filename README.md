# Survey Response Breakdown Tool

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
- Generate sheet-style distribution reports from an uploaded workbook or public Google Sheet.
- Choose response columns to include with checked lists.
- Optionally break those response columns down by one selected column.
- Filter distribution reports by one selected column with fewer than 500 unique values.
- View the generated distribution table in a fixed scrollable panel with percentage heat maps.
- Hide open-ended-style columns from report choices when they have more than 15 unique responses.
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
- A public Google Sheet link that can be opened without signing in. Loaded public sheets can also be charted and used for distribution reports.

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
