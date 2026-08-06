# Date Handling & Formatting Guidelines for Web Scraping, Google Sheets, and CSVs

## 1. Always Parse to ISO 8601 Date Format (`YYYY-MM-DD`)
- When scraping or ingesting dates in `DD/MM/YYYY` format, **always convert them to ISO format `YYYY-MM-DD`** before writing to Google Sheets, CSVs, or databases.
- `YYYY-MM-DD` (e.g. `2026-08-04`) is universally recognized without ambiguity across all locales (US, UK, TH), languages, databases, and spreadsheet engines (Google Sheets, Excel, SQL, Pandas).

## 2. Google Sheets Date Column Formatting
- When writing date strings into Google Sheets, explicitly set the cell/range number format:
  ```javascript
  sheet.getRange(2, dateColIndex, sheet.getLastRow() - 1, 1).setNumberFormat('yyyy-MM-dd');
  ```
- This ensures `=MONTH()` returns `8` (August) and `=DAY()` returns `4` accurately without locale-dependent swapping.

## 3. Clean CSV Export (No Timezone String Dumps)
- When converting Google Sheets data or Javascript `Date` objects to CSV:
  - **Never** rely on raw `cell.toString()` for `Date` instances.
  - Explicitly format date cells using:
    ```javascript
    if (cell instanceof Date) {
      cellStr = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    ```
- This prevents outputs like `Wed Apr 08 2026 00:00:00 GMT+0700 (Indochina Time)` in CSV files.
