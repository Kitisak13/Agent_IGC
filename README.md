# IGC Market at a Glance Web Scraper & Automation System

Automated Web Scraping and Monitoring System for Agricultural Commodity Prices (Wheat, Maize, Barley, Soyabeans, Rice) from the International Grains Council (IGC) website:
`https://www.igc.int/en/default.aspx`

---

## 📌 Features

- 🌾 **Multi-Commodity Support**: Daily graph price extraction for:
  - **Wheat** (EU Grade 1 Rouen, US HRW Gulf, US SRW Gulf)
  - **Maize** (Brazil Paranagua, US 2YC PNW)
  - **Barley** (EU Feed Rouen)
  - **Soyabeans** (Brazil Paranagua, US 2Y Gulf)
  - **Rice** (Thailand 5% Broken Bangkok, Vietnam 5% Broken Ho Chi Minh)
- 🛡️ **Robust Error Handling & Auto-Retries**:
  - **Exponential Backoff Retry Engine**: Retries network requests up to 3 times on connection drops or HTTP 5xx errors.
  - **Failure Alert System**: Automatically sends Error Alert cards to MS Teams Channel and Email if the scraper fails or site structure changes.
- 🤖 **Google Apps Script (`Code.gs`)**:
  - Auto-updates Google Sheets (Consolidated Long-Format & Tabbed Pivot Sheets)
  - **New Date Detection**: Automatically triggers alerts ONLY when new daily market prices are published
  - **MS Teams Channel Notification**: Sends Adaptive Cards with summary prices and direct links to Google Drive CSV file
  - **Automated Emailing**: Sends email to 3 recipients with `.csv` file attached
  - **No Hardcoded Credentials**: Uses `ScriptProperties` (`SPREADSHEET_ID`, `TEAMS_WEBHOOK_URL`, `EMAIL_RECIPIENTS`, etc.)
- 🐍 **Python Module (`igc_scraper.py`)**:
  - High-performance ASP.NET WebForms PostBack scraping engine
  - Exports to `JSON`, `CSV`, and multi-sheet `Excel (XLSX)`

---

## 🛠️ Project Structure

```
.
├── Code.gs             # Google Apps Script code for Google Sheets & Notifications
├── igc_scraper.py      # Python modular scraper script (JSON/CSV/Excel exporter)
└── README.md           # Documentation
```

---

## 🚀 How to Use

### 1. Google Apps Script (`Code.gs`) Deployment

1. Open your Google Sheet -> Go to **Extensions** > **Apps Script**.
2. Copy the contents of [`Code.gs`](Code.gs) into the editor.
3. Configure **Script Properties** (Go to **Project Settings ⚙️** > **Script Properties**):
   - `SPREADSHEET_ID` *(Optional)*: Google Sheet ID (leave empty to use active sheet or auto-create).
   - `SHEET_NAME` *(Optional)*: Sheet name for consolidated data (Default: `All_Data`).
   - `TEAMS_WEBHOOK_URL`: Your MS Teams Channel Webhook URL.
   - `EMAIL_RECIPIENTS`: Comma-separated email addresses (e.g. `user1@email.com, user2@email.com, user3@email.com`).
   - `CSV_FOLDER_ID` *(Optional)*: Google Drive Folder ID for CSV exports.
4. Run `setupDailyTrigger()` once to schedule automated daily scraping at 8:00 AM.

---

### 2. Python Scraper (`igc_scraper.py`)

#### Requirements:
```bash
pip install requests beautifulsoup4 pandas openpyxl
```

#### Run Scraper:
```bash
python igc_scraper.py
```

Outputs (Ignored by Git for security/privacy):
- `igc_market_data.json`
- `igc_market_data.csv`
- `igc_market_data.xlsx`

---

## 📄 License
MIT License
