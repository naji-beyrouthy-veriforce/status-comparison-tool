# Status Comparison Tool

Automated desktop application for comparing Dynamics 365 and SafeContractor status reports. Features a 3-tab dark-themed GUI with integrated Redash API queries — upload D365 files, click one button, and get comparison Excel reports plus a ready-to-send email summary.

## Features

- Dark-themed GUI with drag & drop file upload (ttkbootstrap)
- Auto-detection and classification of Accreditation, WCB, Client, Critical Document, and ESG files
- Automated Redash API integration — extracts IDs, runs queries, downloads results
- Optional automated D365 download via the Dynamics 365 Web API (requires Azure App Registration)
- Two-sheet Excel comparisons with XLOOKUP formulas and red-highlighted headers
- Auto-generated email report with status difference breakdowns
- Auto-cleanup of uploaded files on app close

## Quick Start

### Prerequisites

```bash
pip install -r requirements.txt
```

### 1. Configure secrets

Copy the template and fill in your API keys:

```bat
copy secrets.env.template secrets.env
```

Then open `secrets.env` and fill in:

```env
REDASH_API_KEY=your-redash-api-key-here

# Optional — enables automatic D365 file download (contact IT for these values)
D365_TENANT_ID=
D365_CLIENT_ID=
D365_CLIENT_SECRET=
```

`secrets.env` is gitignored and will never be committed. The `.bat` launchers load it automatically.

### 2. Run

Double-click `Run_GUI.bat` or:
```bash
python gui_app.py
```

Other launchers:
- `Run_CLI.bat` — command-line automated mode
- `Run_Email_Report.bat` — regenerate email report from existing comparison files

## Usage — 3-Tab Workflow

### Tab 1: Upload D365 Files 📁
1. Drag & drop **any combination** of D365 Excel exports — Accreditation, WCB, Client Specific, Critical Document, and/or ESG (you don't need all 5)
2. Files are auto-classified by filename keywords
3. Click **"Save D365 Files & Proceed"**

> The tool will **only run for the report types you uploaded**. If you upload 2 files, only 2 Redash queries execute and 2 comparison files are generated.

### Tab 2: Run Comparison 🚀
1. Click **"Run Full Comparison"** — one button runs the entire pipeline **for the uploaded report types only**:
   - **Detect** which report types have D365 files → builds active types list
   - **Extract IDs** from D365 files (WCB & Accreditation, if uploaded) → saves SQL-formatted lists to `output/query_ids/`
   - **Run Redash queries** automatically for active types only (injects IDs into saved queries, downloads CSV results)
   - **Generate comparisons** — creates Excel workbooks + email report for active types
2. Console shows real-time progress

### Tab 3: Results 📊
- Displays the auto-generated email report
- **"Copy Email Report"** → clipboard, ready to paste into email
- **"Open Output Folder"** → opens the dated comparison directory

## Project Structure

```
status-comparaison-tool/
├── src/                           # Core source package
│   ├── __init__.py
│   ├── config.py                  # All constants, paths, patterns, logging, Redash config
│   ├── utils.py                   # Reusable utilities: UUID cleaning, column detection, Excel formatting
│   ├── redash_api.py              # Redash API: query execution, polling, result downloading
│   └── email_report.py            # Email report: replicates XLOOKUP via pandas merge
├── gui_app.py                     # 3-tab GUI entry point (dark theme, drag & drop)
├── main.py                        # CLI entry point + core comparison logic
├── docs/                          # Documentation
│   ├── README.md
│   └── PROJECT_MEMORY.md
├── Run_GUI.bat                    # GUI launcher (sets API key)
├── Run_CLI.bat                    # CLI launcher (automated mode)
├── Run_Email_Report.bat           # Standalone email report generator
├── requirements.txt               # Python dependencies
├── .gitignore
├── input/
│   ├── dynamics/                  # D365 uploaded files (auto-managed)
│   └── redash/                    # SafeContractor files from Redash (auto-managed)
├── output/
│   ├── query_ids/                 # Extracted SQL ID lists
│   ├── comparison_YYYY-MM-DD/     # Dated comparison Excel workbooks
│   └── email_report.txt           # Generated email report
└── logs/                          # Rotating log files
```

## Output Files

### Comparison Excel Workbooks
Each report type (Accreditation, WCB, Client, Critical Document) produces a workbook with two sheets:
- **SC Sheet**: SafeContractor data + D365 status XLOOKUP + "Is it the same?" column
- **D365 Sheet**: D365 data + SC status XLOOKUP + "Is it the same?" column
- Red headers on key columns (Global Alcumus ID, Status, comparison results)

### Email Report
Auto-generated summary per report type showing:
- Number of status differences between D365 and SC
- Number of records not found in the other system
- Breakdown of not-found records by Status Reason

## Technical Details

### File Detection
Auto-classifies files by keyword in filename:
- **Accreditation**: contains "accreditation"
- **WCB**: contains "wcb"
- **Client**: contains "client" or "cs"
- **Critical Document**: contains "critical" or "cd"

### Key Business Logic
- **Client reports** use the `case` column as status (not `status`) — this is correct per SC Redash query design
- **Deduplication** before merge (`drop_duplicates(subset=['clean_id'], keep='first')`) matches Excel XLOOKUP first-match behavior
- **Email report** replicates XLOOKUP via pandas merge (openpyxl formulas are text until Excel calculates them)

## Dependencies

- Python 3.8+
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- ttkbootstrap >= 1.10.1
- tkinterdnd2 >= 0.3.0
- requests >= 2.31.0

## License

Internal tool for Company use.
