# Project Memory - Status Comparison Tool
**Last Updated:** March 12, 2026  
**Status:** Fully Functional - Automated Redash Integration  
**Code Quality:** Technical Debt Resolved

---

## Recent Critical Updates

### Mar 12, 2026 - Redash Automation (One-Click Pipeline)
**Change:** Fully automated Redash query execution — eliminated all manual copy/paste steps

**What was automated:**
- Extracting IDs from D365 files
- Injecting IDs into Redash queries (accreditation/WCB)
- Executing queries via Redash API
- Downloading results as Excel files
- Generating comparisons and email report

**Technical approach:**
- New `redash_api.py` module: executes raw SQL via `POST /api/query_results`
- Uses `data_source_id` + full SQL text — never modifies saved Redash queries
- Read-only API key is sufficient (no write permissions needed)
- Fetches saved query as SQL template → injects IDs locally → executes directly
- Client query (1277) executes as-is with no ID injection
- Polls `/api/jobs/{job_id}` for async query completion

**GUI restructured:** 4 tabs → 3 tabs
- Tab 1: Upload D365 Files (unchanged)
- Tab 2: Run Comparison (single "Run Full Comparison" button, shows API key status)
- Tab 3: Results & Email Report (email display + copy + open folder)
- Removed: manual Extract IDs tab, manual SC Upload tab

**Config changes:**
- `REDASH_BASE_URL = "https://redash.cognibox.net"`
- `REDASH_API_KEY` sourced from `REDASH_API_KEY` environment variable
- `REDASH_QUERY_IDS = {"accreditation": 1266, "wcb": 1281, "client": 1277}`
- `REDASH_POLL_INTERVAL = 3`, `REDASH_POLL_TIMEOUT = 300`

**Pipeline flow:** `run_automated_workflow()` in main.py orchestrates:
1. `extract_and_save_ids()` — extract IDs from D365 files
2. `run_all_redash_queries()` — execute Redash queries & download SC files
3. `generate_comparisons()` — create Excel comparisons + email report

**Security:** API key stored in environment variable, batch files set it at runtime, `*.bat` added to `.gitignore`

**Initial 403/500 errors resolved:** Original approach tried to modify saved queries (403 forbidden — API key lacked write permission; 500 server error — 71K IDs too large for query update payload). Fixed by switching to direct SQL execution via `/api/query_results` endpoint.

### Feb 24, 2026 - Partial File Support
- **Change:** Not all 3 report types (Accreditation, WCB, Client) are required to run
- Upload/Save/Generate buttons now enable with ANY file uploaded (was: ALL required)
- `save_d365_files()` and `save_sc_files()` skip unuploaded files and report what was saved/skipped
- CLI `main()` in main.py uses `any()` instead of `all()` for SC file check
- `generate_comparisons()` and `extract_and_save_ids()` already gracefully skip missing files
- `generate_email_report.py` already handles partial results (only processes available files)
- Updated UI text: drop zones and instructions no longer say "all 3 required"

### Feb 18, 2026 - Code Cleanup
- Created centralized `find_sc_status_column()` in utils.py
- Removed unused `create_comparison_zip()`, `zipfile`, `defaultdict` imports
- Replaced duplicate column finding logic (~100+ lines eliminated)
- All naming conventions verified (snake_case, PascalCase, UPPER_CASE)

### Feb 16, 2026 - SC Column Detection Fix ⚠️ CRITICAL
**Bug:** Email report showed 0 differences (SC values all false)  
**Cause:** Looking for "Status Reason" in SC sheet (doesn't exist)  
**Fix:** Conditional column logic based on report type:
```python
if report_type == "client":
    sc_status_col = "case"  # Client uses 'case' for status
else:
    sc_status_col = "status"  # WCB/Accreditation use 'status'
# D365 always uses 'Status Reason'
```
**Result:** Email counts now match Excel: Client: 1408, WCB: 31, Accreditation: 13

### Feb 16, 2026 - GUI Tab 4 Unified Output
- Merged separate "Console Output" + "Email Report" into single unified area (28 lines)
- Smart clipboard: Copies only email portion (not generation logs)
- Real-time progress with visual separator and color-coded content

### Feb 11, 2026 - Deduplication Fix
**Bug:** WCB showing 28 differences vs Excel's 25  
**Cause:** D365 duplicates; merge picked random matches  
**Fix:** Added `drop_duplicates(keep='first')` before merge
```python
df_d365_dedup = df_d365.drop_duplicates(subset=['clean_id'], keep='first')
merged = df_sc.merge(df_d365_dedup[['clean_id', 'Status Reason']], on='clean_id', how='left')
```
**Reason:** Matches Excel XLOOKUP first-match behavior

---

## ⚠️ CRITICAL BUSINESS LOGIC

### Client Report Status Column
**DO NOT MODIFY:** Client reports use `case` column (NOT `status`):
- SC Redash query for client-specific global IDs returns status in `case` field
- Accreditation/WCB use standard `status` column
- Comparison logic MUST use correct column per report type
- This is CORRECT per business requirements

---

## System Overview

Compare status records between Dynamics 365 (D365) and SafeContractor (SC) for three report types: Client, WCB, Accreditation.

**Key Components:**
- `main.py` - Creates Excel files with XLOOKUP formulas
- `generate_email_report.py` - Replicates formulas to generate reports
- `gui_app.py` - 3-tab automated workflow interface
- `config.py` - All constants, patterns, Messages class
- `utils.py` - Reusable utilities

---

## Complete Workflow

### Step 1: Upload D365 Files (Tab 1)
**Input:** 3 Excel files from Dynamics 365
- `accreditation_d365.xlsx` (18K-23K rows)
- `wcb_d365.xlsx` (65K-75K rows)
- `client_d365.xlsx` (26K-32K rows)

**Action:** Drag & drop → System auto-detects by filename → Saves to `input/dynamics/`

### Steps 2-4: Automated Pipeline (Tab 2 — "Run Full Comparison")
**Function:** `run_automated_workflow()` → calls 3 sub-steps automatically

**Step 2a — Extract IDs:** `extract_and_save_ids()`
1. Reads D365 files (WCB & Accreditation only; Client uses direct comparison)
2. Finds ID column: keywords `["global", "alcumus", "id"]`
3. Extracts UUIDs: `[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}`
4. Cleans IDs (lowercase, trim, deduplicate)
5. Formats SQL: `'uuid1',\n'uuid2',\n'uuid3'` (no trailing comma)
6. Saves to `output/query_ids/{wcb,accreditation}_ids.sql.txt`

**Step 2b — Redash Queries (Automated):** `run_all_redash_queries()`
1. Verifies Redash API connection (requires VPN)
2. For accreditation/WCB: fetches saved query SQL → injects extracted IDs → executes via `/api/query_results`
3. For client: fetches saved query SQL → executes as-is (NO modification)
4. Polls job status until completion
5. Downloads results as CSV → converts to Excel → saves to `input/redash/`

**Step 2c — Generate Comparisons:** `generate_comparisons()` + `generate_email_report()`

### Step 3: Generate Comparisons (auto-triggered)
**Function:** `generate_comparisons()` + `generate_email_report()`

**Process:**
1. Validates SC files exist
2. Reads D365 + SC files
3. Creates `clean_id` columns (lowercase, trimmed)
4. Creates 2-sheet Excel workbooks per report type
5. Adds XLOOKUP formulas (text strings, uncalculated)
6. Applies red header formatting
7. Saves to `output/comparison_YYYY-MM-DD/`
8. Auto-generates email report
9. Displays in unified output area

**Excel Structure:**
```
SC Sheet: Original Columns → [D365 Status] → [Is it the same?] → Remaining Columns
D365 Sheet: Original Columns → [SC Status] → [Is it the same?]
```

**XLOOKUP Formulas (not calculated until Excel opens):**
```excel
# SC Sheet - D365 Status column:
=_xlfn.XLOOKUP(A2, D365!A:A, D365!B:B, "Not found", 0)

# SC Sheet - Is it the same?:
=G2=H2  # where G=SC Status, H=D365 Status

# D365 Sheet - SC Status column:
=_xlfn.XLOOKUP(A2, SC!A:A, SC!G:G, "Not found", 0)

# D365 Sheet - Is it the same?:
=B2=K2  # where B=D365 Status Reason, K=SC Status
```

**Output:**
```
output/
├── comparison_2026-02-18/
│   ├── Accreditation_Comparison.xlsx
│   ├── WCB_Comparison.xlsx
│   └── Client_Comparison.xlsx
├── query_ids/
│   ├── accreditation_ids.sql.txt
│   └── wcb_ids.sql.txt
└── email_report.txt
```

---

## Email Report Generation Logic

### The Challenge
- Excel formulas created by openpyxl are text strings only
- Formulas aren't calculated until file opens in Excel
- Reading with `data_only=True` returns `NaN` for uncalculated formulas
- **Solution:** Replicate XLOOKUP logic using Python dataframe merges

### SC Sheet Analysis (Critical Fix - Feb 16, 2026)

```python
# 1. Find CORRECT status column (CRITICAL!)
if report_type == "client":
    sc_status_col = "case"  # Client uses 'case'
else:
    sc_status_col = "status"  # WCB/Accreditation use 'status'
d365_status_col = "Status Reason"  # D365 always uses this

# 2. Clean IDs for matching
df_sc['clean_id'] = df_sc['global_alcumus_id'].str.strip().str.lower()
df_d365['clean_id'] = df_d365['Global Alcumus ID'].str.strip().str.lower()

# 3. Deduplicate D365 (Critical! Matches XLOOKUP first-match behavior)
df_d365_dedup = df_d365.drop_duplicates(subset=['clean_id'], keep='first')

# 4. Merge to replicate XLOOKUP
merged = df_sc.merge(df_d365_dedup[['clean_id', d365_status_col]], on='clean_id', how='left')

# 5. Count Not Found (where XLOOKUP returns "Not found")
not_found = merged[d365_status_col].isna().sum()

# 6. Count Differences (replicates: =sc_status=d365_status)
valid_rows = merged[d365_status_col].notna()
differences = (merged.loc[valid_rows, sc_status_col] != merged.loc[valid_rows, d365_status_col]).sum()
```

### D365 Sheet Analysis

```python
# 1. Deduplicate SC
df_sc_dedup = df_sc.drop_duplicates(subset=['clean_id'], keep='first')

# 2. Find D365 not in SC
sc_ids = set(df_sc_dedup['clean_id'])
not_found_df = df_d365[~df_d365['clean_id'].isin(sc_ids)]

# 3. Group by Status Reason
status_breakdown = not_found_df['Status Reason'].value_counts()
```

### Why Deduplication is Critical
- **Without:** Merge creates multiple rows or picks arbitrary D365 record
- **With `keep='first'`:** Matches Excel XLOOKUP (first match wins)
- **Real Impact:** Fixed WCB from 28 to 25 differences (3 false positives eliminated)

---

## Key Functions

### main.py
- `extract_and_save_ids()` - Extracts IDs from D365 files (WCB/Accreditation only)
- `create_comparison_excel()` - Generates 2-sheet workbooks with XLOOKUP formulas
- `generate_comparisons()` - Orchestrates all comparisons
- `run_automated_workflow()` - Full pipeline: extract IDs → Redash queries → comparisons

### generate_email_report.py
- `analyze_sc_sheet()` - Replicates SC sheet XLOOKUP & comparison (with correct column detection)
- `analyze_d365_sheet()` - Replicates D365 sheet XLOOKUP & groups by Status Reason
- `generate_email_report()` - Main report generation & formatting

### utils.py
- `clean_uuid()` - Extracts UUID from mixed text using regex
- `format_ids_for_sql()` - Formats IDs for SQL IN clause
- `find_column_by_keywords()` - Flexible column detection
- `find_sc_status_column()` - Centralized SC status column finder (handles CLIENT 'case' vs WCB/Accreditation 'status')
- `validate_file_format()` - File validation with suggestions
- `safe_read_excel()` - Robust Excel reading
- `apply_header_formatting()` - Excel header styling (red fill)

### redash_api.py
- `verify_connection()` - Tests Redash API connectivity and auth
- `get_query()` - Fetches saved query (SQL template + data_source_id)
- `execute_raw_sql()` - Executes SQL via `/api/query_results` (read-only, never modifies saved queries)
- `_poll_job()` - Polls async Redash job until completion or timeout
- `download_result_by_id()` - Downloads query results as DataFrame by result ID
- `inject_ids_into_sql()` - Regex-replaces IDs in `global_alcumus_id IN (...)` clause
- `read_ids_from_file()` - Reads extracted `.sql.txt` ID files
- `run_redash_query()` - Full flow for one query (fetch → inject → execute → download → save)
- `run_all_redash_queries()` - Orchestrates all 3 queries with error handling

### gui_app.py
- `setup_upload_tab()` - Tab 1: D365 file upload with drag & drop
- `setup_run_tab()` - Tab 2: One-click "Run Full Comparison" with API key status
- `setup_results_tab()` - Tab 3: Email report display, copy, open folder
- `handle_bulk_drop()` - Multi-file drag & drop processing
- `run_automated()` / `run_automated_complete()` - Background automated pipeline
- `copy_email_to_clipboard()` - Smart clipboard (email portion only)

### config.py
- All constants: paths, patterns, validation settings, Excel formatting
- `Messages` class - All UI strings centralized
- `setup_logging()` - Rotating file handler configuration
- `get_dated_comparison_dir()` - Returns dated output folder

---

## Column Detection

**Global Alcumus ID:** `('global', 'alcumus', 'id')` or `('id', 'alcumus')`  
**Status (D365):** `('status', 'reason')` → "Status Reason"  
**Status (SC):** Primary: `('status', 'contractor')`, Fallback: `('current', 'status')`, `('alcumus', 'status')`, `('status',)`  
**Case Status:** `('case', 'status')` or `('status', 'case')` - Client reports only  
**WCB URL:** `('qualification', 'url')` - Optional, WCB only

---

## Configuration Constants

```python
# File Patterns
D365_PATTERNS = {"accreditation": "accreditation", "wcb": "wcb", "client": ["client", "cs"]}
SC_PATTERNS = {"accreditation": "accreditation", "wcb": "wcb", "client": ["client", "cs"]}

# UUID Regex
UUID_PATTERN = r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

# Red Headers
HIGHLIGHT_HEADERS = ['global_alcumus_id', 'status', 'd365 status', 'sc status', 
                     'is it the same?', 'status reason', 'case']

# Redash Configuration
REDASH_BASE_URL = "https://redash.cognibox.net"
REDASH_API_KEY = os.environ.get("REDASH_API_KEY")  # Set via env var or batch file
REDASH_QUERY_IDS = {"accreditation": 1266, "wcb": 1281, "client": 1277}
REDASH_POLL_INTERVAL = 3   # seconds between job status checks
REDASH_POLL_TIMEOUT = 300  # max seconds to wait for query completion
```

---

## Key Design Decisions

### 1. XLOOKUP Replication via Merge
**Why:** openpyxl formulas are text strings until Excel calculates them  
**How:** Use pandas merge to replicate XLOOKUP behavior  
**Benefit:** No need to open/calculate Excel files

### 2. Deduplication with keep='first'
**Why:** Excel XLOOKUP returns first match when duplicates exist  
**How:** `drop_duplicates(subset=['clean_id'], keep='first')` before merge  
**Impact:** Eliminated 3 WCB false positives (28 → 25)

### 3. Client 'case' Column
**Why:** SC Redash query returns client status in `case` field (business requirement)  
**Critical:** DO NOT "fix" - this is correct business logic!

### 4. Two-Sheet Excel Design
**Why:** Bidirectional comparison (SC→D365 and D365→SC perspectives)  
**Benefit:** Comprehensive analysis from both systems

### 5. Automated Redash via Direct SQL Execution
**Why:** Manual copy/paste of IDs was error-prone and time-consuming  
**How:** Execute raw SQL via `POST /api/query_results` with `data_source_id` + full SQL text  
**Benefit:** One-click pipeline, read-only API key sufficient, no saved queries modified  
**History:** Initial attempt to modify saved queries failed (403/500 errors). Direct SQL execution bypasses all limitations.

### 6. Flexible Pattern Matching
**Why:** D365/SC exports have inconsistent naming across time periods  
**How:** Keyword-based detection for files and columns

### 7. Modular Architecture
**Structure:** config.py → utils.py → main.py/generate_email_report.py → gui_app.py  
**Benefit:** Clear separation, testability, reusability

### 8. Background Threading
**Why:** Keep GUI responsive during long operations  
**How:** `threading.Thread(target=func, daemon=True)`

### 9. Logging Infrastructure
**Format:** `TIMESTAMP - LOGGER - LEVEL - MESSAGE`  
**Storage:** `logs/` (git-ignored, auto-rotate at 10MB)  
**Overhead:** <1%

---

## Data Flow

```
D365 Excel → [Tab 1 Upload] → input/dynamics/
           → [Tab 2 Run] → extract_and_save_ids() → output/query_ids/*.sql.txt
                          → run_all_redash_queries() → input/redash/*.xlsx
                          → generate_comparisons() → output/comparison_YYYY-MM-DD/*.xlsx
                          → generate_email_report() → output/email_report.txt
           → [Tab 3 Results] → Email report display + clipboard copy
```

---

## Project Structure

```
status_comparaison_tool/
├── main.py                   # Core logic + automated workflow orchestration
├── redash_api.py             # Redash API integration (query execution & download)
├── generate_email_report.py  # Email report generator
├── config.py                 # Constants, Messages class, Redash config
├── utils.py                  # Reusable utilities
├── gui_app.py                # 3-tab GUI (Upload, Run, Results)
├── requirements.txt          # pandas, openpyxl, requests, tkinterdnd2
├── Run_*.bat                 # Launch scripts (set REDASH_API_KEY env var)
├── .gitignore                # Excludes *.bat (API key protection)
├── README.md                 # User docs
├── PROJECT_MEMORY.md         # This file
├── logs/                     # Auto-rotating logs (git-ignored)
├── input/
│   ├── dynamics/             # D365 files
│   └── redash/               # SC files
└── output/
    ├── comparison_YYYY-MM-DD/
    │   ├── Accreditation_Comparison.xlsx
    │   ├── WCB_Comparison.xlsx
    │   └── Client_Comparison.xlsx
    ├── query_ids/
    │   ├── accreditation_ids.sql.txt
    │   └── wcb_ids.sql.txt
    └── email_report.txt
```

---

## Important Business Rules

1. **Client doesn't extract IDs** - Direct comparison only
2. **UUID cleaning critical** - Removes case numbers: `abc-123-def CAS-39866` → `abc-123-def`
3. **Status matching case-insensitive** - "Active" = "active" = "ACTIVE"
4. **IDs deduplicated before export** - Sorted alphabetically
5. **File validation** - Extension, size, accessibility, required columns

---

## Common Modifications

### Add New Report Type
1. Update `D365_PATTERNS` and `SC_PATTERNS` in main.py
2. Update `D365_FILES` and `SC_FILES` dictionaries
3. Update `gui_app.py` `uploaded_files` dictionary
4. Add Redash query ID if needed

### Change Column Detection
Edit `find_column_by_keywords()` calls with new keywords

### Modify Output Formatting
Edit `HEADER_FILL`, `HEADER_FONT`, `HIGHLIGHT_HEADERS` in config.py

---

## Debugging

**Files Not Found:** Check patterns match actual filenames  
**Column Not Found:** Print `list(df.columns)` to see available columns  
**No Common IDs:** Check `clean_id` values, verify UUID cleaning  
**GUI Not Responding:** Check console, ensure `daemon=True`, add try-except blocks

**Common Errors:**
- "File not found" → Check `input/` folders
- "Column not found" → Update `find_column_by_keywords()`
- "No common IDs" → Check `clean_uuid()` regex
- "Permission denied" → Close Excel files

---

## Quick Start

```bash
Run_GUI.bat
# Tab 1: Upload D365 files (drag & drop)
# Tab 2: Click "Run Full Comparison" (extracts IDs → runs Redash → generates comparisons)
# Tab 3: View email report → copy to clipboard
```
**Requirements:** VPN connected (for Redash access), `REDASH_API_KEY` env var set (batch files handle this)

---

## Success Checklist

- [ ] All 3 tabs visible/functional
- [ ] File uploads show ✓ checkmarks
- [ ] API key status shows "API Key: Configured" on Tab 2
- [ ] VPN connected (Redash reachable)
- [ ] "Run Full Comparison" completes all 3 steps
- [ ] Redash queries execute and download successfully (3/3)
- [ ] Comparison creates 3 Excel files
- [ ] Each Excel has 2 sheets
- [ ] "Is it the same?" column shows matches
- [ ] Red headers applied
- [ ] Email report displays on Tab 3
- [ ] No Python errors

---

## Security

- Redash API key stored in environment variable (`REDASH_API_KEY`), never hardcoded
- Batch files that set the API key are excluded from git (`*.bat` in `.gitignore`)
- No personal data stored (only UUIDs)
- Local processing only
- Saved Redash queries never modified (read-only API approach)
- Input files not modified
- Logs git-ignored

---

## ⚠️ FOR AI ASSISTANTS

**BEFORE any change:**
1. Read CRITICAL BUSINESS LOGIC section
2. Read Complete Workflow section
3. Read Key Design Decisions section
4. Verify understanding matches documented logic

**Critical reminders:**
- Email report replicates XLOOKUP via merge with `keep='first'` deduplication
- Client uses 'case' column (NOT 'status') - CORRECT per business requirements
- Always deduplicate before merging to match Excel XLOOKUP first-match behavior
- Manual Redash step required (API limitations)
- Two-sheet Excel design intentional (bidirectional comparison)

**Update this document when significant code changes are made.**

---

**END OF PROJECT MEMORY**
