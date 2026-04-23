# Project Memory - Status Comparison Tool
**Last Updated:** April 23, 2026  
**Status:** Fully Functional - Automated Redash Integration  
**Code Quality:** Clean

---

## ⚠️ CRITICAL BUSINESS LOGIC

### Client Report Status Column
**DO NOT MODIFY:** Client reports use `case` column (NOT `status`):
- SC Redash query for client-specific global IDs returns status in `case` field
- Accreditation/WCB use standard `status` column
- Comparison logic MUST use correct column per report type
- This is CORRECT per business requirements
- Defined as `CLIENT_STATUS_COLUMN = "case"` in config.py
- `CASE_COLUMN_REPORT_TYPES = frozenset({"client", "critical_document", "esg"})` — all three use `case`

### Client Specific SC Difference Count (Special Filtering)
**DO NOT REMOVE this filtering from `analyze_sc_sheet()` in `email_report.py`:**
- Only rows where `contractor_status == "Active"` AND `client_status == "Active"` are counted
- Rows where the SC `case` column OR the D365 status is `"Cancelled"` are excluded
- "Not found" rows (no matching D365 record) **are included** in the difference count — they count as a mismatch
- This filtering applies **only to the Client report** — all other report types use unfiltered counts
- The email output line for Client SC is: `○ N differences between dynamics and SafeContractor` (no "not found" shown separately)

### Deduplication Before Merge
- `drop_duplicates(subset=['clean_id'], keep='first')` MUST be applied before any merge
- Matches Excel XLOOKUP first-match behavior
- Without this: false positives (WCB showed 28 instead of 25)

### XLOOKUP Replication
- openpyxl formulas are text strings until Excel calculates them
- Email report generator replicates XLOOKUP via pandas merge (not by reading formulas)
- `email_report.py` operates on the raw SC/D365 data, NOT on calculated Excel values

---

## System Overview

Compare status records between Dynamics 365 (D365) and SafeContractor (SC) for five report types: Client, WCB, Accreditation, Critical Document, ESG.

**Architecture:** `config.py` → `utils.py` → `main.py` / `email_report.py` / `redash_api.py` / `dynamics_api.py` → `gui_app.py`

**Key Components:**
| File | Purpose |
|------|------|
| `config.py` | All constants, paths, patterns, `Messages` class, `setup_logging()`, Redash + D365 config |
| `utils.py` | Reusable utilities: UUID cleaning, column detection, file validation, Excel formatting |
| `main.py` | Core logic: ID extraction, Excel comparison generation, automated workflow orchestration |
| `redash_api.py` | Redash API integration: query execution, polling, result downloading |
| `dynamics_api.py` | D365 Web API: OAuth2 auth, saved-view download, column name resolution, pagination |
| `email_report.py` | Email report: replicates XLOOKUP via merge, formats status differences |
| `gui_app.py` | 3-tab GUI: Upload → Run → Results (dark mode, drag & drop) |

---

## Complete Workflow

### Step 0 (Optional): Automated D365 Download
If `D365_TENANT_ID`, `D365_CLIENT_ID`, and `D365_CLIENT_SECRET` are all set, `run_automated_workflow()` automatically downloads D365 files before Step 1 via `run_all_d365_downloads()` in `dynamics_api.py`.

**How it works:**
1. Authenticates via OAuth2 client credentials flow → gets Bearer token (~60 min validity)
2. For each configured report type, fetches saved view's `layoutxml` to get column schema names
3. Calls D365 attribute metadata API to resolve schema names → display names
4. Queries the view via `?savedQuery={view_id}` with pagination (`@odata.nextLink`)
5. Prefers `@OData.Community.Display.V1.FormattedValue` annotations (human-readable labels, not integer codes)
6. Saves each report as Excel to `input/dynamics/<report_type>_d365.xlsx`

If credentials are **not** set, the step is silently skipped and existing files in `input/dynamics/` are used instead (manual upload via Tab 1).

**To use:** Set `D365_TENANT_ID`, `D365_CLIENT_ID`, `D365_CLIENT_SECRET` in `secrets.env` or the `.bat` launchers. Contact IT for an Azure App Registration if one doesn't exist.

---

### Step 1: Upload D365 Files (GUI Tab 1)
**Input:** One or more Excel files from Dynamics 365 (any combination of the 5 report types)
- `accreditation_d365.xlsx` (18K-23K rows)
- `wcb_d365.xlsx` (65K-75K rows)
- `client_d365.xlsx` (26K-32K rows)
- `critical_document_d365.xlsx`
- `esg_d365.xlsx`

**Action:** Drag & drop → System auto-classifies by filename patterns → "Save D365 Files" copies to `input/dynamics/`

⚠️ **The tool only runs for the report types you upload.** If you only upload WCB and Client, only those two are processed.

### Step 2: Automated Pipeline (GUI Tab 2 — "Run Full Comparison")
**Function:** `run_automated_workflow()` orchestrates 3 sub-steps:

**Detection:** `get_uploaded_report_types()` scans `input/dynamics/` at startup and returns only the report types with files present. All subsequent steps receive this filtered list.

**Step 2a — Extract IDs:** `extract_and_save_ids(report_types=active_types)`
1. Reads D365 files (WCB & Accreditation only; Client, CD, ESG don't need ID extraction)
2. Skips any of those two if not in `active_types`
2. Finds ID column via keywords `("global", "alcumus", "id")`
3. Extracts UUIDs via compiled regex pattern
4. Cleans (lowercase, trim), deduplicates, sorts
5. Formats as SQL: `'uuid1',\n'uuid2',\n'uuid3'`
6. Saves to `output/query_ids/{wcb,accreditation}_ids.sql.txt`

**Step 2b — Redash Queries:** `run_all_redash_queries(report_types=active_types)`
1. Verifies API connection (requires VPN + `REDASH_API_KEY` env var)
2. For accreditation/WCB: fetches saved query SQL template → injects extracted IDs into `global_alcumus_id IN (...)` → executes via `POST /api/query_results`
3. For client: fetches saved query SQL → executes as-is (NO ID injection)
4. For critical_document: fetches saved query SQL (1464) → executes as-is (NO ID injection)
5. Accreditation query (1460) also returns `created_at` and `updated_at` columns
5. Polls `/api/jobs/{job_id}` until completion
5. Downloads results as CSV → converts to Excel → saves to `input/redash/`

**Step 2c — Generate Comparisons:** `generate_comparisons(report_types=active_types)` + `generate_email_report()`
1. Reads D365 + SC files, validates structure
2. Creates `clean_id` columns (lowercase UUID extraction)
3. Creates 2-sheet Excel workbooks per report type with XLOOKUP formulas
4. Applies red header formatting, auto-filters
5. Saves to `output/comparison_YYYY-MM-DD_HH-MM-SS/`
6. Auto-generates email report via merge-based XLOOKUP replication

### Step 3: Results (GUI Tab 3)
- Displays email report text
- "Copy Email Report" → clipboard (email portion only)
- "Open Output Folder" → opens comparison directory

**Excel Structure:**
```
SC Sheet: Original Columns → [D365 Status] → [Is it the same?] → Remaining Columns
D365 Sheet: Original Columns → [SC Status] → [Is it the same?]
```

Column placement varies by report type:
- **Client:** Comparison columns inserted after `case` column
- **Accreditation/WCB/Critical Document:** Comparison columns appended at end

**Output (example — only uploaded types appear):**
```
output/
├── comparison_YYYY-MM-DD_HH-MM-SS/     ← new timestamped folder per run
│   ├── Accreditation_Comparison.xlsx   (if uploaded)
│   ├── WCB_Comparison.xlsx             (if uploaded)
│   ├── Client_Comparison.xlsx          (if uploaded)
│   ├── Critical_Document_Comparison.xlsx (if uploaded)
│   └── ESG_Comparison.xlsx             (if uploaded)
├── query_ids/
│   ├── accreditation_ids.sql.txt       (if uploaded — overwritten each run)
│   └── wcb_ids.sql.txt                 (if uploaded — overwritten each run)
└── email_report.txt
```

---

## Email Report Generation Logic

### The Challenge
- Excel formulas created by openpyxl are text strings only
- Formulas aren't calculated until file opens in Excel
- **Solution:** Replicate XLOOKUP logic using Python dataframe merges

### SC Sheet Analysis
```python
# 1. Find CORRECT status column (CRITICAL!)
sc_status_col = find_sc_status_column(df_sc, id_col_sc, report_type)
# → "case" for client/critical_document/esg, "status" for WCB/Accreditation
d365_status_col = find_column_by_keywords(df_d365.columns, ("status", "reason"))

# 2. Clean IDs for matching
df_sc_copy['clean_id'] = df_sc_copy[sc_id_col].astype(str).str.strip().str.lower()
df_d365_copy['clean_id'] = df_d365_copy[d365_id_col].astype(str).str.strip().str.lower()

# 3. Deduplicate D365 (Critical! Matches XLOOKUP first-match behavior)
df_d365_dedup = df_d365_copy.drop_duplicates(subset=['clean_id'], keep='first')

# 4. Merge to replicate XLOOKUP
merged = df_sc_copy.merge(df_d365_dedup[['clean_id', d365_status_col]], on='clean_id', how='left')

# 5. For CLIENT only: filter Active contractor_status + client_status, exclude Cancelled
#    analysis_rows starts as ALL rows (not just found) so not-found count as differences
#    For all other report types: only count rows where D365 status is not null

# 6. Count differences
differences = (sc_statuses != d365_statuses).sum()
not_found = merged[d365_status_col_merged].isna().sum()  # reported for non-client only
```

### D365 Sheet Analysis
```python
# 1. Deduplicate SC
df_sc_dedup = df_sc_copy.drop_duplicates(subset=['clean_id'], keep='first')

# 2. Find D365 not in SC
sc_ids = set(df_sc_dedup['clean_id'].dropna())
not_found_df = df_d365_copy[~df_d365_copy['clean_id'].isin(sc_ids)]

# 3. Group by Status Reason
status_breakdown = not_found_df[status_reason_col].value_counts().to_dict()
```

---

## Key Functions

### main.py
| Function | Purpose |
|----------|---------|
| `get_uploaded_report_types()` | Scans `input/dynamics/` and returns list of report types with D365 files present |
| `extract_and_save_ids(report_types=None)` | Extracts UUIDs from D365 files (WCB/Accreditation only, filtered to uploaded types) |
| `create_comparison_excel(report_type, df_d365, df_sc)` | Generates 2-sheet workbooks with XLOOKUP formulas |
| `generate_comparisons(report_types=None)` | Orchestrates comparisons for uploaded types + triggers email report |
| `run_automated_workflow()` | Full pipeline: (optional D365 download →) extract IDs → Redash queries → comparisons |
| `main()` | Entry point: automated mode (if API key set) or manual fallback |

### email_report.py
| Function | Purpose |
|----------|---------|
| `read_comparison_file(file_path)` | Reads SC and D365 sheets from comparison Excel |
| `analyze_sc_sheet(df_sc, df_d365, report_type)` | Replicates SC XLOOKUP: merge, count differences & not-found |
| `analyze_d365_sheet(df_d365, df_sc, report_type)` | Replicates D365 XLOOKUP: find not-in-SC, group by Status Reason |
| `format_status_name(status)` | Formats status for email display (adds "Statuses" suffix) |
| `generate_email_report()` | Orchestrates full report: read files → analyze → format → save |

### utils.py
| Function | Purpose |
|----------|---------|
| `clean_uuid(value)` | Extracts UUID from mixed text via compiled regex |
| `format_ids_for_sql(ids)` | Formats IDs as `'id1',\n'id2'` for SQL IN clause |
| `find_column_by_keywords(columns, *keyword_groups)` | Flexible column detection: matches ALL keywords in ANY group |
| `find_file_by_pattern(directory, patterns, file_suffix)` | Finds file by keyword match, prioritizes suffix matches |
| `find_sc_status_column(df_sc, id_col_sc, report_type)` | Returns correct SC status column: `case` for client, `status` for others |
| `validate_file_format(file_path)` | Validates existence, extension, size, accessibility |
| `validate_dataframe(df, file_name, required_columns)` | Validates DataFrame structure and required columns |
| `safe_read_excel(file_path)` | Robust Excel reading with detailed error messages |
| `apply_header_formatting(worksheet)` | Applies red fill + bold to specified headers |
| `validate_uuid_data(df, id_column, file_name)` | Validates UUID data quality, returns statistics |
| `check_file_accessibility(file_path, mode)` | Checks read/write access to file |

### redash_api.py
| Function | Purpose |
|----------|---------|
| `get_api_key()` | Returns API key from env var (raises if missing) |
| `verify_connection()` | Tests Redash API connectivity and auth |
| `get_query(query_id)` | Fetches saved query (SQL template + data_source_id) |
| `execute_raw_sql(data_source_id, sql_text)` | Executes SQL via `/api/query_results` (read-only, never modifies saved queries) |
| `_poll_job(job_id)` | Polls async Redash job until completion or timeout |
| `download_result_by_id(query_result_id)` | Downloads query results as DataFrame |
| `inject_ids_into_sql(sql_text, ids_formatted)` | Regex-replaces IDs in `global_alcumus_id IN (...)` clause |
| `read_ids_from_file(report_type)` | Reads extracted `.sql.txt` ID files |
| `run_redash_query(query_id, report_type, ids_formatted)` | Full flow for one query: fetch → inject → execute → download → save |
| `run_all_redash_queries(report_types=None)` | Orchestrates queries for the given report types (all 5 if None) |

### gui_app.py (ComparisonApp class)
| Method | Purpose |
|--------|---------|
| `setup_ui()` | Creates header, notebook tabs, status bar |
| `setup_d365_tab()` | Tab 1: D365 file upload with drag & drop zone + status indicators |
| `setup_run_tab()` | Tab 2: "Run Full Comparison" button, API key status, progress console |
| `setup_results_tab()` | Tab 3: Email report display, copy button, open folder button |
| `classify_file(file_path, file_type_suffix)` | Auto-classifies file by name using `D365_PATTERNS`/`SC_PATTERNS` from config |
| `handle_bulk_drop(event, file_type)` | Multi-file drag & drop processing and classification |
| `parse_dropped_files(data)` | Parses file paths from tkinterdnd2 drop event data |
| `save_d365_files()` | Copies uploaded D365 files to `input/dynamics/` |
| `run_automated()` | Launches automated workflow in background thread |
| `run_automated_complete(output)` | Handles workflow completion: loads email report, switches to Tab 3 |
| `auto_generate_email_report()` | Reads saved `email_report.txt` and displays in results tab |
| `copy_email_to_clipboard()` | Copies email report text to clipboard |
| `cleanup_files()` | Deletes input files on window close |
| `check_existing_files()` | Detects pre-existing D365 files on startup |

### config.py
| Item | Purpose |
|------|---------|
| Directory paths | `BASE_DIR`, `INPUT_DIR`, `OUTPUT_DIR`, `DYNAMICS_DIR`, `REDASH_DIR`, `QUERY_IDS_DIR`, `LOG_DIR` |
| `reset_run_comparison_dir()` | Stamps a fresh `comparison_YYYY-MM-DD_HH-MM-SS` folder for this run; call at start of each `run_automated_workflow()` |
| `get_dated_comparison_dir()` | Returns the cached run folder. In standalone mode (no cache), finds the most-recently-modified `comparison_*` folder in `output/` |
| File patterns | `D365_PATTERNS`, `SC_PATTERNS`, `D365_FILES`, `SC_FILES` |
| Validation | `ALLOWED_FILE_EXTENSIONS`, `MIN_FILE_SIZE_BYTES` |
| `UUID_PATTERN` | Compiled regex for UUID matching |
| Excel formatting | `HIGHLIGHT_HEADERS`, `HEADER_FILL`, `HEADER_FONT` |
| `CLIENT_STATUS_COLUMN` | `"case"` — the status column for client/critical_document/esg reports |
| `CASE_COLUMN_REPORT_TYPES` | `frozenset({"client", "critical_document", "esg"})` |
| Redash config | `REDASH_BASE_URL`, `REDASH_API_KEY`, `REDASH_QUERY_IDS`, polling settings |
| D365 Web API config | `D365_ORG_URL`, `D365_TENANT_ID`, `D365_CLIENT_ID`, `D365_CLIENT_SECRET` (from env vars) |
| `D365_VIEW_IDS` | Per-report-type saved view GUIDs (set in `config.py` or via env vars) |
| `D365_ENTITY` | `"incidents"` — OData entity set name (plural) |
| `D365_ENTITY_LOGICAL_NAME` | `"incident"` — used in metadata API queries |
| `D365_API_VERSION` | `"v9.2"` |
| `D365_PAGE_SIZE` | `5000` — max records per OData page |
| `D365_KNOWN_FIELD_NAMES` | Hardcoded fallback schema→display mappings (e.g. `statuscode` → `"Status Reason"`) |
| `setup_logging()` | Rotating file handler with console/file output options |
| `Messages` class | All UI strings and emoji indicators centralized |

---

## Column Detection

**Global Alcumus ID:** `('global', 'alcumus', 'id')` or `('id', 'alcumus')` — fallback to `df_sc.columns[0]`  
**Status (D365):** `('status', 'reason')` → matches "Status Reason" column  
**Status (SC):** Determined by `find_sc_status_column()`:
- Client → looks for column matching `CLIENT_STATUS_COLUMN` ("case")
- WCB/Accreditation → looks for column containing "status" (excluding ID column)
- Fallback → column after ID column, then first string-type column

---

## Configuration Constants

```python
# File Patterns (in config.py)
D365_PATTERNS = {"accreditation": "accreditation", "wcb": "wcb", "client": ["client", "cs"],
                 "critical_document": ["critical", "cd"], "esg": "esg"}

# UUID Regex (compiled)
UUID_PATTERN = re.compile(r'[0-9a-fA-F]{8}-...-[0-9a-fA-F]{12}')

# Red Headers
HIGHLIGHT_HEADERS = frozenset(['global_alcumus_id', 'global alcumus id', 'status',
    'd365 status', 'sc status', 'is it the same?', 'status reason', 'case'])

# Client/CD/ESG Status Column
CLIENT_STATUS_COLUMN = "case"
CASE_COLUMN_REPORT_TYPES = frozenset({"client", "critical_document", "esg"})

# Redash Configuration
REDASH_BASE_URL = "https://redash.cognibox.net"
REDASH_API_KEY = os.environ.get("REDASH_API_KEY", "")
REDASH_QUERY_IDS = {"accreditation": 1460, "wcb": 1281, "client": 1277,
                    "critical_document": 1464, "esg": 1465}
REDASH_POLL_INTERVAL = 3   # seconds between job status checks
REDASH_POLL_TIMEOUT = 300  # max seconds to wait for query completion

# File Save Retry
MAX_FILE_SAVE_RETRIES = 3
FILE_SAVE_RETRY_DELAY_SECONDS = 1
```

---

## Data Flow

```
[Optional — if D365 credentials set]
dynamics_api.py → OAuth2 token → D365 Web API → input/dynamics/*.xlsx

[Always]
D365 Excel(s) → [Tab 1 Upload] → input/dynamics/  (any subset of 5 types)
             → [Tab 2 Run] → get_uploaded_report_types() → active_types list
                           → extract_and_save_ids(active_types) → output/query_ids/*.sql.txt
                           → run_all_redash_queries(active_types) → input/redash/*.xlsx
                           → generate_comparisons(active_types) → output/comparison_YYYY-MM-DD_HH-MM-SS/*.xlsx
                           → generate_email_report() → output/email_report.txt
             → [Tab 3 Results] → Email report display + clipboard copy
```

---

## Project Structure

```
status-comparaison-tool/
├── main.py                   # Core logic + automated workflow orchestration
├── gui_app.py                # 3-tab GUI (Upload, Run, Results) - dark mode
├── requirements.txt          # pandas, openpyxl, requests, tkinterdnd2, ttkbootstrap
├── Run_*.bat                 # Launch scripts (set REDASH_API_KEY env var)
├── .gitignore                # Excludes *.bat, logs/, __pycache__/
├── src/
│   ├── __init__.py
│   ├── config.py             # Constants, Messages class, Redash config, logging setup
│   ├── utils.py              # Reusable utilities (UUID, columns, validation, formatting)
│   ├── redash_api.py         # Redash API integration (query execution & download)
│   └── email_report.py       # Email report generator (merge-based XLOOKUP replication)
├── docs/
│   ├── README.md             # User documentation
│   └── PROJECT_MEMORY.md     # This file
├── logs/                     # Auto-rotating logs (git-ignored)
├── input/
│   ├── dynamics/             # D365 files (uploaded via GUI)
│   └── redash/               # SC files (downloaded by Redash automation)
└── output/
    ├── comparison_YYYY-MM-DD/  # Dated comparison Excel files
    ├── query_ids/              # Extracted SQL-formatted IDs
    └── email_report.txt        # Generated email report
```

---

## Key Design Decisions

1. **XLOOKUP Replication via Merge** — openpyxl formulas are text-only; email report uses pandas merge instead
2. **Deduplication with `keep='first'`** — matches Excel XLOOKUP first-match behavior
3. **Client `case` Column** — SC Redash query returns client status in `case` field (business requirement)
4. **Two-Sheet Excel Design** — bidirectional comparison (SC→D365 and D365→SC)
5. **Direct SQL Execution** — `POST /api/query_results` with `data_source_id` + SQL text; never modifies saved queries; read-only API key sufficient
6. **Flexible Pattern Matching** — keyword-based detection for files and columns handles naming inconsistencies
7. **Background Threading** — GUI stays responsive during long operations via `threading.Thread(daemon=True)`
8. **Centralized Config** — all constants, messages, and patterns in `config.py`; no magic strings in business logic
9. **Partial File Support** — `get_uploaded_report_types()` detects which D365 files are present; only those types flow through extract → Redash → compare; the tool works with any 1–5 file combination
10. **File Save Retry** — locked files retry 3 times, then save with timestamp suffix as fallback

---

## Important Business Rules

1. **Client doesn't extract IDs** — Client Redash query runs as-is with no modification
2. **UUID cleaning** — Removes trailing case numbers: `abc-123-def CAS-39866` → `abc-123-def`
3. **IDs deduplicated + sorted** before SQL formatting
4. **File validation** — Extension (`.xlsx`/`.xls`/`.csv`), size, accessibility, required columns
5. **Cleanup on exit** — GUI deletes input files when window closes

---

## Common Modifications

### Add New Report Type
1. Add patterns to `D365_PATTERNS` and `SC_PATTERNS` in `config.py`
2. Add filenames to `D365_FILES` and `SC_FILES` in `config.py`
3. Add to `REPORT_TYPES` list in `config.py`
4. Add Redash query ID to `REDASH_QUERY_IDS` in `config.py`
5. Add D365 view ID to `D365_VIEW_IDS` in `config.py` (for automated D365 download)
6. If the new type uses `case` as status column, add to `CASE_COLUMN_REPORT_TYPES` in `config.py`
7. Add D365 key to `uploaded_files` dict in `gui_app.py`
8. Add status indicator row in `setup_d365_tab()` in `gui_app.py`

### Change Column Detection
- Edit keyword tuples in `find_column_by_keywords()` calls
- For SC status column: edit `find_sc_status_column()` in `utils.py`

### Modify Output Formatting
- Edit `HEADER_FILL`, `HEADER_FONT`, `HIGHLIGHT_HEADERS` in `config.py`

---

## Debugging

| Symptom | Cause | Fix |
|---------|-------|-----|
| Files not found | Patterns don't match filenames | Check `D365_PATTERNS`/`SC_PATTERNS` in config.py |
| Column not found | Column name changed | Print `list(df.columns)`, update keyword tuples |
| No common IDs | UUID cleaning mismatch | Check `clean_uuid()` regex, compare sample IDs |
| GUI not responding | Long operation blocking | Ensure background thread with `daemon=True` |
| Permission denied | File open in Excel | Close Excel, retry (auto-retry with timestamp fallback) |
| Redash 403/401 | Bad API key | Check `REDASH_API_KEY` env var |
| Redash connection error | VPN not connected | Connect to VPN, verify with `verify_connection()` |
| 0 differences in email | Wrong status column | Verify `find_sc_status_column()` returns correct column |
| D365 download: 401 | Bad client credentials | Check `D365_CLIENT_ID` / `D365_CLIENT_SECRET` env vars |
| D365 download: 403 | App Registration missing D365 role | Contact IT to assign a D365 Security Role to the App Registration |
| D365 download: 0 rows | Wrong view ID | Open the view in D365 browser → copy `viewid=` from URL → update `D365_VIEW_IDS` |
| D365 download: missing columns | schema name not in metadata | Add to `D365_KNOWN_FIELD_NAMES` in `config.py` as fallback |
| D365 download skipped | Credentials not set | Set `D365_TENANT_ID`, `D365_CLIENT_ID`, `D365_CLIENT_SECRET` in `secrets.env` |

---

## Quick Start

**Manual D365 upload (default):**
```bash
Run_GUI.bat
# Tab 1: Drag & drop D365 Excel exports
# Tab 2: Click "Run Full Comparison" (extracts IDs → runs Redash → generates comparisons)
# Tab 3: View email report → copy to clipboard
```

**Automated D365 download (optional):**
```bash
# Fill in D365_TENANT_ID, D365_CLIENT_ID, D365_CLIENT_SECRET in secrets.env
Run_GUI.bat
# Tab 2: Click "Run Full Comparison"
# Step 0 auto-downloads D365 files before running the rest of the pipeline
```

**Standalone launchers:**
- `Run_CLI.bat` — automated mode without GUI
- `Run_D365_Download.bat` — download D365 files only (no comparison)
- `Run_Email_Report.bat` — regenerate email report from existing comparison files

**Requirements:** VPN connected, `REDASH_API_KEY` env var set (batch files handle this)

---

## Security

- Redash API key stored in environment variable (`REDASH_API_KEY`), never hardcoded
- Batch files that set the API key are excluded from git (`*.bat` in `.gitignore`)
- No personal data stored (only UUIDs)
- Local processing only
- Saved Redash queries never modified (read-only API approach)
- Input files not modified (copied, not moved)
- Logs git-ignored

---

## Change History

| Date | Change | Impact |
|------|--------|--------|
| Apr 23, 2026 | Client SC: filter Active contractor/client_status, exclude Cancelled, include not-found in count | Client difference count now matches business expectation |
| Apr 23, 2026 | Timestamped output folders: `comparison_YYYY-MM-DD_HH-MM-SS` per run | Each run gets its own folder; no overwriting previous results |
| Apr 23, 2026 | `reset_run_comparison_dir()` + `get_dated_comparison_dir()` standalone fallback (most-recent folder) | `Run_Email_Report.bat` now reads from the last real run automatically |
| Apr 23, 2026 | Removed dead imports (`Path`, `sys`) from `email_report.py` | Minor cleanup |
| Mar 12, 2026 | D365 Web API module (`dynamics_api.py`): OAuth2, saved-view download, column resolution, pagination | Fully automated D365 file download (optional — requires Azure App Registration) |
| Mar 12, 2026 | Code cleanup: removed unused imports, dead code, duplicate variables, stale messages | No behavior change |
| Mar 12, 2026 | Redash automation: one-click pipeline via `redash_api.py` | Eliminated manual Redash copy/paste |
| Feb 24, 2026 | Partial file support: any combination of report types works | Upload/generate with 1-3 files |
| Feb 18, 2026 | Centralized `find_sc_status_column()`, removed duplicate logic | ~100 lines eliminated |
| Feb 16, 2026 | SC column detection fix: client uses `case`, not `status` | Email counts now match Excel |
| Feb 11, 2026 | Deduplication fix: `keep='first'` before merge | WCB: 28→25 differences (3 false positives fixed) |

---

**END OF PROJECT MEMORY**