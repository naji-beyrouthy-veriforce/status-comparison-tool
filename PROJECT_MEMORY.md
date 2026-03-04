# Project Memory - Status Comparison Tool
**Last Updated:** February 24, 2026  
**Status:** Fully Functional - Manual Workflow  
**Code Quality:** Technical Debt Resolved

---

## Recent Critical Updates

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
- `gui_app.py` - 4-tab manual workflow interface
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

### Step 2: Extract IDs (Tab 2)
**Function:** `extract_and_save_ids()`

**Process:**
1. Reads D365 files (WCB & Accreditation only; Client uses direct comparison)
2. Finds ID column: keywords `["global", "alcumus", "id"]`
3. Extracts UUIDs: `[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}`
4. Cleans IDs (lowercase, trim, deduplicate)
5. Formats SQL: `'uuid1',\n'uuid2',\n'uuid3'` (no trailing comma)

**Output:** `output/query_ids/{wcb,accreditation}_ids.sql.txt`

### Step 3: Manual Redash Process ⚠️ REQUIRED
**User Actions:**
1. Open Redash queries (IDs: Client=1277, WCB=1281, Accreditation=1266)
2. Copy IDs from `.sql.txt` files
3. Paste into `WHERE global_alcumus_id IN (...)` clause
4. Execute query & download Excel
5. Upload to GUI Tab 3

**Output:** SC files in `input/redash/`

**Why Manual:** Redash API has URI length limits (414 error) for large ID lists

### Step 4: Generate Comparisons (Tab 4)
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

### gui_app.py
- `setup_*_tab()` - Creates 4 tab interfaces
- `handle_bulk_drop()` - Multi-file drag & drop processing
- `generate_comparison()` - Runs comparison in background thread
- `auto_generate_email_report()` - Auto-generates email after comparisons
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

# Redash Query IDs (Reference)
QUERY_IDS = {"accreditation": 1266, "wcb": 1281, "client": 1277}
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

### 5. Manual Redash Execution
**Why:** API URI length limits (414 error) for large datasets  
**Alternative:** Considered API automation (removed due to limitations)

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
           → [Tab 2 Extract] → output/query_ids/*.sql.txt
           → [Manual Redash] → SC Excel
           → [Tab 3 Upload] → input/redash/
           → [Tab 4 Generate] → output/comparison_YYYY-MM-DD/*.xlsx
           → [Email Report] → output/email_report.txt + GUI display
```

---

## Project Structure

```
status_comparaison_tool/
├── main.py                   # Core logic
├── generate_email_report.py  # Email report generator
├── config.py                 # Constants, Messages class
├── utils.py                  # Reusable utilities
├── gui_app.py                # 4-tab GUI
├── requirements.txt          # pandas, openpyxl, tkinterdnd2
├── Run_*.bat                 # Launch scripts
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
# Tab 1: Upload 3 D365 files
# Tab 2: Extract IDs → copy to Redash
# Tab 3: Upload 3 SC files
# Tab 4: Generate comparisons → email report auto-displays
```

---

## Success Checklist

- [ ] All 4 tabs visible/functional
- [ ] File uploads show ✓ checkmarks
- [ ] ID extraction creates `.sql.txt` files
- [ ] SC upload shows success
- [ ] Comparison creates 3 Excel files
- [ ] Each Excel has 2 sheets
- [ ] "Is it the same?" column shows matches
- [ ] Red headers applied
- [ ] Email report displays in unified output
- [ ] No Python errors

---

## Security

- No API keys in code
- No personal data stored (only UUIDs)
- Local processing only
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
