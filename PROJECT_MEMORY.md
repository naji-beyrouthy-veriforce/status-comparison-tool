# Project Memory - Status Comparison Tool
**Last Updated:** February 4, 2026  
**Status:** ✅ Fully Functional - Manual Workflow  
**Code Quality:** ✅ Technical Debt Resolved

---

## 📝 Recent Updates

### Technical Debt Cleanup & Code Quality (Latest)
**Date:** February 4, 2026

✅ **Completed Improvements:**
1. **Eliminated Duplicate Code:** Centralized header formatting in utils.py
2. **Extracted Magic Numbers:** All numeric constants now in config.py
3. **Centralized UI Strings:** Created Messages class for all user-facing text
4. **Moved Imports to Module Level:** datetime and time imports follow PEP 8
5. **GUI Consistency:** GUI now uses Messages class like CLI
6. **Removed Unused Code:** Deleted COLUMN_KEYWORDS and unused dictionaries

**Impact:**
- 50+ hardcoded strings → Messages class
- 100% message consistency across CLI and GUI
- PEP 8 compliant imports
- Cleaner, more maintainable codebase
- Easy internationalization path

---

## ⚠️ CRITICAL BUSINESS LOGIC - READ FIRST

### **Client Report Comparison Logic**
**DO NOT MODIFY WITHOUT UNDERSTANDING THIS:**

For **Client** reports from SafeContractor Redash query:
- The `case` column **IS** the status column for client-specific global IDs
- This is **NOT** the same as a regular `status` column
- Comparison logic **MUST** use `case` column vs D365 Status
- This is the **CORRECT** behavior per business requirements

For **Accreditation/WCB** reports:
- The `status` column is used for comparisons (standard behavior)

**Why this matters:**
- The SafeContractor Redash query structure for client-specific records returns status data in the `case` field
- Attempting to "fix" this by using a generic status column will break client comparisons
- The "Is it the same?" formula correctly compares `case` vs D365 Status for client reports

---

## 🎯 Project Purpose

**Dynamics 365 vs SafeContractor Status Comparison Tool**

Automates the comparison of contractor statuses between Dynamics 365 (D365) and SafeContractor (SC) systems. Creates Excel reports showing status matches, differences, and provides XLOOKUP formulas for data cross-reference.

**Target Users:** Internal team comparing contractor records  
**Processing Volume:** 18K-65K records per report type  
**Output:** Excel files with dual-sheet comparisons and status validation

---

## 📁 Project Structure

```
status_comparaison_tool/
├── automate_comparison.py    # Core logic: ID extraction & comparison generation
├── config.py                 # ⭐ Configuration hub: constants, patterns, Messages class
├── utils.py                  # ⭐ Reusable utilities: validation, formatting, file ops
├── gui_app.py                # GUI interface: 4-tab manual workflow
├── requirements.txt          # Python dependencies
├── Run_CLI.bat              # Run command-line version
├── Run_GUI.bat              # Run GUI version (primary method)
├── README.md                # User documentation
├── PROJECT_MEMORY.md        # THIS FILE - Developer reference
├── TECHNICAL_DEBT_FIXES.md  # ⭐ Technical debt cleanup documentation
├── .gitignore               # ⭐ Git exclusions (enhanced)
├── input/
│   ├── dynamics/            # D365 Excel files (uploaded via GUI or manual)
│   └── redash/              # SafeContractor Excel files (from Redash queries)
└── output/
    ├── query_ids/           # Extracted ID lists for Redash queries
    │   ├── accreditation_ids.sql.txt
    │   └── wcb_ids.sql.txt
    └── *.xlsx               # Final comparison files (3 reports)
```

**⭐ = Recently enhanced/created files**

---

## 🏗️ Code Architecture (Modular Design)

### **Module Separation:**

#### **config.py** - Configuration Hub
- **Purpose:** Single source of truth for all constants
- **Contents:**
  - Directory paths (INPUT_DIR, OUTPUT_DIR, etc.)
  - File patterns (D365_PATTERNS, SC_PATTERNS)
  - Validation settings (MIN_FILE_SIZE_BYTES, ALLOWED_FILE_EXTENSIONS)
  - Retry logic constants (MAX_FILE_SAVE_RETRIES, FILE_SAVE_RETRY_DELAY_SECONDS)
  - Excel formatting (HEADER_FILL, HEADER_FONT, HIGHLIGHT_HEADERS)
  - **Messages class** - All UI strings centralized
- **Benefits:** Change configuration without touching business logic

#### **utils.py** - Reusable Utilities
- **Purpose:** Common functions used across modules
- **Key Functions:**
  - `clean_uuid()` - Extract UUID from mixed text
  - `format_ids_for_sql()` - Format IDs for SQL IN clause
  - `find_column_by_keywords()` - Flexible column detection
  - `validate_file_format()` - File validation with suggestions
  - `validate_dataframe()` - DataFrame structure validation
  - `validate_uuid_data()` - UUID quality checking
  - `safe_read_excel()` - Robust Excel reading
  - `check_file_accessibility()` - Proactive lock detection
  - `apply_header_formatting()` - Excel header styling
- **Benefits:** DRY principle, easy testing, reusability

#### **automate_comparison.py** - Business Logic
- **Purpose:** Core comparison workflow
- **Functions:**
  - `extract_and_save_ids()` - Step 1: ID extraction
  - `create_comparison_excel()` - Create comparison workbooks
  - `generate_comparisons()` - Step 2: Generate all reports
  - `main()` - Entry point
- **Dependencies:** Imports from config.py and utils.py
- **Benefits:** Clean separation, focused functionality

#### **gui_app.py** - User Interface
- **Purpose:** 4-tab drag-and-drop interface
- **Features:**
  - Tab 1: Upload D365 files
  - Tab 2: Extract IDs
  - Tab 3: Upload SC files
  - Tab 4: Generate comparisons
- **Dependencies:** Uses config paths and calls automate_comparison functions

### **Data Flow:**

```
User Input (D365 Files)
    ↓
gui_app.py (File Upload) → Saves to input/dynamics/
    ↓
automate_comparison.py:extract_and_save_ids()
    ├→ utils.py:safe_read_excel()
    ├→ utils.py:validate_uuid_data()
    └→ utils.py:format_ids_for_sql() → output/query_ids/*.sql.txt
        ↓
[MANUAL STEP: User runs Redash queries]
        ↓
User Input (SC Files)
    ↓
gui_app.py (File Upload) → Saves to input/redash/
    ↓
automate_comparison.py:generate_comparisons()
    ├→ automate_comparison.py:create_comparison_excel()
    ├→ utils.py:apply_header_formatting()
    └→ output/*.xlsx (Comparison files)
```

---

## 🔄 Complete Workflow (4 Steps)

### **STEP 1: Upload D365 Files**
- **File:** `gui_app.py` → Tab 1
- **Input:** 3 Excel files from Dynamics 365
  - `accreditation_d365.xlsx` (18K-23K rows)
  - `wcb_d365.xlsx` (65K-75K rows)
  - `client_d365.xlsx` (26K-32K rows)
- **Action:** Drag & drop all 3 files → System auto-detects by filename
- **Output:** Files saved to `input/dynamics/`
- **Next:** Auto-switch to Tab 2

### **STEP 2: Extract IDs**
- **File:** `automate_comparison.py` → `extract_and_save_ids()`
- **Process:**
  1. Reads D365 files from `input/dynamics/`
  2. Finds `Global Alcumus Id` column (case-insensitive, flexible matching)
  3. Extracts UUIDs using regex pattern
  4. Removes duplicates
  5. Formats as SQL-ready list: `'id1',\n'id2',\n'id3'`
- **Output:** `output/query_ids/{type}_ids.sql.txt`
- **Key Function:** `clean_uuid()` extracts UUID from mixed text
- **Next:** Manual Redash query execution

### **STEP 3: Manual Redash Process** ⚠️ CRITICAL
- **User Action Required:**
  1. Open Redash manually
  2. Navigate to queries:
     - Accreditation Query ID: 1266
     - WCB Query ID: 1281
     - Client Query ID: 1277
  3. Copy IDs from `.sql.txt` files
  4. Paste into query `WHERE global_alcumus_id IN (...)` clause
  5. Execute query
  6. Download results as Excel
  7. Rename files:
     - `accreditation_sc.xlsx`
     - `wcb_sc.xlsx`
     - `client_sc.xlsx`
  8. Upload to GUI Tab 3
- **Output:** SC files saved to `input/redash/`

### **STEP 4: Generate Comparisons**
- **File:** `automate_comparison.py` → `generate_comparisons()`
- **Process:**
  1. Validates SC files exist
  2. Reads D365 + SC files
  3. Cleans IDs in both datasets
  4. Merges on `Global Alcumus Id`
  5. Creates comparison Excel with 2 sheets:
     - **SC Sheet:** SafeContractor data + D365 status columns
     - **D365 Sheet:** Dynamics data + SC status columns
  6. Adds calculated columns: "Is it the same?"
  7. Applies red header formatting
- **Output:** 3 Excel files in `output/`
  - `accreditation_comparison.xlsx`
  - `wcb_comparison.xlsx`
  - `client_comparison.xlsx`

---

## 🔑 Key Functions & Their Purpose

### **automate_comparison.py**

| Function | Purpose | Critical Details |
|----------|---------|------------------|
| `find_file_by_pattern()` | Finds files by keyword matching | Case-insensitive, flexible patterns |
| `clean_uuid()` | Extracts UUID from mixed text | Regex: `[0-9a-fA-F]{8}-[0-9a-fA-F]{4}...` |
| `format_ids_for_sql()` | Formats IDs for SQL IN clause | `'id',\n'id2',\n'id3'` (no trailing comma) |
| `find_column_by_keywords()` | Finds columns by partial name match | Handles: "global alcumus id", "Global_Alcumus_Id", etc. |
| `apply_header_formatting()` | Applies red fill to specific headers | Headers: global_alcumus_id, status, status reason |
| `extract_and_save_ids()` | Main extraction logic | Only processes Accreditation & WCB |
| `create_comparison_excel()` | Generates comparison files | Creates 2-sheet workbook with formulas |
| `generate_comparisons()` | Orchestrates all comparisons | Loops through 3 report types |
| `main()` | Entry point | Checks SC files → runs appropriate step |

### **gui_app.py**

| Function | Purpose | Critical Details |
|----------|---------|------------------|
| `setup_d365_tab()` | Creates Tab 1 UI | Bulk file upload zone |
| `setup_extract_tab()` | Creates Tab 2 UI | ID extraction with console output |
| `setup_sc_tab()` | Creates Tab 3 UI | SC file upload zone |
| `setup_compare_tab()` | Creates Tab 4 UI | Comparison generation with console |
| `handle_bulk_drop()` | Processes multi-file drag & drop | Auto-classifies files by name |
| `save_d365_files()` | Saves D365 files to input/dynamics/ | Copies files with validation |
| `save_sc_files()` | Saves SC files to input/redash/ | Copies files with validation |
| `extract_ids()` | Runs extraction in background thread | Captures stdout to console |
| `generate_comparison()` | Runs comparison in background thread | Real-time console updates |
| `check_upload_status()` | Enables/disables buttons | Checks if all 3 files uploaded |

---

## 📊 Data Flow Diagram

```
D365 Excel Files (3)
     ↓
[Tab 1: Upload] → input/dynamics/
     ↓
[Tab 2: Extract IDs] → output/query_ids/*.sql.txt
     ↓
[Manual Redash Process] ← User copies IDs, runs queries
     ↓
SC Excel Files (3)
     ↓
[Tab 3: Upload SC] → input/redash/
     ↓
[Tab 4: Generate] → output/*.xlsx (Comparison Reports)
```

---

## 🎨 File Naming Conventions

### **Expected Input Files:**

**D365 Files** (flexible matching):
- Must contain keywords: `accreditation`, `wcb`, or `client`/`cs`
- Preferably includes `d365` suffix
- Examples: `accreditation_d365.xlsx`, `WCB_Export_D365.xlsx`, `Client_Specific.xlsx`

**SC Files** (flexible matching):
- Must contain keywords: `accreditation`, `wcb`, or `client`/`cs`
- Preferably includes `sc` suffix
- Examples: `accreditation_sc.xlsx`, `WCB_SafeContractor.xlsx`, `CS_Report.xlsx`

### **Generated Output Files:**

```
output/
├── query_ids/
│   ├── accreditation_ids.sql.txt  # Fixed naming
│   └── wcb_ids.sql.txt            # Fixed naming
└── accreditation_comparison.xlsx   # Fixed naming
    wcb_comparison.xlsx             # Fixed naming
    client_comparison.xlsx          # Fixed naming
```

---

## 🔍 Column Detection Logic

### **Required Columns (Flexible Matching):**

**Global Alcumus ID:**
- Searches for: `('global', 'alcumus', 'id')` or `('id', 'alcumus')`
- Matches: "Global Alcumus Id", "global_alcumus_id", "ID_Alcumus", etc.

**Status (D365):**
- Searches for: `('status', 'reason')`
- Matches: "Status Reason", "status_reason", "StatusReason", etc.

**Status (SC):**
- Primary search: `('status', 'contractor')`
- Fallback searches: `('current', 'status')`, `('alcumus', 'status')`, `('status',)`
- Matches: "Status_Contractor", "Current Status", "Alcumus_Status", "Status"

**Case Status:**
- Searches for: `('case', 'status')` or `('status', 'case')`
- Used for Client reports only

**WCB URL (WCB reports only):**
- Searches for: `('qualification', 'url')`
- Optional - only included if found

---

## 🔧 Technical Configuration

### **Python Dependencies:**
```
pandas==2.2.0
openpyxl==3.1.2
tkinterdnd2==0.4.2
```

### **Key Constants:**

```python
# File patterns
D365_PATTERNS = {
    "accreditation": "accreditation",
    "wcb": "wcb",
    "client": ["client", "cs"]
}

SC_PATTERNS = {
    "accreditation": "accreditation",
    "wcb": "wcb",
    "client": ["client", "cs"]
}

# UUID Regex Pattern
UUID_PATTERN = r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

# Highlighted Headers (Red formatting)
HIGHLIGHT_HEADERS = [
    'global_alcumus_id', 'global alcumus id', 
    'status', 'd365 status', 'sc status',
    'is it the same?', 'status reason', 'case'
]
```

### **Redash Query IDs** (For Reference):
```python
QUERY_IDS = {
    "accreditation": 1266,
    "wcb": 1281,
    "client": 1277
}
```

---

## 🎯 Excel Output Structure

### **Each comparison file has 2 sheets:**

#### **Sheet 1: SC (SafeContractor)**
```
Columns:
├── [Original SC columns...]
├── D365 Status          ← Looked up from D365 sheet
├── Is it the same?      ← Comparison formula
└── [Other SC columns]
```

#### **Sheet 2: D365 (Dynamics 365)**
```
Columns:
├── [Original D365 columns...]
├── SC Status            ← Looked up from SC sheet
├── Is it the same?      ← Comparison formula
└── [Other D365 columns]
```

### **Column Insertion Logic:**

**Accreditation & WCB:**
- New columns inserted after "Global Alcumus Id"

**Client:**
- New columns inserted after "Case Status" (if exists)
- Fallback: After "Global Alcumus Id"

---

## ⚙️ Important Business Rules

1. **Client Reports Don't Extract IDs:**
   - Client data is processed directly without ID extraction
   - Only Accreditation and WCB go through ID extraction step

2. **UUID Cleaning is Critical:**
   - D365 IDs may have appended case numbers: `abc-123-def CAS-39866`
   - Regex extracts only UUID portion
   - Both datasets cleaned before comparison

3. **Status Matching is Case-Insensitive:**
   - "Active" = "active" = "ACTIVE"
   - Uses pandas merge for matching

4. **Duplicate Removal:**
   - IDs deduplicated before export
   - Sorted alphabetically for consistency

5. **File Validation:**
   - Checks file existence before processing
   - Validates Excel format (.xlsx, .xls, .csv)
   - Verifies required columns present

---

## 🚨 Known Limitations & Workarounds

### **Limitation 1: Manual Redash Step**
- **Why:** Redash API has URI length limits (414 error) for large datasets
- **Workaround:** Manual copy-paste from `.sql.txt` files
- **Future:** Could implement file upload to Redash if API supports it

### **Limitation 2: Memory Usage**
- **Issue:** Large files (65K rows) loaded entirely into memory
- **Impact:** ~500MB RAM usage during processing
- **Mitigation:** Process files sequentially (not in parallel)

### **Limitation 3: Excel Formula Complexity**
- **Decision:** Using pandas merge instead of Excel XLOOKUP formulas
- **Reason:** More reliable, handles missing data better
- **Benefit:** Faster processing, no formula recalculation needed

### **Limitation 4: Column Name Variations**
- **Challenge:** D365/SC exports have inconsistent column names
- **Solution:** Flexible keyword-based column detection
- **Risk:** If column names change drastically, may need code update

---

## 🛠️ Common Modification Points

### **To Add a New Report Type:**

1. **Update patterns in `automate_comparison.py`:**
```python
D365_PATTERNS = {
    "accreditation": "accreditation",
    "wcb": "wcb",
    "client": ["client", "cs"],
    "new_type": "new_keyword"  # ADD THIS
}

SC_PATTERNS = {
    # Same as above
}
```

2. **Update file dictionaries:**
```python
D365_FILES = {
    "new_type": "new_type_d365.xlsx"  # ADD THIS
}

SC_FILES = {
    "new_type": "new_type_sc.xlsx"  # ADD THIS
}
```

3. **Update GUI file storage in `gui_app.py` `__init__`:**
```python
self.uploaded_files = {
    "new_type_d365": None,  # ADD THIS
    "new_type_sc": None,    # ADD THIS
    # ... existing entries
}
```

4. **Add Redash query ID if needed** (if you re-enable automation)

### **To Change Column Detection:**

Edit `find_column_by_keywords()` calls in `create_comparison_excel()`:

```python
# Example: Change status column detection
status_col_d365 = find_column_by_keywords(
    df_d365.columns, 
    ('status', 'reason'),     # Current
    ('new', 'status', 'name') # Add alternative
)
```

### **To Modify Output Formatting:**

1. **Header colors:** Edit `HEADER_FILL` in `automate_comparison.py`
2. **Header font:** Edit `HEADER_FONT`
3. **Headers to highlight:** Edit `HIGHLIGHT_HEADERS` frozenset

### **To Add New Validation:**

Add to `create_comparison_excel()` before processing:

```python
# Example: Validate row count
if len(df_d365) == 0:
    print(f"     ⚠ Warning: D365 file is empty")
    return None
```

---

## 🐛 Debugging Tips

### **Issue: Files Not Found**
```python
# Check file patterns match actual filenames
print(f"Looking for: {D365_PATTERNS['accreditation']}")
print(f"Found files: {list(input_dir.glob('*.xlsx'))}")
```

### **Issue: Column Not Found**
```python
# Print available columns
print(f"Available columns: {list(df.columns)}")
```

### **Issue: No Common IDs**
```python
# Debug ID cleaning
print(f"Sample D365 IDs: {list(df_d365['clean_id'].head())}")
print(f"Sample SC IDs: {list(df_sc['clean_id'].head())}")
common = set(df_d365['clean_id']) & set(df_sc['clean_id'])
print(f"Common IDs: {len(common)}")
```

### **Issue: GUI Not Responding**
- Check console output - background threads print to stdout
- Ensure `threading.daemon=True` is set
- Add try-except blocks to capture errors

---

## 📝 Code Quality Notes

### **Good Practices Used:**
- ✅ Compiled regex patterns (performance)
- ✅ Frozen sets for constants (immutable)
- ✅ Context managers for file operations
- ✅ Type hints in docstrings
- ✅ Error handling with try-except
- ✅ Real-time progress updates
- ✅ Flexible file/column matching
- ✅ Single Responsibility Principle

### **Areas for Future Improvement:**
- 🔄 Add unit tests for core functions
- 🔄 Add logging to file (not just console)
- 🔄 Add data validation (schema checking)
- 🔄 Add progress bars for long operations
- 🔄 Add configuration file for patterns/paths
- 🔄 Add export to multiple formats (CSV, JSON)

---

## 🎓 Learning Resources

### **Key Libraries Used:**

1. **pandas** - Data manipulation
   - `pd.read_excel()` - Read Excel files
   - `pd.merge()` - Join dataframes
   - `df.apply()` - Apply functions to columns
   - `df.dropna()` - Remove null values

2. **openpyxl** - Excel creation
   - `Workbook()` - Create workbook
   - `dataframe_to_rows()` - Convert pandas to Excel
   - `PatternFill()` - Cell formatting
   - `Font()` - Text styling

3. **tkinter/tkinterdnd2** - GUI
   - `TkinterDnD.Tk()` - Drag & drop support
   - `ttk.Notebook()` - Tab interface
   - `scrolledtext.ScrolledText()` - Console output
   - `threading.Thread()` - Background processing

### **Useful Regex Patterns:**

```python
# UUID: 8-4-4-4-12 hex digits
r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

# Email: simple pattern
r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

# Date: YYYY-MM-DD
r'\d{4}-\d{2}-\d{2}'
```

---

## 🔐 Security & Privacy

- **No API keys in code** (old Redash key can be deleted)
- **No personal data stored** (only UUIDs processed)
- **Local processing only** (no external data transmission)
- **Input files not modified** (copies only)
- **Output files overwritten** (no data accumulation)

---

## 📞 Support Information

### **If Something Breaks:**

1. **Check console output** - errors printed in detail
2. **Verify file names** - must match patterns
3. **Check column names** - flexible but needs keywords
4. **Validate Excel format** - must be `.xlsx` or `.xls`
5. **Clear output folders** - remove old files
6. **Restart GUI** - close and reopen `Run_GUI.bat`

### **Common Error Messages:**

| Error | Cause | Solution |
|-------|-------|----------|
| "File not found" | Wrong filename/location | Check `input/` folders |
| "Column not found" | Column name changed | Update `find_column_by_keywords()` |
| "No common IDs" | UUID format mismatch | Check `clean_uuid()` regex |
| "NameError: REDASH_AVAILABLE" | Import issue | Verify no REDASH references |
| "Permission denied" | File open in Excel | Close Excel files |

---

## 🎉 Success Metrics

**Tool is working correctly when:**

- ✅ All 4 tabs visible and functional
- ✅ File upload shows ✓ checkmarks
- ✅ ID extraction creates `.sql.txt` files
- ✅ SC file upload shows success message
- ✅ Comparison generation creates 3 Excel files
- ✅ Excel files have 2 sheets each
- ✅ "Is it the same?" column shows matches
- ✅ Red headers applied correctly
- ✅ No Python errors in console

---

## 📅 Version History

| Date | Version | Changes |
|------|---------|---------|
| Feb 4, 2026 | 2.0 | Reverted to manual workflow, removed Redash automation |
| Feb 3, 2026 | 1.5 | Added Redash API integration (later removed) |
| Earlier | 1.0 | Initial manual workflow version |

---

## 🚀 Quick Start Reminder

```bash
# 1. Launch GUI
Run_GUI.bat

# 2. Follow the 4 tabs in order
Tab 1: Upload D365 files (3 files)
Tab 2: Extract IDs → copy to Redash
Tab 3: Upload SC files (from Redash)
Tab 4: Generate comparisons

# 3. Find outputs in:
output/query_ids/    # ID lists
output/              # Comparison Excel files
```

---

## 💡 Pro Tips

1. **Keep filenames descriptive** - Helps auto-detection
2. **Process one report at a time** - Easier debugging
3. **Check console output** - Shows real-time progress
4. **Don't close Excel** - While processing files
5. **Backup outputs** - Before reprocessing
6. **Use Ctrl+C in GUI console** - Copy error messages

---

**END OF PROJECT MEMORY**

*This document should be updated whenever significant changes are made to the codebase.*
