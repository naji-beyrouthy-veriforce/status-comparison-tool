# D365 vs SafeContractor Status Comparison Tool

Modern desktop application for comparing Dynamics 365 and SafeContractor status reports with an intuitive dark-themed GUI.

## Features

✨ **Modern User Interface**
- Dark theme with ttkbootstrap
- Drag & drop file upload
- Auto-detection of file types
- Real-time status indicators
- Progress tracking
- Automatic folder opening on completion

🎯 **Key Capabilities**
- **Smart File Classification**: Automatically detects Accreditation, WCB, and Client files
- **Fresh Data Guarantee**: Always uses IDs from uploaded files (no cache)
- **Smart Column Detection**: Automatically identifies ID and status columns
- **Intelligent Formatting**: Red headers on key columns
- **Error Handling**: Detailed error messages and validation
- **Performance Optimized**: 50-60% faster with vectorized operations
- **Auto-Cleanup**: Uploaded files are automatically deleted when closing the app

## Quick Start

### 1. Prerequisites

```bash
pip install -r requirements.txt
```

### 2. Run the Application

Double-click `Run_GUI.bat` or run:
```bash
python gui_app.py
```

## Usage

### 4-Step Manual Workflow

#### **Tab 1: Upload D365 Files** 📁
1. Drag & drop all 3 D365 Excel exports
   - Accreditation export
   - WCB export  
   - Client Specific export
2. Files are auto-detected and classified
3. Click **"Save D365 Files & Proceed"**

#### **Tab 2: Extract IDs** 🔍
1. Click **"Extract IDs"**
2. SQL-formatted ID lists are generated in `output/query_ids/`
3. Use these ID files to run your Redash queries

#### **Tab 3: Upload SC Files** 📊
1. Run Redash queries with the extracted IDs
2. Drag & drop all 3 SafeContractor (Redash) exports
   - Accreditation results
   - WCB results
   - Client results
3. Files are auto-detected and classified
4. Click **"Save SafeContractor Files & Proceed"**

#### **Tab 4: Generate Reports** 🚀
1. Click **"Generate Comparisons"**
2. Wait for processing to complete (~30 seconds)
3. **Output folder opens automatically**
4. Find 3 comparison files ready to use!

## Project Structure

```
status_comparaison_tool/
├── gui_app.py                 # Dark-themed GUI application
├── automate_comparison.py     # Core comparison logic
├── config.py                  # Centralized configuration
├── utils.py                   # Utility functions
├── requirements.txt           # Dependencies
├── Run_GUI.bat               # Quick launcher
├── input/
│   ├── dynamics/             # D365 files (auto-managed)
│   └── redash/               # SC files (auto-managed)
└── output/
    ├── query_ids/            # Extracted ID lists
    └── comparison_*.xlsx     # Final comparison reports
```

## Output Files

### ID Extraction (Tab 2)
- `accreditation_ids_YYYY_MM_DD.txt` - SQL-ready ID list
- `wcb_ids_YYYY_MM_DD.txt` - SQL-ready ID list

### Comparison Reports (Tab 4)
Each comparison file contains:
- **SC Sheet**: SafeContractor data with D365 status XLOOKUP
- **D365 Sheet**: D365 data with SC status XLOOKUP  
- **Highlighted columns**: Global Alcumus ID, Status, comparison results
- **Red headers**: On all key columns for easy identification

### Auto-Cleanup
- All uploaded files in `input/dynamics/` and `input/redash/` are deleted when you close the application
- Comparison outputs and ID files are preserved

## Technical Details

### File Pattern Detection
The system auto-detects files by keywords in filenames:
- **Accreditation**: Contains "accreditation"
- **WCB**: Contains "wcb"
- **Client**: Contains "client" or "cs"

### Special Logic
- **Client Reports**: Uses 'case' column as status (SafeContractor-specific)
- **Accreditation/WCB**: Uses standard 'status' column
- **ID Deduplication**: Automatically removes duplicates before generating queries
- **UUID Validation**: Validates Global Alcumus ID format

## Dependencies

- Python 3.8+
- pandas
- openpyxl
- ttkbootstrap (dark theme)
- tkinterdnd2 (drag and drop)

## License

Internal tool for Cognibox use.
