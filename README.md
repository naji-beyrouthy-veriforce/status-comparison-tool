# D365 vs SafeContractor Status Comparison Tool

Automated tool for comparing Dynamics 365 and SafeContractor status reports with integrated Redash query execution.

## Features

✨ **Fully Automated Workflow**
- Drag & drop D365 exports
- Automatic ID extraction
- Auto-execute Redash queries
- Generate Excel comparison files with XLOOKUP formulas

🎯 **Key Capabilities**
- **Fresh Data Guarantee**: Always uses IDs from uploaded files (no cache)
- **Smart Column Detection**: Automatically identifies ID and status columns
- **Intelligent Formatting**: Red headers on key columns
- **Error Handling**: Retry logic for locked files, detailed error messages
- **Performance Optimized**: 50-60% faster with vectorized operations

## Quick Start

### 1. Prerequisites

```bash
pip install -r requirements.txt
```

### 2. Configure Redash Integration

1. Copy `redash_config_template.py` to `redash_config.py`
2. Fill in your Redash URL and API key
3. Add your query IDs

```python
REDASH_URL = "https://your-redash-instance.com"
API_KEY = "your-api-key-here"

QUERY_IDS = {
    "accreditation": 1266,
    "wcb": 1281,
    "client": 1277
}
```

**⚠️ Important**: `redash_config.py` is gitignored to protect your credentials.

### 3. Run the Application

```bash
python gui_app.py
```

## Usage

### Automated Workflow

1. **Upload D365 Files**: Drag & drop all 3 D365 exports
2. **Click Process**: One button triggers everything
3. **Wait 2-5 minutes**: Extract IDs → Query Redash → Generate comparisons
4. **Done!** Find comparison files in `output/` folder

## Project Structure

```
daily-compare-statuses/
├── automate_comparison.py    # Core processing logic
├── gui_app.py                 # Drag-and-drop GUI
├── redash_integration.py      # Redash API client
├── redash_config_template.py  # Configuration template
└── requirements.txt           # Dependencies
```

## Output

Each comparison file contains:
- **SC Sheet**: SafeContractor data with D365 status XLOOKUP
- **D365 Sheet**: D365 data with SC status XLOOKUP
- **Highlighted columns**: Global Alcumus ID, Status, comparisons

## License

Internal tool for Cognibox use.
