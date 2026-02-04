"""
Configuration file for Status Comparison Tool
Centralizes all constants, patterns, and configuration settings
"""

from pathlib import Path
import re
from openpyxl.styles import Font, PatternFill

# ============================================================================
# DIRECTORY PATHS
# ============================================================================
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
DYNAMICS_DIR = INPUT_DIR / "dynamics"
REDASH_DIR = INPUT_DIR / "redash"
QUERY_IDS_DIR = OUTPUT_DIR / "query_ids"

# ============================================================================
# FILE PATTERNS FOR AUTO-DETECTION
# ============================================================================
# Patterns used to identify files by keywords in their names
D365_PATTERNS = {
    "accreditation": "accreditation",
    "wcb": "wcb",
    "client": ["client", "cs"],  # CS or Client Specific
}

SC_PATTERNS = {"accreditation": "accreditation", "wcb": "wcb", "client": ["client", "cs"]}

# Backwards compatibility - default filenames
D365_FILES = {
    "accreditation": "accreditation_d365.xlsx",
    "wcb": "wcb_d365.xlsx",
    "client": "client_d365.xlsx",
}

SC_FILES = {
    "accreditation": "accreditation_sc.xlsx",
    "wcb": "wcb_sc.xlsx",
    "client": "client_sc.xlsx",
}

# ============================================================================
# VALIDATION SETTINGS
# ============================================================================
ALLOWED_FILE_EXTENSIONS = {".xlsx", ".xls", ".csv"}
MIN_FILE_SIZE_BYTES = 100  # Minimum file size to be considered valid

# ============================================================================
# REGEX PATTERNS
# ============================================================================
# Compiled regex for UUID matching (performance optimization)
UUID_PATTERN = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)

# ============================================================================
# EXCEL FORMATTING
# ============================================================================
# Headers to highlight with red background
HIGHLIGHT_HEADERS = frozenset(
    [
        "global_alcumus_id",
        "global alcumus id",
        "status",
        "d365 status",
        "is it the same?",
        "sc status",
        "status reason",
        "case",
    ]
)

# Header styling
HEADER_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
HEADER_FONT = Font(bold=True, color="000000")

# ============================================================================
# FILE OPERATION SETTINGS
# ============================================================================
# Retry settings for locked files
MAX_FILE_SAVE_RETRIES = 3
FILE_SAVE_RETRY_DELAY_SECONDS = 1

# ============================================================================
# REPORT TYPES
# ============================================================================
REPORT_TYPES = ["accreditation", "wcb", "client"]

# ============================================================================
# CRITICAL BUSINESS LOGIC DOCUMENTATION
# ============================================================================
# ⚠️ IMPORTANT: CLIENT REPORT COMPARISON LOGIC
#
# For CLIENT reports from SafeContractor Redash query:
#   - The 'case' column IS the status column for client-specific global IDs
#   - This is NOT the same as a regular 'status' column
#   - Comparison logic MUST use 'case' column for client reports
#
# For ACCREDITATION/WCB reports:
#   - The 'status' column is used normally
#
# This is the CORRECT behavior per business requirements.
# DO NOT modify this logic without understanding the data structure!
# ============================================================================

CLIENT_STATUS_COLUMN = "case"  # The status column name for client reports

# ============================================================================
# UI MESSAGES & EMOJIS
# ============================================================================
class Messages:
    """Centralized UI messages and emojis for consistent user communication."""

    # Status indicators
    SUCCESS = "✓"
    ERROR = "❌"
    WARNING = "⚠️"
    INFO = "📊"
    DATE = "📅"
    PROCESSING = "▶"
    SUGGESTION = "💡"

    # Common message templates
    @staticmethod
    def processing(report_type: str) -> str:
        return f"\n{Messages.PROCESSING} Processing {report_type.upper()}..."

    @staticmethod
    def error(msg: str) -> str:
        return f"  {Messages.ERROR} Error: {msg}"

    @staticmethod
    def warning(msg: str) -> str:
        return f"  {Messages.WARNING} Warning: {msg}"

    @staticmethod
    def success(msg: str) -> str:
        return f"  {Messages.SUCCESS} {msg}"

    @staticmethod
    def info(msg: str) -> str:
        return f"  {Messages.INFO} {msg}"

    @staticmethod
    def suggestion(msg: str) -> str:
        return f"     {Messages.SUGGESTION} {msg}"

    # Specific error messages
    FILE_NOT_FOUND = "No D365 {report_type} file found, skipping..."
    LOOKING_FOR = "Looking for files containing: {patterns}"
    COLUMN_NOT_FOUND = "'Global Alcumus Id' column not found"
    AVAILABLE_COLUMNS = "Available columns: {columns}"
    ENSURE_EXPORT = "Ensure you exported the correct report with Global Alcumus ID column"
    NO_VALID_UUIDS = "No valid UUIDs extracted from {column} column"
    SAMPLE_VALUES = "Sample values: {values}"
    CHECK_COLUMN = "Check that this column contains proper Global Alcumus IDs"
    MISSING_COLUMNS = "Missing required D365 columns"
    STATUS_COLUMN_MISSING = "Could not find status column in SC data"
    FILE_LOCKED = "File is locked (attempt {attempt}/{max_attempts})"
    CLOSE_FILE = "Close {filename} in Excel and waiting..."
    FILE_STILL_LOCKED = "File still locked after {max_attempts} attempts"
    REMEMBER_TO_CLOSE = "Remember to close {filename} before next run"
    CRITICAL_SAVE_ERROR = "Critical: Cannot save file even with timestamp"
    CHECK_DISK_SPACE = "Check disk space and permissions for: {directory}"
    UNEXPECTED_ERROR = "Unexpected error saving file: {error_type}"
    CHECK_WRITABLE = "Check disk space and ensure output directory is writable"

    # Success messages
    READ_ROWS = "Read {count} rows from {filename}"
    EXTRACTED_IDS = "Extracted and deduplicated {count} unique IDs"
    USING_FRESH_IDS = "Using fresh IDs from today's D365 upload"
    SAVED_TO = "Saved to: {filename}"
    PREVIEW_HEADER = "Preview (first 5 IDs):"
    AND_MORE = "... and {count} more"
    CREATING_COMPARISON = "Creating comparison for {report_type}..."
    ROW_COUNTS = "D365 rows: {d365_count}, SC rows: {sc_count}"
    COLUMN_INFO = "D365 ID column: '{id_col}'"
    STATUS_INFO = "D365 Status column: '{status_col}'"
    UUID_QUALITY = "UUID Quality: {valid}/{total} valid ({null} null, {invalid} invalid)"
    READ_D365 = "Read D365: {count} rows"
    READ_SC = "Read SC: {count} rows"
    CREATED_FILE = "Created: {filename}"
    FAILED_COMPARISON = "Failed to create comparison"
    ALL_FILES_FOUND = "All SC files found - Generating comparisons..."
