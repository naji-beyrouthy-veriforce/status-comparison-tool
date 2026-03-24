"""
Configuration file for Status Comparison Tool
Centralizes all constants, patterns, and configuration settings
"""

from pathlib import Path
import os
import re
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from openpyxl.styles import Font, PatternFill

# ============================================================================
# DIRECTORY PATHS
# ============================================================================
BASE_DIR = Path(__file__).parent.parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
DYNAMICS_DIR = INPUT_DIR / "dynamics"
REDASH_DIR = INPUT_DIR / "redash"
QUERY_IDS_DIR = OUTPUT_DIR / "query_ids"
LOG_DIR = BASE_DIR / "logs"

def get_dated_comparison_dir():
    """
    Get the dated comparison directory for today's comparisons.
    Creates a folder like: output/comparison_2026-02-18/
    
    Returns:
        Path: Path to today's dated comparison folder
    """
    date_stamp = datetime.now().strftime("%Y-%m-%d")
    return OUTPUT_DIR / f"comparison_{date_stamp}"

# ============================================================================
# FILE PATTERNS FOR AUTO-DETECTION
# ============================================================================
# Patterns used to identify files by keywords in their names
D365_PATTERNS = {
    "accreditation": "accreditation",
    "wcb": "wcb",
    "client": ["client", "cs"],  # CS or Client Specific
    "critical_document": ["critical", "cd"],  # Critical Document
}

SC_PATTERNS = {"accreditation": "accreditation", "wcb": "wcb", "client": ["client", "cs"], "critical_document": ["critical", "cd"]}

# Backwards compatibility - default filenames
D365_FILES = {
    "accreditation": "accreditation_d365.xlsx",
    "wcb": "wcb_d365.xlsx",
    "client": "client_d365.xlsx",
    "critical_document": "critical_document_d365.xlsx",
}

SC_FILES = {
    "accreditation": "accreditation_sc.xlsx",
    "wcb": "wcb_sc.xlsx",
    "client": "client_sc.xlsx",
    "critical_document": "critical_document_sc.xlsx",
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
REPORT_TYPES = ["accreditation", "wcb", "client", "critical_document"]

# ============================================================================
# CRITICAL BUSINESS LOGIC DOCUMENTATION
# ============================================================================
# ⚠️ IMPORTANT: CLIENT & CRITICAL DOCUMENT REPORT COMPARISON LOGIC
#
# For CLIENT and CRITICAL DOCUMENT reports from SafeContractor Redash queries:
#   - The 'case' column IS the status column
#   - This is NOT the same as a regular 'status' column
#   - Comparison logic MUST use 'case' column for these reports
#
# For ACCREDITATION/WCB reports:
#   - The 'status' column is used normally
#
# This is the CORRECT behavior per business requirements.
# DO NOT modify this logic without understanding the data structure!
# ============================================================================

CLIENT_STATUS_COLUMN = "case"  # The status column name for client/critical_document reports

# Report types whose SC Redash query returns status in the 'case' column
CASE_COLUMN_REPORT_TYPES = frozenset({"client", "critical_document"})

# ============================================================================
# REDASH API CONFIGURATION
# ============================================================================
REDASH_BASE_URL = "https://redash.cognibox.net"
REDASH_API_KEY = os.environ.get("REDASH_API_KEY", "")

REDASH_QUERY_IDS = {
    "accreditation": 1460,
    "wcb": 1281,
    "client": 1277,
    "critical_document": 1464,
}

# Polling settings for query execution
REDASH_POLL_INTERVAL = 3    # seconds between status checks
REDASH_POLL_TIMEOUT = 300   # 5 minutes max wait

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================
# Log file settings
LOG_LEVEL = logging.INFO  # Can be changed to DEBUG for more detailed logs
LOG_MAX_BYTES = 10 * 1024 * 1024  # 10 MB
LOG_BACKUP_COUNT = 5  # Keep 5 backup files

# Log format
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

def setup_logging(log_name="comparison_tool", console_output=True, file_output=True):
    """
    Setup logging configuration for the application.
    
    Creates a logger that writes to both file and console with rotating file handler.
    Log files are stored in logs/ directory with daily rotation.
    
    Args:
        log_name: Name of the logger (used for log filename)
        console_output: Whether to output logs to console
        file_output: Whether to output logs to file
    
    Returns:
        logging.Logger: Configured logger instance
    """
    # Create logs directory if it doesn't exist
    LOG_DIR.mkdir(exist_ok=True)
    
    # Create logger
    logger = logging.getLogger(log_name)
    logger.setLevel(LOG_LEVEL)
    
    # Remove existing handlers to avoid duplicates
    logger.handlers.clear()
    
    # Create formatter
    formatter = logging.Formatter(LOG_FORMAT, datefmt=LOG_DATE_FORMAT)
    
    # File handler with rotation
    if file_output:
        log_file = LOG_DIR / f"{log_name}_{datetime.now():%Y%m%d}.log"
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=LOG_MAX_BYTES,
            backupCount=LOG_BACKUP_COUNT,
            encoding='utf-8'
        )
        file_handler.setLevel(LOG_LEVEL)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    # Console handler
    if console_output:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(LOG_LEVEL)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    # Prevent propagation to root logger
    logger.propagate = False
    
    return logger

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
