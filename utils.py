"""
Utility functions for Status Comparison Tool
Reusable helper functions for file handling, data cleaning, and validation
"""

import pandas as pd
from pathlib import Path
import zipfile
from config import (
    UUID_PATTERN,
    ALLOWED_FILE_EXTENSIONS,
    MIN_FILE_SIZE_BYTES,
    HEADER_FILL,
    HEADER_FONT,
    HIGHLIGHT_HEADERS,
)


def clean_uuid(value):
    """
    Extract UUID from text using compiled regex pattern.

    Args:
        value: String or value containing a UUID

    Returns:
        str: Lowercase UUID if found, None otherwise

    Example:
        >>> clean_uuid('baa140f6-0511-4819-966b-5d33c2ce7e5a CAS-39866')
        'baa140f6-0511-4819-966b-5d33c2ce7e5a'
    """
    if pd.isna(value):
        return None

    match = UUID_PATTERN.search(str(value))
    return match.group(0).lower() if match else None


def format_ids_for_sql(ids):
    """
    Format cleaned IDs for SQL IN clause.

    Args:
        ids: Iterable of ID strings

    Returns:
        str: SQL-formatted string with one ID per line

    Example:
        >>> format_ids_for_sql(['id1', 'id2', 'id3'])
        "'id1',\\n'id2',\\n'id3'"
    """
    return ",\n".join(f"'{id}'" for id in sorted(ids) if id)


def find_column_by_keywords(columns, *keyword_groups):
    """
    Find a column name that contains keywords from any of the provided keyword groups.

    Args:
        columns: List or Index of column names to search
        *keyword_groups: Variable number of tuples, where each tuple contains keywords to match.
                        A column matches if it contains ALL keywords from ANY group.

    Returns:
        str: The first matching column name, or None if no match found

    Example:
        >>> find_column_by_keywords(df.columns, ('global', 'alcumus', 'id'))
        'Global Alcumus ID'  # Returns column containing all three keywords
    """
    # Pre-lowercase all keywords for efficiency
    keyword_groups_lower = [tuple(kw.lower() for kw in group) for group in keyword_groups]

    for col in columns:
        col_lower = col.lower()
        # Check each keyword group
        for keyword_group in keyword_groups_lower:
            # Check if ALL keywords in this group are present in the column name
            if all(keyword in col_lower for keyword in keyword_group):
                return col
    return None


def find_file_by_pattern(directory, patterns, file_suffix=""):
    """
    Find file in directory matching pattern keywords.

    Args:
        directory: Path object or string to search directory
        patterns: String or list of strings to match in filename
        file_suffix: Optional suffix like '_d365' or '_sc' to prioritize

    Returns:
        Path: Path to matching file, or None if not found

    Example:
        >>> find_file_by_pattern(Path('input/dynamics'), 'accreditation', '_d365')
        Path('input/dynamics/accreditation_d365.xlsx')
    """
    directory = Path(directory)
    
    if not directory.exists():
        return None

    # Convert single pattern to list and pre-lowercase
    patterns_lower = (
        [patterns.lower()] if isinstance(patterns, str) else [p.lower() for p in patterns]
    )
    suffix_lower = file_suffix.lower() if file_suffix else None

    # Single pass: collect matches and prioritize those with suffix
    best_match = None

    for file in directory.iterdir():
        if not file.is_file() or file.suffix not in ALLOWED_FILE_EXTENSIONS:
            continue

        filename_lower = file.name.lower()

        # Check if any pattern matches
        if any(pattern in filename_lower for pattern in patterns_lower):
            # If suffix specified and matches, return immediately (best match)
            if suffix_lower and suffix_lower in filename_lower:
                return file
            # Otherwise, keep as backup match
            if not best_match:
                best_match = file

    return best_match


def validate_file_format(file_path):
    """
    Validate file exists and has correct format with detailed diagnostics.

    Args:
        file_path: Path object or string to validate

    Returns:
        tuple: (is_valid: bool, error_message: str, suggested_fix: str)

    Example:
        >>> is_valid, error, fix = validate_file_format('data.xlsx')
        >>> if not is_valid:
        >>>     print(f"{error}\nSuggestion: {fix}")
    """
    file_path = Path(file_path)

    # Check if file exists
    if not file_path.exists():
        parent_dir = file_path.parent
        if not parent_dir.exists():
            return (
                False,
                f"Directory not found: {parent_dir}",
                f"Create the directory: {parent_dir}",
            )

        # List similar files in directory
        similar_files = [f.name for f in parent_dir.glob("*") if f.is_file()][:5]
        suggestion = f"File not found: {file_path.name}"
        if similar_files:
            suggestion += f"\nFiles in {parent_dir.name}: {', '.join(similar_files)}"
        return False, suggestion, "Check the filename and location are correct"

    # Check if it's a file (not directory)
    if not file_path.is_file():
        return (
            False,
            f"Path is not a file: {file_path.name}",
            "Ensure you're selecting a file, not a folder",
        )

    # Check file extension
    if file_path.suffix.lower() not in ALLOWED_FILE_EXTENSIONS:
        allowed = ", ".join(ALLOWED_FILE_EXTENSIONS)
        return (
            False,
            f"Invalid file format: {file_path.suffix}",
            f"Convert file to one of: {allowed}",
        )

    # Check file size (warn if too large or suspiciously small)
    file_size = file_path.stat().st_size
    if file_size == 0:
        return (
            False,
            f"File is empty (0 bytes): {file_path.name}",
            "Re-export the file from the source system",
        )

    if file_size < MIN_FILE_SIZE_BYTES:
        return (
            False,
            f"File is too small ({file_size} bytes): {file_path.name}",
            "This file appears incomplete - re-export with full data",
        )

    # Check if file is locked
    try:
        with open(file_path, "rb") as f:
            f.read(1)
    except PermissionError:
        return (
            False,
            f"File is locked or in use: {file_path.name}",
            "Close the file in Excel or other programs and try again",
        )
    except Exception as e:
        return False, f"Cannot access file: {file_path.name} - {str(e)}", "Check file permissions"

    return True, "Valid", None


def validate_dataframe(df, file_name, required_columns=None):
    """
    Validate DataFrame structure and content with detailed error messages.

    Args:
        df: DataFrame to validate
        file_name: Name of the file for error messages
        required_columns: List of tuples of keywords that must be present

    Returns:
        tuple: (is_valid: bool, error_message: str, suggested_fix: str)

    Example:
        >>> is_valid, error, fix = validate_dataframe(df, 'data.xlsx', [('id',), ('status',)])
        >>> if not is_valid:
        >>>     print(f"{error}\nSuggestion: {fix}")
    """
    # Check if DataFrame is None
    if df is None:
        return (
            False,
            f"{file_name} could not be loaded",
            "Check if the file is corrupted or in use by another program",
        )

    # Check if DataFrame is empty
    if df.empty:
        return (
            False,
            f"{file_name} contains no data",
            "Ensure the Excel file has data rows (not just headers)",
        )

    # Check if DataFrame has any columns
    if len(df.columns) == 0:
        return False, f"{file_name} has no columns", "Check if the Excel file structure is correct"

    # Check for required columns if specified
    if required_columns:
        missing_cols = []
        for col_keywords in required_columns:
            col = find_column_by_keywords(df.columns, col_keywords)
            if not col:
                keywords_str = " + ".join(col_keywords)
                missing_cols.append(keywords_str)

        if missing_cols:
            available = ", ".join([f"'{col}'" for col in df.columns[:10]])
            if len(df.columns) > 10:
                available += ", ..."

            return (
                False,
                f"{file_name} missing required columns: {', '.join(missing_cols)}",
                f"Available columns are: {available}\nEnsure the file has the correct export format",
            )

    # Check for minimum data
    if len(df) < 1:
        return (
            False,
            f"{file_name} has no data rows",
            "The file only contains headers. Export data with actual records",
        )

    # Check for suspiciously small datasets
    if len(df) < 10:
        return (
            True,
            f"Warning: {file_name} only has {len(df)} rows - this seems unusually small",
            "Verify this is the complete export",
        )

    return True, "Valid", None


def apply_header_formatting(worksheet, highlight_headers=None):
    """
    Apply red fill and black bold text to specified headers.

    Args:
        worksheet: openpyxl worksheet object
        highlight_headers: Set/list of header names to highlight (case-insensitive)
                          If None, uses default HIGHLIGHT_HEADERS

    Example:
        >>> apply_header_formatting(worksheet, {'status', 'id'})
    """
    if highlight_headers is None:
        highlight_headers = HIGHLIGHT_HEADERS

    for col_idx in range(1, worksheet.max_column + 1):
        header_value = worksheet.cell(1, col_idx).value
        if header_value and header_value.lower() in highlight_headers:
            worksheet.cell(1, col_idx).fill = HEADER_FILL
            worksheet.cell(1, col_idx).font = HEADER_FONT


def safe_read_excel(file_path):
    """
    Safely read Excel file with comprehensive error handling.

    Args:
        file_path: Path to Excel file

    Returns:
        tuple: (DataFrame or None, error_message or None, suggested_fix or None)

    Example:
        >>> df, error, fix = safe_read_excel('data.xlsx')
        >>> if error:
        >>>     print(f"Error: {error}\nSuggestion: {fix}")
    """
    try:
        df = pd.read_excel(file_path)
        return df, None, None
    except FileNotFoundError:
        return None, f"File not found: {file_path}", "Check the file path is correct"
    except PermissionError:
        return None, f"Permission denied: {file_path}", "Close the file in Excel and try again"
    except pd.errors.EmptyDataError:
        return None, f"File contains no data: {file_path}", "Ensure the Excel file has data"
    except Exception as e:
        error_type = type(e).__name__
        # Provide specific suggestions based on error type
        if "xlrd" in str(e).lower():
            suggestion = "Install xlrd: pip install xlrd (for .xls files)"
        elif "openpyxl" in str(e).lower():
            suggestion = "Install openpyxl: pip install openpyxl (for .xlsx files)"
        elif "corrupt" in str(e).lower() or "damaged" in str(e).lower():
            suggestion = "File appears corrupted - try re-exporting from source"
        else:
            suggestion = "Try opening the file in Excel to verify it's valid"

        return None, f"{error_type}: {str(e)}", suggestion


def validate_uuid_data(df, id_column, file_name):
    """
    Validate UUID data quality in a DataFrame.

    Args:
        df: DataFrame containing UUID data
        id_column: Name of the column containing UUIDs
        file_name: Name of file for error messages

    Returns:
        tuple: (is_valid: bool, error_message: str, suggested_fix: str, stats: dict)
    """
    from config import UUID_PATTERN

    total_rows = len(df)
    null_count = df[id_column].isna().sum()
    non_null = df[id_column].notna().sum()

    # Extract valid UUIDs
    valid_uuids = (
        df[id_column].dropna().apply(lambda x: UUID_PATTERN.search(str(x)) is not None).sum()
    )

    stats = {
        "total": total_rows,
        "null": null_count,
        "non_null": non_null,
        "valid_uuids": valid_uuids,
        "invalid": non_null - valid_uuids,
    }

    # Critical: No UUIDs at all
    if valid_uuids == 0:
        sample = df[id_column].dropna().head(3).tolist()
        return (
            False,
            f"{file_name}: No valid UUIDs found in '{id_column}' column",
            f"Sample values: {sample}\nEnsure this column contains Global Alcumus IDs",
            stats,
        )

    # Warning: High percentage of invalid UUIDs
    if non_null > 0:
        invalid_percentage = (stats["invalid"] / non_null) * 100
        if invalid_percentage > 10:
            return (
                True,
                f"{file_name}: {invalid_percentage:.1f}% of '{id_column}' values are not valid UUIDs ({stats['invalid']}/{non_null})",
                "This may indicate data quality issues. Review the source data",
                stats,
            )

    return True, "UUID data quality is good", None, stats


def check_file_accessibility(file_path, mode="read"):
    """
    Check if a file can be accessed for reading or writing.

    Args:
        file_path: Path to file
        mode: 'read' or 'write'

    Returns:
        tuple: (is_accessible: bool, error_message: str, suggested_fix: str)
    """
    file_path = Path(file_path)

    try:
        if mode == "read":
            with open(file_path, "rb") as f:
                f.read(1)
            return True, "File is accessible", None
        else:  # write mode
            # Try to open in append mode to check write access without modifying
            with open(file_path, "a") as f:
                pass
            return True, "File is writable", None

    except PermissionError:
        if mode == "read":
            return (
                False,
                f"Cannot read file: {file_path.name}",
                "Check file permissions or close it in other programs",
            )
        else:
            return (
                False,
                f"Cannot write to file: {file_path.name}",
                "Close the file in Excel or other programs",
            )
    except FileNotFoundError:
        return False, f"File not found: {file_path.name}", "Verify the file path is correct"
    except Exception as e:
        return False, f"Access error: {str(e)}", "Check file permissions and try again"


def create_comparison_zip(folders_to_zip, output_zip_path):
    """
    Create a zip file containing the specified folders with maximum LZMA compression.
    
    Args:
        folders_to_zip: List of Path objects representing directories to zip
        output_zip_path: Path object for the output zip file
        
    Returns:
        tuple: (success: bool, message: str, zip_path: Path or None)
        
    Example:
        >>> folders = [Path("output/accreditation"), Path("output/wcb"), Path("output/client")]
        >>> success, msg, path = create_comparison_zip(folders, Path("output/comparison.zip"))
    """
    from config import MAX_ZIP_SIZE_MB
    
    try:
        # Remove existing zip file if it exists
        if output_zip_path.exists():
            output_zip_path.unlink()
        
        # Create the zip file with LZMA compression (better for already-compressed Excel files)
        with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_LZMA) as zipf:
            for folder in folders_to_zip:
                if not folder.exists():
                    continue
                    
                # Add all files in the folder to the zip
                for file_path in folder.rglob('*'):
                    if file_path.is_file():
                        # Calculate the archive name (relative path from parent of folder)
                        arcname = file_path.relative_to(folder.parent)
                        zipf.write(file_path, arcname)
        
        # Verify the zip was created and check size
        if output_zip_path.exists() and output_zip_path.stat().st_size > 0:
            file_size = output_zip_path.stat().st_size
            size_mb = file_size / (1024 * 1024)
            
            # Check if file exceeds maximum size
            if size_mb > MAX_ZIP_SIZE_MB:
                warning_msg = (
                    f"⚠️ WARNING: comparison.zip size ({size_mb:.2f} MB) exceeds {MAX_ZIP_SIZE_MB} MB limit!\n"
                    f"     The file may be too large for email attachments.\n"
                    f"     Consider uploading to a file sharing service instead."
                )
                return (True, warning_msg, output_zip_path)
            else:
                success_msg = f"Successfully created comparison.zip ({size_mb:.2f} MB / {MAX_ZIP_SIZE_MB} MB limit)"
                return (True, success_msg, output_zip_path)
        else:
            return (False, "Zip file was created but is empty or invalid", None)
            
    except PermissionError:
        return (False, f"Cannot create zip file: Permission denied", None)
    except Exception as e:
        return (False, f"Error creating zip file: {str(e)}", None)
