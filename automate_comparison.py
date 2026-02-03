"""
Dynamics 365 vs SafeContractor Status Comparison
Automates ID extraction and status comparison reporting
"""

import pandas as pd
import re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill
import warnings

# Import Redash integration
try:
    from redash_integration import execute_redash_query, save_redash_results
    REDASH_AVAILABLE = True
except ImportError:
    REDASH_AVAILABLE = False
    print("⚠ Warning: Redash integration not available")

# Suppress openpyxl style warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configuration
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

# Compiled regex for UUID matching (performance optimization)
UUID_PATTERN = re.compile(r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}')

# Header formatting constants (created once for reuse)
HIGHLIGHT_HEADERS = frozenset(['global_alcumus_id', 'global alcumus id', 'status', 'd365 status', 
                               'is it the same?', 'sc status', 'status reason', 'case'])
HEADER_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
HEADER_FONT = Font(bold=True, color="000000")

# File search patterns (case-insensitive keyword matching)
D365_PATTERNS = {
    "accreditation": "accreditation",
    "wcb": "wcb",
    "client": ["client", "cs"]  # CS or Client Specific
}

SC_PATTERNS = {
    "accreditation": "accreditation",
    "wcb": "wcb",
    "client": ["client", "cs"]
}

# For backwards compatibility
D365_FILES = {
    "accreditation": "accreditation_d365.xlsx",
    "wcb": "wcb_d365.xlsx",
    "client": "client_d365.xlsx"
}

SC_FILES = {
    "accreditation": "accreditation_sc.xlsx",
    "wcb": "wcb_sc.xlsx",
    "client": "client_sc.xlsx"
}


def find_file_by_pattern(directory, patterns, file_suffix=""):
    """
    Find file in directory matching pattern keywords.
    patterns can be a string or list of strings to match.
    file_suffix can be '_d365' or '_sc' to help differentiate.
    """
    if not directory.exists():
        return None
    
    # Convert single pattern to list and pre-lowercase
    patterns_lower = [patterns.lower()] if isinstance(patterns, str) else [p.lower() for p in patterns]
    suffix_lower = file_suffix.lower() if file_suffix else None
    allowed_extensions = {'.xlsx', '.xls', '.csv'}
    
    # Single pass: collect matches and prioritize those with suffix
    best_match = None
    
    for file in directory.iterdir():
        if not file.is_file() or file.suffix not in allowed_extensions:
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


def clean_uuid(value):
    """
    Extract UUID from text using compiled regex pattern.
    Example: 'baa140f6-0511-4819-966b-5d33c2ce7e5a CAS-39866' -> 'baa140f6-0511-4819-966b-5d33c2ce7e5a'
    """
    if pd.isna(value):
        return None
    
    match = UUID_PATTERN.search(str(value))
    return match.group(0).lower() if match else None


def format_ids_for_sql(ids):
    """
    Format cleaned IDs for SQL IN clause.
    Returns: 'id1',\n'id2',\n'id3' (one per line, no trailing comma)
    """
    # Filter out None/empty and format directly (assuming ids already unique)
    return ',\n'.join(f"'{id}'" for id in sorted(ids) if id)


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
        find_column_by_keywords(df.columns, ('global', 'alcumus', 'id'))
        # Returns column containing 'global' AND 'alcumus' AND 'id'
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


def apply_header_formatting(worksheet, highlight_headers=None):
    """
    Apply red fill and black bold text to specified headers.
    Args:
        worksheet: openpyxl worksheet
        highlight_headers: Set/list of header names to highlight (case-insensitive)
    """
    if highlight_headers is None:
        highlight_headers = HIGHLIGHT_HEADERS
    
    for col_idx in range(1, worksheet.max_column + 1):
        header_value = worksheet.cell(1, col_idx).value
        if header_value and header_value.lower() in highlight_headers:
            worksheet.cell(1, col_idx).fill = HEADER_FILL
            worksheet.cell(1, col_idx).font = HEADER_FONT
            worksheet.cell(1, col_idx).fill = HEADER_FILL
            worksheet.cell(1, col_idx).font = HEADER_FONT


def extract_and_save_ids():
    """
    Step 1: Extract IDs from D365 files and save SQL-ready lists
    """
    print("\n" + "="*70)
    print("STEP 1: EXTRACTING IDs FROM D365 FILES")
    print("="*70)
    
    # Process Accreditation and WCB only (Client doesn't need ID extraction)
    for report_type in ["accreditation", "wcb"]:
        print(f"\n▶ Processing {report_type.upper()}...")
        
        # Find D365 file by pattern in dynamics subdirectory
        dynamics_dir = INPUT_DIR / "dynamics"
        file_path = find_file_by_pattern(dynamics_dir, D365_PATTERNS[report_type], "d365")
        if not file_path:
            # Try without suffix requirement
            file_path = find_file_by_pattern(dynamics_dir, D365_PATTERNS[report_type])
        
        if not file_path:
            print(f"  ⚠ Warning: No D365 {report_type} file found, skipping...")
            print(f"     Looking for files containing: {D365_PATTERNS[report_type]}")
            continue
        
        df = pd.read_excel(file_path)
        print(f"  ✓ Read {len(df)} rows from {file_path.name}")
        
        # Find Global Alcumus Id column
        id_col = find_column_by_keywords(df.columns, ('global', 'alcumus', 'id'))
        
        if not id_col:
            print(f"  ❌ Error: 'Global Alcumus Id' column not found")
            continue
        
        # Extract and clean IDs using vectorized operation
        unique_ids = df[id_col].dropna().map(clean_uuid).dropna().unique()
        unique_ids = sorted(unique_ids)
        
        print(f"  ✓ Extracted and deduplicated {len(unique_ids)} unique IDs")
        print(f"  📅 Using fresh IDs from today's D365 upload")
        
        # Format for SQL
        sql_formatted = format_ids_for_sql(unique_ids)
        
        # Save to file in query_ids subfolder
        query_ids_dir = OUTPUT_DIR / "query_ids"
        query_ids_dir.mkdir(exist_ok=True)
        
        output_file = query_ids_dir / f"{report_type}_ids.sql.txt"
        
        with open(output_file, 'w') as f:
            f.write(sql_formatted)
        
        print(f"  ✓ Saved to: {output_file.name}")
        
        # Show preview
        lines = sql_formatted.split('\n')
        print(f"  Preview (first 5 IDs):")
        for line in lines[:5]:
            print(f"    {line}")
        if len(lines) > 5:
            print(f"    ... and {len(lines) - 5} more")
        
        # Execute Redash query automatically if integration is available
        if REDASH_AVAILABLE:
            try:
                print(f"\n  🚀 Executing Redash query automatically...")
                df_results = execute_redash_query(report_type, unique_ids)
                
                if df_results is not None and len(df_results) > 0:
                    # Save results to redash directory
                    redash_dir = INPUT_DIR / "redash"
                    redash_dir.mkdir(exist_ok=True)
                    save_redash_results(report_type, df_results, redash_dir)
                else:
                    print(f"  ⚠ Warning: Query returned no results")
            except Exception as e:
                print(f"  ⚠ Warning: Redash query failed: {e}")
                print(f"  → You can manually run the query using the IDs in {output_file.name}")

    
    # Execute Client query if Redash is available
    if REDASH_AVAILABLE:
        print(f"\n▶ Processing CLIENT...")
        
        # Find Client D365 file and extract IDs
        dynamics_dir = INPUT_DIR / "dynamics"
        client_file_path = find_file_by_pattern(dynamics_dir, D365_PATTERNS["client"], "d365")
        if not client_file_path:
            client_file_path = find_file_by_pattern(dynamics_dir, D365_PATTERNS["client"])
        
        if client_file_path:
            try:
                df_client = pd.read_excel(client_file_path)
                print(f"  ✓ Read {len(df_client)} rows from {client_file_path.name}")
                
                # Find and extract IDs
                id_col = find_column_by_keywords(df_client.columns, ('global', 'alcumus', 'id'))
                
                if id_col:
                    client_ids = df_client[id_col].dropna().map(clean_uuid).dropna().unique()
                    client_ids = sorted(client_ids)
                    print(f"  ✓ Extracted {len(client_ids)} unique IDs")
                    print(f"  📅 Using fresh IDs from today's D365 Client upload")
                    
                    # Execute Redash query with IDs
                    df_client_results = execute_redash_query("client", client_ids)
                    
                    if df_client_results is not None and len(df_client_results) > 0:
                        redash_dir = INPUT_DIR / "redash"
                        redash_dir.mkdir(exist_ok=True)
                        save_redash_results("client", df_client_results, redash_dir)
                    else:
                        print(f"  ⚠ Warning: Client query returned no results")
                else:
                    print(f"  ⚠ Warning: Could not find ID column in Client file")
                    
            except Exception as e:
                print(f"  ⚠ Warning: Client query failed: {e}")
        else:
            print(f"  ⚠ Warning: No D365 Client file found, skipping client query")
    
    print("\n" + "="*70)
    if REDASH_AVAILABLE:
        print("✅ ID EXTRACTION AND REDASH QUERIES COMPLETED!")
        print("")
        print("All Redash queries have been executed automatically.")
        print("Results saved to input/redash/ folder.")
        print("")
        print("Run this script again to generate comparison files.")
    else:
        print("NEXT STEP (Manual Process):")
        print("1. Copy IDs from output/query_ids/*.sql.txt files")
        print("2. Paste into Redash IN (...) clauses")
        print("3. Download SC results as accreditation_sc.xlsx, wcb_sc.xlsx, client_sc.xlsx")
        print("4. Place SafeContractor (Redash) files in input/redash/ folder")
        print("5. Run this script again to generate comparisons")
    print("="*70 + "\n")


def create_comparison_excel(report_type, df_d365, df_sc, include_qual_url=False):
    """
    Create comparison Excel file with SC and D365 sheets
    Uses pandas merge for status matching (no Excel formulas)
    """
    print(f"\n  📊 Creating comparison for {report_type}...")
    print(f"     D365 rows: {len(df_d365)}, SC rows: {len(df_sc)}")
    
    # Find D365 columns using helper function
    id_col_d365 = find_column_by_keywords(df_d365.columns, ('global', 'alcumus', 'id'))
    status_col_d365 = find_column_by_keywords(df_d365.columns, ('status', 'reason'))
    qual_url_col = find_column_by_keywords(df_d365.columns, ('qualification', 'url')) if include_qual_url else None
    
    if not id_col_d365 or not status_col_d365:
        print(f"     ❌ Missing required D365 columns")
        print(f"        ID column: {id_col_d365}")
        print(f"        Status column: {status_col_d365}")
        return None
    
    print(f"     D365 ID column: '{id_col_d365}'")
    print(f"     D365 Status column: '{status_col_d365}'")
    
    # Find SC columns intelligently
    id_col_sc = (find_column_by_keywords(df_sc.columns, ('global', 'alcumus', 'id'), ('id', 'alcumus')) 
                 or df_sc.columns[0])
    
    # Find status column in SC data (any column with 'status' that isn't the ID column)
    status_col_sc = next((col for col in df_sc.columns 
                         if 'status' in col.lower() and col != id_col_sc), None)
    
    # If status column not found by name, use the column after the ID column
    if not status_col_sc:
        id_col_index = df_sc.columns.get_loc(id_col_sc)
        if id_col_index + 1 < len(df_sc.columns):
            status_col_sc = df_sc.columns[id_col_index + 1]
        else:
            # Fallback: look for a column with string data that might be status
            for col in df_sc.columns:
                if col != id_col_sc and df_sc[col].dtype == 'object':
                    status_col_sc = col
                    break
    
    if not status_col_sc:
        print(f"     ❌ Could not find status column in SC data")
        print(f"        Available columns: {list(df_sc.columns)}")
        return None
    
    print(f"     SC ID column: '{id_col_sc}'")
    print(f"     SC Status column: '{status_col_sc}'")
    
    # Clean IDs in both dataframes
    df_d365['clean_id'] = df_d365[id_col_d365].apply(clean_uuid)
    df_sc['clean_id'] = df_sc[id_col_sc].apply(clean_uuid)
    
    # Verify cleaned IDs
    d365_clean_count = df_d365['clean_id'].notna().sum()
    sc_clean_count = df_sc['clean_id'].notna().sum()
    print(f"     D365 cleaned IDs: {d365_clean_count}/{len(df_d365)}")
    print(f"     SC cleaned IDs: {sc_clean_count}/{len(df_sc)}")
    
    # Check for matches
    common_ids = set(df_d365['clean_id'].dropna()) & set(df_sc['clean_id'].dropna())
    print(f"     Common IDs found: {len(common_ids)}")
    
    if len(common_ids) == 0:
        print(f"     ⚠ WARNING: No matching IDs found between D365 and SC!")
        print(f"     Sample D365 IDs: {list(df_d365['clean_id'].dropna()[:3])}")
        print(f"     Sample SC IDs: {list(df_sc['clean_id'].dropna()[:3])}")
    
    # ===== CREATE EXCEL FILE WITH TWO SHEETS AND XLOOKUP FORMULAS =====
    wb = Workbook()
    wb.remove(wb.active)
    
    # ===== SC SHEET (CREATED FIRST) =====
    ws_sc = wb.create_sheet("SC")
    
    # Write SC data (preserve original column order)
    for r_idx, row in enumerate(dataframe_to_rows(df_sc.drop(columns=['clean_id']), index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws_sc.cell(row=r_idx, column=c_idx, value=value)
    
    # Find SC ID and Status column positions
    sc_cols = list(df_sc.drop(columns=['clean_id']).columns)
    sc_id_col_idx = sc_cols.index(id_col_sc) + 1
    sc_status_col_idx_orig = sc_cols.index(status_col_sc) + 1
    sc_id_col_letter = ws_sc.cell(1, sc_id_col_idx).column_letter
    sc_status_col_letter = ws_sc.cell(1, sc_status_col_idx_orig).column_letter
    
    # Determine where to insert new columns based on report type
    sc_cols_lower = {col.lower(): idx for idx, col in enumerate(sc_cols, 1)}
    is_client = report_type.lower() == 'client'
    
    if is_client and 'case' in sc_cols_lower:
        # For client reports, insert after the 'case' column
        insert_after_idx = sc_cols_lower['case']
        
        # Insert two columns: "D365 Status" and "Is it the same?"
        ws_sc.insert_cols(insert_after_idx + 1, 2)
        
        # Update sc_status_col_letter since we inserted columns before it (if applicable)
        if sc_status_col_idx_orig > insert_after_idx:
            sc_status_col_idx_orig += 2
            sc_status_col_letter = ws_sc.cell(1, sc_status_col_idx_orig).column_letter
        
        # Set headers for inserted columns
        ws_sc.cell(1, insert_after_idx + 1, "D365 Status")
        ws_sc.cell(1, insert_after_idx + 2, "Is it the same?")
    else:
        # For Accreditation and WCB, add columns at the end (no insertion needed)
        insert_after_idx = len(sc_cols)
        
        # Set headers in the last two columns
        ws_sc.cell(1, insert_after_idx + 1, "D365 Status")
        ws_sc.cell(1, insert_after_idx + 2, "Is it the same?")
    
    # Format specific headers (red fill, black bold text)
    apply_header_formatting(ws_sc)
    
    # ===== D365 SHEET (CREATED SECOND) =====
    ws_d365 = wb.create_sheet("D365")
    
    # Write D365 data (preserve original column order)
    for r_idx, row in enumerate(dataframe_to_rows(df_d365.drop(columns=['clean_id']), index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws_d365.cell(row=r_idx, column=c_idx, value=value)
    
    # Find D365 ID and Status column positions (1-indexed for Excel)
    d365_cols = list(df_d365.drop(columns=['clean_id']).columns)
    d365_id_col_idx = d365_cols.index(id_col_d365) + 1
    d365_status_col_idx = d365_cols.index(status_col_d365) + 1
    d365_id_col_letter = ws_d365.cell(1, d365_id_col_idx).column_letter
    
    # Add SC Status column with XLOOKUP formula
    sc_status_col_idx = len(d365_cols) + 1
    ws_d365.cell(1, sc_status_col_idx, "SC Status")
    
    # Add "Is it the same?" column
    is_same_col_idx = sc_status_col_idx + 1
    ws_d365.cell(1, is_same_col_idx, "Is it the same?")
    
    # Format specific headers (red fill, black bold text)
    apply_header_formatting(ws_d365)
    
    # Add XLOOKUP formulas for D365 sheet (row 2 onwards)
    # Cache column letters for better performance
    d365_status_col_letter = ws_d365.cell(1, d365_status_col_idx).column_letter
    sc_status_lookup_col_letter = ws_d365.cell(1, sc_status_col_idx).column_letter
    is_same_col_letter = ws_d365.cell(1, is_same_col_idx).column_letter
    
    for row_idx in range(2, len(df_d365) + 2):
        # XLOOKUP with _xlfn prefix and entire column references
        xlookup_formula = f'=_xlfn.XLOOKUP({d365_id_col_letter}{row_idx},SC!{sc_id_col_letter}:{sc_id_col_letter},SC!{sc_status_col_letter}:{sc_status_col_letter},"Not found",0)'
        ws_d365.cell(row_idx, sc_status_col_idx, xlookup_formula)
        
        # Is it the same? formula
        ws_d365.cell(row_idx, is_same_col_idx, f'={d365_status_col_letter}{row_idx}={sc_status_lookup_col_letter}{row_idx}')
    
    # Add XLOOKUP formulas for SC sheet (row 2 onwards)
    d365_status_col_letter_ref = ws_d365.cell(1, d365_status_col_idx).column_letter
    d365_lookup_col_letter = ws_sc.cell(1, insert_after_idx + 1).column_letter
    is_same_col_letter_sc = ws_sc.cell(1, insert_after_idx + 2).column_letter
    
    # Pre-calculate comparison column for client reports
    comparison_col_letter = sc_status_col_letter
    if report_type.lower() == 'client':
        case_col_idx = next((idx for idx, col in enumerate(sc_cols, 1) if col.lower() == 'case'), None)
        if case_col_idx:
            # Adjust if columns were inserted before the case column
            if case_col_idx > insert_after_idx:
                case_col_idx += 2
            comparison_col_letter = ws_sc.cell(1, case_col_idx).column_letter
    
    for row_idx in range(2, len(df_sc) + 2):
        # XLOOKUP with _xlfn prefix and entire column references
        xlookup_formula = f'=_xlfn.XLOOKUP({sc_id_col_letter}{row_idx},D365!{d365_id_col_letter}:{d365_id_col_letter},D365!{d365_status_col_letter_ref}:{d365_status_col_letter_ref},"Not found",0)'
        ws_sc.cell(row_idx, insert_after_idx + 1, xlookup_formula)
        
        # Is it the same? formula
        ws_sc.cell(row_idx, insert_after_idx + 2, f'={comparison_col_letter}{row_idx}={d365_lookup_col_letter}{row_idx}')
    
    # Save file with retry logic for locked files
    output_file = OUTPUT_DIR / f"{report_type.title()}_Comparison.xlsx"
    
    # Try to save with retries
    max_retries = 3
    retry_delay = 1  # seconds
    
    for attempt in range(max_retries):
        try:
            wb.save(output_file)
            return output_file
            
        except PermissionError:
            if attempt < max_retries - 1:
                # Try again after a short delay
                print(f"     ⚠️  File is locked (attempt {attempt + 1}/{max_retries}), retrying...")
                import time
                time.sleep(retry_delay)
            else:
                # Final attempt failed - save with timestamp
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = OUTPUT_DIR / f"{report_type.title()}_Comparison_{timestamp}.xlsx"
                
                try:
                    wb.save(backup_file)
                    print(f"     ⚠️  Original file locked - saved as: {backup_file.name}")
                    print(f"     💡 Please close {output_file.name} in Excel for next time")
                    return backup_file
                except Exception as e:
                    print(f"     ❌ Failed to save even with timestamp: {e}")
                    raise
        
        except Exception as e:
            print(f"     ❌ Unexpected error saving file: {e}")
            raise
    
    return output_file


def generate_comparisons():
    """
    Step 2: Generate comparison Excel files
    """
    print("\n" + "="*70)
    print("STEP 2: GENERATING COMPARISON FILES")
    print("="*70)
    
    success_count = 0
    
    for report_type in ["accreditation", "wcb", "client"]:
        print(f"\n▶ Processing {report_type.upper()}...")
        
        # Check if files exist in subdirectories
        dynamics_dir = INPUT_DIR / "dynamics"
        redash_dir = INPUT_DIR / "redash"
        d365_file = dynamics_dir / D365_FILES[report_type]
        sc_file = redash_dir / SC_FILES[report_type]
        
        if not d365_file.exists():
            print(f"  ⚠ Warning: {d365_file.name} not found, skipping...")
            continue
        
        if not sc_file.exists():
            print(f"  ⚠ Warning: {sc_file.name} not found, skipping...")
            continue
        
        # Read files
        df_d365 = pd.read_excel(d365_file)
        df_sc = pd.read_excel(sc_file)
        
        print(f"  ✓ Read D365: {len(df_d365)} rows")
        print(f"  ✓ Read SC: {len(df_sc)} rows")
        
        try:
            # Create comparison (include Qualification URL for WCB)
            include_qual_url = (report_type == "wcb")
            output_file = create_comparison_excel(report_type, df_d365, df_sc, include_qual_url)
            
            if output_file:
                print(f"  ✓ Created: {output_file.name}")
                success_count += 1
            else:
                print(f"  ❌ Failed to create comparison")
                
        except Exception as e:
            print(f"❌ Error processing {report_type}: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    print("\n" + "="*70)
    if success_count > 0:
        print(f"SUCCESS! Created {success_count} comparison file(s) in output/")
    else:
        print("No comparison files were created. Check file locations.")
    print("="*70 + "\n")


def main():
    """
    Main execution flow with automatic Redash integration
    """
    print("\n" + "="*70)
    print("DYNAMICS 365 vs SAFECONTRACTOR STATUS COMPARISON")
    if REDASH_AVAILABLE:
        print("✨ Redash Auto-Query: ENABLED")
    print("="*70)
    
    # Check if SC files exist (determines which step to run)
    redash_dir = INPUT_DIR / "redash"
    sc_files_exist = all(
        find_file_by_pattern(redash_dir, SC_PATTERNS[t]) is not None 
        for t in ["accreditation", "wcb", "client"]
    )
    
    if sc_files_exist:
        print("\n✓ All SC files found - Generating comparisons...")
        generate_comparisons()
    else:
        print("\n⚠ SC files not found - Starting with ID extraction and Redash queries...")
        extract_and_save_ids()
        
        # If Redash integration succeeded, automatically generate comparisons
        if REDASH_AVAILABLE:
            # Check again if files were created by Redash
            sc_files_exist = all(
                find_file_by_pattern(redash_dir, SC_PATTERNS[t]) is not None 
                for t in ["accreditation", "wcb", "client"]
            )
            
            if sc_files_exist:
                print("\n" + "="*70)
                print("🎯 All Redash queries completed successfully!")
                print("="*70)
                input("\nPress Enter to generate comparison files...")
                generate_comparisons()
            else:
                print("\n⚠ Note: Some Redash queries may have failed.")
                print("   Check the output above for any errors.")


if __name__ == "__main__":
    main()
