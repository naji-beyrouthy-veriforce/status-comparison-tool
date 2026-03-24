"""
Email Report Generator
Automates the generation of comparison email reports from Excel output files

This script replicates the Excel XLOOKUP formula logic by merging data,
since openpyxl-created formulas aren't calculated until opened in Excel.

Manual verification process replicated:
- SC Sheet: Compares SC status vs D365 status (via merge/XLOOKUP logic)
- D365 Sheet: Finds D365 records not in SC, counts by Status Reason
"""

import pandas as pd
from pathlib import Path
import sys
from datetime import datetime

# Import configuration
from .config import OUTPUT_DIR, setup_logging, get_dated_comparison_dir

# Import utilities  
from .utils import find_sc_status_column, find_column_by_keywords

# Setup logging
logger = setup_logging("email_report", console_output=True, file_output=True)


def read_comparison_file(file_path):
    """
    Read both sheets from a comparison Excel file.
    
    Note: Reads raw data, not calculated formulas. The analysis functions
    will replicate the XLOOKUP formula logic by merging data.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        tuple: (sc_df, d365_df) DataFrames or (None, None) if error
    """
    try:
        logger.info(f"Reading comparison file: {file_path.name}")
        # Read both sheets - formulas won't be calculated, we'll merge data instead
        sc_df = pd.read_excel(file_path, sheet_name="SC")
        logger.debug(f"SC sheet loaded: {len(sc_df)} rows")
        d365_df = pd.read_excel(file_path, sheet_name="D365")
        logger.debug(f"D365 sheet loaded: {len(d365_df)} rows")
        logger.info(f"Successfully read {file_path.name} - SC: {len(sc_df)} rows, D365: {len(d365_df)} rows")
        return sc_df, d365_df
    except Exception as e:
        logger.error(f"Error reading {file_path.name}: {e}")
        return None, None


def analyze_sc_sheet(df_sc, df_d365, report_type="client"):
    """
    Analyze SC sheet by replicating the Excel XLOOKUP formula logic.
    
    Excel Formula Logic:
    - D365 Status column: =XLOOKUP(sc_id, D365!id, D365!status, "Not found")
    - Is it the same?: =sc_status = d365_status
    
    We replicate this by merging SC and D365 data on ID.
    
    Args:
        df_sc: DataFrame from SC sheet
        df_d365: DataFrame from D365 sheet (needed for XLOOKUP replication)
        report_type: Type of report (client, wcb, accreditation, or critical_document)
        
    Returns:
        dict: Statistics about differences
    """
    logger.info(f"Analyzing SC sheet for {report_type}")
    if df_sc is None or df_sc.empty or df_d365 is None or df_d365.empty:
        logger.warning(f"Empty or None dataframes for {report_type} SC analysis")
        return {"differences": 0, "not_found": 0}
    
    # Find ID columns using centralized helper function
    sc_id_col = find_column_by_keywords(df_sc.columns, ("global", "alcumus", "id"))
    d365_id_col = find_column_by_keywords(df_d365.columns, ("global", "alcumus", "id"))
    
    if sc_id_col is None or d365_id_col is None:
        logger.warning(f"Could not find ID columns for {report_type}")
        return {"differences": 0, "not_found": 0}
    
    logger.debug(f"Found ID columns for {report_type} - SC: {sc_id_col}, D365: {d365_id_col}")
    
    # Find status columns
    # D365 always has "Status Reason" column - use centralized helper function
    d365_status_col = find_column_by_keywords(df_d365.columns, ("status", "reason"))
    
    # SC status column varies by report type - use centralized helper function
    sc_status_col = find_sc_status_column(df_sc, sc_id_col, report_type)
    
    if sc_status_col is None or d365_status_col is None:
        logger.warning(f"Could not find status columns for {report_type}: SC={sc_status_col}, D365={d365_status_col}")
        return {"differences": 0, "not_found": 0}
    
    logger.debug(f"Found status columns for {report_type} - SC: {sc_status_col}, D365: {d365_status_col}")
    
    # Replicate XLOOKUP: Merge SC with D365 to get D365 status for each SC record
    df_sc_copy = df_sc.copy()
    df_d365_copy = df_d365.copy()
    
    # Clean IDs for matching
    df_sc_copy['clean_id'] = df_sc_copy[sc_id_col].astype(str).str.strip().str.lower()
    df_d365_copy['clean_id'] = df_d365_copy[d365_id_col].astype(str).str.strip().str.lower()
    
    # Remove duplicates from D365, keeping first match (XLOOKUP behavior)
    df_d365_dedup = df_d365_copy.drop_duplicates(subset=['clean_id'], keep='first')
    
    # Merge to get D365 status (this replicates the XLOOKUP formula)
    # Use suffixes to avoid column name conflicts when both dataframes have same column names
    merged = df_sc_copy.merge(
        df_d365_dedup[['clean_id', d365_status_col]],
        on='clean_id',
        how='left',
        suffixes=('_sc', '_d365')
    )
    
    # Determine the correct column name after merge (may have suffix if column existed in both)
    d365_status_col_merged = f"{d365_status_col}_d365" if f"{d365_status_col}_d365" in merged.columns else d365_status_col
    sc_status_col_merged = f"{sc_status_col}_sc" if f"{sc_status_col}_sc" in merged.columns else sc_status_col
    
    # Count "Not found" (SC records with no matching D365 record)
    not_found = merged[d365_status_col_merged].isna().sum()
    
    # Count differences (replicate "Is it the same?" formula: =sc_status=d365_status)
    valid_rows = merged[d365_status_col_merged].notna()
    if valid_rows.any():
        sc_statuses = merged.loc[valid_rows, sc_status_col_merged].fillna("").astype(str)
        d365_statuses = merged.loc[valid_rows, d365_status_col_merged].fillna("").astype(str)
        differences = (sc_statuses != d365_statuses).sum()
    else:
        differences = 0
    
    logger.info(f"SC sheet analysis for {report_type} complete - Differences: {differences}, Not found: {not_found}")
    
    return {
        "differences": int(differences),
        "not_found": int(not_found)
    }


def analyze_d365_sheet(df_d365, df_sc, report_type="client"):
    """
    Analyze D365 sheet by replicating the Excel XLOOKUP formula logic.
    
    Excel Formula Logic:
    - SC Status column: =XLOOKUP(d365_id, SC!id, SC!status, "Not found")
    
    We replicate this by merging D365 and SC data, then filter for "Not found".
    
    Args:
        df_d365: DataFrame from D365 sheet
        df_sc: DataFrame from SC sheet (needed for XLOOKUP replication)
        report_type: Type of report (client, wcb, accreditation, or critical_document)
        
    Returns:
        dict: Statistics about not found records and status breakdown
    """
    logger.info(f"Analyzing D365 sheet for {report_type}")
    if df_d365 is None or df_d365.empty or df_sc is None or df_sc.empty:
        logger.warning(f"Empty or None dataframes for {report_type} D365 analysis")
        return {"total_not_found": 0, "status_breakdown": {}}
    
    # Find ID columns using centralized helper function
    d365_id_col = find_column_by_keywords(df_d365.columns, ("global", "alcumus", "id"))
    sc_id_col = find_column_by_keywords(df_sc.columns, ("global", "alcumus", "id"))
    
    if d365_id_col is None or sc_id_col is None:
        logger.warning(f"Could not find ID columns for D365 analysis")
        return {"total_not_found": 0, "status_breakdown": {}}
    
    logger.debug(f"Found ID columns for D365 analysis - D365: {d365_id_col}, SC: {sc_id_col}")
    
    # Find Status Reason column using centralized helper function
    status_reason_col = find_column_by_keywords(df_d365.columns, ("status", "reason"))
    
    if status_reason_col is None:
        logger.warning("Could not find 'Status Reason' column in D365 sheet")
        return {"total_not_found": 0, "status_breakdown": {}}
    
    logger.debug(f"Found Status Reason column: {status_reason_col}")
    
    # Replicate XLOOKUP: Find D365 records not in SC
    df_d365_copy = df_d365.copy()
    df_sc_copy = df_sc.copy()
    
    # Clean IDs for matching
    df_d365_copy['clean_id'] = df_d365_copy[d365_id_col].astype(str).str.strip().str.lower()
    df_sc_copy['clean_id'] = df_sc_copy[sc_id_col].astype(str).str.strip().str.lower()
    
    # Remove duplicates from SC, keeping first match (XLOOKUP behavior)
    df_sc_dedup = df_sc_copy.drop_duplicates(subset=['clean_id'], keep='first')
    
    # Find D365 records not in SC (this replicates XLOOKUP returning "Not found")
    sc_ids = set(df_sc_dedup['clean_id'].dropna())
    not_found_df = df_d365_copy[~df_d365_copy['clean_id'].isin(sc_ids)]
    
    total_not_found = len(not_found_df)
    
    # Group by Status Reason
    status_breakdown = {}
    if not not_found_df.empty:
        status_counts = not_found_df[status_reason_col].value_counts()
        status_breakdown = status_counts.to_dict()
        logger.debug(f"Status breakdown for {report_type}: {status_breakdown}")
    
    logger.info(f"D365 sheet analysis for {report_type} complete - Total not found: {total_not_found}, Status types: {len(status_breakdown)}")
    
    return {
        "total_not_found": int(total_not_found),
        "status_breakdown": status_breakdown
    }


def format_status_name(status):
    """
    Format status name for display in email.
    Adds "Statuses" suffix if not present.
    
    Args:
        status: Raw status string
        
    Returns:
        str: Formatted status name
    """
    if pd.isna(status):
        return "Unknown Status"
    
    status_str = str(status).strip()
    
    # Check if it already ends with "Status" or "Statuses"
    if not (status_str.endswith("Status") or status_str.endswith("Statuses")):
        status_str += " Statuses"
    elif status_str.endswith("Status") and not status_str.endswith("Statuses"):
        status_str += "es"
    
    return status_str


def generate_email_report():
    """
    Generate the complete email report by analyzing all comparison files.
    
    Replicates Excel XLOOKUP formula logic by merging data from both sheets.
    """
    logger.info("="*70)
    logger.info("Starting email report generation")
    logger.info(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("="*70)
    print("\n" + "=" * 70)
    print("EMAIL REPORT GENERATOR")
    print("=" * 70)
    
    # Get the dated comparison directory
    comparison_dir = get_dated_comparison_dir()
    
    # Define comparison types and their file paths
    comparisons = {
        "Client": comparison_dir / "Client_Comparison.xlsx",
        "WCB": comparison_dir / "WCB_Comparison.xlsx",
        "Accreditation": comparison_dir / "Accreditation_Comparison.xlsx",
        "Critical_Document": comparison_dir / "Critical_Document_Comparison.xlsx"
    }
    
    # Check which files exist
    logger.info("Checking for comparison files...")
    available_comparisons = {}
    missing_files = []
    
    for name, path in comparisons.items():
        if path.exists():
            available_comparisons[name] = path
            logger.info(f"Found {name} comparison file: {path}")
            print(f"[OK] Found {name} comparison file")
        else:
            missing_files.append(name)
            logger.warning(f"Missing {name} comparison file: {path}")
            print(f"[X] Missing {name} comparison file")
    
    if not available_comparisons:
        logger.error("No comparison files found in output directory")
        print("\n[WARNING] No comparison files found!")
        print("Please run the comparison tool first to generate Excel files.")
        return None
    
    logger.info(f"Found {len(available_comparisons)} comparison file(s)")
    if missing_files:
        logger.warning(f"{len(missing_files)} file(s) missing: {', '.join(missing_files)}")
        print(f"\n[NOTE] {len(missing_files)} file(s) missing: {', '.join(missing_files)}")
    
    print("\n" + "-" * 70)
    print("ANALYZING COMPARISON FILES...")
    print("-" * 70 + "\n")
    
    logger.info("Starting analysis of comparison files...")
    
    # Analyze each comparison
    results = {}
    for name, file_path in available_comparisons.items():
        logger.info(f"Processing {name} comparison...")
        print(f"Analyzing {name}...")
        
        # Read the file
        sc_df, d365_df = read_comparison_file(file_path)
        
        if sc_df is None or d365_df is None:
            logger.error(f"Failed to read {name} file, skipping analysis")
            print(f"  [X] Error reading {name} file\n")
            continue
        
        # Analyze both sheets (pass report type for correct column detection)
        logger.info(f"Analyzing SC sheet for {name}...")
        sc_stats = analyze_sc_sheet(sc_df, d365_df, report_type=name)
        logger.info(f"Analyzing D365 sheet for {name}...")
        d365_stats = analyze_d365_sheet(d365_df, sc_df, report_type=name)
        
        results[name] = {
            "sc": sc_stats,
            "d365": d365_stats
        }
        
        logger.info(f"Completed {name} analysis - SC differences: {sc_stats['differences']}, SC not found: {sc_stats['not_found']}, D365 not found: {d365_stats['total_not_found']}, Status types: {len(d365_stats['status_breakdown'])}")
        print(f"  [OK] SC differences: {sc_stats['differences']}")
        print(f"  [OK] D365 not found: {d365_stats['total_not_found']}")
        print(f"  [OK] Status types: {len(d365_stats['status_breakdown'])}\n")
    
    # Generate email text
    logger.info("Generating email report text...")
    print("-" * 70)
    print("GENERATED EMAIL REPORT")
    print("-" * 70 + "\n")
    
    email_lines = []
    
    # Process in the order: Client, WCB, Accreditation, Critical_Document
    order = ["Client", "WCB", "Accreditation", "Critical_Document"]
    
    for name in order:
        if name not in results:
            continue
        
        data = results[name]
        
        # Section header with display name
        if name == "Client":
            display_name = "Client Specific"
        elif name == "Critical_Document":
            display_name = "Critical Document"
        else:
            display_name = name
        
        # Add blank line between sections (not before the first one)
        if email_lines:
            email_lines.append("")
        
        email_lines.append(f"{display_name}:")
        
        # SC statistics
        sc_diff = data["sc"]["differences"]
        sc_not_found = data["sc"]["not_found"]
        
        email_lines.append("• SC:")
        if sc_not_found > 0:
            email_lines.append(f"\t○ {sc_diff} differences between dynamics and SafeContractor, {sc_not_found} Not found")
        else:
            email_lines.append(f"\t○ {sc_diff} differences between dynamics and SafeContractor")
        
        # D365 statistics
        email_lines.append("")
        email_lines.append("• D365:")
        
        total_not_found = data["d365"]["total_not_found"]
        email_lines.append(f"\t○ {total_not_found} not found in SafeContractor:")
        
        # Sort status breakdown alphabetically for consistency
        status_breakdown = data["d365"]["status_breakdown"]
        if status_breakdown:
            sorted_statuses = sorted(status_breakdown.items(), key=lambda x: x[0])
            for status, count in sorted_statuses:
                formatted_status = format_status_name(status)
                email_lines.append(f"{count} {formatted_status}")
    
    # Join all lines
    email_text = "\n".join(email_lines)
    logger.info(f"Email report generated successfully ({len(email_lines)} lines)")
    
    # Print to console
    print(email_text)
    
    # Save to file
    output_file = OUTPUT_DIR / "email_report.txt"
    logger.info(f"Saving email report to: {output_file}")
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(email_text)
        print("\n" + "-" * 70)
        print(f"[OK] Email report saved to: {output_file}")
        print("-" * 70 + "\n")
        logger.info(f"Email report saved successfully to {output_file}")
        logger.info("Email report generation completed successfully")
    except Exception as e:
        print(f"\n[WARNING] Could not save to file: {e}")
        logger.error(f"Failed to save email report: {e}")
    
    return email_text


def main():
    """Main execution"""
    try:
        logger.info("Email Report Generator started")
        generate_email_report()
        logger.info("Email Report Generator finished successfully")
    except Exception as e:
        logger.exception(f"Unexpected error in email report generation: {e}")
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
