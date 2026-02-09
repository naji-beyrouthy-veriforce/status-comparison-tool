"""
Email Report Generator
Automates the generation of comparison email reports from Excel output files

This script reads the comparison Excel files and generates a formatted email
report with statistics about differences and missing records.
"""

import pandas as pd
from pathlib import Path
import sys
from datetime import datetime
from collections import defaultdict

# Import configuration
from config import OUTPUT_DIR, setup_logging

# Setup logging
logger = setup_logging("email_report", console_output=True, file_output=True)


def read_comparison_file(file_path):
    """
    Read both sheets from a comparison Excel file.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        tuple: (sc_df, d365_df) DataFrames or (None, None) if error
    """
    try:
        # Read both sheets
        sc_df = pd.read_excel(file_path, sheet_name="SC")
        d365_df = pd.read_excel(file_path, sheet_name="D365")
        return sc_df, d365_df
    except Exception as e:
        logger.error(f"Error reading {file_path.name}: {e}")
        return None, None


def analyze_sc_sheet(df_sc, df_d365, report_type="client"):
    """
    Analyze SC sheet for differences by merging with D365 data.
    
    Args:
        df_sc: DataFrame from SC sheet
        df_d365: DataFrame from D365 sheet (needed for lookup)
        report_type: Type of report (client, wcb, or accreditation)
        
    Returns:
        dict: Statistics about differences
    """
    if df_sc is None or df_sc.empty or df_d365 is None or df_d365.empty:
        return {"differences": 0, "not_found": 0}
    
    # Find ID columns
    sc_id_col = None
    for col in df_sc.columns:
        if "global" in str(col).lower() and "id" in str(col).lower():
            sc_id_col = col
            break
    
    d365_id_col = None
    for col in df_d365.columns:
        if "global" in str(col).lower() and "id" in str(col).lower():
            d365_id_col = col
            break
    
    if sc_id_col is None or d365_id_col is None:
        logger.warning(f"Could not find ID columns for {report_type}")
        return {"differences": 0, "not_found": 0}
    
    # Find status columns
    # SC status column varies by report type
    sc_status_col = None
    if report_type.lower() == "client":
        for col in df_sc.columns:
            if col.lower() == "case":
                sc_status_col = col
                break
    else:
        for col in df_sc.columns:
            if "status" in str(col).lower() and "d365" not in str(col).lower() and "contractor" not in str(col).lower() and "client" not in str(col).lower():
                sc_status_col = col
                break
    
    # D365 status column
    d365_status_col = None
    for col in df_d365.columns:
        if "status" in str(col).lower() and "reason" in str(col).lower():
            d365_status_col = col
            break
    
    if sc_status_col is None or d365_status_col is None:
        logger.warning(f"Could not find status columns for {report_type}: SC={sc_status_col}, D365={d365_status_col}")
        return {"differences": 0, "not_found": 0}
    
    # Clean IDs for matching
    df_sc = df_sc.copy()
    df_d365 = df_d365.copy()
    
    df_sc['clean_id'] = df_sc[sc_id_col].astype(str).str.strip().str.lower()
    df_d365['clean_id'] = df_d365[d365_id_col].astype(str).str.strip().str.lower()
    
    # Merge to get D365 status for each SC record
    merged = df_sc.merge(
        df_d365[['clean_id', d365_status_col]],
        on='clean_id',
        how='left',
        suffixes=('_sc', '_d365')
    )
    
    # Count not found (SC records with no matching D365 record)
    not_found = merged[d365_status_col].isna().sum()
    
    # Count differences (where both exist but statuses don't match)
    valid_rows = merged[d365_status_col].notna()
    if valid_rows.any():
        sc_statuses = merged.loc[valid_rows, sc_status_col].fillna("").astype(str).str.strip().str.lower()
        d365_statuses = merged.loc[valid_rows, d365_status_col].fillna("").astype(str).str.strip().str.lower()
        differences = (sc_statuses != d365_statuses).sum()
    else:
        differences = 0
    
    return {
        "differences": int(differences),
        "not_found": int(not_found)
    }


def analyze_d365_sheet(df_d365, df_sc, report_type="client"):
    """
    Analyze D365 sheet for records not found in SafeContractor by merging data.
    Groups by Status Reason for "Not found" entries.
    
    Args:
        df_d365: DataFrame from D365 sheet
        df_sc: DataFrame from SC sheet (needed for lookup)
        report_type: Type of report (client, wcb, or accreditation)
        
    Returns:
        dict: Statistics about not found records and status breakdown
    """
    if df_d365 is None or df_d365.empty or df_sc is None or df_sc.empty:
        return {"total_not_found": 0, "status_breakdown": {}}
    
    # Find ID columns
    d365_id_col = None
    for col in df_d365.columns:
        if "global" in str(col).lower() and "id" in str(col).lower():
            d365_id_col = col
            break
    
    sc_id_col = None
    for col in df_sc.columns:
        if "global" in str(col).lower() and "id" in str(col).lower():
            sc_id_col = col
            break
    
    if d365_id_col is None or sc_id_col is None:
        logger.warning(f"Could not find ID columns for D365 analysis")
        return {"total_not_found": 0, "status_breakdown": {}}
    
    # Find Status Reason column in D365
    status_reason_col = None
    for col in df_d365.columns:
        if "status" in str(col).lower() and "reason" in str(col).lower():
            status_reason_col = col
            break
    
    if status_reason_col is None:
        logger.warning("Could not find 'Status Reason' column in D365 sheet")
        return {"total_not_found": 0, "status_breakdown": {}}
    
    # Clean IDs for matching
    df_d365 = df_d365.copy()
    df_sc = df_sc.copy()
    
    df_d365['clean_id'] = df_d365[d365_id_col].astype(str).str.strip().str.lower()
    df_sc['clean_id'] = df_sc[sc_id_col].astype(str).str.strip().str.lower()
    
    # Find D365 records not in SC (left anti-join)
    sc_ids = set(df_sc['clean_id'].dropna())
    not_found_df = df_d365[~df_d365['clean_id'].isin(sc_ids)]
    
    total_not_found = len(not_found_df)
    
    # Group by Status Reason
    status_breakdown = {}
    if not not_found_df.empty:
        status_counts = not_found_df[status_reason_col].value_counts()
        status_breakdown = status_counts.to_dict()
    
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
    """
    logger.info("Starting email report generation")
    print("\n" + "=" * 70)
    print("EMAIL REPORT GENERATOR")
    print("=" * 70)
    
    # Define comparison types and their file paths
    comparisons = {
        "Client": OUTPUT_DIR / "Client_Comparison.xlsx",
        "WCB": OUTPUT_DIR / "WCB_Comparison.xlsx",
        "Accreditation": OUTPUT_DIR / "Accreditation_Comparison.xlsx"
    }
    
    # Check which files exist
    available_comparisons = {}
    missing_files = []
    
    for name, path in comparisons.items():
        if path.exists():
            available_comparisons[name] = path
            print(f"[OK] Found {name} comparison file")
        else:
            missing_files.append(name)
            print(f"[X] Missing {name} comparison file")
    
    if not available_comparisons:
        print("\n[WARNING] No comparison files found!")
        print("Please run the comparison tool first to generate Excel files.")
        return None
    
    if missing_files:
        print(f"\n[NOTE] {len(missing_files)} file(s) missing: {', '.join(missing_files)}")
    
    print("\n" + "-" * 70)
    print("ANALYZING COMPARISON FILES...")
    print("-" * 70 + "\n")
    
    # Analyze each comparison
    results = {}
    for name, file_path in available_comparisons.items():
        print(f"Analyzing {name}...")
        
        # Read the file
        sc_df, d365_df = read_comparison_file(file_path)
        
        if sc_df is None or d365_df is None:
            print(f"  [X] Error reading {name} file\n")
            continue
        
        # Analyze both sheets (pass report type for correct column detection)
        sc_stats = analyze_sc_sheet(sc_df, d365_df, report_type=name)
        d365_stats = analyze_d365_sheet(d365_df, sc_df, report_type=name)
        
        results[name] = {
            "sc": sc_stats,
            "d365": d365_stats
        }
        
        print(f"  [OK] SC differences: {sc_stats['differences']}")
        print(f"  [OK] D365 not found: {d365_stats['total_not_found']}")
        print(f"  [OK] Status types: {len(d365_stats['status_breakdown'])}\n")
    
    # Generate email text
    print("-" * 70)
    print("GENERATED EMAIL REPORT")
    print("-" * 70 + "\n")
    
    email_lines = []
    
    # Process in the order: Client, WCB, Accreditation
    order = ["Client", "WCB", "Accreditation"]
    
    for name in order:
        if name not in results:
            continue
        
        data = results[name]
        
        # Section header
        if name == "Client":
            email_lines.append("Client specific:\n")
            email_lines.append("SC:")
        else:
            email_lines.append(f"\n{name}:\n")
            email_lines.append("SC:")
        
        # SC statistics
        sc_diff = data["sc"]["differences"]
        sc_not_found = data["sc"]["not_found"]
        
        if sc_not_found > 0:
            email_lines.append(f"{sc_diff} differences between dynamics and SafeContractor, {sc_not_found} Not found")
        else:
            email_lines.append(f"{sc_diff} differences between dynamics and SafeContractor")
        
        # D365 statistics
        email_lines.append("\nD365:")
        
        total_not_found = data["d365"]["total_not_found"]
        email_lines.append(f"{total_not_found} not found in SafeContractor:")
        
        # Sort status breakdown alphabetically for consistency
        status_breakdown = data["d365"]["status_breakdown"]
        if status_breakdown:
            sorted_statuses = sorted(status_breakdown.items(), key=lambda x: x[0])
            for status, count in sorted_statuses:
                formatted_status = format_status_name(status)
                email_lines.append(f"{count} {formatted_status}")
    
    # Join all lines
    email_text = "\n".join(email_lines)
    
    # Print to console
    print(email_text)
    
    # Save to file
    output_file = OUTPUT_DIR / "email_report.txt"
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(email_text)
        print("\n" + "-" * 70)
        print(f"[OK] Email report saved to: {output_file}")
        print("-" * 70 + "\n")
        logger.info(f"Email report saved successfully to {output_file}")
    except Exception as e:
        print(f"\n[WARNING] Could not save to file: {e}")
        logger.error(f"Failed to save email report: {e}")
    
    return email_text


def main():
    """Main execution"""
    try:
        generate_email_report()
    except Exception as e:
        logger.exception(f"Unexpected error in email report generation: {e}")
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
