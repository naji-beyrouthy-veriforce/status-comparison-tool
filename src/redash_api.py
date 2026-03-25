"""
Redash API Integration Module
Automates query execution and result downloading from Redash.

Handles five query types:
- Accreditation (1460): Injects extracted IDs into SQL, executes directly, downloads (includes created_at, updated_at)
- WCB (1281): Injects extracted IDs into SQL, executes directly, downloads
- Client (1277): Executes as-is (no modification), downloads
- Critical Document (1464): Executes as-is (no modification), downloads
- ESG (1465): Executes as-is (no modification), downloads

Approach: Executes raw SQL via /api/query_results with data_source_id.
Never modifies saved Redash queries — read-only API key is sufficient.
"""

import re
import time
import requests
import pandas as pd
from io import StringIO

from .config import (
    REDASH_BASE_URL,
    REDASH_API_KEY,
    REDASH_QUERY_IDS,
    REDASH_POLL_INTERVAL,
    REDASH_POLL_TIMEOUT,
    REDASH_DIR,
    QUERY_IDS_DIR,
    SC_FILES,
    Messages,
    setup_logging,
)

logger = setup_logging("redash_api", console_output=False, file_output=True)


# ============================================================================
# API KEY & HEADERS
# ============================================================================

def get_api_key():
    """
    Get Redash API key from config (sourced from REDASH_API_KEY env var).
    Raises ValueError with setup instructions if not configured.
    """
    if not REDASH_API_KEY:
        raise ValueError(
            "REDASH_API_KEY environment variable not set.\n"
            "  Set it in PowerShell:  $env:REDASH_API_KEY = 'your-key-here'\n"
            "  Or set it permanently in Windows System Environment Variables.\n"
            "  The batch files (Run_CLI.bat, Run_GUI.bat) can also set it automatically."
        )
    return REDASH_API_KEY


def _get_headers():
    """Build authorization headers for Redash API requests."""
    return {"Authorization": f"Key {get_api_key()}"}


# ============================================================================
# CORE API FUNCTIONS
# ============================================================================

def verify_connection():
    """
    Verify Redash API connection and authentication.
    Returns True if connection succeeds, False otherwise.
    """
    try:
        response = requests.get(
            f"{REDASH_BASE_URL}/api/queries/{REDASH_QUERY_IDS['accreditation']}",
            headers=_get_headers(),
            timeout=15,
        )
        response.raise_for_status()
        return True
    except requests.exceptions.ConnectionError:
        print(f"  {Messages.ERROR} Cannot reach Redash server at {REDASH_BASE_URL}")
        return False
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 401:
            print(f"  {Messages.ERROR} Invalid Redash API key")
        else:
            print(f"  {Messages.ERROR} Redash API error: {e.response.status_code}")
        return False
    except Exception as e:
        print(f"  {Messages.ERROR} Connection error: {e}")
        return False


def get_query(query_id):
    """
    Fetch query details from Redash.

    Returns:
        dict with 'query' (SQL text), 'data_source_id', 'name', etc.
    """
    response = requests.get(
        f"{REDASH_BASE_URL}/api/queries/{query_id}",
        headers=_get_headers(),
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def execute_raw_sql(data_source_id, sql_text):
    """
    Execute raw SQL directly via Redash API without modifying any saved query.

    Posts SQL + data_source_id to /api/query_results. This is read-only safe:
    no saved queries are touched.

    Args:
        data_source_id: Redash data source ID (from saved query metadata)
        sql_text: Full SQL query to execute

    Returns:
        query_result_id on success
    """
    headers = _get_headers()

    response = requests.post(
        f"{REDASH_BASE_URL}/api/query_results",
        headers=headers,
        json={
            "data_source_id": data_source_id,
            "query": sql_text,
            "max_age": 0,
        },
        timeout=60,
    )
    response.raise_for_status()
    result = response.json()

    # Result may be immediately available
    if "query_result" in result:
        return result["query_result"]["id"]

    # Otherwise poll the job until completion
    job = result.get("job", {})
    job_id = job.get("id")
    if not job_id:
        raise RuntimeError("No job ID returned from Redash query execution")

    return _poll_job(job_id)


def _poll_job(job_id):
    """
    Poll a Redash job until it completes or times out.

    Returns:
        query_result_id on success

    Raises:
        RuntimeError on query failure
        TimeoutError if polling exceeds REDASH_POLL_TIMEOUT
    """
    headers = _get_headers()
    start_time = time.time()

    while time.time() - start_time < REDASH_POLL_TIMEOUT:
        response = requests.get(
            f"{REDASH_BASE_URL}/api/jobs/{job_id}",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        job = response.json().get("job", {})

        status = job.get("status")
        if status == 3:  # Success
            return job.get("query_result_id")
        elif status == 4:  # Error
            error = job.get("error", "Unknown query execution error")
            raise RuntimeError(f"Redash query failed: {error}")

        # Status 1=pending, 2=started — keep polling
        elapsed = int(time.time() - start_time)
        print(f"  ... waiting for query results ({elapsed}s elapsed)")
        time.sleep(REDASH_POLL_INTERVAL)

    raise TimeoutError(
        f"Redash query timed out after {REDASH_POLL_TIMEOUT}s. "
        f"The query may still be running — check Redash directly."
    )


def download_result_by_id(query_result_id):
    """
    Download query results by result ID as a DataFrame.

    Args:
        query_result_id: Redash query result ID returned from execution

    Returns:
        pandas DataFrame with query results
    """
    response = requests.get(
        f"{REDASH_BASE_URL}/api/query_results/{query_result_id}.csv",
        headers=_get_headers(),
        timeout=120,
    )
    response.raise_for_status()

    df = pd.read_csv(StringIO(response.text))
    if df.empty:
        raise ValueError("Redash query returned no results")

    return df


# ============================================================================
# SQL ID INJECTION
# ============================================================================

def inject_ids_into_sql(sql_text, ids_formatted):
    """
    Replace IDs inside the global_alcumus_id IN (...) clause.

    Handles both formats:
    - Accreditation: global_alcumus_id in (...)
    - WCB: wdc.global_alcumus_id in (...)

    Args:
        sql_text: Original SQL query text
        ids_formatted: SQL-formatted ID string from format_ids_for_sql()
                       e.g. "'id1',\\n'id2',\\n'id3'"

    Returns:
        Updated SQL text with new IDs
    """
    # Pattern matches: [optional_table.]global_alcumus_id in (...)
    pattern = r"((?:\w+\.)?global_alcumus_id\s+in\s*\()([^)]*?)(\))"

    match = re.search(pattern, sql_text, re.IGNORECASE | re.DOTALL)
    if not match:
        raise ValueError(
            "Could not find 'global_alcumus_id in (...)' clause in query SQL. "
            "The query format may have changed."
        )

    new_sql = re.sub(
        pattern,
        lambda m: f"{m.group(1)}\n{ids_formatted}{m.group(3)}",
        sql_text,
        count=1,
        flags=re.IGNORECASE | re.DOTALL,
    )

    return new_sql


def read_ids_from_file(report_type):
    """
    Read SQL-formatted IDs from the extracted .sql.txt file.

    Args:
        report_type: 'accreditation' or 'wcb'

    Returns:
        Formatted ID string ready for SQL injection, or None if file not found
    """
    ids_file = QUERY_IDS_DIR / f"{report_type}_ids.sql.txt"
    if not ids_file.exists():
        return None
    content = ids_file.read_text(encoding="utf-8").strip()
    if not content:
        return None
    return content


# ============================================================================
# ORCHESTRATION — SINGLE QUERY
# ============================================================================

def run_redash_query(query_id, report_type, ids_formatted=None):
    """
    Full flow for a single Redash query:
    1. Fetch saved query to get SQL template and data_source_id
    2. Optionally inject new IDs into SQL (accreditation/wcb)
    3. Execute raw SQL directly (never modifies saved query)
    4. Download results
    5. Save as Excel to input/redash/

    Args:
        query_id: Redash query ID
        report_type: 'accreditation', 'wcb', 'client', 'critical_document', or 'esg'
        ids_formatted: SQL-formatted IDs string, or None for client

    Returns:
        Path to saved Excel file, or None on failure
    """
    try:
        # 1. Fetch saved query for SQL template and data_source_id
        print(f"  📡 Fetching query {query_id} from Redash...")
        query_data = get_query(query_id)
        sql_text = query_data["query"]
        data_source_id = query_data["data_source_id"]
        logger.info(f"Fetched query {query_id} ({report_type}): {len(sql_text)} chars, ds={data_source_id}")

        # 2. Inject IDs if provided (accreditation/wcb only)
        if ids_formatted:
            id_count = ids_formatted.count("'") // 2  # rough count
            print(f"  🔄 Injecting {id_count} extracted IDs into SQL...")
            sql_text = inject_ids_into_sql(sql_text, ids_formatted)
            print(f"  {Messages.SUCCESS} SQL prepared with {id_count} IDs ({len(sql_text):,} chars)")
            logger.info(f"Prepared SQL for {report_type} with {id_count} IDs ({len(sql_text)} chars)")

        # 3. Execute raw SQL directly (no saved query modification)
        print(f"  ⏳ Executing query (this may take a moment)...")
        query_result_id = execute_raw_sql(data_source_id, sql_text)
        print(f"  {Messages.SUCCESS} Query execution completed")

        # 4. Download results
        print(f"  📥 Downloading results...")
        df = download_result_by_id(query_result_id)
        print(f"  {Messages.SUCCESS} Downloaded {len(df)} rows")
        logger.info(f"Downloaded {len(df)} rows for {report_type}")

        # 5. Save as Excel
        REDASH_DIR.mkdir(parents=True, exist_ok=True)
        output_path = REDASH_DIR / SC_FILES[report_type]
        df.to_excel(output_path, index=False)
        print(f"  {Messages.SUCCESS} Saved to {output_path.name}")
        logger.info(f"Saved results to {output_path}")

        return output_path

    except Exception as e:
        print(f"  {Messages.ERROR} Error: {e}")
        logger.error(f"Redash query failed for {report_type} (query {query_id}): {e}")
        raise


# ============================================================================
# ORCHESTRATION — ALL QUERIES
# ============================================================================

def run_all_redash_queries():
    """
    Execute all 5 Redash queries and download results.

    - Accreditation (1460): Injects extracted IDs → execute → download (includes created_at, updated_at)
    - WCB (1281): Injects extracted IDs → execute → download
    - Client (1277): Execute as-is → download (NO modification)
    - Critical Document (1464): Execute as-is → download (NO modification)
    - ESG (1465): Execute as-is → download (NO modification)

    Returns:
        dict: {report_type: output_path} for successful downloads
    """
    print("\n" + "=" * 70)
    print("REDASH: EXECUTING QUERIES & DOWNLOADING RESULTS")
    print("=" * 70)

    # Verify connection first
    print("\n📡 Verifying Redash API connection...")
    if not verify_connection():
        raise ConnectionError(
            "Cannot connect to Redash API. Check your network and API key."
        )
    print(f"{Messages.SUCCESS} Connected to Redash\n")

    results = {}

    # --- Accreditation & WCB: inject IDs, execute, download ---
    for report_type in ["accreditation", "wcb"]:
        print(f"\n{Messages.PROCESSING} Processing {report_type.upper()}...")

        ids_formatted = read_ids_from_file(report_type)
        if not ids_formatted:
            print(f"  {Messages.WARNING} No extracted IDs found for {report_type}, skipping...")
            print(f"  Make sure extract_and_save_ids() ran successfully first.")
            continue

        query_id = REDASH_QUERY_IDS[report_type]

        try:
            output_path = run_redash_query(query_id, report_type, ids_formatted)
            if output_path:
                results[report_type] = output_path
        except Exception as e:
            print(f"  {Messages.ERROR} Failed to process {report_type}: {e}")
            logger.error(f"Failed {report_type} Redash query: {e}")
            continue

    # --- Client: execute as-is, download (NO modification) ---
    print(f"\n{Messages.PROCESSING} Processing CLIENT...")

    try:
        query_id = REDASH_QUERY_IDS["client"]
        output_path = run_redash_query(query_id, "client")
        if output_path:
            results["client"] = output_path
    except Exception as e:
        print(f"  {Messages.ERROR} Failed to process client: {e}")
        logger.error(f"Failed client Redash query: {e}")

    # --- Critical Document: execute as-is, download (NO modification) ---
    print(f"\n{Messages.PROCESSING} Processing CRITICAL DOCUMENT...")

    try:
        query_id = REDASH_QUERY_IDS["critical_document"]
        output_path = run_redash_query(query_id, "critical_document")
        if output_path:
            results["critical_document"] = output_path
    except Exception as e:
        print(f"  {Messages.ERROR} Failed to process critical_document: {e}")
        logger.error(f"Failed critical_document Redash query: {e}")

    # --- ESG: execute as-is, download (NO modification) ---
    print(f"\n{Messages.PROCESSING} Processing ESG...")

    try:
        query_id = REDASH_QUERY_IDS["esg"]
        output_path = run_redash_query(query_id, "esg")
        if output_path:
            results["esg"] = output_path
    except Exception as e:
        print(f"  {Messages.ERROR} Failed to process esg: {e}")
        logger.error(f"Failed esg Redash query: {e}")

    # --- Summary ---
    print("\n" + "-" * 40)
    print(f"Downloaded {len(results)}/5 Redash results:")
    for rt, path in results.items():
        print(f"  {Messages.SUCCESS} {rt.title()}: {path.name}")

    missing = {"accreditation", "wcb", "client", "critical_document", "esg"} - set(results.keys())
    for rt in sorted(missing):
        print(f"  {Messages.ERROR} {rt.title()}: Failed")

    return results


# ============================================================================
# STANDALONE ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    """Run Redash queries standalone (assumes IDs already extracted)."""
    try:
        results = run_all_redash_queries()
        if results:
            print(f"\n✅ Successfully downloaded {len(results)} result(s)")
        else:
            print(f"\n❌ No results were downloaded")
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
