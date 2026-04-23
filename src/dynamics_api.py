"""
Dynamics 365 Web API Integration Module
Downloads D365 incident exports for all 5 report types using saved views.

Authentication: OAuth2 client credentials flow (Azure App Registration)
Data retrieval:  Saved views via D365 OData Web API v9.2

Required environment variables (set in Run_CLI.bat / Run_GUI.bat):
    D365_TENANT_ID       Azure Active Directory tenant ID
    D365_CLIENT_ID       Azure App Registration client (application) ID
    D365_CLIENT_SECRET   Azure App Registration client secret

Required view IDs (set in config.py D365_VIEW_IDS or as env vars):
    D365_VIEW_ID_ACCREDITATION
    D365_VIEW_ID_WCB
    D365_VIEW_ID_CLIENT
    D365_VIEW_ID_CRITICAL_DOCUMENT
    D365_VIEW_ID_ESG

    To find a view ID: open the view in D365 → copy the `viewid=` value from the URL.

Approach:
    1. Authenticate via OAuth2 client credentials → get Bearer token
    2. Fetch the saved view's layoutxml to discover its column schema names
    3. Fetch D365 entity attribute metadata to map schema names → display names
    4. Query the view via ?savedQuery={view_id} with pagination
    5. Prefer formatted annotation values (e.g. "Status Reason" text, not integer codes)
    6. Save each report as Excel to input/dynamics/

Note: This module never modifies any D365 data. It is strictly read-only.
"""

import time
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path

from .config import (
    D365_ORG_URL,
    D365_TENANT_ID,
    D365_CLIENT_ID,
    D365_CLIENT_SECRET,
    D365_VIEW_IDS,
    D365_ENTITY,
    D365_ENTITY_LOGICAL_NAME,
    D365_API_VERSION,
    D365_PAGE_SIZE,
    D365_KNOWN_FIELD_NAMES,
    DYNAMICS_DIR,
    D365_FILES,
    REPORT_TYPE_DISPLAY_NAMES,
    Messages,
    setup_logging,
)

logger = setup_logging("dynamics_api", console_output=False, file_output=True)

# Token endpoint template — tenant_id is filled in at call time
_TOKEN_URL_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

# D365 Web API base URL
_API_BASE = f"{D365_ORG_URL}/api/data/{D365_API_VERSION}"

# OData annotation key for human-readable (formatted) field values
_FORMATTED_VALUE = "@OData.Community.Display.V1.FormattedValue"


# ============================================================================
# AUTHENTICATION
# ============================================================================

def get_access_token():
    """
    Obtain an OAuth2 Bearer token via the client credentials flow.

    Raises:
        ValueError: If any of the three required credentials are missing.
        requests.HTTPError: If the Azure AD token endpoint returns an error.

    Returns:
        str: Bearer access token valid for ~60 minutes.
    """
    if not all([D365_TENANT_ID, D365_CLIENT_ID, D365_CLIENT_SECRET]):
        raise ValueError(
            "D365 credentials not fully configured.\n"
            "  Set the following environment variables (or fill them in the .bat launchers):\n"
            "    D365_TENANT_ID     — Azure AD tenant ID\n"
            "    D365_CLIENT_ID     — App registration client ID\n"
            "    D365_CLIENT_SECRET — App registration client secret\n"
            "  Contact IT to get an Azure App Registration if one does not exist yet."
        )

    token_url = _TOKEN_URL_TEMPLATE.format(tenant_id=D365_TENANT_ID)

    response = requests.post(
        token_url,
        data={
            "grant_type": "client_credentials",
            "client_id": D365_CLIENT_ID,
            "client_secret": D365_CLIENT_SECRET,
            "scope": f"{D365_ORG_URL}/.default",
        },
        timeout=30,
    )
    response.raise_for_status()

    token_data = response.json()
    if "access_token" not in token_data:
        raise RuntimeError(
            f"No access token returned: {token_data.get('error_description', 'Unknown error')}"
        )

    return token_data["access_token"]


def _get_headers(token):
    """Build standard OData request headers with Bearer auth and formatted value annotations."""
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Prefer": (
            f"odata.maxpagesize={D365_PAGE_SIZE},"
            'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
        ),
    }


# ============================================================================
# CONNECTION VERIFICATION
# ============================================================================

def verify_connection():
    """
    Verify D365 API connectivity and authentication using the WhoAmI endpoint.

    Returns:
        bool: True if the connection succeeds, False on any failure.
    """
    try:
        token = get_access_token()
        response = requests.get(
            f"{_API_BASE}/WhoAmI()",
            headers=_get_headers(token),
            timeout=15,
        )
        response.raise_for_status()
        return True

    except ValueError as e:
        print(f"  {Messages.ERROR} {e}")
        return False
    except requests.exceptions.ConnectionError:
        print(f"  {Messages.ERROR} Cannot reach D365 at {D365_ORG_URL}. Check VPN.")
        return False
    except requests.exceptions.HTTPError as e:
        status = e.response.status_code
        if status == 401:
            print(f"  {Messages.ERROR} D365 authentication failed. Check CLIENT_ID / CLIENT_SECRET.")
        elif status == 403:
            print(
                f"  {Messages.ERROR} D365 access denied (403). "
                "Ensure the App Registration has been assigned a D365 Security Role."
            )
        else:
            print(f"  {Messages.ERROR} D365 HTTP error {status}: {e}")
        return False
    except Exception as e:
        print(f"  {Messages.ERROR} D365 connection error: {e}")
        return False


# ============================================================================
# COLUMN NAME RESOLUTION  (schema name → display name)
# ============================================================================

def _get_view_schema_columns(view_id, token):
    """
    Fetch a saved view's layoutxml and return the list of schema column names
    in the order they appear in the view.

    Returns:
        list[str]: Schema column names (e.g. ['cr9cc_globalalcumusid', 'statuscode', ...])
                   Empty list if the view has no layoutxml.
    """
    response = requests.get(
        f"{_API_BASE}/savedqueries({view_id})?$select=layoutxml,returnedtypecode",
        headers=_get_headers(token),
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()

    layout_xml = data.get("layoutxml", "")
    if not layout_xml:
        logger.warning(f"View {view_id} has no layoutxml (may be a personal view — try systemuserqueries)")
        return []

    try:
        root = ET.fromstring(layout_xml)
        return [cell.get("name") for cell in root.iter("cell") if cell.get("name")]
    except ET.ParseError as e:
        logger.warning(f"Could not parse layoutxml for view {view_id}: {e}")
        return []


def _get_attribute_display_names(entity_logical_name, schema_names, token):
    """
    Fetch display names for a list of schema column names via the entity metadata API.

    Returns:
        dict: {schema_name: display_name}  for all resolved columns.
    """
    if not schema_names:
        return {}

    filter_expr = " or ".join(f"LogicalName eq '{name}'" for name in schema_names)
    url = (
        f"{_API_BASE}/EntityDefinitions(LogicalName='{entity_logical_name}')"
        f"/Attributes?$select=LogicalName,DisplayName&$filter={filter_expr}"
    )

    try:
        response = requests.get(url, headers=_get_headers(token), timeout=30)
        response.raise_for_status()
        attrs = response.json().get("value", [])
    except Exception as e:
        logger.warning(f"Could not fetch attribute metadata for {entity_logical_name}: {e}")
        return {}

    mapping = {}
    for attr in attrs:
        schema = attr.get("LogicalName", "")
        label = (
            (attr.get("DisplayName") or {})
            .get("UserLocalizedLabel", {}) or {}
        ).get("Label", "")
        if schema and label:
            mapping[schema] = label

    return mapping


def _build_column_rename_map(view_id, entity_logical_name, token):
    """
    Build a complete schema_name → display_name mapping for a saved view's columns.

    Strategy:
    1. Start with D365_KNOWN_FIELD_NAMES from config (handles critical fields
       like statuscode → 'Status Reason' even if metadata API fails).
    2. Fetch the view's column list from layoutxml.
    3. Enrich with D365 attribute metadata (overrides fallback names where available).

    Returns:
        dict: {schema_name: display_name}
    """
    # Start with hardcoded fallbacks for well-known fields
    rename_map = dict(D365_KNOWN_FIELD_NAMES)

    try:
        schema_cols = _get_view_schema_columns(view_id, token)
        if schema_cols:
            metadata_names = _get_attribute_display_names(entity_logical_name, schema_cols, token)
            rename_map.update(metadata_names)  # metadata takes precedence
            logger.info(f"Resolved {len(metadata_names)} column display names from metadata")
        else:
            logger.warning("No schema columns found in view — using fallback names only")
    except Exception as e:
        logger.warning(f"Column name resolution failed, using fallback names only: {e}")

    return rename_map


# ============================================================================
# DATA FETCHING  (with OData pagination)
# ============================================================================

def _fetch_all_pages(url, headers):
    """
    Fetch all pages of an OData result set, following @odata.nextLink until exhausted.

    Args:
        url: Initial OData query URL.
        headers: Request headers (must include auth + Prefer).

    Returns:
        list[dict]: All records across all pages.

    Raises:
        requests.HTTPError: On any HTTP error response.
    """
    all_records = []
    page = 1

    while url:
        response = requests.get(url, headers=headers, timeout=120)
        response.raise_for_status()

        data = response.json()
        records = data.get("value", [])
        all_records.extend(records)

        url = data.get("@odata.nextLink")  # None when we're on the last page
        if url:
            elapsed_msg = f"{len(all_records):,} rows fetched so far, loading page {page + 1}..."
            print(f"    ... {elapsed_msg}")
            logger.debug(elapsed_msg)

        page += 1

    return all_records


def _records_to_dataframe(records, rename_map):
    """
    Convert a list of OData record dicts to a clean pandas DataFrame.

    Processing rules:
    - Skip OData metadata keys (keys starting with "@").
    - Skip D365 lookup _value fields (pattern: starts with "_", ends with "_value").
    - For option set / lookup fields, use the @FormattedValue annotation when available
      so columns contain human-readable labels (e.g. "Approved") not integer codes.
    - Rename schema column names to display names using rename_map.

    Returns:
        pandas DataFrame, or empty DataFrame if records is empty.
    """
    if not records:
        return pd.DataFrame()

    cleaned = []
    for record in records:
        row = {}
        for key, value in record.items():
            # Skip OData metadata annotations
            if key.startswith("@"):
                continue
            # Skip navigation property _value fields (e.g. _ownerid_value)
            if key.startswith("_") and key.endswith("_value"):
                continue
            # Prefer the human-readable formatted value where available
            formatted_key = f"{key}{_FORMATTED_VALUE}"
            row[key] = record[formatted_key] if formatted_key in record else value
        cleaned.append(row)

    df = pd.DataFrame(cleaned)

    # Rename schema names to display names for any columns we have mappings for
    if rename_map:
        cols_to_rename = {col: rename_map[col] for col in df.columns if col in rename_map}
        if cols_to_rename:
            df = df.rename(columns=cols_to_rename)
            logger.debug(f"Renamed {len(cols_to_rename)} columns to display names")

    return df


# ============================================================================
# ORCHESTRATION — SINGLE REPORT DOWNLOAD
# ============================================================================

def run_single_d365_download(report_type, view_id, token):
    """
    Download one D365 report type using its saved view and save as Excel.

    Steps:
        1. Resolve column display names via view metadata + attribute metadata.
        2. Fetch all rows from the saved view (handles pagination automatically).
        3. Build DataFrame with human-readable column names and values.
        4. Save to input/dynamics/<report_type>_d365.xlsx.

    Args:
        report_type: e.g. 'accreditation', 'wcb', 'client', etc.
        view_id: D365 saved view GUID (from viewid= in the D365 URL).
        token: Valid OAuth2 Bearer token.

    Returns:
        Path to the saved Excel file.

    Raises:
        ValueError: If the view returns no records.
        requests.HTTPError: On API errors.
    """
    display = REPORT_TYPE_DISPLAY_NAMES.get(report_type, report_type.title())

    # 1. Resolve column names
    print(f"  📋 Resolving column names for {display}...")
    rename_map = _build_column_rename_map(view_id, D365_ENTITY_LOGICAL_NAME, token)
    resolved = sum(1 for k in rename_map if k != rename_map.get(k, k))
    print(f"  {Messages.SUCCESS} {len(rename_map)} column name mappings ready")

    # 2. Fetch data via saved view
    print(f"  📡 Fetching rows from D365 view {view_id[:8]}......")
    headers = _get_headers(token)
    url = f"{_API_BASE}/{D365_ENTITY}?savedQuery={view_id}"

    records = _fetch_all_pages(url, headers)

    if not records:
        raise ValueError(
            f"D365 view returned 0 records for {report_type}. "
            "Check that the view ID is correct and the App Registration has read access."
        )

    print(f"  {Messages.SUCCESS} Downloaded {len(records):,} rows")
    logger.info(f"Downloaded {len(records)} rows for {report_type} (view {view_id})")

    # 3. Build DataFrame
    df = _records_to_dataframe(records, rename_map)
    print(f"  {Messages.SUCCESS} Built DataFrame: {len(df):,} rows × {len(df.columns)} columns")

    # 4. Save as Excel
    DYNAMICS_DIR.mkdir(parents=True, exist_ok=True)
    output_path = DYNAMICS_DIR / D365_FILES[report_type]
    df.to_excel(output_path, index=False)
    print(f"  {Messages.SUCCESS} Saved to {output_path.name}")
    logger.info(f"Saved {report_type} D365 data → {output_path}")

    return output_path


# ============================================================================
# ORCHESTRATION — ALL DOWNLOADS
# ============================================================================

def run_all_d365_downloads():
    """
    Download all configured D365 report type views and save as Excel files.

    Only processes report types that have a view ID configured in D365_VIEW_IDS.
    Skips unconfigured types with a warning (does not raise).

    Returns:
        dict: {report_type: output_path} for each successful download.

    Raises:
        ValueError: If no view IDs are configured at all.
        ConnectionError: If the D365 API cannot be reached.
    """
    print("\n" + "=" * 70)
    print("DYNAMICS 365: DOWNLOADING REPORT VIEWS")
    print("=" * 70)

    # Partition view IDs into configured vs still-pending
    configured = {rt: vid for rt, vid in D365_VIEW_IDS.items() if vid}
    not_configured = [rt for rt in D365_VIEW_IDS if not D365_VIEW_IDS[rt]]

    if not configured:
        raise ValueError(
            "No D365 view IDs are configured yet.\n"
            "  Set the following environment variables (or edit D365_VIEW_IDS in src/config.py):\n"
            + "\n".join(f"    D365_VIEW_ID_{rt.upper()}" for rt in D365_VIEW_IDS)
        )

    if not_configured:
        print(
            f"\n  {Messages.WARNING} {len(not_configured)} view ID(s) not yet configured "
            f"(skipping): {', '.join(not_configured)}"
        )

    # Verify connection before attempting downloads
    print("\n📡 Verifying D365 API connection...")
    if not verify_connection():
        raise ConnectionError(
            "Cannot connect to D365 API. Check credentials and VPN connection."
        )
    print(f"  {Messages.SUCCESS} Connected to Dynamics 365\n")

    # Authenticate once — token is valid ~60 min, enough for all downloads
    token = get_access_token()
    results = {}

    for report_type in ["accreditation", "wcb", "client", "critical_document", "esg"]:
        view_id = configured.get(report_type)
        if not view_id:
            continue

        display = REPORT_TYPE_DISPLAY_NAMES.get(report_type, report_type.title())
        print(f"\n{Messages.PROCESSING} Downloading {display.upper()}...")

        try:
            output_path = run_single_d365_download(report_type, view_id, token)
            results[report_type] = output_path
        except Exception as e:
            print(f"  {Messages.ERROR} Failed: {e}")
            logger.error(f"D365 download failed for {report_type}: {e}")
            continue

    # Summary
    print("\n" + "-" * 40)
    print(f"Downloaded {len(results)}/{len(configured)} D365 report(s):")
    for rt, path in results.items():
        print(f"  {Messages.SUCCESS} {REPORT_TYPE_DISPLAY_NAMES.get(rt, rt.title())}: {path.name}")
    failed = set(configured.keys()) - set(results.keys())
    for rt in sorted(failed):
        print(f"  {Messages.ERROR} {REPORT_TYPE_DISPLAY_NAMES.get(rt, rt.title())}: Failed")

    return results


# ============================================================================
# STANDALONE ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    """Run D365 downloads standalone — useful for testing credentials and view IDs."""
    try:
        results = run_all_d365_downloads()
        if results:
            print(f"\n✅ Successfully downloaded {len(results)} D365 report(s)")
        else:
            print("\n❌ No D365 reports were downloaded")
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
