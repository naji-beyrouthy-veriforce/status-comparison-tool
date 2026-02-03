"""
Redash API Integration
Automates query execution and result download from Redash
"""

import requests
import time
import pandas as pd
from pathlib import Path
import json

# Redash Configuration
REDASH_URL = "https://redash.cognibox.net"
API_KEY = "RpWSRcBbV8IHkvXumk442ttCiU2j9XLSa0niHXRD"

# Query IDs
QUERY_IDS = {
    "accreditation": 1266,
    "wcb": 1281,
    "client": 1277
}

# Request headers
HEADERS = {
    "Authorization": f"Key {API_KEY}",
    "Content-Type": "application/json"
}


def get_query_metadata(query_id):
    """
    Fetch query metadata to discover parameter names
    """
    url = f"{REDASH_URL}/api/queries/{query_id}"
    response = requests.get(url, headers=HEADERS)
    
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f"Failed to fetch query metadata: {response.status_code} - {response.text}")


def execute_query_with_params(query_id, parameters):
    """
    Execute a Redash query with parameters
    Forces fresh execution with new IDs (no cache)
    Returns the job ID for polling
    """
    # Use refresh endpoint to force new execution
    url = f"{REDASH_URL}/api/queries/{query_id}/refresh"
    
    # Prepare parameters with p_ prefix (Redash convention)
    params = {}
    for key, value in parameters.items():
        params[f"p_{key}"] = value
    
    # Add max_age=0 to force fresh execution (ignore cache)
    params["max_age"] = 0
    
    # Add parameters to URL query string
    response = requests.post(url, headers=HEADERS, params=params)
    
    if response.status_code in [200, 201]:
        data = response.json()
        job = data.get("job", {})
        
        # Check if we got a job object
        if job and "id" in job:
            return job.get("id")
        
        # Debug: Print response to understand structure
        print(f"     Debug: Response keys: {data.keys()}")
        if "job" in data:
            print(f"     Debug: Job keys: {job.keys()}")
        
        return None
    else:
        raise Exception(f"Failed to execute query: {response.status_code} - {response.text}")


def poll_query_results(job_id, max_attempts=60, delay=2):
    """
    Poll for query results until complete or timeout
    """
    url = f"{REDASH_URL}/api/jobs/{job_id}"
    
    for attempt in range(max_attempts):
        response = requests.get(url, headers=HEADERS)
        
        if response.status_code == 200:
            job = response.json().get("job", {})
            status = job.get("status")
            
            if status == 3:  # Success
                query_result_id = job.get("query_result_id")
                return query_result_id
            elif status == 4:  # Failed
                raise Exception("Query execution failed")
            
            # Still running, wait and retry
            time.sleep(delay)
        else:
            raise Exception(f"Failed to poll job status: {response.status_code}")
    
    raise Exception("Query execution timed out")


def download_query_results(query_result_id):
    """
    Download query results as a pandas DataFrame
    """
    url = f"{REDASH_URL}/api/query_results/{query_result_id}"
    response = requests.get(url, headers=HEADERS)
    
    if response.status_code == 200:
        data = response.json()
        rows = data.get("query_result", {}).get("data", {}).get("rows", [])
        
        if rows:
            df = pd.DataFrame(rows)
            return df
        else:
            return pd.DataFrame()
    else:
        raise Exception(f"Failed to download results: {response.status_code}")


def execute_redash_query(report_type, ids_list):
    """
    Main function to execute a Redash query and return results
    
    Args:
        report_type: 'accreditation', 'wcb', or 'client'
        ids_list: List of contractor IDs to query
    
    Returns:
        pandas DataFrame with query results
    """
    query_id = QUERY_IDS.get(report_type)
    if not query_id:
        raise ValueError(f"Unknown report type: {report_type}")
    
    print(f"\n  🔄 Executing Redash query for {report_type}...")
    print(f"     Query ID: {query_id}")
    print(f"     Using {len(ids_list)} fresh IDs from uploaded D365 file")
    print(f"     ⚡ Forcing fresh execution (ignoring cache)")
    
    # Get query metadata to discover parameter names
    try:
        metadata = get_query_metadata(query_id)
        
        # Extract parameter names from query
        query_text = metadata.get("query", "")
        
        # Common parameter name patterns
        param_patterns = ["contractor_ids", "ids", "global_alcumus_ids", "contractor_id_list"]
        param_name = None
        
        for pattern in param_patterns:
            if f"{{{{{pattern}}}}}" in query_text:
                param_name = pattern
                break
        
        # If no pattern found, try to extract from options
        if not param_name:
            options = metadata.get("options", {}).get("parameters", [])
            if options:
                param_name = options[0].get("name", "ids")
        
        # Default to 'ids' if still not found
        if not param_name:
            param_name = "ids"
        
        print(f"     Using parameter: {param_name}")
        
    except Exception as e:
        print(f"     Warning: Could not fetch metadata, using default parameter 'ids'")
        print(f"     Error: {e}")
        param_name = "ids"
    
    # Format IDs - try without quotes first (Redash might handle this in the query)
    # If ids_list is empty, return empty DataFrame
    if not ids_list:
        print(f"     ⚠ No IDs to query")
        return pd.DataFrame()
    
    ids_string = ",".join(ids_list)  # Try without quotes first
    
    print(f"     Sample IDs: {ids_string[:100]}...")
    
    # Execute query (force fresh execution with uploaded IDs)
    try:
        parameters = {param_name: ids_string}
        job_id = execute_query_with_params(query_id, parameters)
        
        if not job_id:
            raise Exception("No job ID returned from query execution")
        
        print(f"     Job ID: {job_id}")
        print(f"     Executing fresh query with {len(ids_list)} IDs...")
        
        # Poll for results
        query_result_id = poll_query_results(job_id)
        print(f"     Query completed! Result ID: {query_result_id}")
        
        # Download results
        df = download_query_results(query_result_id)
        print(f"     ✓ Downloaded {len(df)} rows")
        
        return df
        
    except Exception as e:
        print(f"     ❌ Error: {e}")
        raise


def save_redash_results(report_type, df, output_dir):
    """
    Save Redash query results to Excel file
    """
    filename = f"{report_type}_sc.xlsx"
    filepath = output_dir / filename
    
    df.to_excel(filepath, index=False)
    print(f"     ✓ Saved to: {filename}")
    
    return filepath


def test_redash_connection():
    """
    Test Redash API connection
    """
    url = f"{REDASH_URL}/api/queries"
    response = requests.get(url, headers=HEADERS)
    
    if response.status_code == 200:
        print("✓ Redash API connection successful")
        return True
    else:
        print(f"❌ Redash API connection failed: {response.status_code}")
        return False


if __name__ == "__main__":
    # Test connection
    test_redash_connection()

