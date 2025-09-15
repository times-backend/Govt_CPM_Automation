from google.cloud import bigquery
import pandas as pd
import os
import sys
from google.oauth2 import service_account

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "credentials.json"
client = bigquery.Client()

def fetch_expresso_data(expresso_id):
    """
    Fetch data from BigQuery for a specific expresso ID
    
    Args:
        expresso_id: The expresso ID to filter results by
        service_account_file: Optional path to service account key file
        project_id: Optional explicit project ID
    
    Returns:
        DataFrame containing the query results
    """
    # Initialize the BigQuery client
    service_account_file="credentials.json"
    
    # Define the SQL query with expresso_id filter
    query = f"""
    SELECT *
    FROM `acoe21.til_strategic.expresso_ro_li_pricing_parent_only`
    WHERE expresso_id = @search_value
    """

    job_config = bigquery.QueryJobConfig(
        query_parameters=[
            bigquery.ScalarQueryParameter("search_value", "INT64", expresso_id)
        ]
    )

    try:
        query_job = client.query(query, job_config=job_config)
        results = query_job.result()
        rows = [dict(row) for row in results]

        print(f"\n Found {len(rows)} row(s) for expresso_id = '{expresso_id}'")
        advertiser_name = rows[0].get("agency_name") if rows[0].get("agency_name") and rows[0].get("agency_name").strip().upper() != "N A" else rows[0].get("client")
        order_name = rows[0].get("ref_no")
        return order_name,advertiser_name

    except Exception as e:
        print(f"\n Error querying BigQuery: {e}")
        return []