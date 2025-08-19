from google.cloud import bigquery
import pandas as pd
import os
import sys
from google.oauth2 import service_account

def fetch_expresso_data(expresso_id, service_account_file=None, project_id=None):
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
    if service_account_file and os.path.exists(service_account_file):
        print(f"Using service account file: {service_account_file}")
        credentials = service_account.Credentials.from_service_account_file(
            service_account_file,
            scopes=["https://www.googleapis.com/auth/cloud-platform"],
        )
        client = bigquery.Client(credentials=credentials, project=project_id)
    else:
        print("Using application default credentials")
        client = bigquery.Client(project=project_id)
    
    # Define the SQL query with expresso_id filter
    query = f"""
    SELECT
        expresso_id,
        til_ro_id,
        campaign_name,
        proposal_type,
        client,
        client_billable,
        sales_industry,
        package_start_date,
        package_end_date,
        pricing_strategy,
        til_ro_package_id,
        til_package_id,
        package_name,
        geo_target,
        quantity,
        sales_person,
        L1_sales,
        L2_sales,
        cs_user,
        ops_user,
        lineitem_comment,
        website_name,
        CASE WHEN audience_abvr_price IS NULL OR TRIM(audience_abvr_price) = '' THEN 'No' ELSE 'Yes' END AS audience_targeting,
        SUM(gross_rate_inr) AS gross_rate,
        SUM(SP_inr) AS RO_value
    FROM acoe21.til_strategic.expresso_ro_li_pricing_parent_only
    WHERE 
        ro_order_create_date >= '2024-04-01' 
        AND ro_order_create_date <= CURRENT_DATE()
        AND pricing_strategy = 'CPM'
        AND expresso_id = '{expresso_id}'  -- Filter for the specific expresso ID
    GROUP BY 
        expresso_id, til_ro_id, campaign_name, proposal_type, client, client_billable,
        sales_industry, package_start_date, package_end_date, pricing_strategy,
        til_ro_package_id, til_package_id, package_name, geo_target, quantity,
        sales_person, L1_sales, L2_sales, cs_user, ops_user, lineitem_comment,
        website_name, audience_targeting
    """
    
    print(f"Fetching BigQuery data for expresso ID: {expresso_id}")
    print(f"Running query...")
    
    # Run the query
    query_job = client.query(query)
    
    # Convert results to DataFrame
    results = query_job.result()
    df = results.to_dataframe()
    
    # Save results to CSV
    output_file = f"expresso_data_{expresso_id}.csv"
    df.to_csv(output_file, index=False)
    print(f"Query results saved to {output_file}")
    
    return df

if __name__ == "__main__":
    # Check for service account file
    service_account_file = None
    if os.path.exists('service-account.json'):
        service_account_file = 'service-account.json'
        print(f"Found service account file: {service_account_file}")
    
    # Get expresso_id from command line or use a default
    if len(sys.argv) < 2:
        print("Usage: python bigquery_fetch.py EXPRESSO_ID [PROJECT_ID] [SERVICE_ACCOUNT_FILE]")
        print("\nArguments:")
        print("  EXPRESSO_ID          - Required: The Expresso ID to query")
        print("  PROJECT_ID           - Optional: Google Cloud project ID")
        print("  SERVICE_ACCOUNT_FILE - Optional: Path to service account key file")
        print("\nExample: python bigquery_fetch.py 272237 my-project-id ./keys/service-account.json")
        expresso_id = input("\nEnter the expresso ID: ")
    else:
        expresso_id = sys.argv[1]
    
    # Get optional project ID
    project_id = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Get optional service account file path
    if len(sys.argv) > 3:
        service_account_file = sys.argv[3]
    
    try:
        # Fetch the data
        result_df = fetch_expresso_data(expresso_id, service_account_file, project_id)
        
        # Display summary of results
        if len(result_df) > 0:
            print(f"\nFound {len(result_df)} matching records for expresso ID {expresso_id}")
            print("\nSummary:")
            if 'campaign_name' in result_df.columns:
                print(f"Campaign: {result_df['campaign_name'].iloc[0]}")
            if 'client' in result_df.columns:
                print(f"Client: {result_df['client'].iloc[0]}")
            if 'RO_value' in result_df.columns:
                print(f"Total RO Value: {result_df['RO_value'].sum()}")
        else:
            print(f"No matching records found for expresso ID {expresso_id}")
    except Exception as e:
        print(f"\nError: {str(e)}")
        print("\nIf this is an authentication error, please run the authentication helper script:")
        print("python authenticate_google_cloud.py")
        sys.exit(1) 