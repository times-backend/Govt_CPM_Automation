import sys
import json
from auth_utils import generate_jwt, fetch_package_details
from bigquery_fetch import fetch_expresso_data

def fetch_full_expresso_details(expresso_id):
    """
    Get comprehensive information about an expresso campaign from multiple sources
    
    Args:
        expresso_id: The expresso ID to lookup
    
    Returns:
        A dictionary containing the extracted campaign details
    """
   # print(f"Fetching comprehensive details for Expresso ID: {expresso_id}")
    
    # 1. Try to get data from BigQuery
   # print("\n===== BigQuery Data =====")
    try:
        bigquery_data = fetch_expresso_data(expresso_id)
        if len(bigquery_data) > 0:
            print(f"Successfully retrieved {len(bigquery_data)} records from BigQuery")
        else:
            print("No matching records found in BigQuery")
    except Exception as e:
        print(f"Error fetching BigQuery data: {str(e)}")
        bigquery_data = None
    
    # 2. Try to get data from Expresso API
    print("\n===== Expresso API Data =====")
    try:
        # User credentials
        username = "ritesh.sanjay@timesinternet.in"
        plaintext_password = "Aouto@2307"
        
        # Generate JWT token
        jwt_token = generate_jwt(username, plaintext_password)
        
        # Fetch package details from API
        api_data = fetch_package_details(jwt_token, expresso_id)
        
        if api_data:
            print("Successfully retrieved data from Expresso API")
        else:
            print("Failed to retrieve data from Expresso API")
    except Exception as e:
        print(f"Error fetching API data: {str(e)}")
        api_data = None
    
    # 3. Combine and save all data to a comprehensive JSON file
    print("\n===== Generating Combined Report =====")
    combined_data = {
        "expresso_id": expresso_id,
        #"bigquery_data": bigquery_data.to_dict('records') if bigquery_data is not None else None,
        "api_data": api_data
    }
    
    # Save combined data
    output_file = f"expresso_{expresso_id}_full_details.json"
    with open(output_file, 'w') as f:
        json.dump(combined_data, f, indent=4, default=str)
    
   # print(f"Combined data saved to {output_file}")
    
    # Extract the campaign details from the nested API response for easier access
    campaign_details =[]
    if api_data and isinstance(api_data, dict):
        for key, package_data in api_data.items():
         if isinstance(package_data, dict):
            campaign_details.append(package_data)
            print (f"campaign_details::::{campaign_details}")
    return campaign_details
  
if __name__ == "__main__":
    # Get expresso_id from command line or prompt
    expresso_id = 277099
    # Fetch comprehensive details
    fetch_full_expresso_details(expresso_id) 