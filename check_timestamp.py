#!/usr/bin/env python3
"""
Check for line items with similar timestamp patterns
"""

import sys
from googleads import ad_manager

def main():
    timestamp = "1753111504359"  # The timestamp from the error
    
    print(f"🔍 Checking for line items with timestamp: {timestamp}")
    
    try:
        # Get authenticated client
        client = ad_manager.AdManagerClient.LoadFromStorage("googleads1.yaml")
        print("✅ Successfully authenticated with Google Ad Manager")
        
        # Get PQL service
        pql_service = client.GetService('PublisherQueryLanguageService', version='v202408')
        
        # Try different timestamp-related searches
        searches = [
            f"SELECT Id, Name, OrderId FROM Line_Item WHERE Name LIKE '%_{timestamp}'",
            f"SELECT Id, Name, OrderId FROM Line_Item WHERE Name LIKE '%_{timestamp}%'",
            # Also check for line items created around the same time
            f"SELECT Id, Name, OrderId, LastModifiedDateTime FROM Line_Item WHERE LastModifiedDateTime >= '2024-03-20T00:00:00' ORDER BY LastModifiedDateTime DESC LIMIT 10"
        ]
        
        for i, query in enumerate(searches, 1):
            print(f"\n🔍 Search {i}: {query}")
            try:
                statement = {'query': query}
                response = pql_service.select(statement)
                
                if hasattr(response, 'rows') and response.rows:
                    print(f"   Found {len(response.rows)} line items:")
                    for row in response.rows:
                        line_id = row.values[0].value
                        name = row.values[1].value
                        order_id = row.values[2].value
                        if len(row.values) > 3:
                            modified = row.values[3].value
                            print(f"      - ID: {line_id}, Name: '{name}', Order: {order_id}, Modified: {modified}")
                        else:
                            print(f"      - ID: {line_id}, Name: '{name}', Order: {order_id}")
                else:
                    print("   No matches found")
                    
            except Exception as e:
                print(f"   ❌ Query failed: {e}")
        
        # Also check if there are any line items with similar base names but different timestamps
        base_name = "27875410DOMEBRANDMTILROSINOUTOFPAGECPMENGNEWSINTERAPPTILINTERSTITIALPKG213475"
        print(f"\n🔍 Checking for line items with same base name but different timestamps:")
        base_query = f"SELECT Id, Name, OrderId FROM Line_Item WHERE Name LIKE '{base_name}%'"
        
        try:
            statement = {'query': base_query}
            response = pql_service.select(statement)
            
            if hasattr(response, 'rows') and response.rows:
                print(f"   Found {len(response.rows)} line items:")
                for row in response.rows:
                    line_id = row.values[0].value
                    name = row.values[1].value
                    order_id = row.values[2].value
                    print(f"      - ID: {line_id}, Name: '{name}', Order: {order_id}")
            else:
                print("   No matches found")
                
        except Exception as e:
            print(f"   ❌ Query failed: {e}")
            
        return 0
            
    except Exception as e:
        print(f"❌ Error during setup: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 