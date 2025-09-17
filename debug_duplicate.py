#!/usr/bin/env python3
"""
Debug script to investigate DUPLICATE_OBJECT errors
"""

import sys
from googleads import ad_manager
from single_line import check_line_item_name_exists

def debug_specific_line_item(client, line_name):
    """Debug function to investigate a specific line item name that's causing DUPLICATE_OBJECT errors"""
    try:
        print(f"\n🔍 DEBUGGING SPECIFIC LINE ITEM: {line_name}")
        print(f"🔍 Line name length: {len(line_name)} characters")
        
        pql_service = client.GetService('PublisherQueryLanguageService', version='v202508')
        
        # Try multiple search strategies including archived/deleted items
        search_strategies = [
            f"SELECT Id, Name, OrderId, Status FROM Line_Item WHERE Name = '{line_name}'",
            f"SELECT Id, Name, OrderId, Status FROM Line_Item WHERE Name LIKE '{line_name}'", 
            f"SELECT Id, Name, OrderId, Status FROM Line_Item WHERE Name LIKE '%{line_name}%'",
            # Search for archived/paused items too
            f"SELECT Id, Name, OrderId, Status FROM Line_Item WHERE Name = '{line_name}' AND Status IN ('PAUSED', 'INACTIVE')",
        ]
        
        for i, query in enumerate(search_strategies, 1):
            print(f"\n🔍 Strategy {i}: {query}")
            try:
                statement = {'query': query}
                response = pql_service.select(statement)
                
                if 'rows' in response and response['rows']:
                    print(f"   ✅ Found {len(response['rows'])} matches:")
                    for row in response['rows']:
                        line_id = row['values'][0]['value']
                        name = row['values'][1]['value']
                        order_id = row['values'][2]['value']
                        status = row['values'][3]['value'] if len(row['values']) > 3 else 'Unknown'
                        print(f"      - ID: {line_id}, Name: '{name}', Order: {order_id}, Status: {status}")
                        
                        # Check if it's an exact match
                        if name == line_name:
                            print(f"      ⚠️ EXACT MATCH! This explains the DUPLICATE_OBJECT error")
                        elif name.upper() == line_name.upper():
                            print(f"      ⚠️ CASE-INSENSITIVE MATCH! This might explain the DUPLICATE_OBJECT error")
                else:
                    print(f"   ❌ No matches found")
                    
            except Exception as e:
                print(f"   ❌ Query failed: {e}")
        
        # Also try to find any line items with similar patterns
        print(f"\n🔍 Searching for line items with similar patterns...")
        pattern_searches = [
            f"SELECT Id, Name, OrderId FROM Line_Item WHERE Name LIKE '%DOMEBRANDMTIL%'",
            f"SELECT Id, Name, OrderId FROM Line_Item WHERE Name LIKE '%PKG213475%'",
            f"SELECT Id, Name, OrderId FROM Line_Item WHERE Name LIKE '%27875410%'"
        ]
        
        for query in pattern_searches:
            try:
                print(f"🔍 Pattern search: {query}")
                statement = {'query': query}
                response = pql_service.select(statement)
                
                if hasattr(response, 'rows') and response.rows:
                    print(f"   Found {len(response.rows)} line items with similar patterns:")
                    for row in response.rows[:5]:  # Show first 5 results
                        line_id = row.values[0].value
                        name = row.values[1].value
                        order_id = row.values[2].value
                        print(f"      - ID: {line_id}, Name: '{name}', Order: {order_id}")
                        if name == line_name:
                            print(f"        ⚠️ EXACT MATCH FOUND!")
                else:
                    print(f"   No matches for this pattern")
            except Exception as e:
                print(f"   Pattern search failed: {e}")
        
        # Check for other objects that might have naming conflicts
        print(f"\n🔍 Checking for potential naming conflicts with other objects...")
        
        # Check for orders with similar names (unlikely but possible)
        try:
            order_query = f"SELECT Id, Name FROM Order WHERE Name LIKE '%{line_name}%'"
            print(f"🔍 Checking orders: {order_query}")
            order_statement = {'query': order_query}
            order_response = pql_service.select(order_statement)
            
            if hasattr(order_response, 'rows') and order_response.rows:
                print(f"   Found {len(order_response.rows)} orders with similar names:")
                for row in order_response.rows:
                    order_id = row.values[0].value
                    name = row.values[1].value
                    print(f"      - Order ID: {order_id}, Name: '{name}'")
            else:
                print(f"   No orders found with similar names")
        except Exception as e:
            print(f"   Order check failed: {e}")
        
        # Also check the specific order we're trying to create in
        try:
            specific_order_query = f"SELECT Id, Name, Status FROM Order WHERE Id = 3811823998"
            print(f"🔍 Checking target order: {specific_order_query}")
            specific_order_statement = {'query': specific_order_query}
            specific_order_response = pql_service.select(specific_order_statement)
            
            if hasattr(specific_order_response, 'rows') and specific_order_response.rows:
                row = specific_order_response.rows[0]
                order_id = row.values[0].value
                name = row.values[1].value
                status = row.values[2].value
                print(f"   Target Order - ID: {order_id}, Name: '{name}', Status: {status}")
            else:
                print(f"   ⚠️ Target order not found! This could be the issue.")
        except Exception as e:
            print(f"   Target order check failed: {e}")
                
    except Exception as e:
        print(f"🔍 Debug function failed: {e}")

def main():
    # The specific line item name that's causing issues
    line_name = "27875410DOMEBRANDMTILROSINOUTOFPAGECPMENGNEWSINTERAPPTILINTERSTITIALPKG213475"
    order_id = "3811823998"
    
    print("🔍 Starting duplicate investigation...")
    print(f"🔍 Line name: {line_name}")
    print(f"🔍 Order ID: {order_id}")
    
    try:
        # Get authenticated client
        client = ad_manager.AdManagerClient.LoadFromStorage("googleads1.yaml")
        print("✅ Successfully authenticated with Google Ad Manager")
        
        # Run the specific debugging function
        debug_specific_line_item(client, line_name)
        
        # Also run the normal duplicate check
        print(f"\n" + "="*80)
        print("🔍 Running normal duplicate check...")
        duplicate_exists = check_line_item_name_exists(client, order_id, line_name)
        
        if duplicate_exists:
            print("✅ Duplicate detection found existing line items!")
            print("🔧 The system should automatically append timestamp to avoid DUPLICATE_OBJECT error")
        else:
            print("✅ Duplicate detection working correctly - no duplicates found")
            print("❓ This means the DUPLICATE_OBJECT error is caused by something else...")
            print("💡 Possible causes:")
            print("   - Archived/deleted line items with same name")
            print("   - Different object type with same name") 
            print("   - Order-level constraints")
            print("   - GAM caching/replication delays")
            print("   - Character encoding issues")
            
    except Exception as e:
        print(f"❌ Error during debug: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 