#!/usr/bin/env python3
"""
Simple script to check if the target order exists and is valid
"""

import sys
from googleads import ad_manager

def main():
    order_id = "3811823998"
    
    print(f"🔍 Checking target order: {order_id}")
    
    try:
        # Get authenticated client
        client = ad_manager.AdManagerClient.LoadFromStorage("googleads1.yaml")
        print("✅ Successfully authenticated with Google Ad Manager")
        
        # Get order service
        order_service = client.GetService('OrderService', version='v202408')
        
        # Try to get the order directly using the service
        try:
            orders = order_service.getOrdersByStatement({
                'query': f'WHERE Id = {order_id}'
            })
            
            if orders and orders.results:
                order = orders.results[0]
                print(f"✅ Order found!")
                print(f"   - ID: {order.id}")
                print(f"   - Name: {order.name}")
                print(f"   - Status: {order.status}")
                print(f"   - Advertiser ID: {order.advertiserId}")
                print(f"   - Currency: {order.currencyCode}")
                print(f"   - Start Date: {order.startDateTime}")
                print(f"   - End Date: {order.endDateTime}")
                return 0
            else:
                print(f"❌ Order {order_id} not found!")
                print("🔧 This could be why line item creation is failing with DUPLICATE_OBJECT")
                return 1
                
        except Exception as e:
            print(f"❌ Error getting order: {e}")
            return 1
            
    except Exception as e:
        print(f"❌ Error during authentication or setup: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 