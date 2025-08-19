from googleads import ad_manager
from create_advertiserId import create_advertiser

def get_adbvertiser_id(client, company_name, company_type):
    """Fetch the ID of a company (advertiser) from Google Ad Manager."""
    try:
        company_service = client.GetService('CompanyService', version='v202408')
        statement = (ad_manager.StatementBuilder()
                     .Where('name = :name AND type = :type')
                     .WithBindVariable('name', company_name)
                     .WithBindVariable('type', company_type))
        response = company_service.getCompaniesByStatement(statement.ToStatement())

        if 'results' in response and len(response['results']) > 0:
            return response['results'][0]['id']
        else:
            return None
        
    except Exception as e:
        print(f"Error fetching company ID: {e}")
        return None

def fetch_trafficker_id(client, trafficker_email):
    """Fetch the ID of a trafficker using email."""
    try:
        user_service = client.GetService('UserService', version='v202408')
        
        # Create a statement to filter by email
        statement = ad_manager.StatementBuilder()
        statement.Where('email = :email')
        statement.WithBindVariable('email', trafficker_email)
        
        response = user_service.getUsersByStatement(statement.ToStatement())
        
        if response and 'results' in response and len(response['results']) > 0:
            user = response['results'][0]
            print(f"✅ Found trafficker: {user['name']} ({user['email']}) - ID: {user['id']}")
            return user['id']
        
        # If not found, get all users to show available options
        print(f"⚠️ No trafficker found with email: {trafficker_email}")
        print("\nAvailable traffickers:")
        all_users = user_service.getUsersByStatement(ad_manager.StatementBuilder().ToStatement())
        if all_users and 'results' in all_users:
            for user in all_users['results']:
                print(f"  • {user['name']} ({user['email']})")
        return None
        
    except Exception as e:
        print(f"❌ Error fetching trafficker ID: {e}")
        print(f"⚠️ Please ensure {trafficker_email} is registered in GAM and has appropriate permissions.")
        return None

def create_order(client, advertiser_name, trafficker_name, order_name, line_item_data):
    """Create an order in Google Ad Manager."""
    try:
        # Fetch IDs
        print(f"\n{'='*50}")
        print(f"📋 Order Creation Details:")
        print(f"  • Order Name: {order_name}")
        print(f"  • Advertiser: {advertiser_name}")
        print(f"  • Trafficker: {trafficker_name}")
        print(f"{'='*50}\n")
        
        # Get advertiser ID
        advertiser_id = get_adbvertiser_id(client, advertiser_name, 'ADVERTISER')
        if advertiser_id is None:
            print(f'⚠️ Advertiser ID not found, creating new advertiser...')
            advertiser_id = create_advertiser(client, advertiser_name)
            if advertiser_id is None:
                print(f"❌ Failed to create advertiser: {advertiser_name}")
                return None
            print(f"✅ Created new advertiser with ID: {advertiser_id}")
        else:
            print(f"✅ Found existing advertiser with ID: {advertiser_id}")
        
        # Get trafficker ID
        trafficker_id = fetch_trafficker_id(client, trafficker_name)
        if trafficker_id is None:
            print(f"❌ Trafficker not found: {trafficker_name}")
            print("⚠️ Please ensure the trafficker email is registered in GAM and has appropriate permissions.")
            return None

        print(f"Advertiser ID: {advertiser_id}, Trafficker ID: {trafficker_id}")

        # Create order
        order_service = client.GetService('OrderService', version='v202408')
        order = {
            'name': order_name,
            'advertiserId': advertiser_id,
            'traffickerId': trafficker_id,
        }

        # Handle label if present
        if line_item_data.get("label"):
            label_name = line_item_data["label"]
            label_service = client.GetService("LabelService", version="v202408")
            
            # Check for "ad exclusion" in label name (case-insensitive)
            if "ad exclusion" in label_name.lower():
                print(f"🚫 Skipping label '{label_name}' because it contains 'Ad exclusion'")
            else:
                statement = {'query': f"WHERE name = '{label_name}'"}
                response = label_service.getLabelsByStatement(statement)

                if 'results' in response and len(response['results']) > 0:
                    matched_label = response['results'][0]
                    label_id = matched_label['id']
                    order["appliedLabels"] = [{"labelId": label_id}]
                    print(f"✅ Order label applied: {label_name} (ID: {label_id})")
                else:
                    print(f"⚠️ Label '{label_name}' not found in the system.")

        # Create the order
        created_order = order_service.createOrders([order])[0]
        print(f"Order '{created_order['name']}' created with ID: {created_order['id']}")
        return created_order['id']

    except Exception as e:
        print(f"Error creating order: {e}")
        return None

if __name__ == "__main__":
    client = ad_manager.AdManagerClient.LoadFromStorage('googleads1.yaml')
    advertiser_name = "IDEAL LIMITED"
    trafficker_name = "Anurag Mishra"  # Replace with actual trafficker name
    order_name = "Test Oer12"

    # Add label for the order
    line_item_data = {
        "Line_label": "Lifestyle"
    }

    # Create the order with the label
    order_id = create_order(client, advertiser_name, trafficker_name, order_name, line_item_data)
    print(f"Created Order ID: {order_id}")