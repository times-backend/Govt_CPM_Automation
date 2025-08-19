from googleads import ad_manager


def fetch_advertiser_id_from_order(client, order_id):
    order_service = client.GetService('OrderService', version='v202408')  # Adjust the version as needed
    statement = ad_manager.StatementBuilder().Where('id = :order_id').WithBindVariable('order_id', order_id)
    try:
        response = order_service.getOrdersByStatement(statement.ToStatement())
        orders = response.results  


        if orders:
            advertiser_id = orders[0].advertiserId
            return advertiser_id
        else:
            raise ValueError(f'No order found with ID: {order_id}')


    except Exception as e:
        print(f'Error occurred while fetching advertiser ID: {e}')
        return None
def get_order_name(client,order_id ):
    """Fetch the name of the order for a given order_id."""
    # Initialize the OrderService
    order_service = client.GetService('OrderService')

    # Create a statement to filter the order by its ID
    statement = ad_manager.StatementBuilder().Where('id = :order_id').WithBindVariable('order_id', order_id)

    # Get the orders matching the statement
    response = order_service.getOrdersByStatement(statement.ToStatement())

    if 'results' in response:
        # If results exist, return the name of the first order
        order = response['results'][0]
        return order['name']
    else:
        print(f'No order found for order_id: {order_id}')
        return None
    


