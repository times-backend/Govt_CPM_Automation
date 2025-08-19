import logging

def create_advertiser(client, advertiser_name):
    company_service = client.GetService('CompanyService', version='v202408')
    advertiser = {
        'name': advertiser_name,
        'type': 'ADVERTISER'
    }
    created_advertiser = company_service.createCompanies([advertiser])
    if created_advertiser:
        advertiser_id = created_advertiser[0]['id']
        logging.info(f"Advertiser '{advertiser_name}' created with ID: {advertiser_id}")
        return advertiser_id
    else:
        logging.error("Failed to create advertiser.")
        return None
    
def hello():
    print("hi")