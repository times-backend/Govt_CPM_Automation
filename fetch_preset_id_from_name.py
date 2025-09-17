from googleads import ad_manager

def get_preset_id_by_name(client, preset_name):
    custom_targeting_values_service = client.GetService('CustomTargetingService', version='v202508')

    statement = (ad_manager.StatementBuilder()
                .Where("name = :presetName")
                .WithBindVariable("presetName", preset_name))

    response = custom_targeting_values_service.getCustomTargetingValuesByStatement(statement.ToStatement())
    print(response)

    if 'results' in response and response['results']:
       
        preset_id = response['results'][0]['id']
        print(f"Preset Name: {preset_name}, Preset ID: {preset_id}")
        return preset_id
    else:
        print(f"No preset found with name: {preset_name}")
        return None

# Initialize Ad Manager client
client = ad_manager.AdManagerClient.LoadFromStorage('googleads.yaml')

# Replace with the preset name you want to search
preset_name = "TOI_Home_New_ATF_300x250"
preset_id = get_preset_id_by_name(client, preset_name)

if preset_id:
    print(f"Found Preset ID: {preset_id}")
else:
    print("Preset not found.")
