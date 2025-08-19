import gspread
from google.oauth2.service_account import Credentials

def fetch_placements_ids(credentials_path, sheet_url, sheet_name, site_filter, platforms_filter, adtype_filters, richmedia_platform_map=None, line_type="standard"):
    creds = Credentials.from_service_account_file(
        credentials_path,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(creds)
    print(f"\n🔍 Opening spreadsheet...")
    spreadsheet = client.open_by_url(sheet_url)
    
    print(f"📑 Available worksheets:")
    all_worksheets = spreadsheet.worksheets()
    sheet_titles = []
    for ws in all_worksheets:
        sheet_titles.append(ws.title)
        print(f"  - '{ws.title}' (id: {ws.id})")
    
    print(f"\n🔍 Looking for worksheet: '{sheet_name}'")
    print(f"🔍 Exact string comparison:")
    for title in sheet_titles:
        print(f"  - '{title}' == '{sheet_name}' ? {title == sheet_name}")
        if title != sheet_name:
            print(f"    Length: {len(title)} vs {len(sheet_name)}")
            print(f"    ASCII values: {[ord(c) for c in title]} vs {[ord(c) for c in sheet_name]}")
    
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        print(f"✅ Found worksheet: '{sheet_name}'")
    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ Worksheet not found: '{sheet_name}'")
        print(f"💡 Please check if the sheet name is correct and case-sensitive")
        print(f"💡 Available similar sheets:")
        import difflib
        matches = difflib.get_close_matches(sheet_name, sheet_titles, n=3, cutoff=0.1)
        for match in matches:
            print(f"   - '{match}'")
        raise
    
    # Get all values instead of using get_all_records to avoid duplicate header issues
    all_values = worksheet.get_all_values()
    if not all_values:
        print("No data found in worksheet")
        return {}
    
    # Get headers from first row and clean them up
    headers = all_values[0]
    print(f"\n📊 Sheet Structure:")
    print(f"Raw headers: {headers}")
    
    # Create a mapping of clean headers to avoid duplicates
    clean_headers = []
    header_count = {}
    for header in headers:
        clean_header = header.strip()
        if not clean_header:
            # Handle empty headers by giving them a unique name
            clean_header = f"empty_col_{len(clean_headers)}"
        elif clean_header in header_count:
            # Handle duplicate headers by appending a number
            header_count[clean_header] += 1
            clean_header = f"{clean_header}_{header_count[clean_header]}"
        else:
            header_count[clean_header] = 0
        clean_headers.append(clean_header)
    
    print(f"Cleaned headers: {clean_headers}")
    
    # Convert data rows to dictionaries manually
    data = []
    for row_values in all_values[1:]:  # Skip header row
        # Pad row with empty strings if it's shorter than headers
        while len(row_values) < len(clean_headers):
            row_values.append('')
        
        row_dict = {}
        for i, header in enumerate(clean_headers):
            if i < len(row_values):
                row_dict[header] = row_values[i]
            else:
                row_dict[header] = ''
        data.append(row_dict)

    # Clone site_filter so we don't modify the original list
    site_filter = list(site_filter)
    print(f"\n🎯 Targeting Criteria:")
    print(f"Sites: {site_filter}")
    print(f"Platforms: {platforms_filter}")

    # Convert site filter to uppercase for case-insensitive matching
    site_filter = [s.upper() for s in site_filter]
    platforms_filter = [p.upper() for p in platforms_filter]

    # Automatically add ETIMES if TOI is in the list
    if "TOI" in site_filter and "ETIMES" not in site_filter:
        site_filter.append("ETIMES")
        print(f"Added ETIMES to site filter: {site_filter}")
    
    # Find the exact column names from the sheet using original headers
    column_mapping = {}
    for col in headers:
        col_upper = col.upper()
        if 'SITE' in col_upper:
            column_mapping['site'] = col
        elif 'PLATFORM' in col_upper:
            column_mapping['platform'] = col
        elif 'AD TYPE' in col_upper or 'ADTYPE' in col_upper:
            column_mapping['adtype'] = col
        elif 'SECTION' in col_upper:
            column_mapping['section'] = col
        elif 'AD SLOT ID' in col_upper or 'AD SLOT' in col_upper or 'Ad slot id' in col:
            column_mapping['ad_slot_id'] = col
        elif 'PLACEMENT' in col_upper:
            column_mapping['placement'] = col
            
    print("\n📋 Using column mapping:")
    print(column_mapping)
    
    # Debug: Check if we found the Ad slot ID column for PSBK lines
    if line_type == "psbk":
        if 'ad_slot_id' in column_mapping:
            print(f"✅ Found Ad slot ID column: '{column_mapping['ad_slot_id']}'")
        else:
            print(f"⚠️ Ad slot ID column not found for PSBK line. Available columns: {list(column_mapping.keys())}")
            print(f"🔍 Looking for columns containing 'Ad slot' or 'AD SLOT'...")
            for col in headers:
                if 'ad slot' in col.lower() or 'ad_slot' in col.lower():
                    print(f"   - Found potential Ad slot column: '{col}'")
    
    print(f"\nFound {len(data)} rows in sheet")
    
    # Print first few rows of data to verify structure
    print("\n📝 Sample data rows:")
    for row in data[:3]:
        print(row)
        
    placement_data = {}

    for adtype, filters in adtype_filters.items():
        print(f"\n🔍 Processing adtype: {adtype}")
        adtype_values = filters.get("adtype_filter", []) or filters.get("adtypes", [])
        section_values = filters.get("section_filter", []) or filters.get("sections", [])
        
        # Convert to uppercase for case-insensitive matching
        adtype_values = [ad.upper() for ad in adtype_values]
        section_values = [section.upper() for section in section_values]
        
        print(f"Looking for ad types: {adtype_values}")
        print(f"Looking for sections: {section_values}")
        
        placement_ids = []
        
        for row in data:
            try:
                # Get values using mapped column names
                row_site = str(row.get(column_mapping['site'], '')).upper()
                row_platform = str(row.get(column_mapping['platform'], '')).upper()
                row_section = str(row.get(column_mapping['section'], '')).upper()
                row_adtype = str(row.get(column_mapping['adtype'], '')).upper()
                # For PSBK lines, use Ad slot ID instead of Placement ID
                if line_type == "psbk":
                    row_ad_slot_id = str(row.get(column_mapping.get('ad_slot_id', ''), '')).strip()
                    row_placement = row_ad_slot_id
                    print(f"🔧 PSBK line: Using Ad slot ID {row_ad_slot_id} instead of Placement ID")
                else:
                    row_placement = str(row.get(column_mapping['placement'], '')).strip()

                # Skip row if no placement ID (or Ad slot ID for PSBK)
                if not row_placement:
                    continue

                # Case-insensitive site matching
                site_match = any(site in row_site for site in site_filter)
                
                # Platform-specific checks with case-insensitive matching
                row_platforms = [p.strip() for p in row_platform.split(',')]
                if adtype == "1260x570":
                    platform_match = "WEB" in row_platforms
                elif adtype == "320x480":
                    platform_match = any(p in ["AMP", "MWEB"] for p in row_platforms)
                elif richmedia_platform_map and adtype in richmedia_platform_map:
                    # Use richmedia-specific platform filtering for this size
                    richmedia_platforms = richmedia_platform_map[adtype]
                    platform_match = any(p.strip().upper() in [plat.upper() for plat in row_platforms] for p in richmedia_platforms)
                    print(f"🎯 Richmedia platform check for {adtype}: required {richmedia_platforms}, row has {row_platforms}, match: {platform_match}")
                else:
                    # For standard lines, apply user platform filtering normally
                    # This ensures user's platform choices (like Mweb, AMP only) are respected
                    platform_match = any(p.strip().upper() in [plat.upper() for plat in row_platforms] for p in platforms_filter)
                
                # Case-insensitive section and ad type matching
                section_match = any(section in row_section for section in section_values)
                adtype_match = any(ad in row_adtype for ad in adtype_values)

                if site_match and platform_match and section_match and adtype_match:
                    if row_placement:
                        placement_ids.append(row_placement)
                        if line_type == "psbk":
                            print(f"✅ Found Ad slot ID {row_placement} for {adtype}")
                        else:
                            print(f"✅ Found placement ID {row_placement} for {adtype}")
            except Exception as e:
                print(f"Error processing row: {e}")
                continue

        # Assign base structure
        placement_data[adtype] = {
            "adtype_filter": adtype_values,
            "section_filter": section_values,
            "placement_ids": placement_ids,  # Contains Ad slot IDs for PSBK lines, Placement IDs for others
            "targeting_type": "ad_slot_id" if line_type == "psbk" else "placement_id"
        }
        # Removed results summary for cleaner output
        # print(f"\n📊 Results for {adtype}:")
        # print(f"Found {len(placement_ids)} placement IDs")
        # if placement_ids:
        #     print(f"Sample IDs: {placement_ids[:5]}")

        # Add size overrides if applicable
        if adtype == "1260x570":
            placement_data[adtype]["additional_sizes"] = ['728x500', '1320x570']
        elif adtype == "980x200":
            placement_data[adtype]["additional_sizes"] = ['728x90']
        elif adtype == "320x100":
            placement_data[adtype]["additional_sizes"] = ['320x50']

    return placement_data


# Example usage
if __name__ == "__main__":
    credentials_path = "credentials.json"
    sheet_url = "https://docs.google.com/spreadsheets/d/11_SZJnn5KALr6zi0JA27lKbmQvA1WSK4snp0UTY2AaY/edit?gid=9815366"
    sheet_name = "Languages Placement/Preset"
    site_filter = ["NBT", "VK"]
    platforms_filter = ["WEB", "MWEB", "AMP"]
    adtype_filters = {
        '300x250': {
            'adtype_filter': ['MREC_ALL', 'MREC'],
            'section_filter': ['ROS', 'HP']
        },
        '320x50': {
            'adtype_filter': ['BANNER'],
            'section_filter': ['HP']
        },
        'out-of-page': {
            'adtype_filter': ['SKIN_OOP'],
            'section_filter': ['ROS', 'HP']
        },
        "1260x570": {
            "adtypes": ["INTERSTITIAL"],
            "sections": ["ROS"]
        }
    }

    placement_data = fetch_placements_ids(
        credentials_path,
        sheet_url,
        sheet_name,
        site_filter,
        platforms_filter,
        adtype_filters
    )

    print("\n🚀 Placement Data:", placement_data)
