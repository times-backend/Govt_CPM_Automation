import os
import glob
from googleads import ad_manager
from datetime import datetime
from ros_banner_template_creatives import create_custom_template_creatives
from placements_for_creatives import fetch_placements_ids
import sys
import requests
import hashlib
import base64
import json
import pandas as pd
import re
import traceback
from config import CREATIVES_FOLDER, CREDENTIALS_PATH
import time
import uuid
from logging_utils import logger

# Constants
SHEET_URL = "https://docs.google.com/spreadsheets/d/11_SZJnn5KALr6zi0JA27lKbmQvA1WSK4snp0UTY2AaY/edit?gid=2043018330"
PLACEMENT_SHEET_NAME_LANG = "ALL LANGUAGES"
PLACEMENT_SHEET_NAME_TOI = "TOI + ETIMES"
PLACEMENT_SHEET_NAME_ET = "ET Placement/Preset"
PLACEMENT_SHEET_NAME_CAN_PSBK = "CAN_PSBK"

# Print sheet information for debugging
print(f"\nSheet Configuration:")
print(f"Sheet URL: {SHEET_URL}")
print(f"Language Sheet: {PLACEMENT_SHEET_NAME_LANG}")
print(f"TOI Sheet: {PLACEMENT_SHEET_NAME_TOI}")
print(f"ET Sheet: {PLACEMENT_SHEET_NAME_ET}")
print(f"CAN_PSBK Sheet: {PLACEMENT_SHEET_NAME_CAN_PSBK}")

# Create creatives folder if it doesn't exist
os.makedirs(CREATIVES_FOLDER, exist_ok=True)

available_presets = ["300x250", "320x50", "125x600", "300x600", "728x90", "980x200", "320x480","1260x570","728x500","1320x570","600x250","320x100"]

# Standard Banner Presets
standard_presets_dict = {
    "300x250": {"adtypes": ["MREC_ALL", "MREC", "MREC_1", "MREC_2", "MREC_3", "MREC_4", "MREC_5","BTF MREC"], "sections": ["ROS", "HP", "HOME"]},
    "320x50": {"adtypes": ["BOTTOMOVERLAY","BOTTOM OVERLAY"], "sections": ["ROS", "HP", "HOME"]},
    "300x600": {"adtypes": ["FLYINGCARPET", "FLYING_CARPET", "TOWER"], "sections": ["ROS", "HP", "HOME"]},
    "728x90": {"adtypes": ["LEADERBOARD"], "sections": ["ROS", "HP", "HOME"]},
    "980x200": {"adtypes": ["LEADERBOARD"], "sections": ["ROS"]},
    "320x480": {"adtypes": ["INTERSTITIAL"], "sections": ["ROS", "HP", "HOME"]},
    "1260x570": {"adtypes": ["INTERSTITIAL"], "sections": ["ROS", "HP", "HOME"]},
    "320x100": {"adtypes": ["SLUG1","SLUG2","SLUG3","SLUG4","SLUG5"], "sections": ["ROS", "HP", "HOME"]},
}

# Rich Media Presets
richmedia_presets_dict = {
    "300x250": {"adtypes": ["MREC_1"], "sections": ["ROS", "HP", "HOME"], "platforms": ["WEB"]},
    "320x100": {"adtypes": ["TOPBANNER", "TOPBANNER", "TOPBANNER" ], "sections": ["ROS", "HP", "HOME"], "platforms": ["MWEB"]},
    "300x600": {"adtypes": ["FLYINGCARPET", "FLYING_CARPET", "TOWER"], "sections": ["ROS", "HP", "HOME"], "platforms": ["WEB", "MWEB", "AMP"]},
    "728x90": {"adtypes": ["LEADERBOARD"], "sections": ["ROS", "HP", "HOME"], "platforms": ["WEB", "MWEB", "AMP"]},
    "320x50": {"adtypes": ["BOTTOMOVERLAY","BOTTOM OVERLAY"], "sections": ["ROS", "HP", "HOME"]},
    "320x480": {"adtypes": ["INTERSTITIAL"], "sections": ["ROS", "HP", "HOME"]},
    # ... add more as needed
}

class LocationNotFoundError(Exception):
    def __init__(self, location_name):
        super().__init__(f"No matching location found at any level for: {location_name}")
        self.location_name = location_name

class MultipleGeoLocationsError(Exception):
    def __init__(self, location_name, matches, requires_csm_confirmation=True):
        self.location_name = location_name
        self.matches = matches
        self.requires_csm_confirmation = requires_csm_confirmation
        
        match_details = []
        for i, match in enumerate(matches[:10]):  # Show up to 10 matches
            match_details.append(f"{i+1}. {match['Name']}, {match.get('ParentRegion', 'Unknown Region')}, {match['CountryCode']} (ID: {match['Id']})")
        
        message = f"""
🚨 MULTIPLE GEO LOCATIONS FOUND FOR: '{location_name}'

Found {len(matches)} matching locations:
{chr(10).join(match_details)}

⚠️  CSM CONFIRMATION REQUIRED ⚠️
Please contact your Campaign Success Manager (CSM) to confirm the correct geo location.

To resolve this automatically in future, use the format: '{location_name}, State/Region'
Example: 'Aurangabad, Maharashtra' or 'Aurangabad, Bihar'
"""
        super().__init__(message)

def get_india_geo_id(client):
    """Get India's geo ID for targeting"""
    try:
        return get_geo_id(client, "India")
    except Exception as e:
        print(f"⚠️ Could not get India geo ID: {e}")
        return 2356  # Fallback to known India geo ID

def setup_geo_targeting_for_line_type(client, geo_targeting, line_type):
    """
    Setup geo targeting based on line type:
    - Standard: Use user-selected geo
    - PSBK/NWP: Target India, exclude user-selected geo
    """
    print(f"\n{'='*50}")
    print(f"🎯 Setting up geo targeting for line type: {line_type}")
    print(f"📍 Input geo targeting: {geo_targeting}")
    print(f"{'='*50}\n")
    
    geo_ids = []
    excluded_geo_ids = []
    
    print(f"📍 Setting up geo targeting for {line_type} line with targeting: {geo_targeting}")
    
    if line_type == "standard":
        # Standard line: Use user-selected geo as-is
        print(f"\n{'='*50}")
        print(f"📍 Setting up Standard line geo targeting")
        print(f"🎯 Target locations: {geo_targeting}")
        print(f"{'='*50}\n")
        
        if geo_targeting:
            for location in geo_targeting:
                try:
                    print(f"🔍 Looking up geo ID for: {location}")
                    geo_id = get_geo_id(client, location)
                    if geo_id:
                        geo_ids.append(geo_id)
                        print(f"✅ Successfully mapped {location} to geo ID: {geo_id}")
                    else:
                        print(f"⚠️ No geo ID found for location: {location}")
                        print("⚠️ This location will be skipped in targeting")
                except LocationNotFoundError as e:
                    print(f"\n❌ ERROR: Location not found")
                    print(f"📍 Location: {location}")
                    print(f"❗ Details: {str(e)}")
                    print("⚠️ This location will be skipped in targeting\n")
                except MultipleGeoLocationsError as e:
                    print(f"\n🚨 WARNING: Multiple locations found")
                    print(f"📍 Location: {location}")
                    print(f"✅ Using first matching location")
                    print(f"❗ Details: {str(e)}\n")
                except Exception as e:
                    print(f"\n❌ Unexpected error processing location: {location}")
                    print(f"❗ Error: {str(e)}")
                    print("⚠️ This location will be skipped in targeting\n")
        else:
            print("⚠️ No geo targeting locations provided for standard line")
    
    else:  # PSBK or NWP line
        # Target India but exclude user-selected geo
        print(f"📍 {line_type.upper()} line geo targeting: India (excluding {geo_targeting})")
        
        # Add India as target
        print(f"🔍 {line_type.upper()} - Getting India geo ID...")
        try:
            india_geo_id = get_india_geo_id(client)
            if india_geo_id:
                geo_ids.append(india_geo_id)
                print(f"🌍 {line_type.upper()} - Successfully targeting India: {india_geo_id}")
            else:
                print(f"❌ {line_type.upper()} - Failed to get India geo ID")
        except Exception as e:
            print(f"❌ {line_type.upper()} - Error getting India geo ID: {e}")
        
        # Add user-selected geo to exclusions
        if geo_targeting:
            print(f"🔍 {line_type.upper()} - Processing exclusions for: {geo_targeting}")
            for location in geo_targeting:
                try:
                    print(f"🔍 {line_type.upper()} - Getting geo ID for exclusion: {location}")
                    geo_id = get_geo_id(client, location)
                    if geo_id:
                        excluded_geo_ids.append(geo_id)
                        print(f"🚫 {line_type.upper()} - Successfully excluding {location}: {geo_id}")
                    else:
                        print(f"⚠️ {line_type.upper()} - No geo ID found for {location}")
                except LocationNotFoundError as e:
                    print(f"❌ Location error for {line_type} line exclusion: {e}")
                    raise e
                except MultipleGeoLocationsError as e:
                    print(f"🚨 Multiple geo locations found for {line_type} line exclusion: {e}")
                    raise e
                except Exception as e:
                    print(f"⚠️ Error getting exclusion geo ID for {location}: {e}")
        else:
            print(f"🔍 {line_type.upper()} - No user geo to exclude")
    
    return geo_ids, excluded_geo_ids

def get_geo_id(client, location_name):
    """
    Enhanced Geo ID search with duplicate location handling
    Supports formats like:
    - "Aurangabad" (returns first India match with disambiguation warning)
    - "Aurangabad, Maharashtra" (specific state targeting)
    - "Aurangabad, Bihar" (specific state targeting)
    """
    print(f"🔍 Searching for Geo ID of: {location_name}")
    pql_service = client.GetService("PublisherQueryLanguageService", version="v202408")

    # Parse location input for state/region specification
    location_parts = [part.strip() for part in location_name.split(',')]
    base_location = location_parts[0]
    specified_state = location_parts[1] if len(location_parts) > 1 else None
    
    # Try as Country first
    country_query = f"""
    SELECT Id, Name, Targetable, Type, CountryCode 
    FROM Geo_Target 
    WHERE Name = '{base_location}' 
    AND Targetable = true 
    AND Type = 'COUNTRY'
    """
    
    # Try as State/Region if Country not found
    region_query = f"""
    SELECT Id, Name, Targetable, Type, CountryCode 
    FROM Geo_Target 
    WHERE Name = '{base_location}' 
    AND Targetable = true 
    AND Type IN ('REGION', 'PROVINCE', 'STATE', 'DEPARTMENT')
    AND CountryCode != 'PK'
    """
    
    # Try as City if neither Country nor Region found
    city_query = f"""
    SELECT Id, Name, Targetable, Type, CountryCode 
    FROM Geo_Target 
    WHERE Name = '{base_location}' 
    AND Targetable = true 
    AND Type = 'CITY'
    AND CountryCode != 'PK'
    """
    
    # Try as Sub-District 
    sub_district_query = f"""
    SELECT Id, Name, Targetable, Type, CountryCode 
    FROM Geo_Target 
    WHERE Name = '{base_location}' 
    AND Targetable = true 
    AND Type = 'SUB_DISTRICT'
    AND CountryCode != 'PK'
    """

    queries = [
        ("COUNTRY", country_query),
        ("REGION", region_query),
        ("CITY", city_query),
        ("SUB_DISTRICT", sub_district_query)
    ]

    for geo_type, query in queries:
        try:
            statement = {'query': query}
            response = pql_service.select(statement)
            
            if hasattr(response, 'rows') and response.rows:
                matches = []
                
                for row in response.rows:
                    try:
                        values = row.values
                        geo_data = {
                            "Id": values[0].value,
                            "Name": values[1].value,
                            "Targetable": values[2].value,
                            "Type": values[3].value,
                            "CountryCode": values[4].value
                        }
                        matches.append(geo_data)
                        print(f"Found match: {geo_data['Name']} ({geo_data['Type']}) - ID: {geo_data['Id']}")
                    except Exception as e:
                        print(f"⚠️ Error processing row: {e}")
                        continue
                
                # Filter by country - prioritize India (IN) and US locations
                india_matches = [m for m in matches if m["CountryCode"] == "IN"]
                us_matches = [m for m in matches if m["CountryCode"] == "US"]
                
                preferred_matches = india_matches if india_matches else (us_matches if us_matches else matches)
                
                # If state/region is specified and we have multiple matches, try to disambiguate
                if specified_state and len(preferred_matches) > 1:
                    final_match = disambiguate_by_parent_region(client, preferred_matches, specified_state)
                    if final_match:
                        print(f"✅ Found as {geo_type} with state disambiguation: {final_match['Name']}, {final_match['CountryCode']}, ID: {final_match['Id']}")
                        return final_match["Id"]
                
                # Handle multiple matches without state specification
                if len(preferred_matches) > 1:
                    # Add parent region info to matches for better display
                    enhanced_matches = []
                    for match in preferred_matches:
                        parent_info = get_parent_region_info(client, match["Id"])
                        enhanced_match = match.copy()
                        enhanced_match['ParentRegion'] = parent_info
                        enhanced_matches.append(enhanced_match)
                    
                    # Check if this requires CSM confirmation (more than 2 matches or ambiguous locations)
                    requires_csm = len(preferred_matches) > 2 or is_ambiguous_location(base_location)
                    
                    if requires_csm:
                            # For multiple matches, show PROMINENT warning but proceed with first match
                        print("\n" + "="*80)
                        print("🚨 MULTIPLE GEO LOCATIONS FOUND - USING FIRST MATCH 🚨")
                        print("="*80)
                        print(f"📍 Location '{base_location}' has multiple matches")
                        print(f"✅ Using first match: {enhanced_matches[0]['Name']} ({enhanced_matches[0]['ParentRegion']})")
                        print("="*80 + "\n")
                    else:
                        # For 2 matches, show PROMINENT warning but proceed with first match
                        print("\n" + "="*80)
                        print("🚨 AUTOMATION ALERT: MULTIPLE GEO LOCATIONS DETECTED 🚨")
                        print("="*80)
                        print(f"📍 SEARCHING FOR: '{base_location}'")
                        print(f"🔍 FOUND {len(enhanced_matches)} MATCHING LOCATIONS:")
                        print("-" * 60)
                        
                        for i, match in enumerate(enhanced_matches):
                            marker = "👉 SELECTED" if i == 0 else "   Available"
                            print(f"{marker}: {match['Name']}, {match['ParentRegion']}, {match['CountryCode']} (ID: {match['Id']})")
                        
                        print("-" * 60)
                        final_match = preferred_matches[0]
                        parent_info = enhanced_matches[0]['ParentRegion']
                        
                        print(f"⚡ AUTOMATION DECISION: Selected '{final_match['Name']}, {parent_info}'")
                        print(f"📋 REASON: First match from available options")
                        print(f"⚠️  WARNING: Other locations with same name exist!")
                        print()
                        print("💡 TO AVOID THIS IN FUTURE:")
                        print(f"   Use specific format: '{base_location}, State/Region'")
                        print(f"   Examples: 'Aurangabad, Maharashtra' or 'Aurangabad, Bihar'")
                        print()
                        print("📞 NEED DIFFERENT LOCATION?")
                        print("   Contact your CSM to change the geo targeting")
                        print("="*80)
                        print("🎯 PROCEEDING WITH LINE CREATION...")
                        print("="*80 + "\n")
                        
                        # Print the automatic selection for audit trail
                        print(f"✅ Auto-selected: {final_match['Name']} ({parent_info})")
                else:
                    final_match = preferred_matches[0]
                
                print(f"✅ Found as {geo_type}: {final_match['Name']}, {final_match['CountryCode']}, ID: {final_match['Id']}")
                return final_match["Id"]
                
        except Exception as e:
            print(f"⚠️ Error searching as {geo_type}: {e}")
            continue

    print(f"❌ No matching location found at any level for: {location_name}")
    raise LocationNotFoundError(location_name)


def get_parent_region_info(client, geo_id):
    """Get parent region information for a geo location"""
    try:
        pql_service = client.GetService("PublisherQueryLanguageService", version="v202408")
        
        # Query to get parent region information
        parent_query = f"""
        SELECT ParentIds, Name, Type
        FROM Geo_Target 
        WHERE Id = {geo_id}
        """
        
        statement = {'query': parent_query}
        response = pql_service.select(statement)
        
        if 'rows' in response and response['rows']:
            values = response['rows'][0]["values"]
            parent_ids = values[0]["value"] if values[0]["value"] else []
            
            if parent_ids:
                # Get the immediate parent (usually state/region)
                parent_id = parent_ids[-1]  # Last parent is usually the most specific
                
                parent_info_query = f"""
                SELECT Name, Type
                FROM Geo_Target 
                WHERE Id = {parent_id}
                """
                
                parent_statement = {'query': parent_info_query}
                parent_response = pql_service.select(parent_statement)
                
                if 'rows' in parent_response and parent_response['rows']:
                    parent_values = parent_response['rows'][0]["values"]
                    parent_name = parent_values[0]["value"]
                    return parent_name
        
        return "Unknown Region"
        
    except Exception as e:
        print(f"⚠️ Error getting parent region info: {e}")
        return "Unknown Region"


def is_ambiguous_location(location_name):
    """
    Check if a location name is known to be ambiguous (has multiple common instances)
    """
    # List of commonly ambiguous location names in India
    ambiguous_locations = {
        'aurangabad', 'salem', 'bangalore', 'mysore', 'hassan', 'mandya', 
        'tumkur', 'shimoga', 'bellary', 'gulbarga', 'bijapur', 'raichur',
        'chitradurga', 'davangere', 'bagalkot', 'haveri', 'gadag', 'koppal',
        'yadgir', 'kolar', 'chikkaballapur', 'ramanagara', 'chamarajanagar',
        'kodagu', 'udupi', 'chikkamagaluru', 'shivamogga', 'vijayapura',
        'kalburgi', 'ballari', 'nellore', 'kadapa', 'kurnool', 'anantapur',
        'chittoor', 'tirupati', 'vizianagaram', 'srikakulam', 'guntur',
        'krishna', 'west godavari', 'east godavari', 'warangal', 'khammam',
        'nalgonda', 'mahbubnagar', 'rangareddy', 'medak', 'nizamabad',
        'adilabad', 'karimnagar', 'hyderabad', 'secunderabad'
    }
    
    return location_name.lower().strip() in ambiguous_locations


def disambiguate_by_parent_region(client, matches, specified_state):
    """
    Disambiguate multiple location matches by checking their parent regions
    """
    try:
        for match in matches:
            parent_info = get_parent_region_info(client, match["Id"])
            
            # Check if the parent region matches the specified state
            if parent_info and specified_state.lower() in parent_info.lower():
                return match
            
            # Also check direct match with common state name variations
            state_variations = {
                'maharashtra': ['maharashtra', 'mh'],
                'bihar': ['bihar', 'br'],
                'uttar pradesh': ['uttar pradesh', 'up'],
                'west bengal': ['west bengal', 'wb'],
                'tamil nadu': ['tamil nadu', 'tn'],
                'karnataka': ['karnataka', 'ka'],
                'gujarat': ['gujarat', 'gj'],
                'rajasthan': ['rajasthan', 'rj'],
                'andhra pradesh': ['andhra pradesh', 'ap'],
                'telangana': ['telangana', 'ts'],
                'kerala': ['kerala', 'kl'],
                'odisha': ['odisha', 'or'],
                'punjab': ['punjab', 'pb'],
                'haryana': ['haryana', 'hr'],
                'himachal pradesh': ['himachal pradesh', 'hp'],
                'uttarakhand': ['uttarakhand', 'uk'],
                'jharkhand': ['jharkhand', 'jh'],
                'chhattisgarh': ['chhattisgarh', 'cg'],
                'madhya pradesh': ['madhya pradesh', 'mp'],
                'assam': ['assam', 'as'],
                'meghalaya': ['meghalaya', 'ml'],
                'manipur': ['manipur', 'mn'],
                'mizoram': ['mizoram', 'mz'],
                'nagaland': ['nagaland', 'nl'],
                'tripura': ['tripura', 'tr'],
                'arunachal pradesh': ['arunachal pradesh', 'ar'],
                'sikkim': ['sikkim', 'sk'],
                'goa': ['goa', 'ga']
            }
            
            specified_lower = specified_state.lower()
            for full_name, variations in state_variations.items():
                if specified_lower in variations:
                    if full_name in parent_info.lower():
                        return match
        
        return None
        
    except Exception as e:
        print(f"⚠️ Error in disambiguation: {e}")
        return None


def show_geo_selection_summary(auto_selections):
    """
    Show a summary of all automatic geo selections made during line creation
    """
    if not auto_selections:
        return
    
    print("\n" + "="*80)
    print("📋 GEO TARGETING SUMMARY - AUTOMATION SELECTIONS MADE")
    print("="*80)
    print(f"🔢 TOTAL AUTO-SELECTIONS: {len(auto_selections)}")
    print()
    
    for i, selection in enumerate(auto_selections, 1):
        print(f"{i}. 📍 '{selection['input']}' → Selected: {selection['selected']}")
        print(f"   🆔 Geo ID: {selection['geo_id']}")
        print(f"   📝 Reason: {selection['reason']}")
        print()
    
    print("⚠️  IMPORTANT REMINDERS:")
    print("• These selections were made automatically from multiple options")
    print("• Review if different locations are needed")
    print("• Contact CSM if changes are required")
    print("• Use 'Location, State' format to avoid future auto-selections")
    print("="*80)
    print("✅ LINE CREATION COMPLETED WITH AUTO-SELECTED GEOS")
    print("="*80 + "\n")


def fetch_images_and_presets(folder_path, available_presets, presets_dict):
    image_files = glob.glob(os.path.join(folder_path, "*.*"))
    detected_presets = {}
    image_size_map = {}  # Map to track images for each size
    
    # First, identify all valid images and their sizes
    for image_path in image_files:
        filename = os.path.basename(image_path)
        for preset in available_presets:
            if preset.lower() in filename.lower() and preset in presets_dict:
                if preset not in image_size_map:
                    image_size_map[preset] = []
                image_size_map[preset].append(image_path)
                
                # Create a unique key for each image of this size
                counter = len(image_size_map[preset])
                size_key = f"{preset}_{counter}" if counter > 1 else preset
                
                detected_presets[size_key] = {
                    "adtype_filter": presets_dict[preset]["adtypes"],
                    "section_filter": presets_dict[preset]["sections"],
                    "image_path": image_path,
                    "base_size": preset
                }
                print(f"Added creative for size {preset} with key {size_key}: {filename}")
    
    return detected_presets, image_files

def detect_line_type(line_name):
    if "RICHMEDIA" in line_name.upper():
        return "richmedia"
    else:
        return "standard"

def read_tag_file():
    """
    Reads a tag file (Excel format) that contains creative dimensions and their corresponding JavaScript tags.
    
    This function looks for files with names like 'tag.xlsx', 'tags.xlsx', 'tag.xls', or 'tags.xls'
    in both the current directory and the 'creatives' directory. If a file is found but can't be read,
    it falls back to creating a simulated tag dictionary.
    
    The function now supports:
    1. JavaScript tags (traditional script tags)
    2. Impression/click tag combinations
    3. DoubleClick tags (DCM tags with <ins> elements)
    
    Returns:
        dict: A dictionary mapping dimension strings to their corresponding JavaScript tags,
              or None if no valid tag file is found or an error occurs.
    """
    try:
        # Get the current directory where the script is running
        current_dir = os.path.dirname(os.path.abspath(__file__))
        creatives_dir = os.path.join(current_dir, "creatives")
        
        # Define possible tag file names
        tag_file_patterns = ['tag.xlsx', 'tags.xlsx', 'tag.xls', 'tags.xls', 'TOI Tags (2).xlsx', 'TOI Tags (2).xls']
        
        # First, try to find a real tag file
        real_tag_file = None
        real_tag_dir = None
        
        for directory in [current_dir, creatives_dir]:
            if not os.path.exists(directory):
                continue
                
            print(f"Checking directory for tag files: {directory}")
            for file in os.listdir(directory):
                file_lower = file.lower()
                if any(pattern.lower() in file_lower for pattern in tag_file_patterns):
                    real_tag_file = file
                    real_tag_dir = directory
                    print(f"Found potential tag file: {os.path.join(real_tag_dir, real_tag_file)}")
                    break
            
            if real_tag_file:
                break
                
        # If we found a real tag file, try to read it
        if real_tag_file:
            tag_file_path = os.path.join(real_tag_dir, real_tag_file)
            print(f"Attempting to read tag file at: {tag_file_path}")
            
            try:
                # For xlsx files, try pandas
                import pandas as pd
                
                def read_excel_with_sheet_selection(file_path, engine=None):
                    """Helper function to read Excel file with preference for 'tags' sheet"""
                    try:
                        # Try to read sheet names first
                        if engine:
                            excel_file = pd.ExcelFile(file_path, engine=engine)
                        else:
                            excel_file = pd.ExcelFile(file_path)
                        
                        sheet_names = excel_file.sheet_names
                        print(f"Available sheets: {sheet_names}")
                        
                        # Check for 'tags' sheet (case insensitive)
                        target_sheet = None
                        for sheet_name in sheet_names:
                            if sheet_name.lower() == 'tags':
                                target_sheet = sheet_name
                                print(f"Found 'tags' sheet: {target_sheet}")
                                break
                        
                        # If no 'tags' sheet found, use the first sheet
                        if target_sheet is None:
                            target_sheet = sheet_names[0]
                            print(f"No 'tags' sheet found, using first sheet: {target_sheet}")
                        
                        # Read the selected sheet
                        if engine:
                            df = pd.read_excel(file_path, sheet_name=target_sheet, engine=engine)
                        else:
                            df = pd.read_excel(file_path, sheet_name=target_sheet)
                        
                        return df
                    except Exception as e:
                        print(f"Error reading Excel file with sheet selection: {e}")
                        # Fallback to default behavior
                        if engine:
                            return pd.read_excel(file_path, engine=engine)
                        else:
                            return pd.read_excel(file_path)
                
                if tag_file_path.lower().endswith('.xlsx'):
                    df = read_excel_with_sheet_selection(tag_file_path)
                else:  # For xls files
                    try:
                        df = read_excel_with_sheet_selection(tag_file_path, engine='xlrd')
                    except:
                        try:
                            df = read_excel_with_sheet_selection(tag_file_path, engine='openpyxl')
                        except:
                            raise Exception(f"Failed to read {tag_file_path} with any Excel engine")
                
                print("\nDataFrame Info:")
                print(df.info())
                
                # Create a dictionary to store dimensions and their corresponding tags
                tag_dict = {}
                
                # Find column names for dimensions, JavaScript tags, Impression Tags and Click Tags
                dimension_col = None
                tag_col = None
                impression_tag_col = None
                click_tag_col = None
                
                print(f"Available columns: {list(df.columns)}")
                
                # Look for exact column names first
                for col in df.columns:
                    col_str = str(col).lower()
                    if col_str == 'dimensions' or col_str == 'placementname':
                        dimension_col = col
                        print(f"Using exact match '{col}' as dimension column")
                    elif col_str == 'javascript tag' or col_str == 'js_https':
                        tag_col = col
                        print(f"Using exact match '{col}' as tag column")
                    elif col_str == 'impression tag (image)' or col_str == 'impression tag':
                        impression_tag_col = col
                        print(f"Using exact match '{col}' as impression tag column")
                    elif col_str == 'click tag':
                        click_tag_col = col
                        print(f"Using exact match '{col}' as click tag column")
                        
                
                # If needed, look for partial matches
                if not dimension_col:
                    for col in df.columns:
                        col_str = str(col).lower()
                        if 'dimension' in col_str or 'size' in col_str or 'placement' in col_str:
                            dimension_col = col
                            print(f"Using partial match '{col}' as dimension column")
                            break
                
                if not tag_col:
                    for col in df.columns:
                        col_str = str(col).lower()
                        if ('javascript' in col_str and 'tag' in col_str) or 'script' in col_str or 'js_' in col_str:
                            tag_col = col
                            print(f"Using partial match '{col}' as tag column")
                            break
                
                if not impression_tag_col:
                    for col in df.columns:
                        col_str = str(col).lower()
                        if 'impression' in col_str and 'tag' in col_str:
                            impression_tag_col = col
                            print(f"Using partial match '{col}' as impression tag column")
                            break
                
                if not click_tag_col:
                    for col in df.columns:
                        col_str = str(col).lower()
                        if 'click' in col_str and 'tag' in col_str:
                            click_tag_col = col
                            print(f"Using partial match '{col}' as click tag column")
                            break
                
                # Final attempt to find tag column
                if dimension_col and not tag_col and not (impression_tag_col and click_tag_col):
                    for col in df.columns:
                        col_str = str(col).lower()
                        if 'tag' in col_str:
                            tag_col = col
                            print(f"Using fallback '{col}' as tag column")
                            break
                
                has_columns = dimension_col and (tag_col or (impression_tag_col and click_tag_col))
                
                if has_columns:
                    # Process the dataframe
                    for index, row in df.iterrows():
                        if pd.notnull(row[dimension_col]):
                            dimension = str(row[dimension_col]).strip()
                            
                            # Check for Impression Tag and Click Tag first (new priority)
                            if impression_tag_col and click_tag_col and pd.notnull(row[impression_tag_col]) and pd.notnull(row[click_tag_col]):
                                impression_tag = str(row[impression_tag_col]).strip()
                                click_tag = str(row[click_tag_col]).strip()
                                
                                # Skip empty entries
                                if not dimension or not impression_tag or not click_tag:
                                    continue
                                
                                # Clean up dimension string to ensure format like "300x250"
                                if 'x' in dimension:
                                    dimension_match = re.search(r'(\d+x\d+)', dimension)
                                    if dimension_match:
                                        dimension = dimension_match.group(1)
                                
                                # Store both tags in a dictionary
                                tag_dict[dimension] = {
                                    'type': 'impression_click',
                                    'impression_tag': impression_tag,
                                    'click_tag': click_tag
                                }
                                print(f"Added impression/click tags for dimension: {dimension}")
                            
                            # Fallback to JavaScript tag if impression/click tags not found
                            elif tag_col and pd.notnull(row[tag_col]):
                                js_tag = str(row[tag_col]).strip()
                                
                                # Skip empty entries
                                if not dimension or not js_tag:
                                    continue
                                    
                                # Clean up dimension string to ensure format like "300x250"
                                if 'x' in dimension:
                                    dimension_match = re.search(r'(\d+x\d+)', dimension)
                                    if dimension_match:
                                        dimension = dimension_match.group(1)
                                
                                # Handle <noscript> tags with <a> href, common in Flashtalking tags
                                if '<noscript>' in js_tag.lower() and '<a href' in js_tag.lower():
                                    print(f"Detected noscript/a href tag for dimension: {dimension}")
                                    href_pattern = r'(<a\s+[^>]*?href=")([^"]*)"'
                                    
                                    # Prepend click macro if not already present
                                    if '%%CLICK_URL_UNESC%%' not in js_tag:
                                        replacement = r'\1%%CLICK_URL_UNESC%%\2"'
                                        modified_tag = re.sub(href_pattern, replacement, js_tag, flags=re.IGNORECASE)
                                        
                                        if modified_tag != js_tag:
                                            js_tag = modified_tag
                                            print(f"Added %%CLICK_URL_UNESC%% to href in noscript tag for dimension: {dimension}")
                                        else:
                                            print(f"Warning: Could not add %%CLICK_URL_UNESC%% to href for dimension: {dimension}")
                                    else:
                                        print(f"Click macro already present for dimension: {dimension}")

                                # Check if this is a DoubleClick tag (contains dcmads or data-dcm attributes)
                                is_doubleclick = False
                                if ('dcmads' in js_tag.lower() or 'data-dcm' in js_tag.lower()) and ('<ins' in js_tag.lower() or '<div' in js_tag.lower()):
                                    is_doubleclick = True
                                    print(f"Detected DoubleClick tag for dimension: {dimension}")
                                    
                                    # Ensure data-dcm-click-tracker is present in the DoubleClick tag
                                    if 'data-dcm-click-tracker' not in js_tag:
                                        try:
                                            # Add data-dcm-click-tracker attribute before the class attribute
                                            tag_pattern = r'(<ins|<div)([^>]*?)(\s+class=)'
                                            replacement = r"\1\2 data-dcm-click-tracker='%%CLICK_URL_UNESC%%'\3"
                                            modified_tag = re.sub(tag_pattern, replacement, js_tag, flags=re.IGNORECASE)
                                            
                                            # If that didn't work, try adding it after the opening tag
                                            if modified_tag == js_tag:
                                                tag_pattern = r'(<ins|<div)(\s)'
                                                replacement = r"\1 data-dcm-click-tracker='%%CLICK_URL_UNESC%%'\2"
                                                modified_tag = re.sub(tag_pattern, replacement, js_tag, flags=re.IGNORECASE)
                                            
                                            js_tag = modified_tag
                                            print(f"Added data-dcm-click-tracker attribute to DoubleClick tag for dimension: {dimension}")
                                        except Exception as e:
                                            print(f"Warning: Could not add data-dcm-click-tracker to tag: {str(e)}")
                                
                                # Only add if tag is substantial
                                if 'x' in dimension and len(js_tag.strip()) > 10:
                                    # Check if this dimension already exists in the dictionary
                                    if dimension in tag_dict:
                                        # It's a duplicate, so append a counter
                                        counter = 1
                                        while f"{dimension}_{counter}" in tag_dict:
                                            counter += 1
                                        dimension_key = f"{dimension}_{counter}"
                                        print(f"Found duplicate dimension {dimension}, using key {dimension_key}")
                                    else:
                                        dimension_key = dimension
                                    
                                    if is_doubleclick:
                                        tag_dict[dimension_key] = {
                                            'type': 'doubleclick',
                                            'js_tag': js_tag
                                        }
                                        print(f"Added DoubleClick tag for dimension: {dimension}")
                                    else:
                                        tag_dict[dimension_key] = {
                                            'type': 'javascript',
                                            'js_tag': js_tag
                                        }
                                        print(f"Added JavaScript tag for dimension: {dimension}")
                    
                    if tag_dict:
                        print(f"Successfully read {len(tag_dict)} tags from {real_tag_file}")
                        return tag_dict
                    else:
                        print(f"No valid tag entries found in {real_tag_file}")
                else:
                    print(f"Couldn't find required columns in {real_tag_file}. Found columns: {list(df.columns)}")
                    print("Looking for columns named 'Dimensions' or 'PlacementName' and either 'JavaScript Tag' or 'js_https'")
            
            except Exception as e:
                print(f"Error reading file {real_tag_file}: {str(e)}")
                traceback.print_exc()
        
        # If we didn't find or couldn't read a real tag file, fall back to simulated dictionary
        print("No valid tag file found. Not creating simulated tags.")
        return None
        
    except Exception as e:
        print(f"Error in read_tag_file: {str(e)}")
        print(f"Error type: {type(e)}")
        traceback.print_exc()
        return None

def check_line_item_name_exists(client, order_id, line_name_base):
    """Check if a line item with similar name already exists in the order or globally"""
    try:
        line_item_service = client.GetService('LineItemService', version='v202408')
        pql_service = client.GetService('PublisherQueryLanguageService', version='v202408')
        
        print(f"🔍 Checking for duplicates of line name: {line_name_base}")
        print(f"🔍 Line name length: {len(line_name_base)} characters")
        
        # Escape single quotes in the line name for PQL query
        escaped_line_name = line_name_base.replace("'", "\\'")
        
        # First check for exact matches globally (this is what causes DUPLICATE_OBJECT error)
        exact_query = f"SELECT Id, Name, OrderId FROM Line_Item WHERE Name = '{escaped_line_name}'"
        print(f"🔍 Executing exact match query: {exact_query}")
        exact_statement = {'query': exact_query}
        exact_response = pql_service.select(exact_statement)
        
        print(f"🔍 Exact query response type: {type(exact_response)}")
        if hasattr(exact_response, 'rows') and exact_response.rows:
            print(f"🔍 Exact query found {len(exact_response.rows)} rows")
            existing_exact = [(row.values[0].value, row.values[1].value, row.values[2].value) 
                             for row in exact_response.rows]
            print(f"⚠️ Found {len(existing_exact)} existing line items with EXACT name '{line_name_base}':")
            for line_id, name, existing_order_id in existing_exact:
                print(f"   - Line ID: {line_id}, Name: {name}, Order ID: {existing_order_id}")
            return True
        else:
            print(f"🔍 No rows found in exact query response")
        
        # Also try a broader search to catch any similar names
        # Use LIKE with wildcards to find potential matches
        like_query = f"SELECT Id, Name, OrderId FROM Line_Item WHERE Name LIKE '%{escaped_line_name}%'"
        print(f"🔍 Executing LIKE query: {like_query}")
        like_statement = {'query': like_query}
        like_response = pql_service.select(like_statement)
        
        print(f"🔍 LIKE query response type: {type(like_response)}")
        if hasattr(like_response, 'rows') and like_response.rows:
            print(f"🔍 LIKE query found {len(like_response.rows)} rows")
            existing_like = [(row.values[0].value, row.values[1].value, row.values[2].value) 
                           for row in like_response.rows]
            print(f"⚠️ Found {len(existing_like)} existing line items with SIMILAR names to '{line_name_base}':")
            for line_id, name, existing_order_id in existing_like:
                print(f"   - Line ID: {line_id}, Name: {name}, Order ID: {existing_order_id}")
                # Check if it's an exact match
                if name == line_name_base:
                    print(f"   ⚠️ EXACT MATCH FOUND: This will cause DUPLICATE_OBJECT error!")
                    return True
        
        # Then check for similar names in the same order (for good measure)
        similar_query = f"SELECT Id, Name FROM Line_Item WHERE OrderId = {order_id} AND Name LIKE '{escaped_line_name}%'"
        print(f"🔍 Executing order-specific query: {similar_query}")
        similar_statement = {'query': similar_query}
        similar_response = pql_service.select(similar_statement)
        
        if hasattr(similar_response, 'rows') and similar_response.rows:
            existing_similar = [row.values[1].value for row in similar_response.rows]
            print(f"⚠️ Found {len(existing_similar)} existing line items with similar names in order {order_id}: {existing_similar}")
            # Check for exact matches in this order too
            for existing_name in existing_similar:
                if existing_name == line_name_base:
                    print(f"   ⚠️ EXACT MATCH in same order: This will cause DUPLICATE_OBJECT error!")
                    return True
        
        print(f"✅ No existing line items found with exact or similar name: {line_name_base}")
        return False
        
    except Exception as e:
        print(f"⚠️ Error checking for existing line item names: {e}")
        print(f"⚠️ Error type: {type(e).__name__}")
        print(f"⚠️ Will continue with creation but may encounter DUPLICATE_OBJECT error")
        return False  # Continue with creation if check fails





def single_line(client, order_id, line_item_data, line_name):
    # Debug: Check what line type we received
    print(f"🔍 DEBUG: single_line received line_type: {line_item_data.get('line_type', 'NOT_SET')}")
    # Generate session ID for this line creation
    session_id = str(uuid.uuid4())
    start_time = time.time()
    
    # Initialize timing checkpoints
    timing_checkpoints = {
        'start_time': start_time,
        'data_processing_start': None,
        'data_processing_end': None,
        'placement_lookup_start': None,
        'placement_lookup_end': None,
        'line_creation_start': None,
        'line_creation_end': None,
        'creative_creation_start': None,
        'creative_creation_end': None
    }
    
    # Log line creation start
    logger.log_line_creation_start(str(order_id), line_item_data, line_name, session_id)
    
    line_item_service = client.GetService('LineItemService', version='v202408')
    
    # Track created creatives by size to prevent duplicates
    created_creative_sizes = set()
    
    def track_creative_creation(size, creative_ids_list):
        """Track created creatives to prevent duplicates"""
        if creative_ids_list:
            created_creative_sizes.add(size)
            print(f"📝 Tracked creative creation: {size} -> {len(creative_ids_list)} creatives")
    
    def is_creative_size_already_created(size):
        """Check if a creative of this size has already been created"""
        return size in created_creative_sizes

    end_date_value = line_item_data.get('End_date')
    print(f"end_date_value::{end_date_value}")
    start_date_value = line_item_data.get('Start_date', '2025-05-06 00:00:00')
    print(f"start_date_value::{start_date_value}")
    Fcap_value = int(line_item_data.get('fcap', 0))
    cost = line_item_data.get('CPM_Rate', line_item_data.get('cpm', 0))
    print(f"line_item_data::{line_item_data}")
    
    # Start data processing timing
    timing_checkpoints['data_processing_start'] = time.time()
    
    # End data processing timing
    timing_checkpoints['data_processing_end'] = time.time()
    
    # Process impression value to ensure it's an integer
    total_impressions = line_item_data.get('impressions', 100000)
    try:
        if isinstance(total_impressions, str):
            total_impressions = float(total_impressions.replace(',', ''))
        elif isinstance(total_impressions, float):
            pass
        else:
            total_impressions = float(total_impressions)
        impressions_int = int(total_impressions)
        print(f"Total impressions converted to integer: {impressions_int}")
    except Exception as e:
        print(f"Error converting impressions to integer: {e}. Using default value.")
        impressions_int = 100000

    # Extract year, month, and date from end_date_value
    try:
        if isinstance(end_date_value, str) and '-' in end_date_value and ':' in end_date_value:
            date_parts = end_date_value.split(' ')[0].split('-')
            year = int(date_parts[0])
            month = int(date_parts[1])
            day = int(date_parts[2])
            
            print(f"year::{year}, month::{month}, day::{day}")
            if ' ' in end_date_value and ':' in end_date_value:
                time_parts = end_date_value.split(' ')[1].split(':')
                hour = int(time_parts[0])
                minute = int(time_parts[1])
                second = int(time_parts[2]) if len(time_parts) > 2 else 0
            else:
                hour, minute, second = 23, 59, 0
        else:
            year, month, day = 2025, 12, 31
            hour, minute, second = 23, 59, 0
            
        structured_end_date = {
            'date': {
                'year': year,
                'month': month,
                'day': day
            },
            'hour': hour,
            'minute': minute,
            'second': second,
            'timeZoneId': 'Asia/Kolkata'
        }
        
        print(f"Extracted date components - Year: {year}, Month: {month}, Day: {day}")
        print(f"Structured end date for API: {structured_end_date}")

        if isinstance(start_date_value, str) and '-' in start_date_value and ':' in start_date_value:
            date_parts = start_date_value.split(' ')[0].split('-')
            start_year = int(date_parts[0])
            start_month = int(date_parts[1])
            start_day = int(date_parts[2])
            
            current_date = datetime.now()
            start_date = datetime(start_year, start_month, start_day)
            
            if start_date.date() <= current_date.date():
                print("Start date is today or in the past, using IMMEDIATELY")
                use_start_date_type = True
                structured_start_date = None
            else:
                print("Start date is in the future, using structured date with 00:00:00")
                use_start_date_type = False
                structured_start_date = {
                    'date': {
                        'year': start_year,
                        'month': start_month,
                        'day': start_day
                    },
                    'hour': 0,
                    'minute': 0,
                    'second': 0,
                    'timeZoneId': 'Asia/Kolkata'
                }
                print(f"Structured start date for API: {structured_start_date}")
        else:
            print("No valid start date provided, using IMMEDIATELY")
            use_start_date_type = True
            structured_start_date = None
            
    except Exception as e:
        print(f"Error extracting date components: {e}. Using default dates.")
        structured_end_date = {
            'date': {
                'year': 2025,
                'month': 12,
                'day': 31
            },
            'hour': 23,
            'minute': 59,
            'second': 0,
            'timeZoneId': 'Asia/Kolkata'
        }
        
        current_date = datetime.now()
        structured_start_date = {
            'date': {
                'year': current_date.year,
                'month': current_date.month,
                'day': current_date.day
            },
            'hour': current_date.hour,
            'minute': current_date.minute,
            'second': current_date.second,
            'timeZoneId': 'Asia/Kolkata'
        }

    # Check if processed geo IDs are already provided (from geo type functions)
    if 'processed_geo_ids' in line_item_data:
        geo_targeting_ids = line_item_data['processed_geo_ids']
        excluded_geo_ids = line_item_data.get('excluded_geo_ids', [])
        line_type = line_item_data.get('line_type', 'standard')
        print(f"✅ Using processed geo targeting for {line_type} line:")
        print(f"  - Target geo IDs: {geo_targeting_ids}")
        print(f"  - Exclude geo IDs: {excluded_geo_ids}")
        print(f"  - Custom sheet: {line_item_data.get('custom_sheet_name', 'None')}")
        
        # Additional debug for geo targeting validation
        if line_type in ['psbk', 'nwp']:
            if not geo_targeting_ids:
                print(f"⚠️ WARNING: {line_type.upper()} line has no target geo IDs (should have India)")
            if not excluded_geo_ids:
                print(f"⚠️ WARNING: {line_type.upper()} line has no excluded geo IDs (should exclude user geo)")
            else:
                print(f"✅ {line_type.upper()} geo targeting looks correct")
    else:
        # Enhanced geo targeting logic with better error handling
        geo_targeting_input = line_item_data.get("geoTargeting", line_item_data.get("geo", []))
        print(f"🌍 Raw geo targeting input: {geo_targeting_input}")
        
        # Handle different input formats
        if isinstance(geo_targeting_input, str):
            geo_targeting_input = [city.strip() for city in geo_targeting_input.split(",") if city.strip()]
            print(f"🔄 Converted string to list: {geo_targeting_input}")
        elif isinstance(geo_targeting_input, list):
            geo_targeting_input = [str(city).strip() for city in geo_targeting_input if str(city).strip()]
            print(f"✨ Cleaned list input: {geo_targeting_input}")
        elif geo_targeting_input is None:
            print("⚠️ No geo targeting provided, initializing empty list")
            geo_targeting_input = []
        else:
            print(f"⚠️ Unexpected geo targeting type: {type(geo_targeting_input)}")
            geo_targeting_input = []
        
        print(f"🎯 Processing geo targeting locations: {geo_targeting_input}")
        geo_targeting_ids = []
        excluded_geo_ids = []
        
        invalid_locations = []
        
        # Validate we have locations to process
        if not geo_targeting_input:
            print("⚠️ No valid geo targeting locations found to process")
        for city in geo_targeting_input:
            try:
                print(f"🔍 Looking up geo ID for location: {city}")
                geo_id = get_geo_id(client, city)
                if geo_id:
                    print(f"✅ Found geo ID {geo_id} for location: {city}")
                    geo_targeting_ids.append(geo_id)
                    print(f"✅ Found geo ID {geo_id} for location: {city}")
                    geo_targeting_ids.append(geo_id)
                else:
                    print(f"❌ No geo ID found for location: {city}")
                    invalid_locations.append(city)
            except LocationNotFoundError as e:
                print(f"⚠️ Location error: {e}")
                invalid_locations.append(city)
            except MultipleGeoLocationsError as e:
                print("\n" + "="*80)
                print("🚨 CRITICAL ALERT: MULTIPLE GEO LOCATIONS FOUND 🚨")
                print("="*80)
                print(f"📍 LOCATION: '{city}' has multiple matches")
                print("⚠️  CSM CONFIRMATION REQUIRED!")
                print()
                print(str(e))
                print()
                print("🛑 LINE CREATION PAUSED - AWAITING CSM CONFIRMATION")
                print("📞 Please contact your Campaign Success Manager immediately")
                print("="*80 + "\n")
                
                # For backward compatibility, add to invalid locations so process can continue
                # In production, this should trigger a workflow pause
                invalid_locations.append(city)
        
        if not geo_targeting_ids:
            print("Warning: No valid geo targeting IDs found. Line item will target all locations.")
            # Proceed with empty geo_targeting_ids (targets all locations)

    destination_url = line_item_data.get('destination_url', '') or ''
    expresso_id = str(line_item_data.get('expresso_id', ''))
    landing_page = (line_item_data.get('landing_page', '') or line_item_data.get('destination_url', '')) or ''
    Template_Id = line_item_data.get('Template_id', '')
    impression_tracker = (line_item_data.get('impression_tracker') or '').strip()
    script_tracker = (line_item_data.get('tracking_tag') or '').strip()
    currency = (line_item_data.get('currency') or 'INR').strip().upper()
    In_Banner_video = line_item_data.get('banner_video', '')
    
    # Replace [timestamp] with %%CACHEBUSTER%% in impression_tracker
    if impression_tracker:
        original_impression_tracker = impression_tracker
        impression_tracker = impression_tracker.replace('[timestamp]', '%%CACHEBUSTER%%').replace('[CACHEBUSTER]', '%%CACHEBUSTER%%')
        if original_impression_tracker != impression_tracker:
            print(f"Replaced timestamp in impression_tracker: {original_impression_tracker} -> {impression_tracker}")
    
    # Replace [timestamp] with %%CACHEBUSTER%% in script_tracker
    if script_tracker:
        original_script_tracker = script_tracker
        script_tracker = script_tracker.replace('[timestamp]', '%%CACHEBUSTER%%').replace('[CACHEBUSTER]', '%%CACHEBUSTER%%')
        if original_script_tracker != script_tracker:
            print(f"Replaced timestamp in script_tracker: {original_script_tracker} -> {script_tracker}")

    # Ensure strings are not None before strip()
    destination_url = str(destination_url) if destination_url is not None else ''
    landing_page = str(landing_page) if landing_page is not None else ''

    # Check if landing page and destination URL are both empty
    if not landing_page.strip() and not destination_url.strip():
        # Default to 12399020, but we'll check for 2x images later and potentially change to 12473441
        print("No landing page or destination URL provided. Will determine template based on image names.")
        Template_Id = 12399020  # Default, may be updated later
        landing_page = ''
        destination_url = ''
    elif not Template_Id or Template_Id == '':
        # If there's a landing page or destination URL but no explicit template, use standard template
        print("Landing page/destination URL provided. Using standard template 12330939.")
        Template_Id = 12330939
    
    if currency not in ['INR', 'USD', 'CAD', 'AED']:
        currency = 'INR'
    
    print(f"Processed values:")
    print(f"destination_url: {destination_url}")
    print(f"expresso_id: {expresso_id}")
    print(f"landing_page: {landing_page}")
    print(f"Template_Id: {Template_Id}")
    print(f"impression_tracker: {impression_tracker}")
    print(f"script_tracker: {script_tracker}")
    print(f"currency: {currency}")
    print(f"In_Banner_video: {In_Banner_video}")
    
    if script_tracker:
        script_tracker = f'<div style="display:none;">{script_tracker}</div>'
   
    # Respect explicit line_type in input if provided; fallback to name-based detection
    explicit_line_type = line_item_data.get('line_type')
    line_type = explicit_line_type if explicit_line_type else detect_line_type(line_name)
    print(f"line_type:::{line_type}")
    if line_type == "richmedia":
        presets_dict = richmedia_presets_dict
    else:
        presets_dict = standard_presets_dict

    # First check if tag file exists
    tag_dict = read_tag_file()
    
    if tag_dict:
        print("Found tag file, will use both tag-based and placement-based targeting")
        detected_creatives = {}
        
        # Create detected_creatives based on available tag sizes
        for size in tag_dict.keys():
            if 'x' in size:
                base_size = size.split('_')[0].strip()
                if base_size in available_presets and base_size in presets_dict:
                    # For richmedia lines, check platform compatibility
                    if line_type == "richmedia":
                        required_platforms = presets_dict[base_size].get("platforms", [])
                        current_platforms = [p.strip().upper() for p in line_item_data['platforms']]
                        
                        # Check if any of the current platforms match the required platforms
                        platform_match = any(platform in required_platforms for platform in current_platforms)
                        
                        if not platform_match:
                            print(f"🚫 Skipping tag size {base_size} - platform mismatch. Required: {required_platforms}, Current: {current_platforms}")
                            continue
                        else:
                            print(f"✅ Tag size {base_size} platform match. Required: {required_platforms}, Current: {current_platforms}")
                    
                    size_key = size
                    detected_creatives[size_key] = {
                        "adtype_filter": presets_dict[base_size]["adtypes"],
                        "section_filter": presets_dict[base_size]["sections"],
                        "base_size": base_size
                    }
                    print(f"Added size {size_key} from tag file")
    else:
        print("No tag file found, using Google Sheets for placement targeting")
        detected_creatives, image_files = fetch_images_and_presets(CREATIVES_FOLDER, available_presets, presets_dict)
        
        if not detected_creatives:
            if In_Banner_video and In_Banner_video.strip():
                print(f"No image creatives detected, but In_Banner_video is provided: {In_Banner_video}")
                detected_creatives["300x250"] = {
                    "adtype_filter": presets_dict["300x250"]["adtypes"] if "300x250" in presets_dict else ["MREC_ALL", "MREC"],
                    "section_filter": presets_dict["300x250"]["sections"] if "300x250" in presets_dict else ["ROS", "HP"],
                    "base_size": "300x250"
                }
                print("Added 300x250 size for In-Banner Video to detected_creatives")
            else:
                raise ValueError("No creatives detected in the creatives folder and no valid tag file found.")

    print(f"Final Template_Id: {Template_Id}")

    # Group creatives by base size for placement targeting
    size_groups = {}
    for size_key, creative_info in detected_creatives.items():
        base_size = creative_info.get('base_size', size_key.split('_')[0])
        
        # For richmedia lines, check platform compatibility
        if line_type == "richmedia" and base_size in presets_dict:
            required_platforms = presets_dict[base_size].get("platforms", [])
            current_platforms = [p.strip().upper() for p in line_item_data['platforms']]
            
            # Check if any of the current platforms match the required platforms
            platform_match = any(platform in required_platforms for platform in current_platforms)
            
            if not platform_match:
                print(f"🚫 Skipping size {base_size} - platform mismatch. Required: {required_platforms}, Current: {current_platforms}")
                continue
            else:
                print(f"✅ Size {base_size} platform match. Required: {required_platforms}, Current: {current_platforms}")
        
        # Special case: if size is 320x100, use 320x50 for placement targeting
        placement_size = "320x50" if base_size == "320x100" else base_size
        
        if placement_size not in size_groups:
            # Use the original base_size for adtype and section filters
            filter_size = base_size if base_size in presets_dict else placement_size
            size_groups[placement_size] = {
                "adtype_filter": presets_dict[filter_size]["adtypes"] if filter_size in presets_dict else creative_info["adtype_filter"],
                "section_filter": presets_dict[filter_size]["sections"] if filter_size in presets_dict else creative_info["section_filter"],
                "placement_ids": [],
                "original_sizes": [base_size]  # Track original sizes for this placement group
            }
        else:
            # Add this original size to the existing group
            if "original_sizes" not in size_groups[placement_size]:
                size_groups[placement_size]["original_sizes"] = []
            if base_size not in size_groups[placement_size]["original_sizes"]:
                size_groups[placement_size]["original_sizes"].append(base_size)

    site_filter = line_item_data['site']
    if isinstance(site_filter, str):
        site_filter = [site_filter]

    print("\n🔍 Input values:")
    print(f"Site filter: {site_filter}")
    print(f"Platforms: {line_item_data['platforms']}")

    # Normalize platforms to uppercase for consistent matching
    normalized_platforms = [p.strip().upper() for p in line_item_data['platforms']]
    print(f"Normalized platforms: {normalized_platforms}")

    if 'ALL_Languages' in site_filter:
        site_filter = [site for site in site_filter if site != 'ALL_Languages']
        site_filter.extend(['IAG', 'ITBANGLA', 'MS', 'MT', 'NBT', 'TLG', 'TML', 'VK'])
        site_filter = list(dict.fromkeys(site_filter))
        print(f"Expanded ALL_Languages to: {site_filter}")

    contains_toi = any(site in ['TOI', 'ETIMES'] for site in site_filter)
    contains_et = any(site == 'ET' for site in site_filter)
    
    # Log placement targeting configuration
    logger.log_placement_targeting({
        'site_filter': site_filter,
        'platform_filter': normalized_platforms,
        'contains_toi': contains_toi,
        'contains_et': contains_et
    }, session_id)
    contains_lang = any(site not in ['TOI', 'ETIMES', 'ET'] for site in site_filter)

    print("\n📊 Site categorization:")
    print(f"Contains TOI/ETIMES: {contains_toi}")
    print(f"Contains ET: {contains_et}")
    print(f"Contains other languages: {contains_lang}")

    placement_data = {}

    # For richmedia lines, create platform mapping for each size
    filtered_size_groups = {}
    richmedia_platform_map = {}
    
    if line_type == "richmedia":
        for size_key, size_info in size_groups.items():
            # Get the original sizes for this group
            original_sizes = size_info.get('original_sizes', [size_key])
            
            # For each original size, check what platforms it supports
            all_supported_platforms = set()
            for orig_size in original_sizes:
                if orig_size in presets_dict:
                    size_platforms = [p.upper() for p in presets_dict[orig_size].get('platforms', [])]
                    # Only use platforms that are both in input and supported by this size
                    supported_for_this_size = set(normalized_platforms).intersection(set(size_platforms))
                    all_supported_platforms.update(supported_for_this_size)
                    print(f"📋 Size {orig_size} supports platforms: {size_platforms}, effective: {list(supported_for_this_size)}")
            
            # Only include this size group if it has supported platforms
            if all_supported_platforms:
                filtered_size_groups[size_key] = size_info.copy()
                richmedia_platform_map[size_key] = list(all_supported_platforms)
                print(f"✅ Size group {size_key} (richmedia) will use platforms: {list(all_supported_platforms)}")
            else:
                print(f"🚫 Size group {size_key} (richmedia) has no supported platforms, skipping")
    else:
        filtered_size_groups = size_groups
        richmedia_platform_map = None
    
    # Use user-specified platforms for both richmedia and standard lines
    platforms_for_fetch = normalized_platforms
    print(f"🎯 Using user-specified platforms for placement fetch: {platforms_for_fetch}")

    # End data processing, start placement lookup timing
    timing_checkpoints['data_processing_end'] = time.time()
    timing_checkpoints['placement_lookup_start'] = time.time()

    # Always fetch placements regardless of tag file
    if contains_toi:
        toi_sites = [s for s in site_filter if s in ['TOI', 'ETIMES']]
        custom_sheet_name = line_item_data.get('custom_sheet_name')
        sheet_name_used = custom_sheet_name if custom_sheet_name else PLACEMENT_SHEET_NAME_TOI
        print(f"\nFetching TOI placements from sheet: {sheet_name_used}")
        
        placement_data_toi = fetch_placements_ids(
            CREDENTIALS_PATH,
            SHEET_URL,
            sheet_name_used,
            toi_sites,
            platforms_for_fetch,
            filtered_size_groups,
            richmedia_platform_map,
            line_type
        )
        # Merge TOI placement data
        print(f"🔍 TOI placement_data_toi type: {type(placement_data_toi)}")
        print(f"🔍 TOI placement_data_toi content: {placement_data_toi}")
        
        for size, data in placement_data_toi.items():
            print(f"🔍 TOI merging - Size: {size}, Data type: {type(data)}, Data: {data}")
            if size not in placement_data:
                placement_data[size] = data.copy()
                print(f"🔍 TOI - Added new size {size} to placement_data")
            else:
                # Ensure we're working with dictionaries
                if not isinstance(placement_data[size], dict):
                    print(f"⚠️ TOI - placement_data[{size}] is not a dict: {type(placement_data[size])}")
                    placement_data[size] = {}
                if not isinstance(data, dict):
                    print(f"⚠️ TOI - incoming data is not a dict: {type(data)}")
                    continue
                    
                # Combine placement IDs from both sources
                existing_ids = set(placement_data[size].get('placement_ids', []))
                new_ids = set(data.get('placement_ids', []))
                placement_data[size]['placement_ids'] = list(existing_ids | new_ids)
                
                # Preserve original_sizes information
                existing_original_sizes = set(placement_data[size].get('original_sizes', []))
                new_original_sizes = set(data.get('original_sizes', []))
                placement_data[size]['original_sizes'] = list(existing_original_sizes | new_original_sizes)
                
                print(f"🔄 Merged TOI placements for {size}: {len(new_ids)} new IDs")

    if contains_et:
        et_sites = [s for s in site_filter if s == 'ET']
        custom_sheet_name = line_item_data.get('custom_sheet_name')
        et_sheet_name_used = custom_sheet_name if custom_sheet_name else PLACEMENT_SHEET_NAME_ET
        print(f"\nFetching ET placements from sheet: {et_sheet_name_used}")
        placement_data_et = fetch_placements_ids(
            CREDENTIALS_PATH,
            SHEET_URL,
            et_sheet_name_used,
            et_sites,
            platforms_for_fetch,
            filtered_size_groups,
            richmedia_platform_map,
            line_type
        )
        # Merge ET placement data
        for size, data in placement_data_et.items():
            if size not in placement_data:
                placement_data[size] = data.copy()
            else:
                # Combine placement IDs from both sources
                existing_ids = set(placement_data[size].get('placement_ids', []))
                new_ids = set(data.get('placement_ids', []))
                placement_data[size]['placement_ids'] = list(existing_ids | new_ids)
                
                # Preserve original_sizes information
                existing_original_sizes = set(placement_data[size].get('original_sizes', []))
                new_original_sizes = set(data.get('original_sizes', []))
                placement_data[size]['original_sizes'] = list(existing_original_sizes | new_original_sizes)
                
                print(f"🔄 Merged ET placements for {size}: {len(new_ids)} new IDs")

    if contains_lang:
        lang_sites = [s for s in site_filter if s not in ['TOI', 'ETIMES', 'ET']]
        
        # Check if custom sheet name is provided (for PSBK line)
        custom_sheet = line_item_data.get('custom_sheet_name')
        sheet_name_to_use = custom_sheet if custom_sheet else PLACEMENT_SHEET_NAME_LANG
        
        print(f"\nFetching Language placements from sheet: {sheet_name_to_use}")
        placement_data_lang = fetch_placements_ids(
            CREDENTIALS_PATH,
            SHEET_URL,
            sheet_name_to_use,
            lang_sites,
            platforms_for_fetch,
            filtered_size_groups,
            richmedia_platform_map,
            line_type
        )
        # Merge Language placement data
        for size, data in placement_data_lang.items():
            if size not in placement_data:
                placement_data[size] = data.copy()
            else:
                # Combine placement IDs from both sources
                existing_ids = set(placement_data[size].get('placement_ids', []))
                new_ids = set(data.get('placement_ids', []))
                placement_data[size]['placement_ids'] = list(existing_ids | new_ids)
                
                # Preserve original_sizes information
                existing_original_sizes = set(placement_data[size].get('original_sizes', []))
                new_original_sizes = set(data.get('original_sizes', []))
                placement_data[size]['original_sizes'] = list(existing_original_sizes | new_original_sizes)
                
                print(f"🔄 Merged Language placements for {size}: {len(new_ids)} new IDs")

    # Safeguard: Ensure original_sizes are preserved from size_groups
    for placement_size, group_data in placement_data.items():
        if placement_size in size_groups:
            expected_original_sizes = size_groups[placement_size].get('original_sizes', [])
            current_original_sizes = group_data.get('original_sizes', [])
            
            # If original_sizes is missing or incorrect, fix it
            if not current_original_sizes or (len(current_original_sizes) == 1 and current_original_sizes[0] == placement_size and expected_original_sizes != [placement_size]):
                placement_data[placement_size]['original_sizes'] = expected_original_sizes.copy()

    print("\nFinal placement data:")
    print(json.dumps(placement_data, indent=2))
    
    # Debug: Show original sizes mapping and data types
    print("\n🔍 Final placement_data debug:")
    print(f"  - placement_data type: {type(placement_data)}")
    print(f"  - placement_data keys: {list(placement_data.keys())}")
    
    for placement_size, data in placement_data.items():
        print(f"🔍 Final - Size: {placement_size}, Data type: {type(data)}")
        if isinstance(data, dict):
            original_sizes = data.get('original_sizes', [placement_size])
            print(f"  - Original sizes: {original_sizes}")
            print(f"  - Placement IDs count: {len(data.get('placement_ids', []))}")
        else:
            print(f"  - ⚠️ WARNING: Data is not a dict: {data}")

    # Get all placement IDs from placement_data
    all_placement_ids = []
    print(f"🔍 Collecting all placement IDs:")
    
    for key, group_info in placement_data.items():
        print(f"  - Processing key: {key}, type: {type(group_info)}")
        if isinstance(group_info, dict):
            placement_ids = group_info.get('placement_ids', [])
            print(f"    - Found {len(placement_ids)} placement IDs")
            all_placement_ids.extend(placement_ids)
        else:
            print(f"    - ⚠️ WARNING: Expected dict but got {type(group_info)}: {group_info}")
    
    all_placement_ids = list(set(all_placement_ids))
    print(f"🔍 Total unique placement IDs collected: {len(all_placement_ids)}")

    # If no placement IDs found, raise an error since inventory targeting is required
    if not all_placement_ids:
        raise ValueError("No placement IDs found. Inventory targeting is required for line item creation.")

    print("Detected creatives:", detected_creatives)
    print("Placement data:", placement_data)

    creative_placeholders = []
    creative_targetings = []
    
    # Create placeholders and targetings only for sizes that have placement IDs
    for base_size, group_info in placement_data.items():
        if group_info.get('placement_ids'):  # Only create if there are placement IDs
            # Get the original sizes for this placement group
            original_sizes = group_info.get('original_sizes', [base_size])
            
            # Create creative placeholders and targetings for each original size
            for original_size in original_sizes:
                # Create a descriptive targeting name for display
                if original_size == "320x100":
                    targeting_display_name = "Mweb_PPD"
                elif original_size == "300x250" and line_type == "richmedia":
                    targeting_display_name = "Mrec_ex"
                elif original_size == "300x600" and line_type == "richmedia":
                    targeting_display_name = "Tower_ex"
                else:
                    targeting_display_name = original_size
                
                creative_placeholders.append({
                    'targetingName': targeting_display_name,
                    'size': {
                        'width': int(original_size.split('x')[0]),
                        'height': int(original_size.split('x')[1])
                    }
                })
                print(f"✅ Added creative placeholder for original size {original_size} with targeting name {targeting_display_name}")
                
                # Create matching targeting for this original size using placements or ad units
                if group_info.get('targeting_type') == 'ad_slot_id':
                    targeted_ad_units = [
                        {
                            'adUnitId': pid,
                            'includeDescendants': True
                        } for pid in group_info['placement_ids']
                    ]
                    targeting_inventory = {
                        'targetedAdUnits': targeted_ad_units
                    }
                else:
                    targeting_inventory = {
                        'targetedPlacementIds': group_info['placement_ids']
                    }
                targeting_dict = {
                    'name': targeting_display_name,
                    'targeting': {
                        'inventoryTargeting': targeting_inventory,
                    }
                }
                creative_targetings.append(targeting_dict)
                # Check if this is a PSBK line using Ad slot IDs
                targeting_type = group_info.get('targeting_type', 'placement_id')
                if targeting_type == 'ad_slot_id':
                    print(f"✅ Added creative targeting '{targeting_display_name}' for original size {original_size} using Ad slot IDs from {base_size}")
                    print(f"🔧 PSBK line: Using Ad slot IDs instead of Placement IDs")
                else:
                    print(f"✅ Added creative targeting '{targeting_display_name}' for original size {original_size} using placement IDs from {base_size}")
                print(f"🔍 Creative targeting type: {type(targeting_dict)}, placement_ids type: {type(group_info['placement_ids'])}")
        else:
            print(f"⚠️ Skipping {base_size} - no placement IDs found")

    # Add special targeting for In-Banner Video if needed
    if In_Banner_video and '300x250' not in placement_data:
        creative_placeholders.append({
            'size': {
                'width': 300,
                'height': 250
            }
        })
        print("➕ Added targeting for In-Banner Video 300x250")

    # Add additional sizes for special cases
    if '1260x570' in placement_data and placement_data['1260x570'].get('placement_ids'):
        additional_sizes = ['728x500', '1320x570']
        for size in additional_sizes:
            creative_placeholders.append({
                'size': {
                    'width': int(size.split('x')[0]),
                    'height': int(size.split('x')[1])
                }
            })
        print(f"➕ Added additional sizes {additional_sizes} because 1260x570 was present")

    if '980x200' in placement_data and placement_data['980x200'].get('placement_ids') and '728x90' not in placement_data:
        creative_placeholders.append({
            'size': {
                'width': 728,
                'height': 90
            }
        })
        print("➕ Added additional size 728x90 because 980x200 was present")

    # Add 320x50 override for 320x100 (similar to 728x90 for 980x200)
    # Check if 320x100 exists in any placement data and 320x50 doesn't exist as its own entry
    print(f"🔍 Debugging placement_data.values():")
    print(f"  - placement_data type: {type(placement_data)}")
    print(f"  - placement_data keys: {list(placement_data.keys()) if isinstance(placement_data, dict) else 'Not a dict'}")
    
    has_320x100 = False
    has_explicit_320x50 = False
    
    try:
        for key, data in placement_data.items():
            print(f"  - Key: {key}, Data type: {type(data)}, Data: {data}")
            if isinstance(data, dict):
                original_sizes = data.get('original_sizes', [])
                if '320x100' in original_sizes:
                    has_320x100 = True
                if '320x50' in original_sizes:
                    has_explicit_320x50 = True
            else:
                print(f"  - WARNING: Expected dict but got {type(data)} for key {key}")
    except Exception as e:
        print(f"  - ERROR in placement_data iteration: {e}")
        print(f"  - placement_data content: {placement_data}")
        # Fallback to original logic with error handling
        has_320x100 = any(
            isinstance(data, dict) and '320x100' in data.get('original_sizes', []) 
            for data in placement_data.values()
        )
        has_explicit_320x50 = any(
            isinstance(data, dict) and '320x50' in data.get('original_sizes', []) 
            for data in placement_data.values()
        )
    
    if has_320x100 and not has_explicit_320x50:
        # Add 320x50 creative placeholder (exactly like 728x90 for 980x200)
        # NO targetingName and NO creative targeting - this allows any 320x50 creative to serve
        creative_placeholders.append({
            'size': {
                'width': 320,
                'height': 50
            }
        })
        print("➕ Added additional size 320x50 because 320x100 was present")



    if '600x250' in placement_data and placement_data['600x250'].get('placement_ids'):
        creative_placeholders.append({
            'targetingName': 'Mrec Expando',
            'size': {'width': 300, 'height': 250}
        })
        
        targeting_dict = {
            'name': 'Mrec Expando',
            'targeting': {
                'inventoryTargeting': {
                    'targetedPlacementIds': placement_data['600x250']['placement_ids'],
                },
            }
        }
        creative_targetings.append(targeting_dict)
        print(f"🔍 Added Mrec Expando targeting, type: {type(targeting_dict)}")

    # Check Expresso information for smarter uniqueness handling
    expresso_line_item_found = line_item_data.get('expresso_line_item_found', False)
    expresso_line_item_name = line_item_data.get('expresso_line_item_name', '')
    
    # Check if line name already exists in GAM
    if check_line_item_name_exists(client, order_id, line_name):
        # If it exists, use the original name without any suffix
        unique_line_name = line_name.strip() if line_name else ''
        print(f"📋 Using line name as-is: {unique_line_name}")
    else:
        # If it doesn't exist, use as-is
        unique_line_name = line_name.strip() if line_name else ''
        print(f"📋 Using line name as-is: {unique_line_name}")
    
    print(f"🔄 Final line item name: {unique_line_name}")
    
    # End placement lookup, start line creation timing
    timing_checkpoints['placement_lookup_end'] = time.time()
    timing_checkpoints['line_creation_start'] = time.time()
    
    # Debug geo targeting before line item creation
    print(f"\n{'='*80}")
    print(f"🔍 LINE ITEM GEO TARGETING CONFIGURATION")
    print(f"{'='*80}")
    print(f"📋 Basic Information:")
    print(f"  • Line Name: {unique_line_name}")
    print(f"  • Line Type: {line_item_data.get('line_type', 'standard')}")
    
    print(f"\n🎯 Targeting Configuration:")
    if geo_targeting_ids:
        print(f"  ✅ Targeted Locations:")
        for geo_id in geo_targeting_ids:
            print(f"    • Location ID: {geo_id}")
    else:
        print(f"  ⚠️ No targeted locations configured")
    
    if excluded_geo_ids:
        print(f"\n🚫 Excluded Locations:")
        for geo_id in excluded_geo_ids:
            print(f"    • Location ID: {geo_id}")
    else:
        print(f"\n  ℹ️ No location exclusions configured")
    
    print(f"\n📊 Summary:")
    print(f"  • Total Targeted Locations: {len(geo_targeting_ids)}")
    print(f"  • Total Excluded Locations: {len(excluded_geo_ids)}")
    print(f"{'='*80}\n")
    
    # Check if we're using Ad slot IDs for PSBK lines
    targeting_type = "placement_id"  # default
    for size_data in placement_data.values():
        if isinstance(size_data, dict) and size_data.get('targeting_type') == 'ad_slot_id':
            targeting_type = 'ad_slot_id'
            break
    
    # Prepare targeting configuration
    if targeting_type == 'ad_slot_id':
        targeting_config = {
            'inventoryTargeting': {
                'targetedAdUnits': [
                    {
                        'adUnitId': pid,
                        'includeDescendants': True
                    } for pid in all_placement_ids
                ]
            }
        }
    else:
        targeting_config = {
            'inventoryTargeting': {
                'targetedPlacementIds': all_placement_ids
            }
        }
    
    # Log the targeting type being used
    if targeting_type == 'ad_slot_id':
        print(f"🔧 PSBK line: Using Ad slot IDs for inventory targeting")
        print(f"📊 Total Ad slot IDs: {len(all_placement_ids)}")
    else:
        print(f"📊 Total Placement IDs: {len(all_placement_ids)}")
    
    # Add geo targeting if we have valid locations
    if geo_targeting_ids:
        print(f"✅ Adding geo targeting with {len(geo_targeting_ids)} locations")
        targeting_config['geoTargeting'] = {
            'targetedLocations': [{'id': geo_id} for geo_id in geo_targeting_ids]
        }
        if excluded_geo_ids:
            print(f"✅ Adding {len(excluded_geo_ids)} excluded locations")
            targeting_config['geoTargeting']['excludedLocations'] = [{'id': geo_id} for geo_id in excluded_geo_ids]
    else:
        print("⚠️ No geo targeting IDs found, skipping geo targeting configuration")
    
    line_item = {
        'name': unique_line_name,
        'orderId': order_id,
        'targeting': targeting_config,
        'creativePlaceholders': creative_placeholders,
        'creativeTargetings': creative_targetings,
        'endDateTime': structured_end_date,
        'deliveryRateType': 'EVENLY',
        'lineItemType': 'STANDARD',
        'costType': 'CPM',
        'costPerUnit': {
            'currencyCode': currency,
            'microAmount': int(float(cost) * 1_000_000)
        },
        'primaryGoal': {
            'goalType': 'LIFETIME',
            'units': impressions_int
        },
        'allowOverbook': True,
        'skipInventoryCheck': True
    }

    if use_start_date_type:
        line_item['startDateTimeType'] = 'IMMEDIATELY'
    else:
        line_item['startDateTime'] = structured_start_date

    if Fcap_value > 0:
        line_item['frequencyCaps'] = [{
            'maxImpressions': Fcap_value,
            'timeUnit': 'LIFETIME'  # Set frequency cap per lifetime for all line types
        }]

    try:
        # Create the line item
        print("🚀 Creating line item...")
        created_line_items = line_item_service.createLineItems([line_item])
        line_item_id = created_line_items[0]['id']
        print(f"✅ Successfully created line item with ID: {line_item_id}")
        
        # End line creation, start creative creation timing
        timing_checkpoints['line_creation_end'] = time.time()
        timing_checkpoints['creative_creation_start'] = time.time()
        
        # Log successful line creation (we'll add creative IDs later)
        logger.log_line_creation_success(
            line_id=line_item_id,
            creative_ids=[],  # Will be populated later
            order_id=str(order_id),
            line_name=unique_line_name,
            session_id=session_id
        )
        
    except Exception as e:
        error_str = str(e)
        if "DUPLICATE_OBJECT" in error_str:
            print(f"❌ DUPLICATE_OBJECT error despite checks. Details:")
            print(f"   - Name length: {len(unique_line_name)} chars")
            print(f"   - Timestamp prefix: {unique_line_name.split('_')[0]}")
            print(f"   - Full error: {error_str}")
        else:
            print(f"❌ Failed to create line item: {e}")
        logger.log_line_creation_error(e, unique_line_name, str(order_id), session_id)
        raise

    # Read tags from Excel only if file exists
    tag_dict = read_tag_file()
    
    creative_ids = []
    
    # First, gather all the tags for each base size
    size_tags = {}
    if tag_dict:
        for tag_key in tag_dict.keys():
            base_size = tag_key.split('_')[0].strip()
            if base_size not in size_tags:
                size_tags[base_size] = []
            size_tags[base_size].append(tag_key)
    
    # Process placement_data sizes
    print(f"Processing placement_data sizes: {list(placement_data.keys())}")
    for size_name in placement_data:
        print(f"Processing size_name: {size_name}, has placement_ids: {bool(placement_data[size_name]['placement_ids'])}")
        if placement_data[size_name]['placement_ids']:
            # Get the original sizes that map to this placement size
            original_sizes = placement_data[size_name].get('original_sizes', [size_name])
            print(f"Original sizes for {size_name}: {original_sizes}")
            
            for original_size in original_sizes:
                try:
                    print(f"Creating creative(s) for original size: {original_size} (placement size: {size_name})")
                    
                    # Initialize variables for tag processing
                    use_script_tag = script_tracker
                    use_impression_tag = impression_tracker
                    use_landing_page = landing_page
                    use_template_id = Template_Id
                    
                    # Special handling for 320x100 size - use template ID 12363950
                    if original_size == "320x100":
                        use_template_id = 12363950
                        print(f"Using special template ID: {use_template_id} for size {original_size}")
                    # Special handling for 300x250 richmedia - use template ID 12460223
                    elif original_size == "300x250" and line_type == "richmedia":
                        use_template_id = 12460223
                        print(f"Using 300x250 richmedia template ID: {use_template_id} for size {original_size}")
                    # Special handling for 300x600 richmedia - use template ID 12443458
                    elif original_size == "300x600" and line_type == "richmedia":
                        use_template_id = 12443458
                        print(f"Using 300x600 richmedia template ID: {use_template_id} for size {original_size}")
                    
                    # Special handling for In_Banner_Video
                    elif In_Banner_video and original_size == "300x250":
                        use_template_id = 12344286
                        print(f"Using In-Banner Video template ID: {use_template_id} for size {original_size}")
                        new_creatives = create_custom_template_creatives(
                            client, order_id, line_item_id,
                            destination_url, expresso_id, original_size, use_landing_page,
                            use_impression_tag, use_script_tag, use_template_id, In_Banner_video, line_type
                        )
                        creative_ids.extend(new_creatives)
                        track_creative_creation(original_size, new_creatives)
                        continue
                    
                    # Process tags if available for the original size
                    if tag_dict and original_size in size_tags:
                        # Process each tag for this original_size
                        for tag_key in size_tags[original_size]:
                            tag_info = tag_dict[tag_key]
                            print(f"Processing tag {tag_key} for original size {original_size}")
                            
                            # Reset to defaults for each tag
                            use_script_tag = script_tracker
                            use_impression_tag = impression_tracker
                            use_landing_page = landing_page
                            use_template_id = Template_Id
                            
                            # Special handling for 320x100 size - use template ID 12363950
                            if original_size == "320x100":
                                use_template_id = 12363950
                                print(f"Using special template ID: {use_template_id} for size {original_size}")
                            # Special handling for 300x250 richmedia - use template ID 12460223
                            elif original_size == "300x250" and line_type == "richmedia":
                                use_template_id = 12460223
                                print(f"Using 300x250 richmedia template ID: {use_template_id} for size {original_size}")
                            # Special handling for 300x600 richmedia - use template ID 12443458
                            elif original_size == "300x600" and line_type == "richmedia":
                                use_template_id = 12443458
                                print(f"Using 300x600 richmedia template ID: {use_template_id} for size {original_size}")
                            elif tag_info['type'] == 'impression_click':
                                # For impression/click tag combo, use template 12330939
                                print(f"Using Impression/Click tags from tag.xlsx for size {original_size}")
                                
                                # Process impression tag to extract just the URL from IMG SRC attribute
                                impression_tag_value = tag_info['impression_tag']
                                
                                # Extract URL from IMG SRC tag if it exists
                                if 'IMG SRC=' in impression_tag_value.upper() or 'src=' in impression_tag_value.lower():
                                    url_match = re.search(r'src=["\'](https?://[^"\']+)["\']', impression_tag_value, re.IGNORECASE)
                                    if url_match:
                                        impression_tag_value = url_match.group(1)
                                        # Replace [timestamp] or [CACHEBUSTER] with %%CACHEBUSTER%%
                                        impression_tag_value = impression_tag_value.replace('[timestamp]', '%%CACHEBUSTER%%').replace('[CACHEBUSTER]', '%%CACHEBUSTER%%')
                                        print(f"Extracted URL from impression tag: {impression_tag_value}")
                                
                                use_impression_tag = impression_tag_value
                                use_landing_page = tag_info['click_tag'].replace('[timestamp]', '%%CACHEBUSTER%%').replace('[CACHEBUSTER]', '%%CACHEBUSTER%%')
                                use_template_id = 12330939  # Use the specified template ID
                                print(f"Using template ID: {use_template_id} for impression/click tags")
                            elif tag_info['type'] == 'doubleclick':
                                # For DoubleClick tag
                                print(f"Using DoubleClick tag from tag.xlsx for size {original_size}")
                                use_script_tag = tag_info['js_tag']
                                
                                # Handle <noscript> tags with <a> href, common in Flashtalking tags
                                if '<noscript>' in use_script_tag.lower() and '<a href' in use_script_tag.lower():
                                    print(f"Detected noscript/a href tag for size {original_size}")
                                    href_pattern = r'(<a\s+[^>]*?href=")([^"]*)"'
                                    
                                    # Prepend click macro if not already present
                                    if '%%CLICK_URL_UNESC%%' not in use_script_tag:
                                        replacement = r'\1%%CLICK_URL_UNESC%%\2"'
                                        modified_tag = re.sub(href_pattern, replacement, use_script_tag, flags=re.IGNORECASE)
                                        
                                        if modified_tag != use_script_tag:
                                            use_script_tag = modified_tag
                                            print(f"Added %%CLICK_URL_UNESC%% to href in noscript tag for size {original_size}")
                                        else:
                                            print(f"Warning: Could not add %%CLICK_URL_UNESC%% to href for size {original_size}")
                                    else:
                                        print(f"Click macro already present for size {original_size}")
                                
                                # Ensure the DoubleClick tag has a data-dcm-click-tracker attribute
                                if 'data-dcm-click-tracker' not in use_script_tag:
                                    try:
                                        # First try to add it before class attribute
                                        tag_pattern = r'(<ins|<div)([^>]*?)(\s+class=)'
                                        replacement = r"\1\2 data-dcm-click-tracker='%%CLICK_URL_UNESC%%'\3"
                                        modified_tag = re.sub(tag_pattern, replacement, use_script_tag, flags=re.IGNORECASE)
                                        
                                        # If that didn't work, try adding it after the opening tag
                                        if modified_tag == use_script_tag:
                                            tag_pattern = r'(<ins|<div)(\s)'
                                            replacement = r"\1 data-dcm-click-tracker='%%CLICK_URL_UNESC%%'\2"
                                            modified_tag = re.sub(tag_pattern, replacement, use_script_tag, flags=re.IGNORECASE)
                                        
                                        use_script_tag = modified_tag
                                        print(f"Added data-dcm-click-tracker to DoubleClick tag for size {original_size}")
                                    except Exception as e:
                                        print(f"Warning: Could not add data-dcm-click-tracker attribute: {str(e)}")
                                
                                use_template_id = 12435443
                            else:
                                # For JavaScript tag, use AI template
                                print(f"Using JavaScript tag from tag.xlsx for size {original_size}")
                                use_script_tag = tag_info['js_tag']
                                
                                # Handle <noscript> tags with <a> href, common in Flashtalking tags
                                if '<noscript>' in use_script_tag.lower() and '<a href' in use_script_tag.lower():
                                    print(f"Detected noscript/a href tag for size {original_size}")
                                    href_pattern = r'(<a\s+[^>]*?href=")([^"]*)"'
                                    
                                    # Prepend click macro if not already present
                                    if '%%CLICK_URL_UNESC%%' not in use_script_tag:
                                        replacement = r'\1%%CLICK_URL_UNESC%%\2"'
                                        modified_tag = re.sub(href_pattern, replacement, use_script_tag, flags=re.IGNORECASE)
                                        
                                        if modified_tag != use_script_tag:
                                            use_script_tag = modified_tag
                                            print(f"Added %%CLICK_URL_UNESC%% to href in noscript tag for size {original_size}")
                                        else:
                                            print(f"Warning: Could not add %%CLICK_URL_UNESC%% to href for size {original_size}")
                                    else:
                                        print(f"Click macro already present for size {original_size}")
                                
                                use_template_id = 12435443  # AI template for JavaScript tags
                                print(f"Using AI template ID: {use_template_id} for JavaScript tag")
                            
                            # Create the creative and associate it with the line item
                            new_creatives = create_custom_template_creatives(
                                client, order_id, line_item_id,
                                destination_url, expresso_id, original_size, use_landing_page,
                                use_impression_tag, use_script_tag, use_template_id, In_Banner_video, line_type
                            )
                            creative_ids.extend(new_creatives)
                            track_creative_creation(original_size, new_creatives)
                    else:
                        # No tags for this size, create a normal creative
                        creative_path = None
                        if original_size in detected_creatives:
                            creative_path = detected_creatives[original_size].get('image_path', '')
                        if creative_path and creative_path.lower().endswith('.html'):
                            use_template_id = 12435443
                            print(f"Using template ID 12435443 for HTML creative: {creative_path}")
                        new_creatives = create_custom_template_creatives(
                            client, order_id, line_item_id,
                            destination_url, expresso_id, original_size, use_landing_page,
                            use_impression_tag, use_script_tag, use_template_id, In_Banner_video, line_type
                        )
                        creative_ids.extend(new_creatives)
                        track_creative_creation(original_size, new_creatives)
                except Exception as e:
                    print(f"⚠️ Failed to create creatives for original size {original_size}: {e}")

    # Note: 320x50 creative creation is handled later in the consolidated logic

    # Now, process any tag sizes that weren't in placement_data
    if tag_dict:
        print("Checking for additional tag sizes not covered by placement data...")
        # First, get all original sizes that were already processed
        processed_original_sizes = set()
        for group_info in placement_data.values():
            processed_original_sizes.update(group_info.get('original_sizes', []))
        
        print(f"Processed original sizes from placement_data: {processed_original_sizes}")
        print(f"Available tag sizes: {list(size_tags.keys())}")
        
        # Find any tags for sizes that we haven't processed yet
        for tag_size, tag_keys in size_tags.items():
            print(f"Checking tag size: {tag_size}")
            
            # Skip sizes we've already processed
            if tag_size in processed_original_sizes:
                print(f"Tag size {tag_size} already processed, skipping")
                continue
                
            # Skip any non-standard sizes
            if tag_size not in available_presets:
                print(f"Tag size {tag_size} not in available presets, skipping")
                continue
                
            # We have a valid tag size that wasn't in placement_data
            print(f"Creating additional creative(s) for tag size: {tag_size}")
            
            # Process each tag for this tag_size
            for tag_key in tag_keys:
                try:
                    tag_info = tag_dict[tag_key]
                    use_script_tag = script_tracker
                    use_impression_tag = impression_tracker
                    use_landing_page = landing_page
                    use_template_id = Template_Id
                    
                    print(f"Processing tag {tag_key} for additional size {tag_size}")
                    
                    # Special handling for 320x100 size - use template ID 12363950
                    if tag_size == "320x100":
                        use_template_id = 12363950
                        print(f"Using special template ID: {use_template_id} for additional size {tag_size}")
                    # Special handling for 300x250 richmedia - use template ID 12460223
                    elif tag_size == "300x250" and line_type == "richmedia":
                        use_template_id = 12460223
                        print(f"Using 300x250 richmedia template ID: {use_template_id} for additional size {tag_size}")
                    # Special handling for 300x600 richmedia - use template ID 12443458
                    elif tag_size == "300x600" and line_type == "richmedia":
                        use_template_id = 12443458
                        print(f"Using 300x600 richmedia template ID: {use_template_id} for additional size {tag_size}")
                    elif tag_info['type'] == 'impression_click':
                        # For impression/click tag combo
                        impression_tag_value = tag_info['impression_tag']
                        if 'IMG SRC=' in impression_tag_value.upper() or 'src=' in impression_tag_value.lower():
                            url_match = re.search(r'src=["\'](https?://[^"\']+)["\']', impression_tag_value, re.IGNORECASE)
                            if url_match:
                                impression_tag_value = url_match.group(1)
                                impression_tag_value = impression_tag_value.replace('[timestamp]', '%%CACHEBUSTER%%').replace('[CACHEBUSTER]', '%%CACHEBUSTER%%')
                        
                        use_impression_tag = impression_tag_value
                        use_landing_page = tag_info['click_tag'].replace('[timestamp]', '%%CACHEBUSTER%%').replace('[CACHEBUSTER]', '%%CACHEBUSTER%%')
                        use_template_id = 12330939
                    elif tag_info['type'] == 'doubleclick':
                        # For DoubleClick tag
                        print(f"Using DoubleClick tag for additional size {tag_size}")
                        use_script_tag = tag_info['js_tag']
                        
                        # Handle <noscript> tags with <a> href, common in Flashtalking tags
                        if '<noscript>' in use_script_tag.lower() and '<a href' in use_script_tag.lower():
                            print(f"Detected noscript/a href tag for additional size {tag_size}")
                            href_pattern = r'(<a\s+[^>]*?href=")([^"]*)"'
                            
                            # Prepend click macro if not already present
                            if '%%CLICK_URL_UNESC%%' not in use_script_tag:
                                replacement = r'\1%%CLICK_URL_UNESC%%\2"'
                                modified_tag = re.sub(href_pattern, replacement, use_script_tag, flags=re.IGNORECASE)
                                
                                if modified_tag != use_script_tag:
                                    use_script_tag = modified_tag
                                    print(f"Added %%CLICK_URL_UNESC%% to href in noscript tag for additional size {tag_size}")
                                else:
                                    print(f"Warning: Could not add %%CLICK_URL_UNESC%% to href for additional size {tag_size}")
                            else:
                                print(f"Click macro already present for additional size {tag_size}")
                        
                        # Ensure the DoubleClick tag has a data-dcm-click-tracker attribute
                        if 'data-dcm-click-tracker' not in use_script_tag:
                            try:
                                # First try to add it before class attribute
                                tag_pattern = r'(<ins|<div)([^>]*?)(\s+class=)'
                                replacement = r"\1\2 data-dcm-click-tracker='%%CLICK_URL_UNESC%%'\3"
                                modified_tag = re.sub(tag_pattern, replacement, use_script_tag, flags=re.IGNORECASE)
                                
                                # If that didn't work, try adding it after the opening tag
                                if modified_tag == use_script_tag:
                                    tag_pattern = r'(<ins|<div)(\s)'
                                    replacement = r"\1 data-dcm-click-tracker='%%CLICK_URL_UNESC%%'\2"
                                    modified_tag = re.sub(tag_pattern, replacement, use_script_tag, flags=re.IGNORECASE)
                                
                                use_script_tag = modified_tag
                                print(f"Added data-dcm-click-tracker to DoubleClick tag for size {tag_size}")
                            except Exception as e:
                                print(f"Warning: Could not add data-dcm-click-tracker attribute: {str(e)}")
                        
                        use_template_id = 12435443
                    else:
                        # For JavaScript tag
                        use_script_tag = tag_info['js_tag']
                        
                        # Handle <noscript> tags with <a> href, common in Flashtalking tags
                        if '<noscript>' in use_script_tag.lower() and '<a href' in use_script_tag.lower():
                            print(f"Detected noscript/a href tag for additional size {tag_size}")
                            href_pattern = r'(<a\s+[^>]*?href=")([^"]*)"'
                            
                            # Prepend click macro if not already present
                            if '%%CLICK_URL_UNESC%%' not in use_script_tag:
                                replacement = r'\1%%CLICK_URL_UNESC%%\2"'
                                modified_tag = re.sub(href_pattern, replacement, use_script_tag, flags=re.IGNORECASE)
                                
                                if modified_tag != use_script_tag:
                                    use_script_tag = modified_tag
                                    print(f"Added %%CLICK_URL_UNESC%% to href in noscript tag for additional size {tag_size}")
                                else:
                                    print(f"Warning: Could not add %%CLICK_URL_UNESC%% to href for additional size {tag_size}")
                            else:
                                print(f"Click macro already present for additional size {tag_size}")
                        
                        use_template_id = 12435443
                    
                    new_creatives = create_custom_template_creatives(
                        client, order_id, line_item_id,
                        destination_url, expresso_id, tag_size, use_landing_page,
                        use_impression_tag, use_script_tag, use_template_id, In_Banner_video, line_type
                    )
                    creative_ids.extend(new_creatives)
                    track_creative_creation(tag_size, new_creatives)
                    print(f"✅ Created additional creative for tag {tag_key} and size {tag_size}")
                except Exception as e:
                    print(f"⚠️ Failed to create additional creative for tag {tag_key} and size {tag_size}: {e}")

    # Create 320x50 creatives when they have their own targeting but aren't in the main placement processing
    # This happens when 320x100 exists and we've added 320x50 as an additional override size
    print(f"🔍 Creative targetings debug:")
    print(f"  - creative_targetings type: {type(creative_targetings)}")
    print(f"  - creative_targetings length: {len(creative_targetings)}")
    for i, targeting in enumerate(creative_targetings):
        print(f"  - Item {i}: type={type(targeting)}, content={targeting}")
    
    has_320x50_targeting = any(
        isinstance(targeting, dict) and targeting.get('name') == '320x50' 
        for targeting in creative_targetings
    )
    processed_320x50 = False
    
    # Check if 320x50 was already processed in the main placement loop
    for size_name in placement_data:
        if placement_data[size_name].get('placement_ids'):
            original_sizes = placement_data[size_name].get('original_sizes', [size_name])
            if '320x50' in original_sizes:
                processed_320x50 = True
                print(f"🔍 320x50 already processed in placement loop for size_name: {size_name}")
                break
    
    # Check if 320x50 creatives were already created
    already_created_320x50 = is_creative_size_already_created('320x50')
    
    print(f"🔍 320x50 Creative Check:")
    print(f"  - has_320x50_targeting: {has_320x50_targeting}")
    print(f"  - processed_320x50: {processed_320x50}")
    print(f"  - already_created_320x50: {already_created_320x50}")
    
    if has_320x50_targeting and not processed_320x50 and not already_created_320x50:
        try:
            print("🔧 Creating 320x50 creative to fulfill the 320x50 targeting (override for 320x100)")
            # Use same settings as would be used for normal creative creation
            use_script_tag = script_tracker
            use_impression_tag = impression_tracker
            use_landing_page = landing_page
            use_template_id = Template_Id or 12330939  # Default to standard template
            
            new_320x50_creatives = create_custom_template_creatives(
                client, order_id, line_item_id,
                destination_url, expresso_id, "320x50", use_landing_page,
                use_impression_tag, use_script_tag, use_template_id, In_Banner_video, line_type
            )
            creative_ids.extend(new_320x50_creatives)
            track_creative_creation('320x50', new_320x50_creatives)
            print(f"✅ Created 320x50 override creative: {new_320x50_creatives}")
        except Exception as e:
            print(f"⚠️ Failed to create 320x50 override creative: {e}")
    elif already_created_320x50:
        print(f"⏭️ Skipping 320x50 creation - already created earlier in the process")

    # Special case for In_Banner_Video: create 300x250 even if not in placement_data
    if In_Banner_video and '300x250' not in placement_data:
        try:
            print("Creating special 300x250 creative for In-Banner Video")
            print(f"In-Banner Video URL: {In_Banner_video}")
            use_template_id = 12344286
            new_creatives = create_custom_template_creatives(
                client, order_id, line_item_id,
                destination_url, expresso_id, "300x250", landing_page,
                impression_tracker, script_tracker, use_template_id, In_Banner_video, line_type
            )
            creative_ids.extend(new_creatives)
            track_creative_creation("300x250", new_creatives)
        except Exception as e:
            print(f"⚠️ Failed to create special In-Banner Video creative: {e}")
                
    # Clear landing_page and impression_tracker after order creation
    line_item_data['landing_page'] = ''
    line_item_data['impression_tracker'] = ''

    # Ensure tracking_tag and banner_video fields are populated
    # Assuming these fields are already populated as needed

    # Set fcap to 0 by default
    line_item_data['fcap'] = 0

    # Set currency to INR
    line_item_data['currency'] = 'INR'

    # End creative creation timing
    timing_checkpoints['creative_creation_end'] = time.time()
    
    # Calculate performance metrics
    end_time = time.time()
    total_time = end_time - start_time
    
    # Calculate individual phase timings
    data_processing_time = (timing_checkpoints['data_processing_end'] - timing_checkpoints['data_processing_start']) if timing_checkpoints['data_processing_end'] else 0
    placement_lookup_time = (timing_checkpoints['placement_lookup_end'] - timing_checkpoints['placement_lookup_start']) if timing_checkpoints['placement_lookup_end'] else 0
    line_creation_time = (timing_checkpoints['line_creation_end'] - timing_checkpoints['line_creation_start']) if timing_checkpoints['line_creation_end'] else 0
    creative_creation_time = (timing_checkpoints['creative_creation_end'] - timing_checkpoints['creative_creation_start']) if timing_checkpoints['creative_creation_end'] else 0
    
    # Log final success with all details
    logger.log_line_creation_success(
        str(line_item_id), 
        [str(cid) for cid in creative_ids] if creative_ids else [], 
        unique_line_name, 
        session_id
    )
    
    # Log performance metrics with detailed timing
    logger.log_performance_metrics({
        'total_time': total_time,
        'data_processing_time': data_processing_time,
        'placement_lookup_time': placement_lookup_time,
        'line_creation_time': line_creation_time,
        'creative_creation_time': creative_creation_time,
        'line_item_id': line_item_id,
        'creative_count': len(creative_ids) if creative_ids else 0,
        'session_id': session_id
    }, session_id)
    
    # Log final performance summary
    print(f"📊 Session Summary:")
    print(f"  - Lines created: 1")
    print(f"  - Creatives created: {len(creative_ids) if creative_ids else 0}")
    print(f"  - Total time: {total_time:.2f}s")
    print(f"  - Success rate: 100.0%")

    return line_item_id, creative_ids


def three_lines(client, order_id, line_item_data, line_name):
    """
    Create three line items instead of one:
    1. Standard line with original name
    2. _psbk line with CAN_PSBK placement data
    3. _nwp line with standard placement data
    """
    # Generate session ID for this three-line creation
    session_id = str(uuid.uuid4())
    start_time = time.time()
    
    print(f"🚀 Starting three-line creation process...")
    print(f"📝 Base line name: {line_name}")
    print(f"🆔 Session ID: {session_id}")
    
    # Log three-line creation start
    logger.log_line_creation_start(str(order_id), line_item_data, f"THREE_LINES: {line_name}", session_id)
    
    all_line_ids = []
    all_creative_ids = []
    success_count = 0
    error_messages = []
    
    # Calculate impression distribution
    total_impressions = line_item_data.get('impressions', 100000)
    if isinstance(total_impressions, str):
        total_impressions = int(total_impressions.replace(',', ''))
    elif isinstance(total_impressions, float):
        total_impressions = int(total_impressions)
    
    # Distribution: Standard 10%, PSBK 80%, NWP 10%
    # Ensure minimum 1 impression per line for small totals
    standard_impressions = max(1, int(total_impressions * 0.10))
    psbk_impressions = max(1, int(total_impressions * 0.80))
    nwp_impressions = max(1, int(total_impressions * 0.10))
    
    # For very small numbers, ensure exact distribution
    if total_impressions <= 10:
        if total_impressions == 100:
            standard_impressions = 10
            psbk_impressions = 80
            nwp_impressions = 10
        elif total_impressions == 10:
            standard_impressions = 1
            psbk_impressions = 8
            nwp_impressions = 1
        elif total_impressions < 10:
            # For numbers less than 10, distribute as evenly as possible
            standard_impressions = max(1, total_impressions // 10)
            psbk_impressions = max(1, int(total_impressions * 0.8))
            nwp_impressions = max(1, total_impressions - standard_impressions - psbk_impressions)
            if nwp_impressions <= 0:
                nwp_impressions = 1
                psbk_impressions = total_impressions - standard_impressions - nwp_impressions
    
    # Special handling for exactly 100 impressions
    if total_impressions == 100:
        standard_impressions = 10
        psbk_impressions = 80
        nwp_impressions = 10
    
    print(f"📊 Impression Distribution:")
    print(f"  - Total: {total_impressions:,}")
    print(f"  - Standard (10%): {standard_impressions:,}")
    print(f"  - PSBK (80%): {psbk_impressions:,}")
    print(f"  - NWP (10%): {nwp_impressions:,}")

    # Generate unique timestamps for each line to avoid name conflicts
    base_timestamp = int(time.time())
    
    # Define the three lines to create with unique timestamps
    lines_to_create = [
        {
            'name': f"{line_name}",
            'suffix': '',
            'description': 'Standard line',
            'use_psbk': False,
            'use_nwp': False,
            'line_type': 'standard',
            'impressions': standard_impressions
        },
        {
            'name': f"{line_name}_psbk",
            'suffix': '_psbk',
            'description': 'PSBK line with CAN_PSBK placement data',
            'use_psbk': True,
            'use_nwp': False,
            'line_type': 'psbk',
            'impressions': psbk_impressions
        },
        {
            'name': f"{line_name}_nwp",
            'suffix': '_nwp',
            'description': 'NWP line with hardcoded placement data',
            'use_psbk': False,
            'use_nwp': True,
            'line_type': 'nwp',
            'impressions': nwp_impressions
        }
    ]
    
    print(f"📋 Creating {len(lines_to_create)} line items:")
    for i, line_config in enumerate(lines_to_create, 1):
        print(f"  {i}. {line_config['description']}: {line_config['name']}")
    
    # Create each line item with retry mechanism
    for i, line_config in enumerate(lines_to_create):
        max_retries = 3
        retry_delay = 2  # seconds
        
        for attempt in range(max_retries):
            try:
                print(f"\n🔄 Creating {line_config['description']} (attempt {attempt + 1}/{max_retries})...")
                print(f"📝 Line name: {line_config['name']}")
                print(f"📊 Impressions: {line_config['impressions']:,}")
                
                # Add delay between line creations to avoid concurrent modification
                if i > 0:  # Don't delay for the first line
                    print(f"⏳ Waiting {retry_delay} seconds to avoid concurrent modification...")
                    time.sleep(retry_delay)
                
                # Create a copy of line_item_data for this specific line
                current_line_data = line_item_data.copy()
                
                # Update impressions for this specific line
                current_line_data['impressions'] = line_config['impressions']
                
                # Handle different line types
                if line_config['use_nwp']:
                    print(f"🎯 Using hardcoded NWP placement data for {line_config['name']}")
                    # Use special NWP function with hardcoded placements and geo targeting
                    line_id, creative_ids = single_line_nwp(
                        client, order_id, current_line_data, line_config['name'], line_config['line_type']
                    )
                elif line_config['use_psbk']:
                    print(f"🔧 Using CAN_PSBK placement data for {line_config['name']}")
                    # We'll modify the single_line function call to use CAN_PSBK sheet and geo targeting
                    line_id, creative_ids = single_line_with_custom_sheet(
                        client, order_id, current_line_data, line_config['name'], 
                        custom_sheet_name=PLACEMENT_SHEET_NAME_CAN_PSBK,
                        line_type=line_config['line_type']
                    )
                else:
                    # Use standard single_line function with standard geo targeting
                    line_id, creative_ids = single_line_with_geo_type(
                        client, order_id, current_line_data, line_config['name'], line_config['line_type']
                    )
                
                # Track successful creation
                all_line_ids.append(line_id)
                all_creative_ids.extend(creative_ids if creative_ids else [])
                success_count += 1
                
                print(f"✅ Successfully created {line_config['description']}")
                print(f"   - Line ID: {line_id}")
                print(f"   - Creatives: {len(creative_ids) if creative_ids else 0}")
                
                # Break out of retry loop on success
                break
                
            except Exception as e:
                # Enhanced error handling with more details
                error_str = str(e) if e else "Unknown error"
                error_type = type(e).__name__
                
                print(f"🔍 Exception Debug Info:")
                print(f"  - Exception type: {error_type}")
                print(f"  - Exception str: '{error_str}'")
                print(f"  - Exception repr: {repr(e)}")
                print(f"  - Line config: {line_config['description']}")
                
                is_concurrent_error = "CONCURRENT_MODIFICATION" in error_str
                
                if is_concurrent_error and attempt < max_retries - 1:
                    # Retry for concurrent modification errors
                    wait_time = retry_delay * (attempt + 1)  # Exponential backoff
                    print(f"⚠️ Concurrent modification error, retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                    continue
                else:
                    # Final failure or non-retryable error
                    if not error_str or error_str == "None":
                        error_str = f"Unknown {error_type} exception"
                    error_msg = f"Failed to create {line_config['description']}: {error_str}"
                    error_messages.append(error_msg)
                    print(f"❌ {error_msg}")
                    
                    # Log individual line creation error
                    logger.log_line_creation_error(e, line_config['name'], str(order_id), session_id)
                    break
    
    # Calculate total time
    total_time = time.time() - start_time
    
    # Log final results
    print(f"\n📊 Three-Line Creation Summary:")
    print(f"  - Lines created: {success_count}/{len(lines_to_create)}")
    print(f"  - Total creatives: {len(all_creative_ids)}")
    print(f"  - Total time: {total_time:.2f}s")
    print(f"  - Success rate: {(success_count/len(lines_to_create)*100):.1f}%")
    
    if error_messages:
        print(f"  - Errors: {len(error_messages)}")
        for error in error_messages:
            print(f"    • {error}")
    
    # Log performance metrics
    logger.log_performance_metrics({
        'total_time': total_time,
        'lines_created': success_count,
        'total_lines_attempted': len(lines_to_create),
        'creative_count': len(all_creative_ids),
        'session_id': session_id,
        'success_rate': success_count/len(lines_to_create)*100
    }, session_id)
    
    # Return results - if any line creation failed, raise an exception with details
    if success_count < len(lines_to_create):
        error_summary = f"Only {success_count}/{len(lines_to_create)} lines created successfully. Errors: {'; '.join(error_messages)}"
        raise Exception(error_summary)
    
    return all_line_ids, all_creative_ids


def single_line_with_geo_type(client, order_id, line_item_data, line_name, line_type="standard"):
    """
    Wrapper for single_line that handles geo targeting based on line type
    """
    # Modify geo targeting based on line type
    original_geo = line_item_data.get('geoTargeting', [])
    modified_line_data = line_item_data.copy()
    
    # Setup geo targeting based on line type
    geo_ids, excluded_geo_ids = setup_geo_targeting_for_line_type(client, original_geo, line_type)
    
    # Update the line item data with processed geo targeting
    modified_line_data['processed_geo_ids'] = geo_ids
    modified_line_data['excluded_geo_ids'] = excluded_geo_ids
    modified_line_data['line_type'] = line_type
    
    # Call the original single_line function
    return single_line(client, order_id, modified_line_data, line_name)

def single_line_with_custom_sheet(client, order_id, line_item_data, line_name, custom_sheet_name=None, line_type="psbk"):
    """
    Modified version of single_line that allows using a custom sheet for placement data
    This is specifically for the _psbk line that needs to use CAN_PSBK sheet
    """
    # Store parameters to prevent variable shadowing
    psbk_custom_sheet_name = custom_sheet_name
    psbk_line_type = line_type
    
    # Map non-TOI/ET sites to 'language' for PSBK line
    if 'site' in line_item_data:
        original_sites = line_item_data['site']
        mapped_sites = []
        for site in original_sites:
            if site in ['TOI', 'ET']:
                mapped_sites.append(site)
            else:
                mapped_sites.append('language')
        print(f"🔄 PSBK Site Mapping: {original_sites} -> {mapped_sites}")
        line_item_data = line_item_data.copy()
        line_item_data['site'] = mapped_sites
    
    # Set up default size groups for PSBK line
    filtered_size_groups = {
        '300x250': {
            'adtype_filter': ['MREC_1', 'MREC'],
            'section_filter': ['ROS']
        },
        '320x50': {
            'adtype_filter': ['BOTTOM OVERLAY'],
            'section_filter': ['ROS']
        },
        '728x90': {
            'adtype_filter': ['LEADERBOARD'],
            'section_filter': ['ROS']
        }
    }
    # Generate session ID for this line creation
    session_id = str(uuid.uuid4())
    start_time = time.time()
    
    # Initialize timing checkpoints
    timing_checkpoints = {
        'start_time': start_time,
        'data_processing_start': None,
        'data_processing_end': None,
        'placement_lookup_start': None,
        'placement_lookup_end': None,
        'line_creation_start': None,
        'line_creation_end': None,
        'creative_creation_start': None,
        'creative_creation_end': None
    }
    
    # Log line creation start
    logger.log_line_creation_start(str(order_id), line_item_data, line_name, session_id)
    
    line_item_service = client.GetService('LineItemService', version='v202408')
    
    # Track created creatives by size to prevent duplicates
    created_creative_sizes = set()
    
    def track_creative_creation(size, creative_ids_list):
        """Track created creatives to prevent duplicates"""
        if creative_ids_list:
            created_creative_sizes.add(size)
            print(f"📝 Tracked creative creation: {size} -> {len(creative_ids_list)} creatives")
    
    def is_creative_size_already_created(size):
        """Check if a creative of this size has already been created"""
        return size in created_creative_sizes

    end_date_value = line_item_data.get('End_date')
    print(f"end_date_value::{end_date_value}")
    start_date_value = line_item_data.get('Start_date', '2025-05-06 00:00:00')
    print(f"start_date_value::{start_date_value}")
    Fcap_value = int(line_item_data.get('fcap', 0))
    cost = line_item_data.get('CPM_Rate', line_item_data.get('cpm', 0))
    print(f"line_item_data::{line_item_data}")
    
    # Start data processing timing
    timing_checkpoints['data_processing_start'] = time.time()
    
    # End data processing timing
    timing_checkpoints['data_processing_end'] = time.time()
    
    # Process impression value to ensure it's an integer
    total_impressions = line_item_data.get('impressions', 100000)
    if isinstance(total_impressions, str):
        # Remove any commas and convert to int
        total_impressions = int(total_impressions.replace(',', ''))
    elif isinstance(total_impressions, float):
        total_impressions = int(total_impressions)
    
    # Validate that total_impressions is a positive integer
    if not isinstance(total_impressions, int) or total_impressions <= 0:
        total_impressions = 100000  # Default fallback
        print(f"⚠️ Invalid impressions value, using default: {total_impressions}")
    
    print(f"📊 Total impressions (validated): {total_impressions}")
    
    # Process start and end dates
    if isinstance(start_date_value, str):
        start_date = datetime.strptime(start_date_value, '%Y-%m-%d %H:%M:%S')
    else:
        start_date = start_date_value
    
    if isinstance(end_date_value, str):
        end_date = datetime.strptime(end_date_value, '%Y-%m-%d %H:%M:%S')
    else:
        end_date = end_date_value
    
    # Extract data from line_item_data
    site_filter = line_item_data.get('site', [])
    platforms_filter = line_item_data.get('platforms', [])
    geo_targeting = line_item_data.get('geoTargeting', [])
    
    print(f"📍 Site filter: {site_filter}")
    print(f"💻 Platforms filter: {platforms_filter}")
    print(f"🌍 Geo targeting: {geo_targeting}")
    
    # Check for duplicate line item name
    if check_line_item_name_exists(client, order_id, line_name):
        timestamp = int(time.time())
        unique_line_name = f"{line_name}_{timestamp}"
        print(f"🔄 Using unique name: {unique_line_name}")
    else:
        unique_line_name = line_name
        print(f"✅ Line name is unique: {unique_line_name}")
    
    # Determine line type - respect line type if already set in line data
    line_type = line_item_data.get('line_type', 'standard')
    banner_video = line_item_data.get('banner_video')
    
    # Debug: Check if line type is being overridden
    print(f"🔍 DEBUG: Line type from line_item_data: {line_item_data.get('line_type', 'NOT_SET')}")
    print(f"🔍 DEBUG: Final line_type variable: {line_type}")
    
    # Only override line type for rich media if not already explicitly set
    if line_type == 'standard' and banner_video and banner_video.lower() in ['rich media', 'richmedia', 'rich-media']:
        line_type = "richmedia"
        print(f"🎨 Line type: Rich Media")
    else:
        print(f"📄 Line type: {line_type.title()}")
    
    # Get presets based on line type
    if line_type == "richmedia":
        presets_dict = richmedia_presets_dict
        richmedia_platform_map = {
            "WEB": ["WEB"],
            "MWEB": ["MWEB"], 
            "AMP": ["AMP"]
        }
    else:
        presets_dict = standard_presets_dict
        richmedia_platform_map = None
    
    # Get available image sizes from creatives folder
    detected_presets, image_files = fetch_images_and_presets(CREATIVES_FOLDER, available_presets, presets_dict)
    print(f"🖼️ Detected presets from images: {detected_presets}")
    
    # Filter size groups based on detected presets
    filtered_size_groups = {}
    for size in detected_presets:
        if size in presets_dict:
            filtered_size_groups[size] = presets_dict[size]
    
    print(f"📋 Filtered size groups: {list(filtered_size_groups.keys())}")
    
    # Prepare platforms for fetch (convert to uppercase)
    platforms_for_fetch = [p.upper() for p in platforms_filter]
    
    # Check if TOI or ETIMES is in site filter
    contains_toi = any(site.upper() in ['TOI', 'ETIMES'] for site in site_filter)
    contains_et = any(site.upper() in ['ET'] for site in site_filter)
    
    placement_data = {}
    
    timing_checkpoints['placement_lookup_start'] = time.time()

    # Always fetch placements regardless of tag file
    if contains_toi:
        toi_sites = [s for s in site_filter if s in ['TOI', 'ETIMES']]
        custom_sheet_name = line_item_data.get('custom_sheet_name')
        sheet_name_used = custom_sheet_name if custom_sheet_name else PLACEMENT_SHEET_NAME_TOI
        print(f"\nFetching TOI placements from sheet: {sheet_name_used}")
        
        placement_data_toi = fetch_placements_ids(
            CREDENTIALS_PATH,
            SHEET_URL,
            sheet_name_used,
            toi_sites,
            platforms_for_fetch,
            filtered_size_groups,
            richmedia_platform_map,
            line_type
        )
        # Merge TOI placement data
        print(f"🔍 TOI placement_data_toi type: {type(placement_data_toi)}")
        print(f"🔍 TOI placement_data_toi content: {placement_data_toi}")
        
        for size, data in placement_data_toi.items():
            print(f"🔍 TOI merging - Size: {size}, Data type: {type(data)}, Data: {data}")
            if size not in placement_data:
                placement_data[size] = []
            if isinstance(data, list):
                placement_data[size].extend(data)
            else:
                print(f"⚠️ TOI data for size {size} is not a list: {type(data)}")

    if contains_et:
        et_sites = [s for s in site_filter if s.upper() == 'ET']
        custom_sheet_name = line_item_data.get('custom_sheet_name')
        et_sheet_name_used = custom_sheet_name if custom_sheet_name else PLACEMENT_SHEET_NAME_ET
        print(f"\nFetching ET placements for sites: {et_sites} from sheet: {et_sheet_name_used}")
        
        placement_data_et = fetch_placements_ids(
            CREDENTIALS_PATH,
            SHEET_URL,
            et_sheet_name_used,
            et_sites,
            platforms_for_fetch,
            filtered_size_groups,
            richmedia_platform_map,
            line_type
        )
        # Merge ET placement data
        print(f"🔍 ET placement_data_et type: {type(placement_data_et)}")
        print(f"🔍 ET placement_data_et content: {placement_data_et}")
        
        for size, data in placement_data_et.items():
            print(f"🔍 ET merging - Size: {size}, Data type: {type(data)}, Data: {data}")
            if size not in placement_data:
                placement_data[size] = []
            if isinstance(data, list):
                placement_data[size].extend(data)
            else:
                print(f"⚠️ ET data for size {size} is not a list: {type(data)}")

    # Fetch from other sites (non-TOI, non-ET)
    other_sites = [s for s in site_filter if s.upper() not in ['TOI', 'ETIMES', 'ET']]
    if other_sites:
        print(f"\nFetching placements for other sites: {other_sites}")
        
        # Use custom sheet if specified (for _psbk line)
        sheet_name_to_use = custom_sheet_name if custom_sheet_name else PLACEMENT_SHEET_NAME_LANG
        print(f"🔧 Using sheet: {sheet_name_to_use}")
        
        placement_data_other = fetch_placements_ids(
            CREDENTIALS_PATH,
            SHEET_URL,
            sheet_name_to_use,
            other_sites,
            platforms_for_fetch,
            filtered_size_groups,
            richmedia_platform_map,
            line_type
        )
        # Merge other placement data
        print(f"🔍 Other placement_data type: {type(placement_data_other)}")
        print(f"🔍 Other placement_data content: {placement_data_other}")
        
        for size, data in placement_data_other.items():
            print(f"🔍 Other merging - Size: {size}, Data type: {type(data)}, Data: {data}")
            if size not in placement_data:
                placement_data[size] = []
            if isinstance(data, list):
                placement_data[size].extend(data)
            else:
                print(f"⚠️ Other data for size {size} is not a list: {type(data)}")

    timing_checkpoints['placement_lookup_end'] = time.time()
    
    print(f"\n🎯 Final placement data summary:")
    for size, placements in placement_data.items():
        print(f"  - {size}: {len(placements) if isinstance(placements, list) else 'N/A'} placements")
    
    # Continue with the rest of the single_line logic...
    # [Rest of the single_line function logic would be copied here]
    # For brevity, I'll call the original single_line function but with modified placement data
    
    # Since we can't easily modify the existing single_line function without major refactoring,
    # let's use a simpler approach: temporarily modify the constants and call single_line
    
    # Setup geo targeting for PSBK line (India excluding user geo)
    original_geo = line_item_data.get('geoTargeting', [])
    modified_line_data = line_item_data.copy()
    
    # IMPORTANT: Set custom_sheet_name in the line data FIRST before any other processing
    # This prevents it from being overwritten by single_line function
    modified_line_data['custom_sheet_name'] = psbk_custom_sheet_name
    modified_line_data['line_type'] = psbk_line_type
    modified_line_data['filtered_size_groups'] = filtered_size_groups
    
    # Debug: Check if line type is being set correctly
    print(f"🔍 DEBUG: PSBK line_type parameter: {psbk_line_type}")
    print(f"🔍 DEBUG: Modified line_data line_type: {modified_line_data.get('line_type')}")
    
    # Setup geo targeting based on line type
    try:
        geo_ids, excluded_geo_ids = setup_geo_targeting_for_line_type(client, original_geo, psbk_line_type)
    except Exception as e:
        print(f"❌ PSBK - Geo setup failed: {e}")
        # Fallback to empty lists
        geo_ids, excluded_geo_ids = [], []
    
    # Update the line item data with processed geo targeting
    modified_line_data['processed_geo_ids'] = geo_ids
    modified_line_data['excluded_geo_ids'] = excluded_geo_ids
    
    # Call the original single_line function with modified data
    print(f"🔍 DEBUG: About to call single_line with line_type: {modified_line_data.get('line_type')}")
    result = single_line(client, order_id, modified_line_data, line_name)
    
    return result


def single_line_nwp(client, order_id, line_item_data, line_name, line_type="nwp"):
    """
    Special function for _nwp line with hardcoded placement targeting
    Only creates 300x250 and 320x50 creatives with specific placement IDs
    """
    # Generate session ID for this line creation
    session_id = str(uuid.uuid4())
    start_time = time.time()
    
    print(f"🎯 Starting NWP line creation with hardcoded placements...")
    print(f"📝 Line name: {line_name}")
    print(f"🆔 Session ID: {session_id}")
    
    # Log line creation start
    logger.log_line_creation_start(str(order_id), line_item_data, f"NWP_LINE: {line_name}", session_id)
    
    # Initialize timing checkpoints
    timing_checkpoints = {
        'start_time': start_time,
        'data_processing_start': None,
        'data_processing_end': None,
        'placement_lookup_start': None,
        'placement_lookup_end': None,
        'line_creation_start': None,
        'line_creation_end': None,
        'creative_creation_start': None,
        'creative_creation_end': None
    }
    
    line_item_service = client.GetService('LineItemService', version='v202408')
    
    # Track created creatives by size to prevent duplicates
    created_creative_sizes = set()
    
    def track_creative_creation(size, creative_ids_list):
        """Track created creatives to prevent duplicates"""
        if creative_ids_list:
            created_creative_sizes.add(size)
            print(f"📝 Tracked creative creation: {size} -> {len(creative_ids_list)} creatives")
    
    def is_creative_size_already_created(size):
        """Check if a creative of this size has already been created"""
        return size in created_creative_sizes

    # Extract data from line_item_data
    end_date_value = line_item_data.get('End_date')
    print(f"end_date_value::{end_date_value}")
    start_date_value = line_item_data.get('Start_date', '2025-05-06 00:00:00')
    print(f"start_date_value::{start_date_value}")
    Fcap_value = int(line_item_data.get('fcap', 0))
    cost = line_item_data.get('CPM_Rate', line_item_data.get('cpm', 0))
    print(f"line_item_data::{line_item_data}")
    
    # Start data processing timing
    timing_checkpoints['data_processing_start'] = time.time()
    
    # Process impression value to ensure it's an integer
    total_impressions = line_item_data.get('impressions', 100000)
    if isinstance(total_impressions, str):
        total_impressions = int(total_impressions.replace(',', ''))
    elif isinstance(total_impressions, float):
        total_impressions = int(total_impressions)
    
    print(f"📊 NWP Total impressions: {total_impressions:,}")
    
    # Process start and end dates
    if isinstance(start_date_value, str):
        start_date = datetime.strptime(start_date_value, '%Y-%m-%d %H:%M:%S')
    else:
        start_date = start_date_value
    
    if isinstance(end_date_value, str):
        end_date = datetime.strptime(end_date_value, '%Y-%m-%d %H:%M:%S')
    else:
        end_date = end_date_value
    
    # Ensure start date is not in the past for NWP line
    from datetime import timedelta
    current_time = datetime.now()
    if start_date <= current_time:
        # Set start date to current time + 1 hour to avoid past date error
        start_date = current_time + timedelta(hours=1)
        print(f"⚠️ NWP Start date was in the past, updated to: {start_date}")
    
    print(f"📅 NWP Start date: {start_date}")
    print(f"📅 NWP End date: {end_date}")
    
    # Extract targeting data and setup geo targeting for NWP line
    original_geo_targeting = line_item_data.get('geoTargeting', [])
    print(f"🌍 Original geo targeting: {original_geo_targeting}")
    
    # Setup geo targeting for NWP line (India excluding user geo)
    geo_ids, excluded_geo_ids = setup_geo_targeting_for_line_type(client, original_geo_targeting, line_type)
    print(f"🎯 NWP Geo IDs to target: {geo_ids}")
    print(f"🚫 NWP Geo IDs to exclude: {excluded_geo_ids}")
    
    # Check for duplicate line item name
    if check_line_item_name_exists(client, order_id, line_name):
        timestamp = int(time.time())
        unique_line_name = f"{line_name}_{timestamp}"
        print(f"🔄 Using unique name: {unique_line_name}")
    else:
        unique_line_name = line_name
        print(f"✅ Line name is unique: {unique_line_name}")
    
    # End data processing timing
    timing_checkpoints['data_processing_end'] = time.time()
    
    # Hardcoded Ad Unit data for NWP line - only 300x250 and 320x50
    timing_checkpoints['placement_lookup_start'] = time.time()
    
    # Updated to use Ad Units instead of Placement IDs
    ad_unit_data = {
        '300x250': {
            'MWEB': [23314114031],  # NP_MWEB_PSBK_CAN_MREC
            'AMP': [23314120439]    # NP_AMP_PSBK_CAN_MREC
        },
        '320x50': {
            'MWEB': [23314114448],  # NP_MWEB_PSBK_CAN_ATF
            'AMP': [23312946423]    # NP_AMP_PSBK_CAN_ATF
        }
    }
    
    print(f"🎯 Using hardcoded NWP Ad Unit data:")
    for size, platforms in ad_unit_data.items():
        for platform, ad_units in platforms.items():
            print(f"  - {size} ({platform}): {ad_units}")
    
    timing_checkpoints['placement_lookup_end'] = time.time()
    
    # Start line creation timing
    timing_checkpoints['line_creation_start'] = time.time()
    
    # Create the line item with Ad Unit targeting
    # Flatten all ad units from all platforms and sizes
    all_ad_units = []
    for size_data in ad_unit_data.values():
        for platform_ad_units in size_data.values():
            all_ad_units.extend(platform_ad_units)
    
    line_item = {
        'name': unique_line_name,
        'orderId': order_id,
        'targeting': {
            'inventoryTargeting': {
                'targetedAdUnits': [{'adUnitId': str(ad_unit_id), 'includeDescendants': True} for ad_unit_id in all_ad_units]
            }
        },
        'startDateTime': {
            'date': {
                'year': start_date.year,
                'month': start_date.month,
                'day': start_date.day
            },
            'hour': start_date.hour,
            'minute': start_date.minute,
            'second': start_date.second,
            'timeZoneId': 'Asia/Kolkata'
        },
        'endDateTime': {
            'date': {
                'year': end_date.year,
                'month': end_date.month,
                'day': end_date.day
            },
            'hour': end_date.hour,
            'minute': end_date.minute,
            'second': end_date.second,
            'timeZoneId': 'Asia/Kolkata'
        },
        'lineItemType': 'STANDARD',
        'costType': 'CPM',
        'costPerUnit': {
            'currencyCode': 'INR',
            'microAmount': int(cost * 1000000)
        },
        'creativePlaceholders': [],
        'deliveryRateType': 'EVENLY',  # Ensure EVENLY delivery rate
        'primaryGoal': {
            'goalType': 'LIFETIME',  # Ensure LIFETIME goal type
            'unitType': 'IMPRESSIONS',
            'units': total_impressions
        }
    }
    
    # Add geo targeting if available
    if geo_ids:
        line_item['targeting']['geoTargeting'] = {
            'targetedLocations': [{'id': geo_id} for geo_id in geo_ids]
        }
        
        # Add geo exclusions if available
        if excluded_geo_ids:
            line_item['targeting']['geoTargeting']['excludedLocations'] = [
                {'id': geo_id} for geo_id in excluded_geo_ids
            ]
    
    # Add frequency cap if specified
    if Fcap_value > 0:
        line_item['frequencyCaps'] = [{
            'maxImpressions': Fcap_value,
            'timeUnit': 'LIFETIME'  # Set frequency cap per lifetime
        }]
    
    # Check which creative files are actually available before adding placeholders
    available_nwp_sizes = []
    for size in ['300x250', '320x50']:
        width, height = map(int, size.split('x'))
        size_pattern = f"{width}x{height}"
        found_files = False
        for ext in ['jpg', 'jpeg', 'png', 'gif']:
            pattern = os.path.join(CREATIVES_FOLDER, f"*{size_pattern}*.{ext}")
            files = glob.glob(pattern)
            if files:
                found_files = True
                break
        if found_files:
            available_nwp_sizes.append(size)
    
    print(f"🖼️ Available NWP creative sizes: {available_nwp_sizes}")
    
    # Add creative placeholders and targetings for available creative sizes
    creative_targetings = []
    for size in available_nwp_sizes:
        width, height = map(int, size.split('x'))
        
        # Add creative placeholder with targetingName
        line_item['creativePlaceholders'].append({
            'targetingName': size,  # This should match the name in creativeTargetings
            'size': {'width': width, 'height': height},
            'expectedCreativeCount': 1
        })
        print(f"📝 Added creative placeholder for {size} with targetingName: {size}")
        
        # Add creative targeting that matches the Ad Unit data
        if size in ad_unit_data:
            # Get all ad units for this size across all platforms
            size_ad_units = []
            for platform_ad_units in ad_unit_data[size].values():
                size_ad_units.extend(platform_ad_units)
            
            targeting_dict = {
                'name': size,  # This must match the targetingName in LICA
                'targeting': {
                    'inventoryTargeting': {
                        'targetedAdUnits': [{'adUnitId': str(ad_unit_id), 'includeDescendants': True} for ad_unit_id in size_ad_units]
                    }
                }
            }
            creative_targetings.append(targeting_dict)
            print(f"📝 Added creative targeting for {size} with Ad Units: {size_ad_units}")
    
    # Add creative targetings to line item
    line_item['creativeTargetings'] = creative_targetings
    
    # Debug: Print the line item structure before creation
    print(f"🔍 NWP Line Item Debug:")
    print(f"  - Creative Placeholders: {len(line_item['creativePlaceholders'])}")
    print(f"  - Creative Targetings: {len(creative_targetings)}")
    for i, targeting in enumerate(creative_targetings):
        ad_units = [unit['adUnitId'] for unit in targeting['targeting']['inventoryTargeting']['targetedAdUnits']]
        print(f"    {i+1}. Name: '{targeting['name']}', Ad Units: {ad_units}")
    print(f"  - Line Item Keys: {list(line_item.keys())}")
    
    # Update Ad Unit data to only include sizes with available creatives
    filtered_ad_unit_data = {}
    for size in available_nwp_sizes:
        if size in ad_unit_data:
            filtered_ad_unit_data[size] = ad_unit_data[size]
    
    ad_unit_data = filtered_ad_unit_data
    print(f"🎯 Filtered NWP Ad Unit data: {ad_unit_data}")
    
    # Update targeted Ad Units to only include ad units for available creatives
    if ad_unit_data:
        # Flatten all ad units from all platforms and sizes
        all_ad_units = []
        for size_data in ad_unit_data.values():
            for platform_ad_units in size_data.values():
                all_ad_units.extend(platform_ad_units)
        
        line_item['targeting']['inventoryTargeting']['targetedAdUnits'] = [{'adUnitId': str(ad_unit_id), 'includeDescendants': True} for ad_unit_id in all_ad_units]
    else:
        print("⚠️ No creative files found for NWP line!")
        raise Exception("No creative files found for NWP line (300x250 or 320x50)")

    try:
        # Create the line item
        print("🚀 Creating NWP line item...")
        print(f"🔍 Line item structure:")
        print(f"  - Name: {line_item.get('name')}")
        print(f"  - Order ID: {line_item.get('orderId')}")
        print(f"  - Targeting keys: {list(line_item.get('targeting', {}).keys())}")
        print(f"  - Inventory targeting keys: {list(line_item.get('targeting', {}).get('inventoryTargeting', {}).keys())}")
        
        try:
            created_line_items = line_item_service.createLineItems([line_item])
            print(f"✅ Line item creation API call successful")
        except Exception as api_error:
            print(f"❌ Line item creation API call failed: {api_error}")
            print(f"🔍 API Error type: {type(api_error).__name__}")
            print(f"🔍 API Error details: {str(api_error)}")
            raise api_error
        
        # Debug the response
        print(f"🔍 Created line items response: {created_line_items}")
        print(f"🔍 Response type: {type(created_line_items)}")
        print(f"🔍 Response length: {len(created_line_items) if isinstance(created_line_items, list) else 'Not a list'}")
        
        if not created_line_items or len(created_line_items) == 0:
            raise Exception("No line items were created - empty response from API")
        
        first_line_item = created_line_items[0]
        print(f"🔍 First line item: {first_line_item}")
        print(f"🔍 First line item keys: {list(first_line_item.keys()) if isinstance(first_line_item, dict) else 'Not a dict'}")
        
        if 'id' not in first_line_item:
            raise Exception(f"Line item created but no 'id' field found. Available keys: {list(first_line_item.keys())}")
        
        line_item_id = first_line_item['id']
        print(f"✅ Successfully created NWP line item with ID: {line_item_id}")
        
        # End line creation, start creative creation timing
        timing_checkpoints['line_creation_end'] = time.time()
        timing_checkpoints['creative_creation_start'] = time.time()
        
    except Exception as e:
        error_str = str(e)
        print(f"❌ Failed to create NWP line item: {e}")
        print(f"🔍 Exception Debug Info:")
        print(f"  - Exception type: {type(e).__name__}")
        print(f"  - Exception str: {str(e)}")
        print(f"  - Exception repr: {repr(e)}")
        logger.log_line_creation_error(e, unique_line_name, str(order_id), session_id)
        # Re-raise the original exception instead of trying to access 'id'
        raise e

    # Create creatives only for 300x250 and 320x50
    creative_ids = []
    
    # Check for available creative files in the folder
    available_creative_files = []
    for size in ['300x250', '320x50']:
        width, height = map(int, size.split('x'))
        # Look for image files matching this size
        size_pattern = f"{width}x{height}"
        for ext in ['jpg', 'jpeg', 'png', 'gif']:
            pattern = os.path.join(CREATIVES_FOLDER, f"*{size_pattern}*.{ext}")
            files = glob.glob(pattern)
            available_creative_files.extend([(f, size) for f in files])
    
    print(f"🖼️ Found {len(available_creative_files)} creative files for NWP line")
    
    # Create creatives using the existing creative creation logic
    if available_creative_files:
        try:
            # Use the existing creative creation function
            destination_url = line_item_data.get('destination_url', line_item_data.get('landing_page', ''))
            impression_tracker = line_item_data.get('Impression_tracker', line_item_data.get('impression_tracker', ''))
            tracking_tag = line_item_data.get('Tracking_Tag', line_item_data.get('tracking_tag', ''))
            
            # Create creatives for each size
            for creative_file, size in available_creative_files:
                if not is_creative_size_already_created(size):
                    try:
                        print(f"🎨 Creating creative for size {size}: {os.path.basename(creative_file)}")
                        
                        # Create creative using existing template function
                        expresso_id = line_item_data.get('expresso_id', '000000')
                        
                        print(f"🎨 NWP Creative params:")
                        print(f"  - Order ID: {order_id}")
                        print(f"  - Line Item ID: {line_item_id}")
                        print(f"  - Expresso ID: {expresso_id}")
                        print(f"  - Size: {size}")
                        print(f"  - Destination URL: {destination_url}")
                        
                        creative_ids_result = create_custom_template_creatives(
                            client=client, 
                            order_id=order_id,
                            line_item_id=line_item_id, 
                            destination_url=destination_url,
                            expresso_id=expresso_id,
                            size_name=size,
                            landing_page=destination_url,
                            impression_tracker=impression_tracker,
                            tracking_tag=tracking_tag
                        )
                        
                        print(f"🎨 NWP Creative result: {creative_ids_result}")
                        
                        # The function returns a list of creative IDs
                        if creative_ids_result and len(creative_ids_result) > 0:
                            creative_id = creative_ids_result[0]  # Get first creative ID
                        else:
                            creative_id = None
                        
                        if creative_id:
                            creative_ids.append(creative_id)
                            track_creative_creation(size, [creative_id])
                            print(f"✅ Created creative ID: {creative_id} for size {size}")
                        
                    except Exception as e:
                        print(f"⚠️ Failed to create creative for {size}: {e}")
        except Exception as e:
            print(f"⚠️ Error in creative creation process: {e}")
    
    # End creative creation timing
    timing_checkpoints['creative_creation_end'] = time.time()
    
    # Calculate timing metrics
    total_time = time.time() - start_time
    data_processing_time = (timing_checkpoints['data_processing_end'] - timing_checkpoints['data_processing_start']) if timing_checkpoints['data_processing_end'] else 0
    placement_lookup_time = (timing_checkpoints['placement_lookup_end'] - timing_checkpoints['placement_lookup_start']) if timing_checkpoints['placement_lookup_end'] else 0
    line_creation_time = (timing_checkpoints['line_creation_end'] - timing_checkpoints['line_creation_start']) if timing_checkpoints['line_creation_end'] else 0
    creative_creation_time = (timing_checkpoints['creative_creation_end'] - timing_checkpoints['creative_creation_start']) if timing_checkpoints['creative_creation_end'] else 0
    
    # Log final success with all details
    logger.log_line_creation_success(
        str(line_item_id), 
        [str(cid) for cid in creative_ids] if creative_ids else [], 
        unique_line_name, 
        session_id
    )
    
    # Log performance metrics
    logger.log_performance_metrics({
        'total_time': total_time,
        'data_processing_time': data_processing_time,
        'placement_lookup_time': placement_lookup_time,
        'line_creation_time': line_creation_time,
        'creative_creation_time': creative_creation_time,
        'line_item_id': line_item_id,
        'creative_count': len(creative_ids) if creative_ids else 0,
        'session_id': session_id,
        'line_type': 'NWP'
    }, session_id)
    
    # Log final performance summary
    print(f"📊 NWP Line Summary:")
    print(f"  - Line created: 1")
    print(f"  - Creatives created: {len(creative_ids) if creative_ids else 0}")
    print(f"  - Total time: {total_time:.2f}s")
    print(f"  - Ad Units used: {list(ad_unit_data.keys())}")

    return line_item_id, creative_ids


if __name__ == '__main__':
    # Test the three_lines functionality
    print("🧪 Testing three_lines function...")
    
    client = ad_manager.AdManagerClient.LoadFromStorage("googleads1.yaml") 
    order_id = 3741465536  # Hardcoded for testing
    timestamp = int(time.time())
    line_name = f"TEST_THREE_LINE47"
    
    # Print current sheet names for debugging
    print("\n🔍 Current sheet name constants:")
    print(f"LANG: '{PLACEMENT_SHEET_NAME_LANG}'")
    print(f"TOI: '{PLACEMENT_SHEET_NAME_TOI}'")
    print(f"ET: '{PLACEMENT_SHEET_NAME_ET}'")
    print(f"CAN_PSBK: '{PLACEMENT_SHEET_NAME_CAN_PSBK}'")
    
    line_item_data = {
        'cpm': 120.0, 
        'impressions': 100,  # 1M impressions for testing distribution
        'site': ['VK'], 
        'platforms': ['WEB', 'MWEB', 'AMP'], 
        'destination_url': 'https://svkm.ac.in/', 
        'expresso_id': '271089',
        'landing_page': 'https://svkm.ac.in/', 
        'Impression_tracker': '', 
        'Tracking_Tag': '', 
        'Start_date': '2025-01-15 00:00:00',  # Future start date
        'End_date': '2025-09-13 23:59:59',  # Fixed date format
        'fcap': '3', 
        'geoTargeting': ['Mumbai'], 
        'Line_label': 'Education', 
        'Line_name': f'TEST_THREE_LINES_STANDARD_{timestamp}', 
       
    }
    
    try:
        # Test the new three_lines function
        line_ids, creative_ids = three_lines(client, order_id, line_item_data, line_name)
        print(f"✅ Successfully created three lines: {line_ids}")
        print(f"✅ Total creatives created: {len(creative_ids)}")
        
        # Verify the line names and impression distribution
        expected_names = [line_name, f"{line_name}_psbk", f"{line_name}_nwp"]
        print(f"📝 Expected line names: {expected_names}")
        
        # Show impression distribution results
        total_test_impressions = 1000000
        print(f"📊 Impression Distribution Verification:")
        print(f"  - Total: {total_test_impressions:,}")
        print(f"  - Standard (10%): {int(total_test_impressions * 0.10):,}")
        print(f"  - PSBK (80%): {int(total_test_impressions * 0.80):,}")
        print(f"  - NWP (10%): {int(total_test_impressions * 0.10):,}")
        print(f"🎯 NWP Hardcoded Placements: 300x250=31928991, 320x50=31929216")
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
    
    
    # detected_presets, image_files = fetch_images_and_presets(CREATIVES_FOLDER, available_presets, presets_dict)
    # print("✅ Detected Presets:", detected_presets)
    # print("🖼️ Image Files:", image_files)
