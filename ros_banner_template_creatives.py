import os
import logging
import time
import copy
import re
import requests
from googleads import ad_manager
from get_order_name import fetch_advertiser_id_from_order
from config import CREATIVES_FOLDER

# Add retry and timeout handling
import socket
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter

"""
Creates custom template creatives for various banner sizes and associates them with a line item.

This module provides functionality to create different types of display creatives based on:
- Standard banners (Template ID: 12330939)
- 2x density banners (Template ID: 12459443)
- AI/script-based creatives (Template ID: 12435443)
- No landing page creatives (Template ID: 12399020)
- Expandable creatives (Template ID: 12460223)
- In-Banner Video creatives (Template ID: 12344286)
- 320x100 Special creatives (Template ID: 12363950)
- 300x250 Richmedia creatives (Template ID: 12460223)
- No destination URL creatives (Template ID: 12399020)

The function automatically detects the appropriate template to use based on:
1. Script content from tag files (uses AI template)
2. Banner filename (checks for 'ai', '2x', 'nolp' indicators)
3. Banner size (special handling for 600x250 as expandable, 320x100 as special)
"""

def setup_retry_session():
    """Set up a requests session with retry logic"""
    session = requests.Session()
    retry_strategy = Retry(
        total=3,
        status_forcelist=[429, 500, 502, 503, 504],
        method_whitelist=["HEAD", "GET", "POST", "PUT", "DELETE", "OPTIONS", "TRACE"],
        backoff_factor=1
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

def create_lica_with_retry(lica_service, lica, max_retries=3, initial_delay=2):
    """Create Line Item Creative Association with retry logic"""
    for attempt in range(max_retries):
        try:
            print(f"Attempting to create LICA (attempt {attempt + 1}/{max_retries})...")
            
            # Set a shorter timeout for the specific operation
            original_socket_timeout = socket.getdefaulttimeout()
            socket.setdefaulttimeout(60)  # 60 seconds instead of 3600
            
            result = lica_service.createLineItemCreativeAssociations([lica])
            
            # Restore original timeout
            socket.setdefaulttimeout(original_socket_timeout)
            
            print(f"✅ Successfully created LICA on attempt {attempt + 1}")
            return result
            
        except Exception as e:
            error_msg = str(e)
            print(f"❌ LICA creation attempt {attempt + 1} failed: {error_msg}")
            
            # Restore original timeout in case of error
            if 'original_socket_timeout' in locals():
                socket.setdefaulttimeout(original_socket_timeout)
            
            # Check if it's a timeout or connection error
            if "timeout" in error_msg.lower() or "connection" in error_msg.lower():
                if attempt < max_retries - 1:
                    delay = initial_delay * (2 ** attempt)  # Exponential backoff
                    print(f"⏳ Waiting {delay} seconds before retry...")
                    time.sleep(delay)
                    continue
            
            # If it's not a timeout/connection error or we've exhausted retries
            if attempt == max_retries - 1:
                print(f"⚠️ All {max_retries} attempts failed. Last error: {error_msg}")
                raise e
            
    raise Exception(f"Failed to create LICA after {max_retries} attempts")

def process_html_creative(html_path, landing_page_url, impression_tracker=None):
    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    # 1. Replace all instances of the default URL with the landing page URL
    if landing_page_url:
        html_content = html_content.replace('https://www.google.co.in', landing_page_url)

    # 2. Inject impression tracker after <!--NO_REFRESH-->
    if impression_tracker:
        tracker_block = f'<!--NO_REFRESH-->' + '\n' + '<div style="display:none;">' + f'\n<IMG SRC="{impression_tracker}" attributionsrc BORDER="0" HEIGHT="1" WIDTH="1" ALT="Advertisement">\n</div>'
        if '<!--NO_REFRESH-->' in html_content:
            html_content = html_content.replace('<!--NO_REFRESH-->', tracker_block, 1)
        else:
            # If <!--NO_REFRESH--> is not present, just prepend the block
            html_content = tracker_block + '\n' + html_content

    # Overwrite the original file
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

def create_custom_template_creatives(client, order_id, line_item_id, destination_url, expresso_id,
                                     size_name, landing_page=None,
                                     impression_tracker=None, script_code=None, template_id=None, In_Banner_video=None, line_type=None, tracking_tag=None):
    """
    Creates custom template creatives and associates them with a line item.
    
    Args:
        client (AdManagerClient): The Google Ad Manager client
        order_id (str): Order ID to associate the creative with
        line_item_id (str): Line item ID to associate the creative with
        destination_url (str): Default click-through URL
        expresso_id (str): Expresso tracking ID
        size_name (str): Size name in format 'WIDTHxHEIGHT' (e.g., '300x250')
        landing_page (str, optional): Landing page URL (falls back to destination_url if not provided)
        impression_tracker (str, optional): Third-party impression tracker URL
        script_code (str, optional): JavaScript or HTML code from tag file for AI template
        template_id (int, optional): Custom template ID to use, overrides auto-detection
        
    Returns:
        list: List of created creative IDs
        
    Raises:
        ValueError: If required fields are missing or no creatives are found
    """
    creative_service = client.GetService('CreativeService', version='v202508')
    lica_service = client.GetService('LineItemCreativeAssociationService', version='v202508')
    logging.info(f"Creating creatives for size: {size_name}")

    # For AI template (12435443), impression/click template (12330939), In-Banner Video template (12344286), 320x100 special template (12363950), 300x250 richmedia template (12460223), no destination template (12473441), and no landing page template (12399020), destination_url is not strictly required
    if template_id in [12435443, 12330939, 12344286, 12363950, 12460223, 12473441, 12399020]:
        if not all([order_id, line_item_id, expresso_id]):
            raise ValueError("Required fields must be provided (order_id, line_item_id, expresso_id)")
    else:
        if not all([order_id, line_item_id, destination_url, expresso_id]):
            raise ValueError("Required fields must be provided (order_id, line_item_id, destination_url, expresso_id)")
     
    advertiser_id = fetch_advertiser_id_from_order(client, order_id)
    base_size = size_name.split('_')[0]
    # Handle both uppercase and lowercase 'x' in dimensions (e.g., "600X250" or "600x250")
    width, height = map(int, base_size.lower().split('x'))
    # Define file extensions
    image_extensions = ('.png', '.jpeg', '.jpg', '.webp', '.gif')
    script_extensions = ('.html', '.xlsx', '.xls')
    valid_extensions = image_extensions + script_extensions
    
    # Check for both 300x250 and 600x250 images
    # has_300x250 = any('300x250' in f.lower() for f in os.listdir(CREATIVES_FOLDER) if f.lower().endswith(valid_extensions))
    # has_600x250 = any('600x250' in f.lower() for f in os.listdir(CREATIVES_FOLDER) if f.lower().endswith(valid_extensions))
    
    banner_files = [f for f in os.listdir(CREATIVES_FOLDER) if base_size in f.lower() and f.lower().endswith(valid_extensions)]
    
    # If no image files found but script_code is provided, we can still create a creative using the AI template
    if not banner_files and script_code and len(script_code.strip()) > 10:
        creative_ids = []
        
        # Generate a unique name for the creative
        timestamp = int(time.time() * 1000)
        unique_creative_name = f"{order_id}_{base_size}_script_{timestamp}"
        
        # If template_id is provided, use it; otherwise choose based on available data
        if template_id:
            current_template_id = template_id
        elif landing_page or impression_tracker:
            # If there's a landing page or impression tracker, use standard template
            current_template_id = 12330939
        else:
            # Default to AI template for script-only creatives
            current_template_id = 12435443
        
        # Use AI template for script-only creative
        template_variables = [
            {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
            {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ScriptCode', 'value': script_code}
        ]
        
        # For template 12330939, add landing page and impression tracker if available
        if current_template_id == 12330939 and landing_page:
            template_variables.append(
                {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page}
            )
        
        if current_template_id == 12330939 and impression_tracker:
            template_variables.append(
                {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'ImpressionTracker', 'value': impression_tracker}
            )
        
        # Add ScriptCode for specified templates if tracking_tag is provided
        if current_template_id in [ 12459443, 12399020, 12460223, 12344286, 12363950] and tracking_tag and tracking_tag.strip():
            template_variables.append(
                {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ScriptCode', 'value': tracking_tag}
            )
        if current_template_id in [12330939] and script_code and script_code.strip():
            template_variables.append(
                {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ScriptCode', 'value': script_code}
            )
        
        template_creative = {
            'xsi_type': 'TemplateCreative',
            'name': unique_creative_name,
            'advertiserId': advertiser_id,
            'creativeTemplateId': current_template_id,
            'size': {'width': width, 'height': height, 'isAspectRatio': False},
            'creativeTemplateVariableValues': template_variables
        }
        
        # Only add destinationUrl if it's not empty (AI template doesn't require it)
        if destination_url and destination_url.strip():
            template_creative['destinationUrl'] = destination_url
        
        try:
            created_creative = creative_service.createCreatives([template_creative])
            if created_creative:
                creative_id = created_creative[0]['id']
                creative_ids.append(creative_id)

                lica = {
                    'creativeId': creative_id,
                    'lineItemId': line_item_id,
                    'targetingName': "Mweb_PPD" if (width == 320 and height == 100) else ("Mrec_ex" if (width == 300 and height == 250 and line_type == "richmedia") else f'{width}x{height}'),
                    'sizes': [{'width': width, 'height': height}]
                }

                create_lica_with_retry(lica_service, lica)
                log_msg = f"✅ Created script-only creative ID: {creative_id} for {base_size} using AI template"
                if impression_tracker:
                    log_msg += " with impression tracker"
                logging.info(log_msg)
                print(log_msg)
                
                # Log creative creation with actual template_id and creative_id
                from logging_utils import logger
                logger.log_creative_creation(
                    template_id=str(current_template_id),
                    creative_id=str(creative_id),
                    size=base_size,
                    asset_files=["Script-only creative"]
                )
                
                return creative_ids
        except Exception as e:
            logging.error(f"⚠️ Failed to create script-only creative for size {base_size}: {str(e)}")
            raise
    
    # Check if In_Banner_video is provided and not empty
    if not banner_files and In_Banner_video and In_Banner_video.strip():
        # Validate that we have either landing_page or destination_url for In-Banner Video template
        if not (landing_page or destination_url):
            raise ValueError("In-Banner Video template requires either landing_page or destination_url")
        
        creative_ids = []
        
        # Generate a unique name for the creative
        timestamp = int(time.time() * 1000)
        unique_creative_name = f"{order_id}_inbanner_video_{timestamp}"
        
        # Use In-Banner video template
        current_template_id = 12344286
        
        template_variables = [
            {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
            {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page or destination_url},
            {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'VideoUrl', 'value': In_Banner_video},
            {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'AutoPlay', 'value': 'Yes'}
        ]
        
        # Add ScriptCode for specified templates if tracking_tag is provided
        if current_template_id in [12330939, 12459443, 12399020, 12460223, 12344286, 12363950] and tracking_tag and tracking_tag.strip():
            template_variables.append(
                {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ScriptCode', 'value': tracking_tag}
            )
        
        template_creative = {
            'xsi_type': 'TemplateCreative',
            'name': unique_creative_name,
            'advertiserId': advertiser_id,
            'creativeTemplateId': current_template_id,
            'size': {'width': 300, 'height': 250, 'isAspectRatio': False},
            'destinationUrl': landing_page or destination_url,
            'creativeTemplateVariableValues': template_variables
        }
        
        try:
            created_creative = creative_service.createCreatives([template_creative])
            if created_creative:
                creative_id = created_creative[0]['id']
                creative_ids.append(creative_id)

                lica = {
                    'creativeId': creative_id,
                    'lineItemId': line_item_id,
                    'targetingName': "Mweb_PPD" if (width == 320 and height == 100) else ("Mrec_ex" if (width == 300 and height == 250 and line_type == "richmedia") else f'{width}x{height}'),
                    'sizes': [{'width': width, 'height': height}]
                }

                create_lica_with_retry(lica_service, lica)
                log_msg = f"✅ Created In-Banner video creative ID: {creative_id} with size {width}x{height}"
                logging.info(log_msg)
                print(log_msg)
                
                # Log creative creation with actual template_id and creative_id
                from logging_utils import logger
                logger.log_creative_creation(
                    template_id=str(current_template_id),
                    creative_id=str(creative_id),
                    size=f"{width}x{height}",
                    asset_files=["In-Banner Video"]
                )
                
                return creative_ids
        except Exception as e:
            logging.error(f"⚠️ Failed to create In-Banner video creative: {str(e)}")
            raise
    
    if not banner_files:
        raise ValueError("No creatives detected in the creatives folder and no valid tag file found")
    
    creative_ids = []

    for banner_filename in banner_files:
        banner_file_path = os.path.join(CREATIVES_FOLDER, banner_filename)
        
        # Determine file type and read accordingly
        is_image_file = banner_filename.lower().endswith(image_extensions)
        is_script_file = banner_filename.lower().endswith(script_extensions)
        
        # Initialize variables
        banner_byte_array = None
        ai_html_content = None
        
        # Read file content based on type
        if is_image_file:
            # Read image files as binary
            with open(banner_file_path, 'rb') as file:
                banner_byte_array = file.read()
        elif is_script_file and banner_filename.lower().endswith('.html'):
            # Process HTML creative before reading
            process_html_creative(banner_file_path, landing_page or destination_url, impression_tracker)
            # Read HTML files as text
            with open(banner_file_path, 'r', encoding='utf-8') as file:
                ai_html_content = file.read()
        else:
            print(f"⚠️ Skipping unsupported file: {banner_filename}")
            continue
        
        # Skip files that don't have required content
        if is_image_file and not banner_byte_array:
            print(f"⚠️ Skipping image file {banner_filename} - no image data")
            continue
        elif is_script_file and not ai_html_content:
            print(f"⚠️ Skipping HTML file {banner_filename} - no HTML content")
            continue
        
        # Generate unique names with timestamp
        timestamp = int(time.time() * 1000)
        unique_creative_name = f"{order_id}_{banner_filename.split('.')[0]}_{timestamp}"
        unique_asset_name = f"{unique_creative_name}.png"
        
        # Detect banner characteristics
        is_ai = 'ai' in banner_filename.lower() or banner_filename.lower().endswith('.html')
        is_2x = '2x' in banner_filename.lower() or '2x' in size_name.lower()
        is_NoLP = 'nolp' in banner_filename.lower() or 'nolp' in size_name.lower()
        
        # Debug logging for 2x detection
        if is_2x:
            print(f"🔍 2x detected for file: {banner_filename}, size: {size_name}")
        if template_id:
            print(f"🔍 Template ID provided: {template_id}")
        
        # Initialize template_creative variable 
        template_creative = None
        
        if base_size == '600x250':
            # Read both images
            banner_300x250 = None
            banner_600x250 = None
            
            for f in os.listdir(CREATIVES_FOLDER):
                if f.lower().endswith(image_extensions):
                    if '300x250' in f.lower():
                        with open(os.path.join(CREATIVES_FOLDER, f), 'rb') as file:
                            banner_300x250 = file.read()
                    elif '600x250' in f.lower():
                        with open(os.path.join(CREATIVES_FOLDER, f), 'rb') as file:
                            banner_600x250 = file.read()
            
            # If we don't have both images, use the available one for both
            if not banner_300x250:
                banner_300x250 = banner_byte_array
            if not banner_600x250:
                banner_600x250 = banner_byte_array
            
            # Generate unique names for both images
            unique_300x250_name = f"{order_id}_300x250_{timestamp}.jpg"
            unique_600x250_name = f"{order_id}_600x250_{timestamp}.webp"
            
            template_variables = [
                {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'SmallBanner',
                 'asset': {'assetByteArray': banner_300x250, 'fileName': unique_300x250_name}},
                {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'BigBanner',
                 'asset': {'assetByteArray': banner_600x250, 'fileName': unique_600x250_name}},
                {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'CookieName', 'value': 'expndo'},
                {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CookieTime', 'value': 1},
                {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
                {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page or destination_url},
                {'xsi_type': 'StringCreativeTemplateVariableValue','uniqueName': 'AutoExpand','value': 'Yes' }
            ]
            template_creative = {
                'xsi_type': 'TemplateCreative',
                'name': unique_creative_name,
                'advertiserId': advertiser_id,
                'creativeTemplateId': 12460223,
                'size': {'width': 300, 'height': 250, 'isAspectRatio': False},
                'destinationUrl': landing_page or destination_url,
                'creativeTemplateVariableValues': template_variables
            }
            current_template_id = 12460223
            
        # Template ID precedence: provided template_id > script_code > filename indicators
        elif template_id:
            # Special case: if template is 12399020 (no destination URL) and individual image has 2x, use 12473441 for this creative
            if template_id == 12399020 and is_2x:
                current_template_id = 12473443
                print(f"🔍 Using 2x template 12473443 (NoLP + 2x case)")
            # Special case: if image has 2x indicator, use 2x template (12459443) regardless of provided template_id
            elif is_2x:
                current_template_id = 12459443
                print(f"🔍 Using 2x template 12459443 (2x override case)")
            else:
                current_template_id = template_id
                print(f"🔍 Using provided template ID: {template_id}")
        elif script_code and len(script_code.strip()) > 10 and not landing_page and not impression_tracker:
            # Only use AI template if there's no landing page or impression tracker
            # If there's a landing page or impression tracker, use standard template instead
            current_template_id = 12435443
        elif is_ai:
            # Use AI template for files with "ai" in filename or if .html file exists for this size
            current_template_id = 12435443
            # For the is_ai case, we keep the existing ai_html_content from .html file
            # But if ai_html_content is None or empty, fall back to standard template
            if not ai_html_content or not ai_html_content.strip():
                print("⚠️ AI template selected but no HTML content found, falling back to standard template")
                current_template_id = 12330939
            
        elif is_NoLP:
            # Use No Landing Page template when specified
            # This is triggered when "nolp" is in the filename or size_name
            current_template_id = 12399020
        
        elif In_Banner_video and In_Banner_video.strip():
            # Use In-Banner Video template when In_Banner_video is provided and not empty
            current_template_id = 12344286
        elif is_2x:
            # Special case: if template is 12399020 (no destination URL) and individual image has 2x, use 12473441 for this creative
            if template_id == 12399020:
                current_template_id = 12473441
                print(f"🔍 Using 2x template 12473441 (NoLP + 2x case, no template_id override)")
            else:
                # Use 2x template for files with "2x" in filename or size_name
                current_template_id = 12459443
                print(f"🔍 Using 2x template 12459443 (2x case, no template_id override)")
        else:
            # Special case: if template is 12399020 (no destination URL) and image doesn't have 2x, keep using 12399020
            if template_id == 12399020:
                current_template_id = 12399020
                print(f"🔍 Using NoLP template 12399020 (default case)")
            else:
                # Default to standard template
                current_template_id = 12330939
                print(f"🔍 Using standard template 12330939 (default case)")

        # Initialize default size overrides
        size_overrides = []

        if width == 1260 and height == 570:
            size_overrides = [
                {'width': 728, 'height': 500},
                {'width': 1320, 'height': 570}
            ]
        elif width == 980 and height == 200:
            size_overrides = [
                {'width': 728, 'height': 90}
            ]
        elif width == 320 and height == 100:
            size_overrides = [
                {'width': 320, 'height': 50}
            ]
         
        # Log the final template selection
        print(f"🔍 Final template selection for {banner_filename}: Template ID {current_template_id}")
        
        # Set up template variables and creative based on template type (only if not already created for 600x250)
        if template_creative is None:
            if current_template_id == 12363950:  # 320x100 Special template
                # For 320x100 special template, we need both small and big banner
                # SmallBanner: Use 320x100 image
                # BigBanner: Look for 320x250 image, fallback to 320x100 if not found
                
                small_banner_bytes = banner_byte_array  # 320x100 image for SmallBanner
                big_banner_bytes = banner_byte_array    # Default fallback
                
                # Look for 320x250 image for BigBanner
                big_banner_size_info = "320x100"  # Default fallback size info
                for f in os.listdir(CREATIVES_FOLDER):
                    if f.lower().endswith(image_extensions) and '320x250' in f.lower():
                        big_banner_path = os.path.join(CREATIVES_FOLDER, f)
                        with open(big_banner_path, 'rb') as file:
                            big_banner_bytes = file.read()
                            big_banner_size_info = "320x250"  # Found 320x250 image
                        break
                
                # Generate unique asset names with size information
                small_banner_asset_name = f"{unique_creative_name}_small_320x100.png"
                big_banner_asset_name = f"{unique_creative_name}_big_{big_banner_size_info}.png"
                
                template_variables = [
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'SmallBanner',
                     'asset': {'assetByteArray': small_banner_bytes, 'fileName': small_banner_asset_name}},
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'BigBanner',
                     'asset': {'assetByteArray': big_banner_bytes, 'fileName': big_banner_asset_name}},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'AutoExpand', 'value': 'Yes'},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'CookieName', 'value': 'expndo'},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CookieTime', 'value': 1},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
                    {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page or destination_url},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'BigCreativeViewTime', 'value': 10000}
                ]
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'destinationUrl': landing_page or destination_url,
                    'creativeTemplateVariableValues': template_variables
                }
            elif current_template_id == 12460223:  # 300x250 Richmedia template
                # For 300x250 richmedia template, we need both small and big banner
                # SmallBanner: Use 300x250 image
                # BigBanner: Look for 600x250 image, fallback to 300x250 if not found
                
                small_banner_bytes = banner_byte_array  # 300x250 image for SmallBanner
                big_banner_bytes = banner_byte_array    # Default fallback
                
                # Look for 600x250 image for BigBanner
                big_banner_size_info = "300x250"  # Default fallback size info
                for f in os.listdir(CREATIVES_FOLDER):
                    if f.lower().endswith(image_extensions) and '600x250' in f.lower():
                        big_banner_path = os.path.join(CREATIVES_FOLDER, f)
                        with open(big_banner_path, 'rb') as file:
                            big_banner_bytes = file.read()
                            big_banner_size_info = "600x250"  # Found 600x250 image
                        break
                
                # Generate unique asset names with size information
                small_banner_asset_name = f"{unique_creative_name}_small_300x250.png"
                big_banner_asset_name = f"{unique_creative_name}_big_{big_banner_size_info}.png"
                
                template_variables = [
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'SmallBanner',
                     'asset': {'assetByteArray': small_banner_bytes, 'fileName': small_banner_asset_name}},
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'BigBanner',
                     'asset': {'assetByteArray': big_banner_bytes, 'fileName': big_banner_asset_name}},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'AutoExpand', 'value': 'Yes'},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'CookieName', 'value': 'expndo'},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CookieTime', 'value': 1},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
                    {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page or destination_url},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'BigCreativeViewTime', 'value': 10000}
                ]
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'destinationUrl': landing_page or destination_url,
                    'creativeTemplateVariableValues': template_variables
                }
            elif current_template_id == 12443458:  # 300x600 Richmedia template
                # For 300x600 richmedia template, we need both small and big banner
                # SmallBanner: Use 300x600 image
                # BigBanner: Look for 450x600 image, fallback to 300x600 if not found
                
                small_banner_bytes = banner_byte_array  # 300x600 image for SmallBanner
                big_banner_bytes = banner_byte_array    # Default fallback
                
                # Look for 450x600 image for BigBanner
                big_banner_size_info = "300x600"  # Default fallback size info
                for f in os.listdir(CREATIVES_FOLDER):
                    if f.lower().endswith(image_extensions) and '450x600' in f.lower():
                        big_banner_path = os.path.join(CREATIVES_FOLDER, f)
                        with open(big_banner_path, 'rb') as file:
                            big_banner_bytes = file.read()
                            big_banner_size_info = "450x600"  # Found 450x600 image
                        break
                
                # Generate unique asset names with size information
                small_banner_asset_name = f"{unique_creative_name}_small_300x600.png"
                big_banner_asset_name = f"{unique_creative_name}_big_{big_banner_size_info}.png"
                
                template_variables = [
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'SmallBanner',
                     'asset': {'assetByteArray': small_banner_bytes, 'fileName': small_banner_asset_name}},
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'BigBanner',
                     'asset': {'assetByteArray': big_banner_bytes, 'fileName': big_banner_asset_name}},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'AutoExpand', 'value': 'Yes'},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'CookieName', 'value': 'expndo'},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CookieTime', 'value': 1},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
                    {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page or destination_url},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'BigCreativeViewTime', 'value': 10000}
                ]
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'destinationUrl': landing_page or destination_url,
                    'creativeTemplateVariableValues': template_variables
                }
            elif current_template_id == 12435443:  # AI HTML template
                html_var_name = get_html_variable_name(client, current_template_id)
                print(f"Using variable '{html_var_name}' for HTML content in template {current_template_id}")
                template_variables = [
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': html_var_name, 'value': ai_html_content},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id}
                ]
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'destinationUrl': landing_page or destination_url,
                    'creativeTemplateVariableValues': template_variables
                }
                
                # Only add destinationUrl if provided (AI template doesn't require it)
                if destination_url and destination_url.strip():
                    template_creative['destinationUrl'] = destination_url
            elif current_template_id == 12344286:  # In-Banner Video template
                # Validate that we have either landing_page or destination_url for In-Banner Video template
                if not (landing_page or destination_url):
                    raise ValueError("In-Banner Video template requires either landing_page or destination_url")
                    
                # The In-Banner Video template requires VideoUrl variable
                template_variables = [
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
                    {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page or destination_url},
                    {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'VideoUrl', 'value': In_Banner_video},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'AutoPlay', 'value': 'Yes'}
                ]
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'destinationUrl': landing_page or destination_url,
                    'creativeTemplateVariableValues': template_variables
                }
            elif current_template_id == 12399020:  # No Landing Page template
                # No Landing Page template requires image data
                if not banner_byte_array:
                    raise ValueError(f"No Landing Page template requires image data for file: {banner_filename}")
                
                template_variables = [
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'Banner',
                     'asset': {'assetByteArray': banner_byte_array, 'fileName': unique_asset_name}},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CreativeWidth', 'value': width},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CreativeHeight', 'value': height},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id}
                ]
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'creativeTemplateVariableValues': template_variables
                }
            elif current_template_id == 12459443:  # 2x template
                # 2x template requires image data
                if not banner_byte_array:
                    raise ValueError(f"2x template requires image data for file: {banner_filename}")
                
                template_variables = [
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'Banner',
                     'asset': {'assetByteArray': banner_byte_array, 'fileName': unique_asset_name}},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CreativeWidth', 'value': width},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CreativeHeight', 'value': height},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
                    {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page or destination_url}
                ]
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'destinationUrl': landing_page or destination_url,
                    'creativeTemplateVariableValues': template_variables
                }
            elif current_template_id == 12473441:  # No destination URL template
                # No destination URL template requires image data
                if not banner_byte_array:
                    raise ValueError(f"No destination URL template requires image data for file: {banner_filename}")
                
                template_variables = [
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'Banner',
                     'asset': {'assetByteArray': banner_byte_array, 'fileName': unique_asset_name}},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CreativeWidth', 'value': width},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CreativeHeight', 'value': height},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id}
                ]
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'creativeTemplateVariableValues': template_variables
                }
            else:  # Default template (12330939)
                # Standard template requires image data
                if not banner_byte_array:
                    raise ValueError(f"Standard template requires image data for file: {banner_filename}")
                
                template_variables = [
                    {'xsi_type': 'AssetCreativeTemplateVariableValue', 'uniqueName': 'Banner',
                     'asset': {'assetByteArray': banner_byte_array, 'fileName': unique_asset_name}},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CreativeWidth', 'value': width},
                    {'xsi_type': 'LongCreativeTemplateVariableValue', 'uniqueName': 'CreativeHeight', 'value': height},
                    {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ExpressoID', 'value': expresso_id},
                    {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'LandingPage', 'value': landing_page or destination_url}
                ]
                
                template_creative = {
                    'xsi_type': 'TemplateCreative',
                    'name': unique_creative_name,
                    'advertiserId': advertiser_id,
                    'creativeTemplateId': current_template_id,
                    'size': {'width': width, 'height': height, 'isAspectRatio': False},
                    'destinationUrl': landing_page or destination_url,
                    'creativeTemplateVariableValues': template_variables
                }

        # Add ImpressionTracker template variable to all templates except AI/HTML (12435443), No Landing Page (12399020), 320x100 Special (12363950), 300x250 Richmedia (12460223), and No destination URL (12473441)
        if impression_tracker and current_template_id not in [12435443, 12399020, 12363950, 12460223, 12473441]:
            template_variables.append(
                {'xsi_type': 'UrlCreativeTemplateVariableValue', 'uniqueName': 'ImpressionTracker', 'value': impression_tracker}
            )

        # Add ScriptCode for specified templates if tracking_tag is provided
        if current_template_id in [12330939, 12459443, 12399020, 12460223, 12344286, 12363950] and tracking_tag and tracking_tag.strip():
            template_variables.append(
                {'xsi_type': 'StringCreativeTemplateVariableValue', 'uniqueName': 'ScriptCode', 'value': tracking_tag}
            )

        try:
            created_creative = creative_service.createCreatives([template_creative])
            if created_creative:
                creative_id = created_creative[0]['id']
                creative_ids.append(creative_id)

                # Determine the correct targeting name to match line item creative targeting
                if width == 320 and height == 100:
                    targeting_name = "Mweb_PPD"  # Must match the targeting name used in line item creation
                elif width == 300 and height == 250 and line_type == "richmedia":
                    targeting_name = "Mrec_ex"  # Must match the targeting name used in line item creation for 300x250 richmedia
                else:
                    targeting_name = f'{width}x{height}'

                # Include both the original size and any override sizes
                sizes_for_lica = [{'width': width, 'height': height}]
                if size_overrides:
                    sizes_for_lica.extend(size_overrides)
                
                lica = {
                    'creativeId': creative_id,
                    'lineItemId': line_item_id,
                    'targetingName': targeting_name,
                    'sizes': sizes_for_lica
                }

                create_lica_with_retry(lica_service, lica)
                log_msg = f"✅ Created creative ID: {creative_id} for {base_size} with targeting name: {targeting_name}"
                if impression_tracker:
                    log_msg += " with impression tracker"
                logging.info(log_msg)
                print(log_msg)
                
                # Log creative creation with actual template_id and creative_id
                from logging_utils import logger
                asset_files = [banner_filename] if banner_filename else []
                logger.log_creative_creation(
                    template_id=str(current_template_id),
                    creative_id=str(creative_id),
                    size=base_size,
                    asset_files=asset_files
                )
        except Exception as e:
            logging.error(f"⚠️ Failed to create creatives for size {base_size}: {str(e)}")

    return creative_ids

def get_html_variable_name(client, template_id):
    creative_template_service = client.GetService('CreativeTemplateService', version='v202508')
    # Use a filter statement to fetch the template by ID
    filter_statement = {
        'query': 'WHERE id = :id',
        'values': [
            {'key': 'id', 'value': {'xsi_type': 'NumberValue', 'value': template_id}}
        ]
    }
    response = creative_template_service.getCreativeTemplatesByStatement(filter_statement)
    templates = getattr(response, 'results', [])
    if not templates:
        print(f"⚠️ No template found with ID {template_id}")
        return 'ScriptCode'  # Common fallback
    
    template = templates[0]
    template_variables = getattr(template, 'variables', [])
    
    # Debug: Print all available variables
    print(f"🔍 Available variables in template {template_id}:")
    for var in template_variables:
        var_name = getattr(var, 'uniqueName', 'Unknown')
        var_type = getattr(var, 'type', 'Unknown')
        print(f"  - {var_name} ({var_type})")
    
    # Look for HTML-related variables first
    for var in template_variables:
        var_name = getattr(var, 'uniqueName', '')
        var_type = getattr(var, 'type', '')
        if var_type == 'StringCreativeTemplateVariable' and 'html' in var_name.lower():
            print(f"✅ Found HTML variable: {var_name}")
            return var_name
    
    # Look for ScriptCode variable (common for AI templates)
    for var in template_variables:
        var_name = getattr(var, 'uniqueName', '')
        var_type = getattr(var, 'type', '')
        if var_type == 'StringCreativeTemplateVariable' and var_name == 'ScriptCode':
            print(f"✅ Found ScriptCode variable: {var_name}")
            return var_name
    
    # If no HTML or ScriptCode found, use the first StringCreativeTemplateVariable
    for var in template_variables:
        var_name = getattr(var, 'uniqueName', '')
        var_type = getattr(var, 'type', '')
        if var_type == 'StringCreativeTemplateVariable':
            print(f"✅ Using first string variable: {var_name}")
            return var_name
    
    # Last resort fallback
    print(f"⚠️ No suitable variable found, using ScriptCode as fallback")
    return 'ScriptCode'


# 🔍 Example usage
if __name__ == '__main__':
    client = ad_manager.AdManagerClient.LoadFromStorage("googleads1.yaml")
    order_id = "3741471098"
    line_item_id = "6976336920"
    destination_url = "https://www.phoenixmarketcity.com/chennai"
    expresso_id = "EXP123"
    landing_page = "https://www.phoenixmarketcity.com/chennai"
    impression_tracker = ""
    script = ' '
    script_code = ''
    ai_script = '<html>fff</html>'


    creative_ids = create_custom_template_creatives(
        client, order_id, line_item_id,
        destination_url, expresso_id,
        '600X250', 12399020,
        landing_page, impression_tracker,
        script_code, ai_script
    )

    print(f"✅ Created creatives: {creative_ids}")
