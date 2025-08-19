#!/usr/bin/env python3
"""
Google Cloud Authentication Helper

This script handles authentication for Google Cloud services using service account credentials.
"""

import os
from google.oauth2.service_account import Credentials
from googleads import ad_manager

def get_ads_client():
    """Get authenticated Google Ads client using service account"""
    try:
        client = ad_manager.AdManagerClient.LoadFromStorage("googleads1.yaml")
        print("✅ Successfully authenticated with Google Ad Manager")
        return client
    except Exception as e:
        print(f"❌ Error authenticating with Google Ad Manager: {e}")
        raise

def setup_authentication():
    """Initialize authentication using service account"""
    try:
        # Check if googleads1.yaml exists
        if not os.path.exists("googleads1.yaml"):
            print("❌ googleads1.yaml not found. Please ensure it exists in the current directory.")
            return False
            
        # Check if credentials.json exists (for sheets API)
        if not os.path.exists("credentials.json"):
            print("❌ credentials.json not found. Please ensure it exists in the current directory.")
            return False
            
        print("✅ Found required authentication files")
        return True
        
    except Exception as e:
        print(f"❌ Error during authentication setup: {e}")
        return False

if __name__ == "__main__":
    setup_authentication() 