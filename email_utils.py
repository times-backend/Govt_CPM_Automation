#!/usr/bin/env python3
"""
Email detection utility that gets the user's email from Chrome browser data
"""

import os
import json
import platform
import sqlite3
import base64
from pathlib import Path

def get_chrome_user_data_dir():
    """Get Chrome's user data directory based on OS"""
    system = platform.system()
    if system == "Darwin":  # macOS
        return os.path.expanduser("~/Library/Application Support/Google/Chrome/Default")
    elif system == "Windows":
        return os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
    elif system == "Linux":
        return os.path.expanduser("~/.config/google-chrome/Default")
    else:
        raise OSError(f"Unsupported operating system: {system}")

def get_chrome_email():
    """Get email from Chrome's logged-in account"""
    try:
        chrome_dir = get_chrome_user_data_dir()
        
        # Try to get email from Preferences file
        preferences_file = os.path.join(chrome_dir, "Preferences")
        if os.path.exists(preferences_file):
            with open(preferences_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Look for email in account info
                account_info = data.get('account_info', [])
                if account_info:
                    email = account_info[0].get('email')
                    if email and '@' in email:
                        print(f"🌐 Found Chrome account: {email}")
                        return email
                
                # Also check sync accounts
                sync_accounts = data.get('sync', {}).get('account_info', {}).get('email')
                if sync_accounts and '@' in sync_accounts:
                    print(f"🌐 Found Chrome sync account: {sync_accounts}")
                    return sync_accounts
        
        # Try to get email from Web Data
        web_data_file = os.path.join(chrome_dir, "Web Data")
        if os.path.exists(web_data_file):
            # Create a copy of the database since Chrome might have it locked
            import shutil
            temp_file = os.path.join(os.path.dirname(web_data_file), "Web_Data_temp")
            shutil.copy2(web_data_file, temp_file)
            
            try:
                conn = sqlite3.connect(temp_file)
                cursor = conn.cursor()
                cursor.execute("SELECT email FROM autofill_profile_emails")
                results = cursor.fetchall()
                conn.close()
                
                if results:
                    email = results[0][0]
                    if email and '@' in email:
                        print(f"🌐 Found Chrome autofill email: {email}")
                        return email
            except:
                pass
            finally:
                try:
                    os.remove(temp_file)
                except:
                    pass
        
        return None
        
    except Exception as e:
        print(f"⚠️ Error getting Chrome email: {e}")
        return None

def get_default_email_with_fallback():
    """Get user's email with fallbacks"""
    # Try Chrome first
    chrome_email = get_chrome_email()
    if chrome_email:
        print(f"📧 Auto-detected email: {chrome_email}")
        return chrome_email
    
    # Fallback to system username
    try:
        import getpass
        username = getpass.getuser()
        default_email = f"{username}@timesinternet.in"
        print(f"📧 Using system username email: {default_email}")
        return default_email
    except:
        # Final fallback
        fallback_email = "user@timesinternet.in"
        print(f"📧 Using fallback email: {fallback_email}")
        return fallback_email

if __name__ == "__main__":
    email = get_default_email_with_fallback()
    print(f"Detected email: {email}") 