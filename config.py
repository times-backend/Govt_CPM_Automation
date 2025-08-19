import os

# Constants
WORKSPACE_ROOT = os.path.dirname(os.path.abspath(__file__))
CREATIVES_FOLDER = os.path.join(WORKSPACE_ROOT, "creatives")
CREDENTIALS_PATH = os.path.join(WORKSPACE_ROOT, "credentials.json")

# Create creatives folder if it doesn't exist
os.makedirs(CREATIVES_FOLDER, exist_ok=True) 