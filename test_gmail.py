"""
Test Gmail API connection
Run this to verify everything is set up correctly
"""

import pickle
from pathlib import Path
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
]

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = ".token.pickle"


def test_gmail_auth():
    """Test Gmail API authentication."""
    creds = None

    # Check if we have a saved token
    if Path(TOKEN_FILE).exists():
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)

    # If no valid token, get a new one
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Refreshing token...")
            creds.refresh(Request())
        else:
            print("📱 Opening browser for Gmail login...")
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Save token for next time
        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)

    print("✅ Gmail authentication successful!")
    print(f"Token saved to: {Path(TOKEN_FILE).absolute()}")
    return True


if __name__ == "__main__":
    try:
        test_gmail_auth()
    except Exception as e:
        print(f"❌ Error: {e}")
