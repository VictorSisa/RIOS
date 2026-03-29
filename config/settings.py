"""
RIOS Configuration
Environment-based settings for paths, email, etc.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# Paths
PROJECT_ROOT = Path(__file__).parent.parent
OUTPUTS_DIR = Path(os.getenv("OUTPUTS_DIR", PROJECT_ROOT / "outputs"))
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

# Email Configuration
PERSONAL_EMAIL = "victor.r.sisa@gmail.com"
WORK_EMAIL = "sisav@aafes.com"
REPORT_SUBJECT_FILTER = "Weekly Business Review"

# Gmail API Scopes
GMAIL_SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify',
]

# Credentials & Token Storage
CREDENTIALS_FILE = PROJECT_ROOT / "credentials.json"
TOKEN_FILE = PROJECT_ROOT / ".token.pickle"

# Logging
LOG_FILE = OUTPUTS_DIR / "rios.log"
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
