"""
Logging setup for RIOS
"""

import logging
from config.settings import LOG_FILE, LOG_LEVEL

# Create logger
logger = logging.getLogger("RIOS")
logger.setLevel(LOG_LEVEL)

# File handler
file_handler = logging.FileHandler(LOG_FILE)
file_handler.setLevel(LOG_LEVEL)

# Console handler
console_handler = logging.StreamHandler()
console_handler.setLevel(LOG_LEVEL)

# Formatter
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add handlers
logger.addHandler(file_handler)
logger.addHandler(console_handler)
