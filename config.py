"""
Configuration loader — reads all settings from .env file.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env from the same directory as this script
_env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(_env_path)


def _get(key: str, default: str = "") -> str:
    """Get an environment variable or return default."""
    return os.getenv(key, default).strip()


def _get_int(key: str, default: int = 0) -> int:
    """Get an environment variable as integer."""
    try:
        return int(_get(key, str(default)))
    except ValueError:
        return default


# ── API ──────────────────────────────────────────────
API_LOGIN_URL = _get("API_LOGIN_URL")
API_UPLOAD_URL = _get("API_UPLOAD_URL")
API_USERNAME = _get("API_USERNAME")
API_PASSWORD = _get("API_PASSWORD")

# ── File Paths ───────────────────────────────────────
CSV_FILE_PATH = _get("CSV_FILE_PATH")
TEMPLATE_FILE_PATH = _get("TEMPLATE_FILE_PATH")
OUTPUT_FOLDER_PATH = _get("OUTPUT_FOLDER_PATH")

# ── Excel Settings ───────────────────────────────────
ACCORD_CODE_CELL = _get("ACCORD_CODE_CELL", "A1")
REFRESH_WAIT_SECONDS = _get_int("REFRESH_WAIT_SECONDS", 30)
EXCEL_MAX_RETRIES = _get_int("EXCEL_MAX_RETRIES", 3)

# ── Upload Settings ──────────────────────────────────
UPLOAD_MAX_RETRIES = _get_int("UPLOAD_MAX_RETRIES", 3)
UPLOAD_RETRY_DELAY = _get_int("UPLOAD_RETRY_DELAY", 5)

# ── Logging ──────────────────────────────────────────
LOG_FILE_PATH = _get("LOG_FILE_PATH", "automation.log")
ERROR_LOG_FILE_PATH = _get("ERROR_LOG_FILE_PATH", "errors.log")
