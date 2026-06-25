"""
API Client — handles authentication and file upload to tickercharcha.com.

Upload endpoint: POST /tickertest/upload-files
Form fields:
    - company_name  (str, required)
    - uploader_name (str, default "-")
    - file          (UploadFile, required)
    - upload_type   (int, default 0)
        0 = Excel upload (what we use)
        1 = Annual consensus
        2 = BB Q4
        3 = Report upload

Auth: Bearer token via get_current_user dependency.
"""

import os
import time
import logging
import requests

from config import (
    API_LOGIN_URL,
    API_UPLOAD_URL,
    API_USER_ID,
    API_PASSWORD,
    UPLOAD_MAX_RETRIES,
    UPLOAD_RETRY_DELAY,
)


class APIClient:
    """Manages authentication and file uploads to tickercharcha.com."""

    def __init__(self, logger: logging.Logger, error_logger: logging.Logger):
        self.logger = logger
        self.error_logger = error_logger
        self.token: str | None = None
        self.session = requests.Session()

    # ── Authentication ───────────────────────────────

    def login(self) -> None:
        """
        Authenticate with the API and store the Bearer token.
        Raises on failure.
        """
        self.logger.info(f"Authenticating with API at {API_LOGIN_URL}...")
        try:
            response = self.session.post(
                API_LOGIN_URL,
                json={
                    "user_id": API_USER_ID,
                    "password": API_PASSWORD,
                },
                timeout=30,
            )
            response.raise_for_status()

            data = response.json()

            # Try common token field names from the response
            self.token = (
                data.get("token")
                or data.get("access_token")
                or data.get("auth_token")
                or data.get("data", {}).get("token")
            )

            if not self.token:
                raise ValueError(
                    f"Token not found in login response. "
                    f"Response keys: {list(data.keys())}. "
                    f"Full response: {data}"
                )

            self.session.headers.update({
                "Authorization": f"Bearer {self.token}"
            })
            self.logger.info("Authentication successful. Token acquired.")

        except requests.exceptions.RequestException as e:
            self.logger.error(f"Authentication failed: {e}")
            raise
        except (ValueError, KeyError) as e:
            self.logger.error(f"Failed to parse auth response: {e}")
            raise

    def _ensure_authenticated(self) -> None:
        """Re-authenticate if token is missing."""
        if not self.token:
            self.login()

    # ── File Upload ──────────────────────────────────

    def upload_file(
        self, file_path: str, company_name: str, accord_code: str
    ) -> dict:
        """
        Upload a generated Excel file to tickercharcha.com.

        Endpoint: POST /tickertest/upload-files
        Form data:
            - company_name:  company name from ace.csv
            - uploader_name: identifier for who uploaded (default: "ace_automation")
            - file:          the .xlsx file
            - upload_type:   0 (Excel upload)

        Returns the API response as a dict.
        Raises on unrecoverable failure.
        """
        self._ensure_authenticated()

        filename = os.path.basename(file_path)

        for attempt in range(1, UPLOAD_MAX_RETRIES + 1):
            try:
                self.logger.info(
                    f"[{company_name}] Uploading '{filename}' "
                    f"(attempt {attempt}/{UPLOAD_MAX_RETRIES})..."
                )

                with open(file_path, "rb") as f:
                    response = self.session.post(
                        API_UPLOAD_URL,
                        files={
                            "file": (
                                filename,
                                f,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                        },
                        data={
                            "company_name": company_name,
                            "uploader_name": "ace_automation",
                            "upload_type": 0,
                        },
                        timeout=120,
                    )

                # Handle token expiration (401/403) → re-auth and retry
                if response.status_code in (401, 403):
                    self.logger.warning(
                        f"[{company_name}] Token expired (HTTP {response.status_code}). "
                        f"Re-authenticating..."
                    )
                    self.token = None
                    self.login()
                    continue

                response.raise_for_status()
                result = response.json()

                self.logger.info(
                    f"[{company_name}] Upload successful. Response: {result}"
                )
                return result

            except requests.exceptions.RequestException as e:
                self.logger.warning(
                    f"[{company_name}] Upload attempt {attempt} failed: {e}"
                )
                self.error_logger.warning(
                    f"[{company_name}] Upload attempt {attempt} failed: {e}"
                )

                if attempt == UPLOAD_MAX_RETRIES:
                    raise RuntimeError(
                        f"Upload failed for '{company_name}' after "
                        f"{UPLOAD_MAX_RETRIES} attempts"
                    ) from e

                time.sleep(UPLOAD_RETRY_DELAY)

            except Exception as e:
                self.logger.error(
                    f"[{company_name}] Unexpected upload error: {e}"
                )
                raise

    # ── Cleanup ──────────────────────────────────────

    def close(self) -> None:
        """Close the HTTP session."""
        self.session.close()
        self.logger.debug("HTTP session closed.")
