"""
Excel COM Automation — controls Microsoft Excel via pywin32 to:
  1. Open the template workbook
  2. Write Accord Code to a cell
  3. Trigger RefreshAll (for AccEquity XL NXT add-in)
  4. Wait for data population
  5. Save as a new file
"""

import os
import time
import shutil
import logging
import pythoncom
import win32com.client as win32

from config import (
    TEMPLATE_FILE_PATH,
    OUTPUT_FOLDER_PATH,
    ACCORD_CODE_CELL,
    REFRESH_WAIT_SECONDS,
    EXCEL_MAX_RETRIES,
)


class ExcelAutomation:
    """Manages the lifecycle of an Excel COM session."""

    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.excel_app = None

    # ── Lifecycle ────────────────────────────────────

    def start_excel(self) -> None:
        """Launch Excel application via COM."""
        try:
            pythoncom.CoInitialize()
            self.excel_app = win32.DispatchEx("Excel.Application")
            self.excel_app.Visible = False          # Run headless
            self.excel_app.DisplayAlerts = False    # Suppress dialogs
            self.excel_app.AskToUpdateLinks = False # Don't prompt for links
            self.logger.info("Excel application started successfully.")
        except Exception as e:
            self.logger.error(f"Failed to start Excel: {e}")
            raise

    def quit_excel(self) -> None:
        """Gracefully close Excel."""
        try:
            if self.excel_app:
                self.excel_app.Quit()
                self.logger.info("Excel application closed.")
        except Exception as e:
            self.logger.warning(f"Error closing Excel: {e}")
        finally:
            self.excel_app = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # ── Core Processing ──────────────────────────────

    def process_company(
        self, accord_code: str, company_name: str
    ) -> str:
        """
        Process a single company:
          1. Copy template → working file
          2. Open working file in Excel
          3. Write accord_code to ACCORD_CODE_CELL
          4. RefreshAll + wait
          5. Save as <Company Name>.xlsx
          
        Returns the path to the saved output file.
        Raises on unrecoverable failure.
        """
        os.makedirs(OUTPUT_FOLDER_PATH, exist_ok=True)

        # Sanitize company name for filename
        safe_name = self._sanitize_filename(company_name)
        output_path = os.path.join(OUTPUT_FOLDER_PATH, f"{safe_name}.xlsx")
        
        # Create a temporary working copy of the template
        working_copy = os.path.join(OUTPUT_FOLDER_PATH, f"_working_{safe_name}.xlsx")
        shutil.copy2(TEMPLATE_FILE_PATH, working_copy)
        self.logger.debug(f"Template copied to working file: {working_copy}")

        workbook = None
        try:
            # Open workbook
            workbook = self.excel_app.Workbooks.Open(
                os.path.abspath(working_copy),
                UpdateLinks=0,  # Don't update external links on open
            )
            self.logger.info(
                f"[{company_name}] Workbook opened. Writing Accord Code '{accord_code}' to {ACCORD_CODE_CELL}."
            )

            # Write Accord Code to the designated cell
            sheet = workbook.Sheets(1)
            sheet.Range(ACCORD_CODE_CELL).Value = int(accord_code)
            self.logger.debug(f"[{company_name}] Accord Code written.")

            # Refresh with retries
            self._refresh_with_retry(workbook, company_name)

            # Save as output file
            # FileFormat=51 → .xlsx (xlOpenXMLWorkbook)
            workbook.SaveAs(
                os.path.abspath(output_path),
                FileFormat=51,
            )
            self.logger.info(f"[{company_name}] Saved → {output_path}")

            return output_path

        except Exception as e:
            self.logger.error(f"[{company_name}] Excel processing failed: {e}")
            raise

        finally:
            # Close workbook without saving again
            if workbook:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            # Clean up working copy
            if os.path.exists(working_copy):
                try:
                    os.remove(working_copy)
                except Exception:
                    pass

    # ── Refresh Logic ────────────────────────────────
    #
    # AccEquity XL NXT is an Excel add-in, NOT a standard data connection.
    # workbook.RefreshAll() alone only refreshes ODBC/Power Query/Web queries.
    # Add-ins typically use one (or more) of these mechanisms:
    #
    #   1. RTD (Real-Time Data) server  → needs RTD throttle reset
    #   2. Custom UDF worksheet functions → needs CalculateFullRebuild
    #   3. Ribbon "Refresh" button hook  → needs SendKeys Ctrl+Alt+F5
    #   4. COM Add-in with its own macro → needs Application.Run (if known)
    #
    # We fire ALL strategies so that whichever mechanism the add-in uses,
    # it gets triggered.

    def _refresh_with_retry(self, workbook, company_name: str) -> None:
        """
        Multi-strategy refresh to ensure the AccEquity XL NXT add-in
        recalculates all data. Retries on failure.
        """
        for attempt in range(1, EXCEL_MAX_RETRIES + 1):
            try:
                self.logger.info(
                    f"[{company_name}] Refreshing all data (attempt {attempt}/{EXCEL_MAX_RETRIES})..."
                )

                # ── Strategy 1: Reset RTD Throttle ───────────────
                # If the add-in uses an RTD server, resetting the throttle
                # interval forces Excel to re-query the server immediately
                # instead of waiting for the next polling cycle.
                try:
                    original_throttle = self.excel_app.RTD.ThrottleInterval
                    self.excel_app.RTD.ThrottleInterval = 0
                    self.logger.debug(
                        f"[{company_name}] RTD throttle reset (was {original_throttle}ms → 0ms)."
                    )
                except Exception:
                    self.logger.debug(f"[{company_name}] RTD throttle reset not applicable.")

                # ── Strategy 2: CalculateFullRebuild ─────────────
                # Forces Excel to rebuild the dependency tree and recalculate
                # ALL formulas, including custom UDF functions from add-ins.
                # This is stronger than Calculate() or CalculateFull().
                try:
                    self.excel_app.CalculateFullRebuild()
                    self.logger.debug(f"[{company_name}] CalculateFullRebuild triggered.")
                except Exception as e:
                    self.logger.debug(f"[{company_name}] CalculateFullRebuild skipped: {e}")

                # ── Strategy 3: RefreshAll ───────────────────────
                # Standard refresh for data connections, pivot tables,
                # and any query tables. Some add-ins also hook into this.
                workbook.RefreshAll()
                self.logger.debug(f"[{company_name}] RefreshAll triggered.")

                # ── Strategy 4: SendKeys Ctrl+Alt+F5 ─────────────
                # Simulates pressing the keyboard shortcut for
                # "Refresh All" in the Data ribbon tab. Many add-ins
                # intercept this event even if they ignore RefreshAll().
                # SendKeys requires the Excel window to exist (it can
                # be minimized but must be Visible=True briefly).
                try:
                    self.excel_app.Visible = True
                    time.sleep(1)  # Let window render
                    self.excel_app.SendKeys("^%{F5}", Wait=True)  # Ctrl+Alt+F5
                    self.logger.debug(f"[{company_name}] SendKeys Ctrl+Alt+F5 sent.")
                    self.excel_app.Visible = False
                except Exception as e:
                    self.logger.debug(f"[{company_name}] SendKeys skipped: {e}")
                    try:
                        self.excel_app.Visible = False
                    except Exception:
                        pass

                # ── Wait for add-in data ─────────────────────────
                self.logger.info(
                    f"[{company_name}] Waiting {REFRESH_WAIT_SECONDS}s for AccEquity XL NXT data..."
                )
                time.sleep(REFRESH_WAIT_SECONDS)

                # ── Smart async wait ─────────────────────────────
                # CalculateUntilAsyncQueriesDone blocks until all async
                # operations complete (available in Excel 2010+).
                try:
                    self.excel_app.CalculateUntilAsyncQueriesDone()
                    self.logger.debug(f"[{company_name}] Async queries done.")
                except Exception:
                    pass

                # ── Restore RTD throttle ─────────────────────────
                try:
                    self.excel_app.RTD.ThrottleInterval = original_throttle
                except Exception:
                    pass

                self.logger.info(f"[{company_name}] Refresh completed (all strategies fired).")
                return  # Success

            except Exception as e:
                self.logger.warning(
                    f"[{company_name}] Refresh attempt {attempt} failed: {e}"
                )
                if attempt == EXCEL_MAX_RETRIES:
                    raise RuntimeError(
                        f"Refresh failed after {EXCEL_MAX_RETRIES} attempts"
                    ) from e
                time.sleep(5)

    # ── Helpers ──────────────────────────────────────

    @staticmethod
    def _sanitize_filename(name: str) -> str:
        """Remove characters that are invalid in Windows filenames."""
        invalid = '<>:"/\\|?*'
        sanitized = "".join(c for c in name if c not in invalid)
        return sanitized.strip()
