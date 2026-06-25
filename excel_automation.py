"""
Excel COM Automation — controls Microsoft Excel via pywin32 to:
  1. Open the template workbook
  2. Write Accord Code to a cell
  3. Trigger the ACEEQ XL NXT add-in's "Refresh All Sheets" ribbon button
  4. Wait for data population
  5. Save as a new file

The ACEEQ XL NXT add-in has its own ribbon tab with:
  - Settings, Import Data, Template Builder, Single Function
  - Company Search, Check Compatibility
  - "Refresh All Sheets" ← THIS is what we need to click
  - "Refresh"

Since these are custom add-in buttons (not standard Excel data connections),
workbook.RefreshAll() alone will NOT trigger them. We use multiple strategies:
  1. pywinauto UI automation to find & click the ribbon button
  2. Application.Run() to invoke the add-in's macros directly
  3. SendKeys as a fallback
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
            
            # Use Dispatch instead of DispatchEx.
            # Dispatch will attach to an already running Excel instance if one exists,
            # which ensures the add-in is loaded if the user has Excel open.
            self.excel_app = win32.Dispatch("Excel.Application")
            self.excel_app.Visible = True           # Must be visible for ribbon clicks
            self.excel_app.DisplayAlerts = False     # Suppress dialogs
            self.excel_app.AskToUpdateLinks = False  # Don't prompt for links
            
            # CRITICAL: When launched via COM, Excel often does not load Add-ins.
            # We must explicitly force-connect COM add-ins.
            try:
                for addin in self.excel_app.COMAddIns:
                    # If it's not connected, force connect it
                    if not addin.Connect:
                        addin.Connect = True
                        self.logger.debug(f"Force-connected COM Add-in: {addin.Description}")
            except Exception as e:
                self.logger.debug(f"Could not iterate COMAddIns: {e}")

            # Also ensure standard add-ins (like .xla / .xlam) are installed
            try:
                for addin in self.excel_app.AddIns:
                    if not addin.Installed:
                        addin.Installed = True
            except Exception:
                pass

            self.logger.info("Excel application started successfully (Visible=True, Add-ins forced).")
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
          4. Trigger ACEEQ XL NXT refresh + wait
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

            # Write Accord Code to the designated cell on the Active Sheet
            # Using ActiveSheet is safer than Sheets(1) in case the yellow cell is on sheet 2 or 3
            sheet = workbook.ActiveSheet
            sheet.Range(ACCORD_CODE_CELL).Value = int(accord_code)
            self.logger.debug(f"[{company_name}] Accord Code written to {sheet.Name}!{ACCORD_CODE_CELL}.")

            # Trigger ACEEQ XL NXT refresh with retries
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
    # The ACEEQ XL NXT add-in has its OWN "Refresh All Sheets" button
    # on its custom ribbon tab. Standard workbook.RefreshAll() does NOT
    # trigger it. We use multiple strategies in order of reliability:
    #
    #   1. pywinauto → find & click the actual ribbon button
    #   2. Application.Run → call the add-in's VBA macro directly
    #   3. SendKeys → keyboard-navigate the ribbon tab
    #   4. RefreshAll + CalculateFullRebuild → standard fallbacks

    def _refresh_with_retry(self, workbook, company_name: str) -> None:
        """
        Multi-strategy refresh targeting the ACEEQ XL NXT add-in.
        Retries up to EXCEL_MAX_RETRIES times.
        """
        for attempt in range(1, EXCEL_MAX_RETRIES + 1):
            try:
                self.logger.info(
                    f"[{company_name}] Triggering ACEEQ XL NXT refresh "
                    f"(attempt {attempt}/{EXCEL_MAX_RETRIES})..."
                )

                refresh_triggered = False

                # ── Strategy 1: pywinauto — click the ribbon button ──
                refresh_triggered = self._try_pywinauto_refresh(company_name)

                # ── Strategy 2: Application.Run — call add-in macro ──
                if not refresh_triggered:
                    refresh_triggered = self._try_application_run(company_name)

                # ── Strategy 3: SendKeys — keyboard-navigate ribbon ──
                if not refresh_triggered:
                    refresh_triggered = self._try_sendkeys_refresh(company_name)

                # ── Strategy 4: Standard fallbacks ───────────────────
                # Always fire these as supplementary triggers
                self._standard_refresh(workbook, company_name)

                # ── Wait for data ────────────────────────────────────
                self.logger.info(
                    f"[{company_name}] Waiting {REFRESH_WAIT_SECONDS}s for "
                    f"ACEEQ XL NXT to populate data..."
                )
                time.sleep(REFRESH_WAIT_SECONDS)

                # Smart async wait
                try:
                    self.excel_app.CalculateUntilAsyncQueriesDone()
                    self.logger.debug(f"[{company_name}] Async queries done.")
                except Exception:
                    pass

                self.logger.info(f"[{company_name}] Refresh completed.")
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

    # ── Strategy 1: pywinauto (most reliable) ────────

    def _try_pywinauto_refresh(self, company_name: str) -> bool:
        """
        Use pywinauto to find and click the 'Refresh All Sheets' button
        on the ACEEQ XL NXT ribbon tab.
        """
        try:
            from pywinauto import Application, findwindows

            self.logger.debug(f"[{company_name}] Trying pywinauto ribbon click...")

            # Connect to the running Excel instance
            app = Application(backend="uia").connect(class_name="XLMAIN")
            excel_window = app.window(class_name="XLMAIN")

            # Ensure window is in focus
            if excel_window.exists():
                excel_window.set_focus()
                time.sleep(0.5)

            # Find the ACEEQ XL NXT ribbon tab and click it
            try:
                aceeq_tab = excel_window.child_window(title="ACEEQ XL NXT", control_type="TabItem")
                aceeq_tab.click_input()
                time.sleep(0.5)
                self.logger.debug(f"[{company_name}] Clicked ACEEQ XL NXT tab.")
            except Exception as e:
                self.logger.debug(f"[{company_name}] Could not find ACEEQ XL NXT tab: {e}")
                # Try alternate name patterns
                try:
                    aceeq_tab = excel_window.child_window(title_re=".*ACEEQ.*|.*AccEquity.*", control_type="TabItem")
                    aceeq_tab.click_input()
                    time.sleep(0.5)
                    self.logger.debug(f"[{company_name}] Clicked ACEEQ tab (via regex).")
                except Exception:
                    self.logger.debug(f"[{company_name}] ACEEQ tab not found via regex either.")
                    return False

            # Find and click "Refresh All Sheets" button
            try:
                refresh_btn = excel_window.child_window(
                    title="Refresh All Sheets", control_type="Button"
                )
                refresh_btn.click_input()
                self.logger.info(f"[{company_name}] ✓ Clicked 'Refresh All Sheets' via pywinauto.")
                return True
            except Exception:
                pass

            # Fallback: try just "Refresh All Sheets" as a menu item or other control
            try:
                refresh_btn = excel_window.child_window(title_re=".*Refresh All Sheets.*")
                refresh_btn.click_input()
                self.logger.info(f"[{company_name}] ✓ Clicked 'Refresh All Sheets' (fallback match).")
                return True
            except Exception:
                pass

            # Last resort: try the "Refresh" button
            try:
                refresh_btn = excel_window.child_window(
                    title="Refresh", control_type="Button"
                )
                refresh_btn.click_input()
                self.logger.info(f"[{company_name}] ✓ Clicked 'Refresh' button via pywinauto.")
                return True
            except Exception as e:
                self.logger.debug(f"[{company_name}] pywinauto could not find refresh button: {e}")
                return False

        except ImportError:
            self.logger.warning("pywinauto not installed. Skipping Strategy 1.")
            return False
        except Exception as e:
            self.logger.debug(f"[{company_name}] pywinauto strategy failed: {e}")
            return False

    # ── Strategy 2: Application.Run (if macro is registered) ──

    def _try_application_run(self, company_name: str) -> bool:
        """
        Try calling the add-in's refresh macro directly.
        The exact macro name depends on how ACEEQ XL NXT registers itself.
        We try several common patterns.
        """
        # Common macro name patterns for Excel add-ins
        macro_names = [
            "RefreshAllSheets",
            "ACEEQXLNXT.RefreshAllSheets",
            "AceeqXLNXT.RefreshAllSheets",
            "AccEquity.RefreshAllSheets",
            "RefreshAll",
            "ACEEQXLNXT.RefreshAll",
            "AceeqXLNXT.Refresh",
            "ACEEQ.Refresh",
        ]

        for macro in macro_names:
            try:
                self.excel_app.Run(macro)
                self.logger.info(
                    f"[{company_name}] ✓ Application.Run('{macro}') succeeded."
                )
                return True
            except Exception:
                continue

        self.logger.debug(f"[{company_name}] No Application.Run macro worked.")
        return False

    # ── Strategy 3: SendKeys ribbon navigation ───────

    def _try_sendkeys_refresh(self, company_name: str) -> bool:
        """
        Use keyboard shortcuts to navigate to the ACEEQ XL NXT tab
        and trigger Refresh All Sheets.

        Alt → activates ribbon keytips → navigate to add-in tab → click button.
        """
        try:
            self.logger.debug(f"[{company_name}] Trying SendKeys ribbon navigation...")

            # Press Alt to show keytips, then Escape to reset state
            self.excel_app.SendKeys("%", Wait=True)   # Alt — activate ribbon
            time.sleep(0.5)
            self.excel_app.SendKeys("{ESC}", Wait=True)
            time.sleep(0.3)

            # Try Ctrl+Alt+F5 (standard Excel "Refresh All" shortcut)
            # Some add-ins hook into this
            self.excel_app.SendKeys("^%{F5}", Wait=True)
            time.sleep(0.5)

            self.logger.info(f"[{company_name}] ✓ SendKeys Ctrl+Alt+F5 sent.")
            return True

        except Exception as e:
            self.logger.debug(f"[{company_name}] SendKeys strategy failed: {e}")
            return False

    # ── Strategy 4: Standard Excel fallbacks ─────────

    def _standard_refresh(self, workbook, company_name: str) -> None:
        """
        Fire standard Excel refresh mechanisms as supplementary triggers.
        These won't trigger the ACEEQ add-in directly but help with
        any standard data connections the template may also contain.
        """
        # RTD throttle reset
        try:
            original_throttle = self.excel_app.RTD.ThrottleInterval
            self.excel_app.RTD.ThrottleInterval = 0
            time.sleep(0.5)
            self.excel_app.RTD.ThrottleInterval = original_throttle
            self.logger.debug(f"[{company_name}] RTD throttle cycled.")
        except Exception:
            pass

        # CalculateFullRebuild — forces recalculation of all formulas
        try:
            self.excel_app.CalculateFullRebuild()
            self.logger.debug(f"[{company_name}] CalculateFullRebuild triggered.")
        except Exception:
            pass

        # Standard RefreshAll
        try:
            workbook.RefreshAll()
            self.logger.debug(f"[{company_name}] RefreshAll triggered.")
        except Exception:
            pass

    # ── Helpers ──────────────────────────────────────

    @staticmethod
    def _sanitize_filename(name: str) -> str:
        """Remove characters that are invalid in Windows filenames."""
        invalid = '<>:"/\\|?*'
        sanitized = "".join(c for c in name if c not in invalid)
        return sanitized.strip()
