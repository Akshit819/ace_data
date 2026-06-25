"""
Main Orchestrator — ties everything together.

Usage:
    python main.py
"""

import os
import sys
import pandas as pd
from datetime import datetime

from config import CSV_FILE_PATH, TEMPLATE_FILE_PATH, OUTPUT_FOLDER_PATH
from logger_setup import setup_logger
from excel_automation import ExcelAutomation
# from api_client import APIClient  # TODO: Uncomment when ready for uploads


def load_companies(csv_path: str, logger) -> pd.DataFrame:
    """Load and validate the ace.csv file."""
    if not os.path.isfile(csv_path):
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    df = pd.read_csv(csv_path)

    # Normalize column names: strip whitespace and convert to title case
    # so "Company name" and "Company Name" both become "Company Name"
    df.columns = df.columns.str.strip().str.title()

    required_cols = {"Accord Code", "Company Name"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"CSV is missing required columns: {missing}")

    # Drop rows with missing Accord Code (Company Name can have blanks for Rights etc.)
    initial_count = len(df)
    df = df.dropna(subset=["Accord Code", "Company Name"])
    dropped = initial_count - len(df)
    if dropped:
        logger.warning(f"Dropped {dropped} rows with missing Accord Code or Company Name.")

    # Filter out Rights Entitlements and Partly Paid-up entries
    mask = df["Company Name"].str.contains(
        r"Rights Entitlements|Partly Paid-up|Amalgamated|Merged",
        case=False, na=False
    )
    filtered = mask.sum()
    if filtered:
        df = df[~mask]
        logger.info(f"Filtered out {filtered} Rights/Partly-Paid/Amalgamated entries.")

    # Convert Accord Code to string (for consistent handling)
    df["Accord Code"] = df["Accord Code"].astype(int).astype(str)
    df["Company Name"] = df["Company Name"].astype(str).str.strip()

    logger.info(f"Loaded {len(df)} companies from {csv_path}.")
    return df


def validate_prerequisites(logger) -> None:
    """Check that required files exist before starting."""
    errors = []

    if not os.path.isfile(CSV_FILE_PATH):
        errors.append(f"CSV file not found: {CSV_FILE_PATH}")

    if not os.path.isfile(TEMPLATE_FILE_PATH):
        errors.append(f"Template file not found: {TEMPLATE_FILE_PATH}")

    if errors:
        for err in errors:
            logger.error(err)
        raise FileNotFoundError("\n".join(errors))

    os.makedirs(OUTPUT_FOLDER_PATH, exist_ok=True)
    logger.info("All prerequisites validated.")


def main():
    # ── Setup ────────────────────────────────────────
    main_logger, error_logger = setup_logger()
    start_time = datetime.now()

    main_logger.info("=" * 60)
    main_logger.info("AccEquity Excel Automation — START")
    main_logger.info(f"Start Time: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    main_logger.info("=" * 60)

    # Validate
    try:
        validate_prerequisites(main_logger)
    except FileNotFoundError:
        main_logger.critical("Prerequisites check failed. Exiting.")
        sys.exit(1)

    # Load companies
    try:
        companies = load_companies(CSV_FILE_PATH, main_logger)
    except Exception as e:
        main_logger.critical(f"Failed to load CSV: {e}")
        sys.exit(1)

    # ── Initialize Services ──────────────────────────
    excel = ExcelAutomation(main_logger)
    # TODO: Uncomment when ready for uploads
    # api = APIClient(main_logger, error_logger)
    #
    # # Authenticate once upfront
    # try:
    #     api.login()
    # except Exception as e:
    #     main_logger.critical(f"Initial authentication failed: {e}")
    #     sys.exit(1)

    # Start Excel
    try:
        excel.start_excel()
    except Exception as e:
        main_logger.critical(f"Cannot start Excel: {e}")
        sys.exit(1)

    # ── Process Each Company ─────────────────────────
    results = {
        "success": 0,
        "excel_fail": 0,
        "upload_fail": 0,
    }

    total = len(companies)

    for idx, row in companies.iterrows():
        accord_code = row["Accord Code"]
        company_name = row["Company Name"]
        company_num = idx + 1

        main_logger.info("-" * 50)
        main_logger.info(f"Processing [{company_num}/{total}]: {company_name} (Code: {accord_code})")

        # Step 1: Excel Processing
        output_file = None
        try:
            output_file = excel.process_company(accord_code, company_name)
        except Exception as e:
            main_logger.error(f"[{company_name}] EXCEL FAILED: {e}")
            error_logger.error(
                f"EXCEL_FAIL | {company_name} | {accord_code} | {e}"
            )
            results["excel_fail"] += 1
            continue  # Skip to next company

        # Step 2: Upload (disabled — testing Excel only)
        # TODO: Uncomment when ready for uploads
        # try:
        #     response = api.upload_file(output_file, company_name, accord_code)
        #     main_logger.info(
        #         f"[{company_name}] UPLOAD SUCCESS | Response: {response}"
        #     )
        #     results["success"] += 1
        # except Exception as e:
        #     main_logger.error(f"[{company_name}] UPLOAD FAILED: {e}")
        #     error_logger.error(
        #         f"UPLOAD_FAIL | {company_name} | {accord_code} | File: {output_file} | {e}"
        #     )
        #     results["upload_fail"] += 1
        #     continue
        results["success"] += 1  # Count as success if Excel processing worked

    # ── Cleanup ──────────────────────────────────────
    excel.quit_excel()
    # api.close()  # TODO: Uncomment when ready for uploads

    # ── Summary ──────────────────────────────────────
    end_time = datetime.now()
    duration = end_time - start_time

    main_logger.info("=" * 60)
    main_logger.info("AccEquity Excel Automation — COMPLETE")
    main_logger.info(f"End Time  : {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    main_logger.info(f"Duration  : {duration}")
    main_logger.info(f"Total     : {total}")
    main_logger.info(f"Success   : {results['success']}")
    main_logger.info(f"Excel Fail: {results['excel_fail']}")
    main_logger.info(f"Upload Fail: {results['upload_fail']}")
    main_logger.info("=" * 60)

    # Exit with error code if any failures
    if results["excel_fail"] + results["upload_fail"] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
