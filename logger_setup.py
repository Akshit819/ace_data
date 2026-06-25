"""
Logger setup — creates dual loggers for general operations and errors.
"""

import logging
from datetime import datetime
from config import LOG_FILE_PATH, ERROR_LOG_FILE_PATH


def setup_logger() -> tuple[logging.Logger, logging.Logger]:
    """
    Returns (main_logger, error_logger).
    
    main_logger  → logs everything (DEBUG+) to console + log file
    error_logger → logs only failures to a separate error file
    """

    # ── Formatter ────────────────────────────────────
    fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # ── Main Logger ──────────────────────────────────
    main_logger = logging.getLogger("ace_automation")
    main_logger.setLevel(logging.DEBUG)

    # Console handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    main_logger.addHandler(ch)

    # File handler
    fh = logging.FileHandler(LOG_FILE_PATH, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    main_logger.addHandler(fh)

    # ── Error Logger ─────────────────────────────────
    error_logger = logging.getLogger("ace_errors")
    error_logger.setLevel(logging.WARNING)

    efh = logging.FileHandler(ERROR_LOG_FILE_PATH, encoding="utf-8")
    efh.setLevel(logging.WARNING)
    efh.setFormatter(fmt)
    error_logger.addHandler(efh)

    return main_logger, error_logger
