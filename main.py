"""
main.py — Entry point PDF vs Excel Comparison Tool
AMFS Life Insurance | QA Automation Team

Cara pakai:
    python main.py
    python main.py --testcase data/testcase/testcase.xlsx

Flow per eksekusi:
    1. Baca test case Excel
    2. Buat folder results/YYYYMMDD_HHMMSS/
    3. Duplikat test case → testcase_result.xlsx
    4. Per policy: extract DB → cari PDF → extract PDF → compare → report → update status
"""

import argparse
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

import pandas as pd

from src.comparator import PolicyResult, compare_policy
from src.config_loader import load_config
from src.db_extractor import extract_policy_data, extract_transaction_data
from src.excel_handler import (
    duplicate_testcase,
    fill_db_data,
    fill_transaction_data,
    read_testcase,
    update_row_status,
)
from src.pdf_extractor import extract_proposal
from src.pdf_locator import find_pdf, validate_pdf
from src.report_generator import build_report

# ── Constants ─────────────────────────────────────────────────────────────────
DEFAULT_TESTCASE   = os.getenv("TESTCASE_FILE",  "data/testcase/testcase.xlsx")
DEFAULT_PDF_FOLDER = os.getenv("PDF_FOLDER",     "data/pdf")
DEFAULT_RESULTS    = os.getenv("RESULTS_FOLDER", "results")


# ── Logging Setup ─────────────────────────────────────────────────────────────

def _setup_logging(log_path: str) -> None:
    """Setup logging ke console dan file secara bersamaan."""
    log_format = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_path, encoding="utf-8"),
        ],
    )


# ── Folder Setup ──────────────────────────────────────────────────────────────

def prepare_result_folder(base_dir: str) -> tuple[str, str]:
    """
    Buat folder hasil eksekusi dengan prefix datetime.

    Args:
        base_dir: Folder results/ utama

    Returns:
        tuple (result_folder, reports_folder)
    """
    timestamp     = datetime.now().strftime("%Y%m%d_%H%M%S")
    result_folder = Path(base_dir) / timestamp
    reports_folder = result_folder / "reports"

    result_folder.mkdir(parents=True, exist_ok=True)
    reports_folder.mkdir(parents=True, exist_ok=True)

    return str(result_folder), str(reports_folder)


# ── Per-Policy Runner ─────────────────────────────────────────────────────────

def run_single_policy(
    row: pd.Series,
    result_path: str,
    reports_folder: str,
    pdf_folder: str,
    config_cache: dict,
    logger: logging.Logger,
) -> str:
    """
    Orkestrasi compare untuk satu policy number.

    Args:
        row: Satu baris dari DataFrame test case
        result_path: Path testcase_result.xlsx
        reports_folder: Path folder reports/
        pdf_folder: Path folder PDF
        config_cache: Cache config per product_code
        logger: Logger instance

    Returns:
        Status akhir: "PASS", "FAIL", atau "ERROR"
    """
    test_case_id  = str(row["test_case_id"]).strip()
    policy_number = str(row["policy_number"]).strip()
    product_code  = str(row["product_code"]).strip()

    # ── Load config (gunakan cache) ──
    if product_code not in config_cache:
        config_cache[product_code] = load_config(product_code)
    config = config_cache[product_code]

    # ── Step 2: Extract DB ──
    try:
        db_data = extract_policy_data(policy_number)
        transactions = extract_transaction_data(policy_number)
        db_data["transactions"] = transactions

        # Tulis data header polis ke Sheet 1 (Test Cases)
        fill_db_data(result_path, policy_number, db_data)

        # Tulis data transaksi ke Sheet "DB Transactions" (terpisah)
        if transactions:
            fill_transaction_data(result_path, policy_number, transactions)
        else:
            logger.warning(f"[{policy_number}] Tidak ada data transaksi di DB.")

    except Exception as e:
        logger.error(f"[{policy_number}] Gagal extract DB: {e}")
        update_row_status(result_path, test_case_id, "ERROR", notes=f"DB Error: {e}")
        return "ERROR"

    # ── Step 3: Locate PDF ──
    pdf_path = find_pdf(policy_number, pdf_folder)
    if not pdf_path:
        logger.warning(f"[{policy_number}] File PDF tidak ditemukan.")
        update_row_status(
            result_path, test_case_id,
            status="FAIL",
            notes="File PDF tidak ditemukan",
        )
        return "FAIL"

    if not validate_pdf(pdf_path):
        logger.error(f"[{policy_number}] File PDF corrupt: {pdf_path}")
        update_row_status(
            result_path, test_case_id,
            status="ERROR",
            notes="File PDF corrupt atau tidak valid",
        )
        return "ERROR"

    # ── Step 4: Extract PDF ──
    try:
        pdf_data = extract_proposal(pdf_path, config)
    except Exception as e:
        logger.error(f"[{policy_number}] Gagal extract PDF: {e}")
        update_row_status(result_path, test_case_id, "ERROR", notes=f"PDF Extract Error: {e}")
        return "ERROR"

    # Check mapping: pastikan semua field wajib berhasil di-extract
    missing_fields = [
        f for f in config.get("excel_column_map", {})
        if f not in ("summary_table", "detail_table") and pdf_data.get(f) is None
    ]
    if missing_fields:
        logger.warning(f"[{policy_number}] Field tidak ditemukan di PDF: {missing_fields}")
        update_row_status(
            result_path, test_case_id,
            status="ERROR",
            notes=f"Field tidak ditemukan di PDF: {', '.join(missing_fields)}",
        )
        return "ERROR"

    # ── Step 5: Compare ──
    policy_result: PolicyResult = compare_policy(
        pdf_data      = pdf_data,
        db_data       = db_data,
        config        = config,
        policy_number = policy_number,
        test_case_id  = test_case_id,
    )

    # ── Step 6: Generate Report ──
    build_report(policy_result, reports_folder)

    # ── Step 7: Update Status ──
    update_row_status(
        result_path    = result_path,
        test_case_id   = test_case_id,
        status         = policy_result.status,
        mismatch_fields= policy_result.mismatch_fields or None,
        notes          = policy_result.notes,
    )

    return policy_result.status


# ── Main Runner ───────────────────────────────────────────────────────────────

def run(testcase_path: str) -> None:
    """
    Jalankan seluruh pipeline comparison untuk semua policy di test case.

    Args:
        testcase_path: Path ke file testcase.xlsx
    """
    # Setup folder hasil
    result_folder, reports_folder = prepare_result_folder(DEFAULT_RESULTS)
    log_path = str(Path(result_folder) / "run.log")
    _setup_logging(log_path)
    logger = logging.getLogger("main")

    logger.info("=" * 60)
    logger.info("PDF vs Excel Comparison Tool — AMFS Life Insurance")
    logger.info(f"Execution folder: {result_folder}")
    logger.info("=" * 60)

    # Baca test case
    df = read_testcase(testcase_path)
    total = len(df)

    # Duplikat test case ke folder result
    result_path = duplicate_testcase(testcase_path, result_folder)

    # Proses per policy
    config_cache: dict = {}
    counts = {"PASS": 0, "FAIL": 0, "ERROR": 0}

    for idx, row in df.iterrows():
        policy_number = str(row.get("policy_number", "")).strip()
        logger.info(f"\nProcessing [{idx + 1}/{total}]: policy {policy_number}")

        try:
            status = run_single_policy(
                row            = row,
                result_path    = result_path,
                reports_folder = reports_folder,
                pdf_folder     = DEFAULT_PDF_FOLDER,
                config_cache   = config_cache,
                logger         = logger,
            )
        except Exception as e:
            # Error tak terduga — log dan lanjut ke policy berikutnya
            logger.error(f"[{policy_number}] Unexpected error: {e}", exc_info=True)
            test_case_id = str(row.get("test_case_id", "")).strip()
            update_row_status(result_path, test_case_id, "ERROR", notes=f"Unexpected: {e}")
            status = "ERROR"

        counts[status] = counts.get(status, 0) + 1

    # Summary akhir
    logger.info("\n" + "=" * 60)
    logger.info("EXECUTION COMPLETE")
    logger.info(f"  Total  : {total}")
    logger.info(f"  PASS   : {counts['PASS']}")
    logger.info(f"  FAIL   : {counts['FAIL']}")
    logger.info(f"  ERROR  : {counts['ERROR']}")
    logger.info(f"  Result : {result_folder}")
    logger.info("=" * 60)


# ── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="PDF vs Excel Comparison Tool — AMFS Life Insurance",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--testcase",
        default=DEFAULT_TESTCASE,
        help=f"Path ke file Excel test case\n(default: {DEFAULT_TESTCASE})",
    )
    args = parser.parse_args()
    run(args.testcase)


if __name__ == "__main__":
    main()
