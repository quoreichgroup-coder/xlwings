"""
main.py — SAP Analyst Framework Python pipeline entry point.

Called by VBA via:
    python main.py <workbook_path> <export_file_path>

The script orchestrates cleaning, validation, KPI calculation, and writing
results back to the workbook — all without requiring admin installation.
Requires only: pandas, openpyxl  (pip install pandas openpyxl)
"""

from __future__ import annotations

import sys
import logging
from pathlib import Path
from datetime import datetime

# ── Bootstrap: ensure sibling modules are importable regardless of cwd ────────
_HERE = Path(__file__).parent
sys.path.insert(0, str(_HERE))

from excel_bridge import ExcelBridge
from cleaner import clean_export
from quality_check import validate_data
from kpi_engine import calculate_kpis, write_results

# ── Logging setup ─────────────────────────────────────────────────────────────
def _setup_logging(workbook_path: Path) -> logging.Logger:
    log_dir = workbook_path.parent / "Logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"python_{ts}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("main")


def run(workbook_path: str, export_file_path: str) -> None:
    wb_path  = Path(workbook_path)
    exp_path = Path(export_file_path)

    log = _setup_logging(wb_path)
    log.info("Python pipeline started")
    log.info("  Workbook : %s", wb_path)
    log.info("  Export   : %s", exp_path)

    if not exp_path.exists():
        log.error("Export file not found: %s", exp_path)
        sys.exit(1)

    bridge = ExcelBridge(wb_path)

    # 1. Clean raw SAP export
    log.info("Step 1 — Cleaning export data")
    df = clean_export(exp_path, log)

    # 2. Validate data quality
    log.info("Step 2 — Validating data")
    issues = validate_data(df, log)
    if issues:
        log.warning("Validation issues found: %d", len(issues))
        for issue in issues:
            log.warning("  • %s", issue)

    # 3. Read REPORT_CONFIG from workbook
    log.info("Step 3 — Reading REPORT_CONFIG")
    report_config = bridge.read_sheet("REPORT_CONFIG")

    # 4. Calculate KPIs
    log.info("Step 4 — Calculating KPIs")
    results = calculate_kpis(df, report_config, log)

    # 5. Write results back to workbook
    log.info("Step 5 — Writing results to workbook")
    write_results(wb_path, results, log)

    log.info("Python pipeline completed successfully")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python main.py <workbook_path> <export_file_path>")
        sys.exit(1)

    run(sys.argv[1], sys.argv[2])
