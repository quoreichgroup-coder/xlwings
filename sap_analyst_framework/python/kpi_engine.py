"""
kpi_engine.py — Configuration-driven KPI engine for the SAP Analyst Framework.

Reads the REPORT_CONFIG DataFrame (from ExcelBridge) and evaluates each KPI
against the cleaned data DataFrame.  Mirrors the VBA modFormulaEngine logic.

Supported formula types:
  SUM, COUNT, COUNT_UNIQUE, AVERAGE, MIN, MAX, FORMULA, CUSTOM
"""

from __future__ import annotations

import logging
from typing import Any

import pandas as pd

log = logging.getLogger(__name__)


# ── Public interface ──────────────────────────────────────────────────────────

def calculate_kpis(
    data_df: pd.DataFrame,
    report_config: pd.DataFrame,
    logger: logging.Logger | None = None,
) -> dict[str, Any]:
    """
    Calculate all KPIs defined in report_config against data_df.

    report_config columns (case-insensitive match):
        kpi_name, formula_type, column_name, filter_column, filter_value,
        target_sheet, target_cell, label

    Returns a dict keyed by target_sheet + "!" + target_cell → result value.
    """
    _log = logger or log
    results: dict[str, Any] = {}

    if report_config.empty:
        _log.warning("REPORT_CONFIG is empty — no KPIs to calculate")
        return results

    # Normalise column names
    rc = report_config.copy()
    rc.columns = [str(c).strip().lower().replace(" ", "_") for c in rc.columns]

    required = {"kpi_name", "formula_type", "column_name"}
    missing = required - set(rc.columns)
    if missing:
        _log.error("REPORT_CONFIG missing columns: %s", missing)
        return results

    for _, row in rc.iterrows():
        kpi_name     = str(row.get("kpi_name", "")).strip()
        formula_type = str(row.get("formula_type", "")).strip().upper()
        col_name     = str(row.get("column_name", "")).strip()
        filter_col   = str(row.get("filter_column", "")).strip()
        filter_val   = str(row.get("filter_value", "")).strip()
        target_sheet = str(row.get("target_sheet", "")).strip()
        target_cell  = str(row.get("target_cell", "")).strip()

        if not kpi_name or not formula_type:
            continue

        try:
            result = _evaluate(data_df, formula_type, col_name, filter_col, filter_val)
            _log.info("  KPI %-30s = %s", kpi_name, result)
        except Exception as exc:
            result = f"ERROR: {exc}"
            _log.error("  KPI %s failed: %s", kpi_name, exc)

        if target_sheet and target_cell:
            key = f"{target_sheet}!{target_cell}"
            results[key] = result

    return results


def write_results(
    workbook_path,
    results: dict[str, Any],
    logger: logging.Logger | None = None,
) -> None:
    """Write KPI results back to the workbook via ExcelBridge."""
    from excel_bridge import ExcelBridge
    from pathlib import Path

    _log = logger or log
    bridge = ExcelBridge(Path(workbook_path))

    for key, value in results.items():
        try:
            sheet, cell = key.split("!", 1)
            bridge.write_cell(sheet, cell, value)
        except Exception as exc:
            _log.error("write_results(%s) failed: %s", key, exc)


# ── Formula evaluators ────────────────────────────────────────────────────────

def _evaluate(
    df: pd.DataFrame,
    formula_type: str,
    col_name: str,
    filter_col: str = "",
    filter_val: str = "",
) -> Any:
    series = _get_series(df, col_name, filter_col, filter_val)

    match formula_type:
        case "SUM":
            return _to_numeric(series).sum()
        case "COUNT":
            return len(series)
        case "COUNT_UNIQUE":
            return series.nunique()
        case "AVERAGE":
            s = _to_numeric(series)
            return s.mean() if len(s) > 0 else 0
        case "MIN":
            return _to_numeric(series).min()
        case "MAX":
            return _to_numeric(series).max()
        case "FORMULA":
            # col_name is treated as a pandas eval expression on the full df
            return df.eval(col_name).sum()
        case "CUSTOM":
            return f"[CUSTOM: {col_name}]"
        case _:
            raise ValueError(f"Unknown formula type: {formula_type}")


def _get_series(
    df: pd.DataFrame,
    col_name: str,
    filter_col: str,
    filter_val: str,
) -> pd.Series:
    matched_col = _find_column(df, col_name)
    if matched_col is None:
        raise KeyError(f"Column not found: '{col_name}'")

    if filter_col and filter_val:
        matched_filter = _find_column(df, filter_col)
        if matched_filter is None:
            raise KeyError(f"Filter column not found: '{filter_col}'")
        mask = df[matched_filter].astype(str).str.strip().str.lower() == filter_val.lower()
        return df.loc[mask, matched_col]

    return df[matched_col]


def _find_column(df: pd.DataFrame, name: str) -> str | None:
    """Case-insensitive column name lookup."""
    needle = name.strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == needle:
            return col
    return None


def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)
