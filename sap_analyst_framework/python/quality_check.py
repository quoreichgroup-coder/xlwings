"""
quality_check.py — Data quality validation for SAP exports.

Returns a list of human-readable issue strings so VBA or main.py
can log them or surface them in the LOG sheet.
"""

from __future__ import annotations

import logging
from typing import Any

import pandas as pd

log = logging.getLogger(__name__)


def validate_data(df: pd.DataFrame, logger: logging.Logger | None = None) -> list[str]:
    """
    Run all quality checks against a cleaned DataFrame.
    Returns a (possibly empty) list of issue description strings.
    """
    _log = logger or log
    issues: list[str] = []

    if df.empty:
        issues.append("DataFrame is empty — no data to validate.")
        return issues

    issues += _check_empty_rows(df)
    issues += _check_duplicate_rows(df)
    issues += _check_null_counts(df)
    issues += _check_negative_quantities(df)
    issues += _check_date_range(df)

    _log.info("Validation complete — %d issue(s) found", len(issues))
    return issues


# ── Individual checks ─────────────────────────────────────────────────────────

def _check_empty_rows(df: pd.DataFrame) -> list[str]:
    n = df.isna().all(axis=1).sum()
    if n > 0:
        return [f"{n} fully empty row(s) found."]
    return []


def _check_duplicate_rows(df: pd.DataFrame) -> list[str]:
    n = df.duplicated().sum()
    if n > 0:
        return [f"{n} duplicate row(s) found."]
    return []


def _check_null_counts(df: pd.DataFrame, threshold: float = 0.3) -> list[str]:
    """Flag columns where more than `threshold` fraction of values are null."""
    issues = []
    for col in df.columns:
        null_rate = df[col].isna().mean()
        if null_rate > threshold:
            issues.append(
                f"Column '{col}' has {null_rate:.0%} null values (threshold {threshold:.0%})."
            )
    return issues


def _check_negative_quantities(df: pd.DataFrame) -> list[str]:
    """Warn if numeric columns that look like quantities contain negatives."""
    issues = []
    qty_keywords = ("quantity", "qty", "menge", "amount", "betrag", "count")
    for col in df.select_dtypes(include="number").columns:
        if any(kw in col.lower() for kw in qty_keywords):
            n_neg = (df[col] < 0).sum()
            if n_neg > 0:
                issues.append(f"Column '{col}' contains {n_neg} negative value(s).")
    return issues


def _check_date_range(df: pd.DataFrame,
                      min_year: int = 2000,
                      max_year: int = 2100) -> list[str]:
    """Flag date columns with out-of-range values."""
    issues = []
    for col in df.select_dtypes(include=["datetime64[ns]", "datetime"]).columns:
        out_of_range = df[col].dropna().apply(
            lambda d: d.year < min_year or d.year > max_year
        ).sum()
        if out_of_range > 0:
            issues.append(
                f"Column '{col}' has {out_of_range} date(s) outside "
                f"{min_year}–{max_year}."
            )
    return issues


# ── Utility ───────────────────────────────────────────────────────────────────

def quality_summary(df: pd.DataFrame) -> dict[str, Any]:
    """Return a dict summary suitable for writing to the workbook."""
    return {
        "row_count":       len(df),
        "column_count":    len(df.columns),
        "null_cells":      int(df.isna().sum().sum()),
        "duplicate_rows":  int(df.duplicated().sum()),
        "issues":          validate_data(df),
    }
