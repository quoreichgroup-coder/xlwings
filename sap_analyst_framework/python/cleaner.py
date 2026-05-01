"""
cleaner.py — Cleans raw SAP ALV Excel exports before KPI calculation.

Design principles:
- All operations are non-destructive (returns new DataFrame).
- Each step is a separate function for easy unit testing.
- Logging is explicit so the analyst can trace every transformation.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)


def clean_export(file_path: Path, logger: logging.Logger | None = None) -> pd.DataFrame:
    """
    Full cleaning pipeline for a SAP ALV export file.

    Steps:
      1. Load file (xlsx or csv)
      2. Drop SAP header/footer junk rows
      3. Normalise column names
      4. Strip whitespace from all string values
      5. Convert numeric columns (SAP uses commas as decimal separators)
      6. Parse date columns
      7. Drop fully empty rows / columns

    Returns a clean DataFrame.
    """
    _log = logger or log

    # 1. Load
    _log.info("Loading export: %s", file_path)
    df = _load_file(file_path)
    _log.info("  Raw shape: %s", df.shape)

    # 2. Drop SAP junk
    df = _drop_sap_junk_rows(df)

    # 3. Normalise headers
    df = _normalise_columns(df)
    _log.info("  Columns: %s", list(df.columns))

    # 4. Strip whitespace
    df = _strip_strings(df)

    # 5. Fix numeric columns
    df = _fix_numeric_columns(df)

    # 6. Parse date columns
    df = _parse_date_columns(df)

    # 7. Drop empties
    df = df.dropna(how="all").reset_index(drop=True)
    df = df.loc[:, ~(df.isna().all())]

    _log.info("  Clean shape: %s", df.shape)
    return df


# ── Step implementations ──────────────────────────────────────────────────────

def _load_file(file_path: Path) -> pd.DataFrame:
    ext = file_path.suffix.lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        return pd.read_excel(file_path, header=0, dtype=str)
    elif ext == ".csv":
        # Try semicolon first (common SAP export locale)
        try:
            return pd.read_csv(file_path, sep=";", dtype=str, encoding="utf-8")
        except Exception:
            return pd.read_csv(file_path, sep=",", dtype=str, encoding="utf-8")
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def _drop_sap_junk_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Remove SAP header lines that appear before the real data table."""
    # Heuristic: drop rows where more than 80% of values are NaN
    threshold = 0.8
    mask = df.isna().mean(axis=1) < threshold
    return df[mask].reset_index(drop=True)


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Lowercase, replace spaces/special chars with underscores."""
    df = df.copy()
    df.columns = [
        re.sub(r"[^a-zA-Z0-9]+", "_", str(c)).strip("_").lower()
        for c in df.columns
    ]
    # De-duplicate column names
    seen: dict[str, int] = {}
    new_cols = []
    for col in df.columns:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)
    df.columns = new_cols
    return df


def _strip_strings(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip().replace("nan", pd.NA)
    return df


def _fix_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Convert SAP-formatted numbers (e.g. '1.234,56' → 1234.56)."""
    df = df.copy()
    for col in df.columns:
        sample = df[col].dropna().head(20)
        looks_numeric = sample.apply(_is_sap_number).mean() > 0.5
        if looks_numeric:
            df[col] = df[col].apply(_parse_sap_number)
    return df


def _is_sap_number(val: str) -> bool:
    cleaned = re.sub(r"[.,\s\-+]", "", str(val))
    return cleaned.isdigit()


def _parse_sap_number(val) -> float | None:
    if pd.isna(val) or str(val).strip() in ("", "nan", "None"):
        return None
    s = str(val).strip()
    # SAP: 1.234,56  →  1234.56
    if "," in s and "." in s:
        if s.rindex(",") > s.rindex("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _parse_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Detect columns that look like SAP date strings and convert them."""
    df = df.copy()
    date_pattern = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")

    for col in df.columns:
        sample = df[col].dropna().head(20)
        if sample.apply(lambda v: bool(date_pattern.match(str(v)))).mean() > 0.7:
            df[col] = pd.to_datetime(df[col], format="%d.%m.%Y", errors="coerce")
    return df


def validate_data(df: pd.DataFrame, logger: logging.Logger | None = None) -> list[str]:
    """
    validate_data — basic quality checks; returns a list of issue strings.
    Delegates to quality_check.validate_data for the full implementation.
    """
    from quality_check import validate_data as _validate
    return _validate(df, logger)
