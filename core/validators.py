# core/validators.py

from dataclasses import dataclass
from typing import List, Optional, Tuple
import pandas as pd

@dataclass
class CheckConfig:
    status_column:   Optional[str]           = None
    status_value:    Optional[str]           = None
    project_column:  Optional[str]           = None
    project_value:   Optional[str]           = None
    custom_checks:   List[Tuple[str, str]]   = None
    filename_column: Optional[str]           = None

def append_comment(existing: str, new_comment: str) -> str:
    """Appends a new_comment to existing, separated by newline if needed."""
    new_comment = new_comment.strip()
    if not new_comment:
        return existing or ""
    if existing:
        return f"{existing}\n{new_comment}"
    return new_comment

def apply_validators(df: pd.DataFrame, config: CheckConfig) -> pd.DataFrame:
    """
    Applies status, project, custom, and filename checks to populate a 'Comments_1' column.
    """
    # Ensure Comments_1 column exists
    if "Comments_1" not in df.columns:
        df["Comments_1"] = ""

    # 1. Build unified list of (column, expected_value) checks
    all_checks: List[Tuple[str, str]] = []

    if config.status_column and config.status_value is not None:
        all_checks.append((config.status_column, config.status_value))

    if config.project_column and config.project_value is not None:
        all_checks.append((config.project_column, config.project_value))

    if config.custom_checks:
        all_checks.extend(config.custom_checks)

    # 2. Apply all column/value mismatch checks
    for col_name, expected in all_checks:
        def check_row(row):
            # Skip rows without a reference number
            ref = row.get("number_1")
            if pd.isna(ref) or not str(ref).strip():
                return row["Comments_1"]
            actual = str(row.get(col_name, "")).strip()
            if actual.lower() != str(expected).strip().lower():
                comment = f"{col_name} Mismatch: {actual} <--> {expected}"
                return append_comment(row["Comments_1"], comment)
            return row["Comments_1"]

        df["Comments_1"] = df.apply(check_row, axis=1)

    # 3. Filename vs reference number check
    if config.filename_column:
        filename_col = config.filename_column
        def filename_check(row):
            ref = str(row.get("number_1", "")).strip()
            filename = str(row.get(filename_col, "")).strip()
            if ref and filename and not filename.startswith(ref):
                comment = f"Filename & Drawing Number Mismatch"
                return append_comment(row["Comments_1"], comment)
            return row["Comments_1"]

        df["Comments_1"] = df.apply(filename_check, axis=1)

    return df
