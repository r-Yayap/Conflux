# core/merge_service.py

from typing import List, Optional, Dict
import pandas as pd

def merge_dataframes(
    dfs: List[pd.DataFrame],
    ref_columns: List[str],
    title_columns: List[Optional[str]],
    metadata: Dict
) -> pd.DataFrame:
    """
    Merge a list of DataFrames on their reference columns, optionally carrying along title columns.

    Args:
        dfs: list of DataFrames to merge (in the same order as ref_columns).
        ref_columns: list of column names in each DataFrame to join on.
        title_columns: list of optional title‐column names for each DataFrame.
        metadata: dict carrying extra info (e.g. hyperlinks, original_row_index).

    Returns:
        A single merged DataFrame, with columns renamed to:
          - number_1, number_2, … for each ref column
          - common_ref
          - title_excel1, title_excel2, … for any title_columns provided
    """
    # Validate inputs
    if len(dfs) != len(ref_columns):
        raise ValueError("dfs and ref_columns must have the same length")

    # 1) Prepare each DataFrame: rename ref col, compute counts, set common_ref, rename title col
    prepared = []
    for idx, (df, ref_col) in enumerate(zip(dfs, ref_columns), start=1):
        if ref_col not in df.columns:
            raise KeyError(f"Reference column '{ref_col}' not found in DataFrame #{idx}")

        working = df.copy().fillna("")  # preserve original

        # Rename reference column to a uniform name
        num_col = f"number_{idx}"
        working = working.rename(columns={ref_col: num_col})

        # Compute occurrence count for duplicates of the same ref
        working["refno_count"] = working.groupby(num_col).cumcount()

        # Set common_ref from this number
        working["common_ref"] = working[num_col]

        # Rename title column if provided
        title_col = title_columns[idx-1] if idx-1 < len(title_columns) else None
        if title_col and title_col in working.columns:
            working = working.rename(columns={title_col: f"title_excel{idx}"})

        prepared.append(working)

    # 2) Merge all DataFrames pairwise
    merged = prepared[0]
    for next_df in prepared[1:]:
        merged = pd.merge(
            merged,
            next_df,
            on=["common_ref", "refno_count"],
            how="outer",
            suffixes=("", "")
        ).fillna("")

    # 3) Clean up
    # Drop the temporary refno_count column
    if "refno_count" in merged.columns:
        merged = merged.drop(columns=["refno_count"])

    # Ensure original_row_index is numeric if present
    if "original_row_index" in merged.columns:
        merged["original_row_index"] = pd.to_numeric(
            merged["original_row_index"], errors="coerce"
        ).fillna(0).astype(int)

    return merged
