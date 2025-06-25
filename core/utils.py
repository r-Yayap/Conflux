# core/utils.py

import re
import pandas as pd
from typing import List, Optional

def add_title_match_column(
    df: pd.DataFrame,
    title_columns: List[Optional[str]]
) -> pd.DataFrame:
    """
    Create a 'title_match' column comparing title_excel1 to each subsequent
    title_excelN, ignoring punctuation and whitespace in the comparison.
    title_columns is the list of original header names (or None) the user chose,
    in the same order as the input files.
    """
    # build the list of *renamed* columns for comparison:
    # for each idx, if title_columns[idx] is truthy, then there's a title_excel{idx+1} column
    renamed = [
        f"title_excel{idx+1}"
        for idx, col in enumerate(title_columns)
        if col  # they ticked this comparison on
    ]

    # nothing to compare if fewer than 2 columns
    if len(renamed) < 2:
        df["title_match"] = "N/A"
        return df

    # normalize helper: strip out non-alphanumeric, lowercase
    def normalize(s: str) -> str:
        return re.sub(r'[^A-Za-z0-9]+', '', s).lower()

    base, *others = renamed

    def compare_row(row):
        base_norm = normalize(str(row.get(base, "")))
        out = []
        for oc in others:
            other_norm = normalize(str(row.get(oc, "")))
            out.append("True" if base_norm == other_norm else "False")
        return out[0] if len(out) == 1 else ", ".join(out)

    df["title_match"] = df.apply(compare_row, axis=1)
    return df
