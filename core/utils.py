# core/utils.py
import os
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


def extract_drawing_from_filename(fn: str, num_tokens: int) -> str:
    """
    Given a filename like "A-B-C-D-extra.pdf" and num_tokens=4,
    return "A-B-C-D". If fn has fewer tokens, just return fn.
    """
    parts = fn.split("-")
    return "-".join(parts[:num_tokens]) if len(parts) >= num_tokens else fn


def _is_empty(val) -> bool:
    """True if value is None, NaN, empty or whitespace-only string."""
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False

def remerge_by_filename(
    df: pd.DataFrame,
    filename_col: Optional[str]
) -> pd.DataFrame:
    """
    For rows where number_1 exists but number_2/3 are blank,
    extract a candidate from the filename, find that in number_2/3
    **only among rows that originally had no number_1**, copy into
    truly-empty target cells, mark them Remerged=True, then drop
    the original orphan row.
    """
    if not filename_col or filename_col not in df.columns:
        df["Remerged"] = False
        return df

    df = df.copy()
    df["Remerged"] = False
    to_drop = []
    has3 = "number_3" in df.columns

    for idx, row in df.iterrows():
        n1 = str(row.get("number_1", "")).strip()
        fn = str(row.get(filename_col, "")).strip()
        # strip extension (so `.pdf` / etc. doesn’t pollute our tokenization)
        fn_base, _ = os.path.splitext(fn)

        # only look at orphans: have a number_1 but no number_2/3 yet
        if not n1 or (not _is_empty(row.get("number_2"))) or (has3 and not _is_empty(row.get("number_3"))):
            continue

        # build our “drawing” candidate from the filename
        token_count = len(n1.split("-"))
        cand = extract_drawing_from_filename(fn_base, token_count)

        # now look in number_2 → number_3, but only among rows whose number_1 was blank
        for num_col in ("number_2", "number_3") if has3 else ("number_2",):
            matches = df[
                (df[num_col].astype(str).str.strip() == cand) &
                (df["number_1"].astype(str).str.strip() == "")
            ]
            if matches.empty:
                continue

            target = matches.index[0]
            # copy **only** into truly-empty cells
            for col in df.columns:
                if col in (
                    "number_1", "number_2", "number_3",
                    "common_ref", "Remerged",
                    "refno_count", "original_row_index", "title_match"
                ):
                    continue
                src = row[col]
                tgt = df.at[target, col]
                if not _is_empty(src) and _is_empty(tgt):
                    df.at[target, col] = src

            # now fix up our numbers & flags
            df.at[target, "number_1"] = n1
            df.at[target, "common_ref"] = cand
            df.at[target, "Remerged"] = True
            # **preserve** the original_row_index so hyperlinks stick
            df.at[target, "original_row_index"] = row["original_row_index"]

            to_drop.append(idx)
            break

    if to_drop:
        df = df.drop(index=to_drop).reset_index(drop=True)

    return df