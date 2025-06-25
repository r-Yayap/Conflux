"""
Reader module: loads Excel files into DataFrames and extracts metadata (hyperlinks, original row indices).
"""
from typing import List, Tuple, Dict
import pandas as pd
from openpyxl import load_workbook


def extract_hyperlinks(path: str) -> Dict[int, Dict[str, str]]:
    """
    Extracts hyperlinks from the given Excel file.

    Returns a dict mapping original row numbers to a dict of {column_header: hyperlink_target}.
    """
    wb = load_workbook(path, data_only=False)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(max_row=1))]

    hyperlinks: Dict[int, Dict[str, str]] = {}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_idx = row[0].row
        row_links: Dict[str, str] = {}
        for idx, cell in enumerate(row):
            if cell.hyperlink:
                col_name = headers[idx]
                row_links[col_name] = cell.hyperlink.target
        if row_links:
            hyperlinks[row_idx] = row_links
    wb.close()
    return hyperlinks


def extract_original_row_indices(path: str) -> List[int]:
    """
    Returns a list of Excel row numbers for all non-blank data rows (excluding header).
    """
    wb = load_workbook(path, data_only=False)
    ws = wb.active
    original_indices: List[int] = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        values = [cell.value for cell in row]
        if any(v is not None and str(v).strip() != "" for v in values):
            original_indices.append(row[0].row)
    wb.close()
    return original_indices


def read_excels(
    paths: List[str],
    ref_columns: List[str]
) -> Tuple[List[pd.DataFrame], Dict]:
    """
    Loads multiple Excel files into pandas DataFrames, and extracts metadata from the first file.

    Args:
        paths: List of Excel file paths (2 or 3 elements).
        ref_columns: Corresponding reference column names (for validation downstream).

    Returns:
        dfs: List of DataFrames loaded from each path.
        metadata: Dict containing:
            - "hyperlinks": mapping of original row number to hyperlinks dict
            - "original_row_indices": list of original row numbers for data rows
            - "headers": list of column headers from the first workbook
    """
    # Load all files into DataFrames
    dfs: List[pd.DataFrame] = []
    for path in paths:
        df = pd.read_excel(path, engine="openpyxl", dtype=str).fillna("")
        dfs.append(df)

    # Extract metadata from the first file
    first_path = paths[0]
    headers = list(dfs[0].columns)
    hyperlinks = extract_hyperlinks(first_path)
    original_row_indices = extract_original_row_indices(first_path)

    metadata: Dict = {
        "hyperlinks": hyperlinks,
        "original_row_indices": original_row_indices,
        "headers": headers
    }

    return dfs, metadata
