# core/formatter.py

from typing import List, Optional, Dict
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


from core.validators import CheckConfig

def write_styled_excel(
    merged_df: pd.DataFrame,
    metadata: Dict,
    output_path: str,
    title_columns: List[Optional[str]],
    check_config: CheckConfig
) -> None:
    """
    Orchestrates writing the merged_df to Excel at output_path,
    then re-opening and applying styling, hyperlinks, and rich text.
    """
    # 1. Save raw DataFrame
    merged_df.to_excel(output_path, index=False, header=True)

    # 2. Load workbook & sheet
    wb = load_workbook(output_path, data_only=False, rich_text=True)
    ws = wb.active

    # 3. Reorder columns
    reorder_columns(ws)

    # 4. Apply formatting, fills, duplicates, instance/case, hyperlinks
    apply_formatting_and_hyperlinks(ws, metadata, merged_df, check_config)

    # 5. Apply title_match rich-text coloring
    apply_title_match_highlighting(ws, merged_df)

    # 6. Apply title-to-title highlighting
    apply_title_highlighting(ws, merged_df, title_columns)

    # 7. Save changes
    wb.save(output_path)
    wb.close()


def reorder_columns(ws: Worksheet) -> None:
    """Reorder columns to: data cols, common_ref, title_excels, title_match, Comments_1, Instance/Case."""
    original_headers = [cell.value for cell in ws[1]]
    data = list(ws.iter_rows(min_row=2, values_only=True))

    # Identify columns
    title_cols = [c for c in ["title_excel1", "title_excel2", "title_excel3"] if c in original_headers]
    final_cols = [c for c in ["title_match", "Comments_1", "Instance", "Case"] if c in original_headers]
    excluded = set(title_cols) | set(final_cols) | {"common_ref"}

    data_cols = [c for c in original_headers if c not in excluded]
    new_order = data_cols + (["common_ref"] if "common_ref" in original_headers else []) + title_cols + final_cols

    # Clear existing
    ws.delete_cols(1, ws.max_column)
    # Write headers
    for idx, header in enumerate(new_order, start=1):
        ws.cell(row=1, column=idx, value=header)
    # Write data
    header_index = {h: i for i, h in enumerate(original_headers)}
    for row_i, row_vals in enumerate(data, start=2):
        for col_i, header in enumerate(new_order, start=1):
            orig_idx = header_index.get(header)
            ws.cell(row=row_i, column=col_i, value=row_vals[orig_idx] if orig_idx is not None else None)


from typing import Dict
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill
from core.validators import CheckConfig

def apply_formatting_and_hyperlinks(
    ws: Worksheet,
    metadata: Dict,
    merged_df: pd.DataFrame,
    config: CheckConfig
) -> None:
    """Apply cell fills, fonts for duplicates, Instance/Case, and re-insert hyperlinks."""
    # Build header → column-index map
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i+1 for i, h in enumerate(headers)}

    # Font for duplicates
    dup_font = Font(bold=True, color="FF3300")

    # Determine if we're merging 3 files
    has3 = 'number_3' in merged_df.columns

    # Fill style mappings
    if has3:
        fills = {
            (True, False, False): PatternFill("solid", fgColor="CC99FF"),
            (False, True, False): PatternFill("solid", fgColor="FFCC66"),
            (False, False, True): PatternFill("solid", fgColor="AACCFF"),
            (True, True, False): PatternFill("solid", fgColor="FFC0CB"),
            (True, False, True): PatternFill("solid", fgColor="90EE90"),
            (False, True, True): PatternFill("solid", fgColor="FFD700"),
        }
    else:
        fills = {
            (True, False): PatternFill("solid", fgColor="CC99FF"),
            (False, True): PatternFill("solid", fgColor="FFCC66"),
        }

    # Duplicate value sets
    dup_sets = {
        col: set(merged_df[col][merged_df[col].duplicated()])
        for col in (['number_1', 'number_2', 'number_3', 'common_ref'] if has3
                    else ['number_1', 'number_2', 'common_ref'])
    }

    # Add the Instance/Case column header
    inst_col = ws.max_column + 1
    inst_header = "Case" if has3 else "Instance"
    ws.cell(row=1, column=inst_col, value=inst_header)

    # Build the presence → text mapping once
    if has3:
        instance_mapping = {
            (True, True, True): ""   # all three present → blank
        }
    else:
        instance_mapping = {
            (True, True): ""         # both present → blank
        }

    # Iterate rows
    for i, row in merged_df.iterrows():
        r = i + 2  # Excel row index

        # Compute presence tuple
        if has3:
            presence = (
                bool(ws.cell(r, col_idx['number_1']).value),
                bool(ws.cell(r, col_idx['number_2']).value),
                bool(ws.cell(r, col_idx['number_3']).value)
            )
            number_cols = ['number_1', 'number_2', 'number_3']
        else:
            presence = (
                bool(ws.cell(r, col_idx['number_1']).value),
                bool(ws.cell(r, col_idx['number_2']).value),
            )
            number_cols = ['number_1', 'number_2']

        # Apply fill if this presence combination is mapped
        if presence in fills:
            fill = fills[presence]
            # Fill each number column and common_ref
            for colname, present in zip(number_cols, presence):
                if present:
                    ws.cell(r, col_idx[colname]).fill = fill
            # common_ref
            if 'common_ref' in col_idx:
                ws.cell(r, col_idx['common_ref']).fill = fill

        # Apply duplicate-font for any duplicated values
        for colname, dset in dup_sets.items():
            if colname in col_idx:
                val = ws.cell(r, col_idx[colname]).value
                if val in dset:
                    ws.cell(r, col_idx[colname]).font = dup_font

        # Set Instance/Case text
        ws.cell(r, inst_col).value = instance_mapping.get(presence, "None")

    # Re-insert hyperlinks from metadata
    for orig_idx, cols in metadata.get('hyperlinks', {}).items():
        match = merged_df[merged_df['original_row_index'] == orig_idx]
        if not match.empty:
            r = match.index[0] + 2
            for colname, url in cols.items():
                if colname in col_idx:
                    cell = ws.cell(r, col_idx[colname])
                    cell.hyperlink = url
                    cell.style = "Hyperlink"

    # Highlight failed checks in light red
    light_red = PatternFill("solid", fgColor="FFCCCC")
    check_items = [
        (config.status_column, config.status_value),
        (config.project_column, config.project_value)
    ] + (config.custom_checks or [])

    for colname, expected in check_items:
        if not colname:
            continue
        cidx = col_idx.get(colname)
        if not cidx:
            continue
        for i, row in merged_df.iterrows():
            actual = row.get(colname, "")
            if pd.notna(row.get('number_1')) and str(actual).strip().lower() != str(expected).strip().lower():
                ws.cell(i+2, cidx).fill = light_red


def get_ws_column_index(ws: Worksheet, header_name: str) -> Optional[int]:
    """Case-insensitive find of header in first row → column index."""
    for cell in ws[1]:
        if isinstance(cell.value, str) and cell.value.strip().lower() == header_name.strip().lower():
            return cell.column
    return None


def apply_title_match_highlighting(ws: Worksheet, merged_df: pd.DataFrame) -> None:
    """Color 'True' green and 'False' red in the title_match column."""
    col_idx = get_ws_column_index(ws, "title_match")
    if not col_idx:
        return
    for i, row in merged_df.iterrows():
        text = str(row.get("title_match", ""))
        rich = CellRichText()
        for j, token in enumerate(text.split(",")):
            tok = token.strip()
            font = InlineFont(rFont="Calibri", sz=11)
            font.color = "008000" if tok.lower()=="true" else ("FF0000" if tok.lower()=="false" else "000000")
            rich.append(TextBlock(font, tok))
            if j < len(text.split(",")) - 1:
                sep = InlineFont(rFont="Calibri", sz=11, color="000000")
                rich.append(TextBlock(sep, ", "))
        ws.cell(i+2, col_idx).value = rich


def tokenize_with_indices(text: str):
    import re
    return [(m.group(), m.start()) for m in re.finditer(r'[A-Za-z0-9]+|[_–—-]|[^\w\s]', text)]


def dp_align_tokens(tokens1, tokens2):
    from difflib import SequenceMatcher
    s1 = [t[0] for t in tokens1]
    s2 = [t[0] for t in tokens2]
    n, m = len(s1), len(s2)
    dp = [[0]*(m+1) for _ in range(n+1)]
    back = [[None]*(m+1) for _ in range(n+1)]
    for i in range(1, n+1):
        dp[i][0]=i; back[i][0]='UP'
    for j in range(1, m+1):
        dp[0][j]=j; back[0][j]='LEFT'
    for i in range(1,n+1):
        for j in range(1,m+1):
            sim = SequenceMatcher(None, s1[i-1], s2[j-1]).ratio()
            if s1[i-1]==s2[j-1]:
                cost=dp[i-1][j-1]; flag='DIAG'
            elif s1[i-1].lower()==s2[j-1].lower():
                cost=dp[i-1][j-1]+0.5; flag='DIAG'
            elif sim>=0.8:
                cost=dp[i-1][j-1]+(1-sim); flag='DIAG'
            else:
                cost=float('inf')
            delete=dp[i-1][j]+1; insert=dp[i][j-1]+1
            best=min(cost, delete, insert)
            dp[i][j]=best
            if best==cost: back[i][j]='DIAG'
            elif best==delete: back[i][j]='UP'
            else: back[i][j]='LEFT'
    # backtrack
    i,j=n,m
    a1,a2,flags=[],[],[]
    while i>0 or j>0:
        move=back[i][j]
        if move=='DIAG':
            a1.append(tokens1[i-1]); a2.append(tokens2[j-1]); flags.append("MATCH")
            i-=1; j-=1
        elif move=='UP':
            a1.append(tokens1[i-1]); a2.append((None,None)); flags.append("DEL")
            i-=1
        else:
            a1.append((None,None)); a2.append(tokens2[j-1]); flags.append("INS")
            j-=1
    return a1[::-1], a2[::-1], flags[::-1]


def create_rich_text(original: str, aligned1, aligned2, flags):
    rich = CellRichText()
    idx=0
    for (tok, pos), (_, _), flag in zip(aligned1, aligned2, flags):
        if tok is None: continue
        if pos>idx:
            rich.append(original[idx:pos])
        color="000000"
        if flag=="MATCH":
            color="000000"
        elif flag=="DEL":
            color="FF0000"
        elif flag=="INS":
            color="FFA500"
        font=InlineFont(rFont="Calibri", sz=11, color=color)
        rich.append(TextBlock(font, tok))
        idx=pos+len(tok)
    if idx<len(original):
        rich.append(original[idx:])
    return rich

def apply_title_highlighting(
    ws: Worksheet,
    merged_df: pd.DataFrame,
    title_columns: List[Optional[str]]
) -> None:
    """
    For each provided title column after the first, compare it to title_columns[0]
    and write rich‐text diffs back into both columns.
    """
    # The “baseline” title is always the first one
    baseline_col = title_columns[0] if title_columns else None
    if not baseline_col:
        return

    base_idx = get_ws_column_index(ws, baseline_col)
    if base_idx is None:
        return

    # For each other title column
    for other_col in title_columns[1:]:
        if not other_col:
            continue
        other_idx = get_ws_column_index(ws, other_col)
        if other_idx is None:
            continue

        # Compare row by row
        for i, row in merged_df.iterrows():
            excel_row = i + 2  # account for header

            text1 = str(row.get(baseline_col, ""))
            text2 = str(row.get(other_col, ""))

            # Tokenize and align
            tokens1 = tokenize_with_indices(text1)
            tokens2 = tokenize_with_indices(text2)
            aligned1, aligned2, flags = dp_align_tokens(tokens1, tokens2)

            # Build rich text
            rich1 = create_rich_text(text1, aligned1, aligned2, flags)
            rich2 = create_rich_text(text2, aligned2, aligned1, flags)

            # Write back (baseline only if you really want to overwrite original)
            ws.cell(row=excel_row, column=base_idx).value = rich1
            ws.cell(row=excel_row, column=other_idx).value = rich2

