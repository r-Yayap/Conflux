# core/formatter.py

from typing import List, Optional, Dict
import pandas as pd
import re
from difflib import SequenceMatcher
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
    # 1. Save raw DataFrame
    merged_df.to_excel(output_path, index=False, header=True)

    # 2. Load workbook & sheet
    wb = load_workbook(output_path, data_only=False, rich_text=True)
    ws = wb.active

    # 3. Reorder columns
    reorder_columns(ws)

    # 4. Apply fills, duplicates, instance/case, hyperlinks
    apply_formatting_and_hyperlinks(ws, metadata, merged_df, check_config)

    # 5. Title-match rich-text coloring ("True"/"False")
    apply_title_match_highlighting(ws, merged_df)

    # 6. Title-to-title rich-text diffing
    apply_title_highlighting(ws, merged_df, title_columns)

    # 7. Save & close
    wb.save(output_path)
    wb.close()


def reorder_columns(ws: Worksheet) -> None:
    original_headers = [cell.value for cell in ws[1]]
    data = list(ws.iter_rows(min_row=2, values_only=True))

    title_cols = [c for c in ["title_excel1", "title_excel2", "title_excel3"] if c in original_headers]
    final_cols = [c for c in ["title_match", "Comments_1", "Instance", "Case"] if c in original_headers]
    excluded = set(title_cols) | set(final_cols) | {"common_ref"}

    data_cols = [c for c in original_headers if c not in excluded]
    new_order = (
        data_cols
        + (["common_ref"] if "common_ref" in original_headers else [])
        + title_cols
        + final_cols
    )

    ws.delete_cols(1, ws.max_column)
    for idx, header in enumerate(new_order, start=1):
        ws.cell(row=1, column=idx, value=header)

    header_index = {h: i for i, h in enumerate(original_headers)}
    for row_i, row_vals in enumerate(data, start=2):
        for col_i, header in enumerate(new_order, start=1):
            orig_idx = header_index.get(header)
            ws.cell(row=row_i, column=col_i, value=row_vals[orig_idx] if orig_idx is not None else None)


def apply_formatting_and_hyperlinks(
    ws: Worksheet,
    metadata: Dict,
    merged_df: pd.DataFrame,
    config: CheckConfig
) -> None:
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i+1 for i, h in enumerate(headers)}

    dup_font = Font(bold=True, color="FF3300")
    has3 = 'number_3' in merged_df.columns

    fills = {
        (True, False, False): PatternFill("solid", fgColor="CC99FF"),
        (False, True, False): PatternFill("solid", fgColor="FFCC66"),
        (False, False, True): PatternFill("solid", fgColor="AACCFF"),
        (True, True, False): PatternFill("solid", fgColor="FFC0CB"),
        (True, False, True): PatternFill("solid", fgColor="90EE90"),
        (False, True, True): PatternFill("solid", fgColor="FFD700"),
    } if has3 else {
        (True, False): PatternFill("solid", fgColor="CC99FF"),
        (False, True): PatternFill("solid", fgColor="FFCC66"),
    }

    dup_cols = ['number_1', 'number_2', 'common_ref'] + (['number_3'] if has3 else [])
    dup_sets = {col: set(merged_df[col][merged_df[col].duplicated()]) for col in dup_cols}

    inst_col = ws.max_column + 1
    inst_header = "Case" if has3 else "Instance"
    ws.cell(row=1, column=inst_col, value=inst_header)

    instance_mapping = {
        (True, True, True): ""   # all three present
    } if has3 else {
        (True, True): ""         # both present
    }

    for i, _ in merged_df.iterrows():
        r = i + 2
        if has3:
            presence = (
                bool(ws.cell(r, col_idx['number_1']).value),
                bool(ws.cell(r, col_idx['number_2']).value),
                bool(ws.cell(r, col_idx['number_3']).value),
            )
            number_cols = ['number_1', 'number_2', 'number_3']
        else:
            presence = (
                bool(ws.cell(r, col_idx['number_1']).value),
                bool(ws.cell(r, col_idx['number_2']).value),
            )
            number_cols = ['number_1', 'number_2']

        # fill
        if presence in fills:
            fill = fills[presence]
            for colname, present in zip(number_cols, presence):
                if present:
                    ws.cell(r, col_idx[colname]).fill = fill
            if 'common_ref' in col_idx:
                ws.cell(r, col_idx['common_ref']).fill = fill

        # duplicate font
        for colname, dset in dup_sets.items():
            if colname in col_idx:
                val = ws.cell(r, col_idx[colname]).value
                if val in dset:
                    ws.cell(r, col_idx[colname]).font = dup_font

        # instance/case text
        ws.cell(r, inst_col).value = instance_mapping.get(presence, "None")

    # hyperlinks
    for orig_idx, cols in metadata.get('hyperlinks', {}).items():
        match = merged_df[merged_df.get('original_row_index', -1) == orig_idx]
        if not match.empty:
            r = match.index[0] + 2
            for colname, url in cols.items():
                if colname in col_idx:
                    cell = ws.cell(r, col_idx[colname])
                    cell.hyperlink = url
                    cell.style = "Hyperlink"

    # highlight failed checks
    light_red = PatternFill("solid", fgColor="FFCCCC")
    checks = [
        (config.status_column, config.status_value),
        (config.project_column, config.project_value)
    ] + (config.custom_checks or [])

    for colname, expected in checks:
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
    for cell in ws[1]:
        if isinstance(cell.value, str) and cell.value.strip().lower() == header_name.strip().lower():
            return cell.column
    return None


def apply_title_match_highlighting(ws: Worksheet, merged_df: pd.DataFrame) -> None:
    col_idx = get_ws_column_index(ws, "title_match")
    if not col_idx:
        return

    for i, row in merged_df.iterrows():
        text = str(row.get("title_match", ""))
        rich = CellRichText()
        for j, token in enumerate(text.split(",")):
            tok = token.strip()
            font = InlineFont(rFont="Calibri", sz=11)
            if tok.lower() == "true":
                font.color = "008000"
            elif tok.lower() == "false":
                font.color = "FF0000"
            rich.append(TextBlock(font, tok))
            if j < len(text.split(",")) - 1:
                sep = InlineFont(rFont="Calibri", sz=11, color="000000")
                rich.append(TextBlock(sep, ", "))
        ws.cell(i+2, col_idx).value = rich


def tokenize_with_indices(text: str):
    return [(m.group(), m.start())
            for m in re.finditer(r'[A-Za-z0-9]+|[_–—-]|[^\w\s]', text)]


def dp_align_tokens(tokens1, tokens2):
    from difflib import SequenceMatcher

    s1 = [t[0] for t in tokens1]
    s2 = [t[0] for t in tokens2]
    n, m = len(s1), len(s2)

    # cost matrix & backpointers
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    bt = [[None] * (m + 1) for _ in range(n + 1)]

    # initialize
    for i in range(1, n + 1):
        dp[i][0] = i
        bt[i][0] = "DEL"
    for j in range(1, m + 1):
        dp[0][j] = j
        bt[0][j] = "INS"

    # fill
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            a, b = s1[i - 1], s2[j - 1]
            if a == b:
                cost_sub, flag = dp[i - 1][j - 1], "EXACT"
            elif a.lower() == b.lower():
                cost_sub, flag = dp[i - 1][j - 1] + 0.5, "CASE_ONLY"
            else:
                sim = SequenceMatcher(None, a, b).ratio()
                if sim >= 0.8:
                    cost_sub, flag = dp[i - 1][j - 1] + (1 - sim), "CHAR_LEVEL"
                else:
                    cost_sub, flag = float("inf"), None

            cost_del = dp[i - 1][j] + 1
            cost_ins = dp[i][j - 1] + 1

            best = min(cost_sub, cost_del, cost_ins)
            dp[i][j] = best

            if best == cost_sub:
                bt[i][j] = flag
            elif best == cost_del:
                bt[i][j] = "DEL"
            else:
                bt[i][j] = "INS"

    # backtrack
    aligned1, aligned2, flags = [], [], []
    i, j = n, m
    while i > 0 or j > 0:
        move = bt[i][j]
        if move in ("EXACT", "CASE_ONLY", "CHAR_LEVEL"):
            aligned1.append(tokens1[i - 1])
            aligned2.append(tokens2[j - 1])
            flags.append(move)
            i, j = i - 1, j - 1
        elif move == "DEL":
            aligned1.append(tokens1[i - 1])
            aligned2.append((None, None))
            flags.append("DEL")
            i -= 1
        else:  # INS
            aligned1.append((None, None))
            aligned2.append(tokens2[j - 1])
            flags.append("INS")
            j -= 1

    return aligned1[::-1], aligned2[::-1], flags[::-1]

def create_rich_text(original, aligned1, aligned2, flags):
    """
    Build a CellRichText so that:
      - EXACT → black
      - CASE_ONLY → gray
      - CHAR_LEVEL → per-char: match=black, diff=orange
      - DEL/INS → red
    """
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    rich = CellRichText()
    idx = 0

    for (tok1, pos), (tok2, _), flag in zip(aligned1, aligned2, flags):
        if tok1 is None:
            continue

        # write any intervening literal text
        if pos > idx:
            rich.append(original[idx:pos])

        # choose style by flag
        if flag == "EXACT":
            font = InlineFont(rFont="Calibri", sz=11, color="000000")
            rich.append(TextBlock(font, tok1))

        elif flag == "CASE_ONLY":
            font = InlineFont(rFont="Calibri", sz=11, color="808080")
            rich.append(TextBlock(font, tok1))

        elif flag == "CHAR_LEVEL":
            # compare char by char
            for c_idx, ch in enumerate(tok1):
                other_ch = tok2[c_idx] if tok2 and c_idx < len(tok2) else None
                color = "000000" if ch == other_ch else "FFA500"
                font = InlineFont(rFont="Calibri", sz=11, color=color)
                rich.append(TextBlock(font, ch))

        elif flag == "DEL" or flag == "INS":
            font = InlineFont(rFont="Calibri", sz=11, color="FF0000")
            rich.append(TextBlock(font, tok1))

        else:
            # fallback
            font = InlineFont(rFont="Calibri", sz=11, color="000000")
            rich.append(TextBlock(font, tok1))

        idx = pos + len(tok1)

    # any trailing text
    if idx < len(original):
        rich.append(original[idx:])

    return rich

def apply_title_highlighting(
    ws: Worksheet,
    merged_df: pd.DataFrame,
    title_columns: List[Optional[str]]
) -> None:
    from openpyxl.cell.rich_text import CellRichText

    # baseline is first title
    baseline = title_columns[0] if title_columns else None
    if not baseline:
        return

    base_idx = get_ws_column_index(ws, baseline)
    if base_idx is None:
        return

    # for each other title
    for other in title_columns[1:]:
        if not other:
            continue
        other_idx = get_ws_column_index(ws, other)
        if other_idx is None:
            continue

        for i, row in merged_df.iterrows():
            excel_row = i + 2
            t1 = str(row.get(baseline, ""))
            t2 = str(row.get(other, ""))

            # tokenize
            tok1 = tokenize_with_indices(t1)
            tok2 = tokenize_with_indices(t2)

            # align
            aligned1, aligned2, flags = dp_align_tokens(tok1, tok2)

            # build rich text
            rich1 = create_rich_text(t1, aligned1, aligned2, flags)
            rich2 = create_rich_text(t2, aligned2, aligned1, flags)

            # write back
            ws.cell(row=excel_row, column=base_idx).value = rich1
            ws.cell(row=excel_row, column=other_idx).value = rich2