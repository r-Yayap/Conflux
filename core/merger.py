# core/merger.py

"""
Facade for Excel merging: orchestrates reading, merging, validating, and styling.
"""
from typing import List, Optional
import pandas as pd

from .reader import read_excels  # loads List[DataFrame], metadata (hyperlinks, original indices)
from .merge_service import merge_dataframes  # pure merge of DataFrames
from .validators import apply_validators, CheckConfig  # populates Comments_1
from .formatter import write_styled_excel  # writes Excel with formatting, rich-text, hyperlinks
from .utils import add_title_match_column, remerge_by_filename
from .revision_checker import (
    RevCheckSettings,
    apply_revision_checks,
)

class MergerFacade:
    """
    High-level facade for merging Excel files with validation and styling.
    """

    @staticmethod
    def run_merge(
        paths: List[str],
        ref_columns: List[str],
        output_path: str,
        *,
        title_columns: Optional[List[str]]  = None,
        check_config:  Optional[CheckConfig] = None,
        rev_check_settings: Optional[RevCheckSettings] = None,
    ) -> pd.DataFrame:

        # 1) Read input files (and extract hyperlinks & original_row_index)
        dfs, metadata = read_excels(paths, ref_columns)

        # 2) Merge into a single DataFrame
        merged_df = merge_dataframes(
            dfs,
            ref_columns   = ref_columns,
            title_columns = title_columns or [],
            metadata      = metadata
        )

        # ──────────────────────────────────────────
        # Added Steps
        # ──────────────────────────────────────────

        # 2.a: re‐merge by filename when number_1 had no match
        merged_df = remerge_by_filename(merged_df, check_config.filename_column if check_config else None)

        # 2.b: Add true or false for title match
        merged_df = add_title_match_column(merged_df, title_columns)

        # 2.c: Detect duplicates in each number_★ column and label them
        n1 = merged_df["number_1"].fillna("").astype(str)
        mask1 = n1 != ""
        dup1 = mask1 & n1.duplicated(keep=False)

        n2 = merged_df["number_2"].fillna("").astype(str)
        mask2 = n2 != ""
        dup2 = mask2 & n2.duplicated(keep=False)

        has3 = "number_3" in merged_df.columns
        if has3:
            n3 = merged_df["number_3"].fillna("").astype(str)
            mask3 = n3 != ""
            dup3 = mask3 & n3.duplicated(keep=False)
        else:
            dup3 = None

        def make_dup_label(i):
            labels = []
            if dup1.iloc[i]:
                labels.append("Duplicate in PDF Title block")
            if dup2.iloc[i]:
                labels.append("Duplicate in LOD 2")
            if has3 and dup3.iloc[i]:
                labels.append("Duplicate in LOD 3")
            return "; ".join(labels)

        merged_df["Duplicate"] = [make_dup_label(i) for i in range(len(merged_df))]
        # ──────────────────────────────────────────
        # ──────────────────────────────────────────

        # 3) Apply validation rules (Comments_1, status/project/custom/filename checks)
        if check_config:
            merged_df = apply_validators(merged_df, check_config)

        # 3.b) Apply revision checker logic (Comments-Revision + highlights)
        merged_df, revision_highlights = apply_revision_checks(merged_df, rev_check_settings)
        metadata["revision_highlights"] = revision_highlights

        # 4) Prepare the *renamed* title column names for styling
        renamed_titles: List[str] = []
        for idx, orig in enumerate(title_columns or [], start=1):
            if orig:
                renamed_titles.append(f"title_excel{idx}")

        # 5) Write out styled Excel (fills, hyperlinks, rich-text, diffs)
        write_styled_excel(
            merged_df     = merged_df,
            metadata      = metadata,
            output_path   = output_path,
            title_columns = renamed_titles,
            check_config  = check_config or CheckConfig()
        )

        return merged_df
