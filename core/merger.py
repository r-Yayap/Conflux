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
from .utils import add_title_match_column

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
        check_config:  Optional[CheckConfig] = None
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

        # Add true or false
        merged_df = add_title_match_column(merged_df, title_columns)

        # 3) Apply validation rules (Comments_1, status/project/custom/filename checks)
        if check_config:
            merged_df = apply_validators(merged_df, check_config)

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
