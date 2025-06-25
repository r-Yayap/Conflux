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


class MergerFacade:
    """
    High-level facade for merging Excel files with validation and styling.

    Example usage:
        merged_df = MergerFacade.run_merge(
            paths=["file1.xlsx", "file2.xlsx"],
            ref_columns=["DrawingNo", "DrawingNo"],
            output_path="merged_output.xlsx",
            title_columns=["Title1", "Title2"],
            check_config=CheckConfig(
                status_column="Status", status_value="OK",
                project_column="Project", project_value="Alpha",
                custom_checks=[("Phase", "P1")],
                filename_column="Filename"
            )
        )
    """


    @staticmethod
    def run_merge(
            paths: List[str],
            ref_columns: List[str],
            output_path: str,
            *,
            title_columns: Optional[List[str]] = None,
            check_config: Optional[CheckConfig] = None
    ) -> pd.DataFrame:
        # 1) Read all Excels
        dfs, metadata = read_excels(paths, ref_columns)

        # 2) Merge with the correct signature
        merged_df = merge_dataframes(
            dfs,
            ref_columns=ref_columns,
            title_columns=title_columns or [],
            metadata=metadata
        )

        # 3) Validate
        if check_config:
            merged_df = apply_validators(merged_df, check_config)

        # 4) Style & write out
        write_styled_excel(
            merged_df=merged_df,
            metadata=metadata,
            output_path=output_path,
            title_columns=title_columns or [],
            check_config=check_config or CheckConfig()
        )

        return merged_df