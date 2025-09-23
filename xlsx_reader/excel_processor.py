"""Excel processing module for reading XLSX files.

This module provides simple functions to read Excel files and get basic
sheet information using pandas.
"""

from collections.abc import Callable

import pandas as pd


def get_sheet_names(file_path: str) -> list[str]:
    """Get all sheet names from an Excel file.

    Args:
        file_path (str): Path to the Excel file (.xlsx)

    Returns:
        List[str]: List of sheet names in the workbook

    Note:
        Students should implement this function to return a list of all
        sheet names in the Excel file. Use pandas.ExcelFile to read
        the sheet names without loading all data.
    """
    excel_file = pd.ExcelFile(file_path)
    sheet_name = excel_file.sheet_names
    return sheet_name
    # raise NotImplementedError


def get_sheet_row_count(file_path: str, sheet_name: str) -> int:
    """Get the number of rows in a specific sheet.

    Args:
        file_path (str): Path to the Excel file (.xlsx)
        sheet_name (str): Name of the sheet to analyze

    Returns:
        int: Number of rows in the sheet

    Note:
        Students should implement this function to return the number of
        rows in the specified sheet. Use pandas.read_excel() to read
        the sheet and return len(dataframe).
    """
    excel_file = pd.read_excel(file_path, sheet_name)
    row_no = excel_file.shape[0]
    return row_no
    # raise NotImplementedError


def process_excel_file(
    file_path: str, progress_callback: Callable[[int, int, str], None] | None = None
) -> dict[str, int]:
    """Process an Excel file and return row counts for each sheet.

    Args:
        file_path (str): Path to the Excel file (.xlsx)
        progress_callback (Optional[Callable]): Function to call for progress updates.
            Receives (current_sheet_index, total_sheets, sheet_name)

    Returns:
        Dict[str, int]: Dictionary mapping sheet names to their row counts

    Note:
        1. Get all sheet names using get_sheet_names()
        2. For each sheet, call get_sheet_row_count()
        3. Call progress_callback if provided
        4. Return a dictionary with sheet names as keys and row counts as values
    """
    sheet_names = get_sheet_names(file_path)
    total_sheets = len(sheet_names)
    results = {}

    for current_index, sheet_name in enumerate(sheet_names):
        if progress_callback:
            progress_callback(current_index, total_sheets, sheet_name)

        row_count = get_sheet_row_count(file_path, sheet_name)
        results[sheet_name] = row_count

    return results
