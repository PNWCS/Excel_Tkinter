"""Tests for Excel processing functions.

This module tests the core Excel processing functionality using simple
test cases with pre-created Excel files.
"""

from pathlib import Path

import pandas as pd
import pytest

# from pathlib import Path
from xlsx_reader.excel_processor import get_sheet_names, get_sheet_row_count, process_excel_file


def create_test_excel(file_path: Path) -> None:
    """Create a simple test Excel file with multiple sheets."""
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        df1 = pd.DataFrame({"A": range(1, 11), "B": [f"Row {i}" for i in range(1, 11)]})
        df1.to_excel(writer, sheet_name="Sheet1", index=False)

        df2 = pd.DataFrame({"X": range(1, 6), "Y": [f"Data {i}" for i in range(1, 6)]})
        df2.to_excel(writer, sheet_name="Sheet2", index=False)

        df3 = pd.DataFrame(columns=["Col1", "Col2"])
        df3.to_excel(writer, sheet_name="EmptySheet", index=False)


class TestExcelProcessor:
    @pytest.fixture
    def test_excel_file(self, tmp_path: Path):
        """Create an Excel file in pytest's temp directory."""
        file_path = tmp_path / "test.xlsx"
        create_test_excel(file_path)
        return str(file_path)  # just return path, no manual unlink

    def test_get_sheet_names(self, test_excel_file):
        """Test getting sheet names from Excel file."""
        sheet_names = get_sheet_names(test_excel_file)

        assert isinstance(sheet_names, list)
        assert len(sheet_names) == 3
        assert "Sheet1" in sheet_names
        assert "Sheet2" in sheet_names
        assert "EmptySheet" in sheet_names

    def test_get_sheet_row_count_normal_sheet(self, test_excel_file):
        """Test getting row count from a normal sheet."""
        row_count = get_sheet_row_count(test_excel_file, "Sheet1")
        assert row_count == 10

    def test_get_sheet_row_count_small_sheet(self, test_excel_file):
        """Test getting row count from a smaller sheet."""
        row_count = get_sheet_row_count(test_excel_file, "Sheet2")
        assert row_count == 5

    def test_get_sheet_row_count_empty_sheet(self, test_excel_file):
        """Test getting row count from an empty sheet."""
        row_count = get_sheet_row_count(test_excel_file, "EmptySheet")
        assert row_count == 0

    def test_process_excel_file(self, test_excel_file):
        """Test processing entire Excel file."""
        results = process_excel_file(test_excel_file)

        assert isinstance(results, dict)
        assert len(results) == 3
        assert results["Sheet1"] == 10
        assert results["Sheet2"] == 5
        assert results["EmptySheet"] == 0

    def test_process_excel_file_with_callback(self, test_excel_file):
        """Test processing Excel file with progress callback."""
        progress_calls = []

        def progress_callback(current, total, sheet_name):
            progress_calls.append((current, total, sheet_name))

        process_excel_file(test_excel_file, progress_callback)

        # Check that callback was called for each sheet
        assert len(progress_calls) == 3
        assert progress_calls[0] == (0, 3, "Sheet1")
        assert progress_calls[1] == (1, 3, "Sheet2")
        assert progress_calls[2] == (2, 3, "EmptySheet")

    def test_file_not_found(self):
        """Test handling of non-existent file."""
        with pytest.raises(FileNotFoundError):
            get_sheet_names("nonexistent.xlsx")

        with pytest.raises(FileNotFoundError):
            get_sheet_row_count("nonexistent.xlsx", "Sheet1")

        with pytest.raises(FileNotFoundError):
            process_excel_file("nonexistent.xlsx")
