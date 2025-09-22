# XLSX Reader

A simple Python application with tkinter GUI for reading Excel files and counting rows in each sheet.

## Features

- **Simple GUI**: Basic tkinter interface with file selection
- **Progress Bar**: Shows progress while processing sheets
- **Row Counting**: Counts rows in each sheet of the Excel file
- **Function-based**: Uses simple functions (no OOP)

## Installation

```bash
# Install dependencies
poetry install
```

## Usage

```bash
# Run the application
poetry run python -m xlsx_reader.main
```

## Project Structure

```
xlsx_reader/
├── __init__.py
├── main.py              # Main entry point
├── gui.py               # Simple tkinter GUI
└── excel_processor.py   # Excel processing functions

tests/
├── __init__.py
└── test_excel_processor.py  # Tests for functions
```

## Student Assignment

Students need to implement the following functions (marked with `pass`):

### excel_processor.py:
- `get_sheet_names()` - Get all sheet names from Excel file
- `get_sheet_row_count()` - Count rows in a specific sheet

### gui.py:
- `select_excel_file()` - Open file dialog for Excel file selection
- `update_progress()` - Update progress bar

## Dependencies

- **pandas**: For reading Excel files
- **tkinter**: For GUI (included with Python)
- **pytest**: For testing (dev dependency)
- **xlsxwriter**: For creating test Excel files (dev dependency)

## Running Tests

```bash
poetry run pytest
```

Tests will create temporary Excel files automatically for testing your implementations.
