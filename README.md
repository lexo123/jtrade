# Excel Template Generator

A Python application that reads an Excel template file and generates new Excel files with specified cell changes while preserving all other content.

## Features

- **Read Template**: Loads `template.xls` without modifications
- **Apply Changes**: Specify which cells to modify and their new values
- **Preserve Format**: Keeps all formatting, formulas, and structure intact
- **Single & Batch**: Generate one file or multiple files at once
- **Easy to Use**: Simple Python API and interactive CLI

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Option 1: Interactive CLI

```bash
python cli.py
```

Follow the prompts to:
1. Choose to generate a single file or multiple files
2. Enter the output filename
3. Specify cell changes (e.g., `A1=New Title` or `B2=100`)

### Option 2: Python Script

```python
from excel_generator import ExcelTemplateGenerator

# Initialize generator
generator = ExcelTemplateGenerator("template.xls")

# Single file generation
changes = {
    'A1': 'New Title',
    'B2': 'Updated Value',
    'C3': 100
}
generator.generate("output.xlsx", changes)

# Multiple files
changes_list = [
    ('file1.xlsx', {'A1': 'Value1', 'B2': 123}),
    ('file2.xlsx', {'A1': 'Value2', 'B2': 456}),
]
generator.generate_multiple("output_folder", changes_list)
```

## Cell Reference Format

- **Cell Address**: Use Excel standard format (e.g., `A1`, `B2`, `Z100`)
- **Case Insensitive**: Both `a1` and `A1` work
- **Value Types**: Automatically detects strings, integers, and floats

## Example

Input template with:
```
| Header    | Value  |
|-----------|--------|
| Name      | ...    |
| Amount    | ...    |
```

Generate output:
```python
changes = {'A2': 'John Doe', 'B2': 5000}
```

Result: Excel file with name and amount updated, everything else unchanged.

## Notes

- Original template file is never modified
- All cell formatting, colors, and formulas are preserved
- Output files are created in Excel 2007+ format (.xlsx)
