# Work Order Duplicate Checker

A Python application that analyzes work order files to identify duplicate tasks across different work orders, preventing double work in maintenance operations.

## Features

- **Multi-format Support**: Handles TXT, CSV, JSON, PDF, HTML, XML, XLS, XLSX, DOC, and DOCX work order files
- **Smart Task Detection**: Recognizes task patterns with IDs in brackets (e.g., `[212934]`)
- **Duplicate Detection**: Identifies when the same task appears in multiple work orders
- **Detailed Reporting**: Provides comprehensive duplicate analysis reports
- **Batch Processing**: Can process entire directories of work order files
- **Robust Parsing**: Extracts tasks from various document structures and formats

## Installation

1. Clone or download this repository
2. Install Python 3.6+ if not already installed
3. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Quick Start

1. **Test with sample data:**
   ```bash
   python main.py sample_data/
   ```

2. **Use with your own files:**
   ```bash
   python main.py /path/to/your/workorders/
   ```

## VS Code Integration

This project includes VS Code tasks for easy execution:

1. **Run with sample data:** Press `Ctrl+Shift+P` → "Tasks: Run Task" → "Run Work Order Checker"
2. **Run with custom files:** Press `Ctrl+Shift+P` → "Tasks: Run Task" → "Run Work Order Checker - Custom Files"
3. **Install dependencies:** Press `Ctrl+Shift+P` → "Tasks: Run Task" → "Install Dependencies"

## Usage

### Basic Usage

Check specific work order files:
```bash
python main.py workorder1.pdf workorder2.xlsx workorder3.html
```

Check all work orders in a directory:
```bash
python main.py /path/to/workorders/
```

### Example Output

```
Checking 5 work order files for duplicates...
✓ Loaded: exit_lights_general.txt
✓ Loaded: mob_b_exit_lights.txt
✓ Loaded: building_maintenance.txt
✓ Loaded: maintenance_html.html
✓ Loaded: emergency_systems.xml

🚨 Found 2 duplicate tasks:
================================================================================

Duplicate #1:
Task: exit lights [212934] MOB B - ground floor - HRC back door
Found in work orders:
  - exit_lights_general
  - mob_b_exit_lights
  - maintenance_html
  - emergency_systems
----------------------------------------

Duplicate #2:
Task: emergency lighting [445566] MOB B - 2nd floor - main corridor
Found in work orders:
  - exit_lights_general
  - building_maintenance
----------------------------------------
```

## Supported File Formats

### Text Files (.txt)
Plain text files with tasks listed line by line. The program automatically detects patterns like:
- `exit lights [212934] MOB B - ground floor - HRC back door`
- `emergency lighting [445566] MOB B - 2nd floor - main corridor`

### CSV Files (.csv)
Comma-separated files with task information in columns. The program handles various CSV formats automatically.

### Excel Files (.xls, .xlsx)
Excel spreadsheets with task data in rows and columns. Supports both legacy (.xls) and modern (.xlsx) formats.

### PDF Files (.pdf)
Extracts text from PDF documents and parses task information. Works with most text-based PDFs.

### HTML Files (.html, .htm)
Parses HTML documents and extracts task information from various HTML elements (paragraphs, divs, list items, table cells).

### XML Files (.xml)
Processes XML documents and extracts task information from XML elements and their text content.

### Word Documents (.docx)
Reads Microsoft Word documents and extracts tasks from paragraphs and tables. Note: Legacy .doc files are not supported - please convert to .docx format.

### JSON Files (.json)
JSON files with task objects. Supports multiple structures:
```json
{
  "tasks": [
    {
      "id": "212934",
      "description": "exit lights",
      "location": "MOB B - ground floor - HRC back door"
    }
  ]
}
```

## How It Works

1. **File Parsing**: The program reads work order files and extracts task information
2. **Task Normalization**: Tasks are normalized (whitespace, case) for accurate comparison
3. **Duplicate Detection**: Tasks are compared across all work orders to find duplicates
4. **Reporting**: Results are displayed with clear identification of duplicate tasks and their locations

## Example Work Order Files

Create test files to try the program:

**exit_lights_general.txt:**
```
exit lights [212934] MOB B - ground floor - HRC back door
exit lights [212935] MOB B - 1st floor - main entrance
emergency lighting [445566] MOB B - 2nd floor - main corridor
```

**mob_b_exit_lights.txt:**
```
exit lights [212934] MOB B - ground floor - HRC back door
exit lights [778899] MOB B - basement - utility room
```

Run: `python main.py exit_lights_general.txt mob_b_exit_lights.txt`

## Contributing

Feel free to submit issues or pull requests to improve the duplicate detection algorithms or add support for additional file formats.

## License

This project is provided as-is for maintenance management purposes.