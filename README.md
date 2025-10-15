# Work Order Duplicate Checker

A Python application that analyzes work order files to identify duplicate work orders by detecting the same equipment at the same location across different work orders, preventing double work in maintenance operations. Available in both **command-line** and **user-friendly GUI** versions.

## Features

- **Multi-format Support**: Handles TXT, CSV, JSON, PDF, HTML, XML, XLS, XLSX, DOC, and DOCX work order files
- **Smart Task Detection**: Recognizes task patterns with equipment/part IDs in brackets (e.g., `[212934]`)
- **Location-Specific Detection**: Identifies duplicates based on equipment ID and exact location when the same location and part number appear in multiple work orders, indicating potential duplicate work assignments
- **Detailed Reporting**: Provides comprehensive duplicate analysis reports
- **Batch Processing**: Can process entire directories of work order files
- **Robust Parsing**: Extracts tasks from various document structures and formats
- **GUI Interface**: Easy-to-use graphical interface with drag-and-drop support
- **Cross-Platform**: Works on Windows, macOS, and Linux

## Installation

### Windows Users (Easy Installation)

1. **Download** or clone this repository
2. **Run** `install_windows.bat` - This will automatically install Python dependencies
3. **Launch** the GUI: `python gui.py` or double-click `launcher.py`

### Manual Installation (All Platforms)

1. **Install Python 3.8+** if not already installed
2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the application:**
   ```bash
   # GUI Version (recommended)
   python gui.py
   
   # Command Line Version
   python main.py [files or directory]
   ```

### Building Windows Executable

To create a standalone Windows executable that doesn't require Python installation:

```bash
python build_windows.py
```

This creates `WorkOrderChecker.exe` in the `dist` folder.

## Usage

### GUI Version (Recommended)

1. **Launch the application:**
   ```bash
   python gui.py
   ```

2. **Add work order files:**
   - Click "Add Files" to select individual files
   - Click "Add Folder" to add all supported files from a directory
   - **Drag and drop** files directly into the file list

3. **Check for duplicates:**
   - Click "Check for Duplicates"
   - View results in the results panel
   - Export results to a text file if needed

4. **Supported operations:**
   - Remove selected files from the list
   - Clear all files
   - View detailed duplicate information
   - Export results for reporting

### Command Line Version

**Check specific work order files:**
```bash
python main.py workorder1.pdf workorder2.xlsx workorder3.html
```

**Check all work orders in a directory:**
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

🚨 Found 1 duplicate tasks:
================================================================================

Duplicate #1:
Task: Exit Light [212934] MOB B - ground floor - HRC back door
Found in work orders:
  - exit_lights_general
  - mob_b_exit_lights

<<<<<<< HEAD
=======
Duplicate #2:
Task: Emergency Light [445566] MOB B - 2nd floor - main corridor
Found in work orders:
  - exit_lights_general
  - building_maintenance
>>>>>>> 4e04e64 (Add GUI version and Windows installer)
----------------------------------------
```

## Supported File Formats

### Text Files (.txt)
Plain text files with tasks listed line by line. The program automatically detects patterns like:
- `emergency lighting [445566] MOB B - 2nd floor - main corridor`
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

1. **File Parsing**: The program reads work order files and extracts task information including equipment/part numbers and locations
2. **Task Normalization**: Tasks are normalized (whitespace, case) for accurate comparison
3. **Duplicate Detection**: Work orders are compared by matching exact locations AND part numbers (equipment IDs in brackets). When the same part number is found at the exact same location across multiple work orders, this indicates potential duplicate work assignments.
4. **Reporting**: Results are displayed with clear identification of work orders that may be duplicates based on matching locations and part numbers

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
