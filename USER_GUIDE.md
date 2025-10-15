# Work Order Duplicate Checker - User Guide

## GUI Version Features

### Main Interface
- **File Management Area**: Add, remove, and manage work order files
- **Drag-and-Drop Support**: Simply drag files from Windows Explorer into the application
- **Results Display**: View detailed duplicate analysis with clear formatting
- **Export Functionality**: Save results to text files for reporting

### Supported File Operations
- **Add Files**: Select individual work order files
- **Add Folder**: Process entire directories of work orders
- **Remove Selected**: Remove specific files from the analysis list
- **Clear All**: Start fresh with a new set of files

### Analysis Features
- **Real-time Progress**: See loading progress as files are processed
- **Detailed Results**: View equipment IDs, locations, and affected work orders
- **Error Handling**: Clear error messages for problematic files
- **Statistics Summary**: Overview of processed files and duplication rates

## Command Line Usage

For advanced users or automation:

```bash
# Basic usage
python main.py file1.pdf file2.xlsx

# Process entire directory
python main.py "C:\Work Orders\October 2025\"

# Multiple directories
python main.py "C:\Building A\" "C:\Building B\"
```

## File Format Support

### Fully Supported
- **Text Files** (.txt): Plain text work orders
- **CSV Files** (.csv): Comma-separated data
- **Excel Files** (.xls, .xlsx): Microsoft Excel spreadsheets
- **HTML Files** (.html, .htm, .HTM): Web pages and Maintenance Connection exports
- **JSON Files** (.json): Structured data files
- **XML Files** (.xml): Structured markup data

### Experimental Support
- **PDF Files** (.pdf): Text-based PDFs (not scanned images)
- **Word Documents** (.docx): Microsoft Word documents (not .doc)

## Tips for Best Results

1. **File Organization**: Keep work orders in separate folders by date or building
2. **Naming Convention**: Use consistent file naming for easier identification
3. **Regular Checks**: Run duplicate checks before finalizing work schedules
4. **Export Results**: Save analysis results for record-keeping and reporting

## Troubleshooting

### Common Issues
- **File Not Loading**: Check if the file format is supported
- **No Duplicates Found**: Verify that equipment IDs are in [bracket] format
- **Application Won't Start**: Ensure all dependencies are installed

### Getting Help
- Check the GitHub repository: https://github.com/built-by-Beck/work-order-checker
- Review sample data files for format examples
- Ensure Python 3.8+ is installed for best compatibility