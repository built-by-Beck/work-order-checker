<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# Work Order Duplicate Checker - Copilot Instructions

## Project Overview
This is a Python application that analyzes work order files to identify duplicate tasks across different work orders, preventing double work in maintenance operations.

## Key Features
- Multi-format support: TXT, CSV, JSON, PDF, HTML, XML, XLS, XLSX, DOC, DOCX
- Smart task detection with ID patterns in brackets (e.g., `[212934]`)
- Duplicate detection across multiple work orders
- Batch processing of entire directories

## Development Guidelines
- Use Python 3.6+ with virtual environment located in `.venv/`
- Main entry point: `main.py`
- Core logic: `work_order_checker.py`
- Sample data available in `sample_data/` directory
- Dependencies listed in `requirements.txt`

## Testing
- Run with sample data: `python main.py sample_data/`
- VS Code tasks available for easy execution
- Test files include various formats demonstrating duplicate detection

## Code Style
- Follow PEP 8 Python style guidelines
- Use type hints where appropriate
- Include docstrings for all functions and classes
- Handle file format errors gracefully with informative messages