# Work Order Duplicate Checker - Windows Build Script
# This script creates a standalone Windows executable using PyInstaller

import os
import sys
import subprocess
import shutil
from pathlib import Path

def install_requirements():
    """Install required packages for building."""
    print("Installing build requirements...")
    requirements = [
        "pyinstaller",
        "pandas",
        "openpyxl", 
        "xlrd",
        "PyPDF2",
        "beautifulsoup4",
        "lxml",
        "python-docx",
        "tkinterdnd2"
    ]
    
    for req in requirements:
        print(f"Installing {req}...")
        subprocess.run([sys.executable, "-m", "pip", "install", req], check=True)

def build_executable():
    """Build the Windows executable."""
    print("Building Windows executable...")
    
    # PyInstaller command for GUI version
    gui_cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name", "WorkOrderChecker",
        "--icon", "icon.ico" if Path("icon.ico").exists() else None,
        "--add-data", "sample_data;sample_data",
        "gui.py"
    ]
    
    # Remove None values (in case icon doesn't exist)
    gui_cmd = [arg for arg in gui_cmd if arg is not None]
    
    # PyInstaller command for console version
    console_cmd = [
        "pyinstaller",
        "--onefile",
        "--name", "WorkOrderChecker-Console",
        "--add-data", "sample_data;sample_data",
        "main.py"
    ]
    
    # Build GUI version
    print("Building GUI version...")
    subprocess.run(gui_cmd, check=True)
    
    # Build console version
    print("Building console version...")
    subprocess.run(console_cmd, check=True)
    
    print("Build completed!")
    print("Executables are in the 'dist' folder:")
    print("  - WorkOrderChecker.exe (GUI version)")
    print("  - WorkOrderChecker-Console.exe (Command line version)")

def create_installer_package():
    """Create a complete installer package."""
    print("Creating installer package...")
    
    # Create distribution folder
    dist_folder = Path("dist")
    package_folder = dist_folder / "WorkOrderChecker-Package"
    package_folder.mkdir(exist_ok=True)
    
    # Copy executables
    if (dist_folder / "WorkOrderChecker.exe").exists():
        shutil.copy2(dist_folder / "WorkOrderChecker.exe", package_folder)
    
    if (dist_folder / "WorkOrderChecker-Console.exe").exists():
        shutil.copy2(dist_folder / "WorkOrderChecker-Console.exe", package_folder)
    
    # Copy documentation
    for file in ["README.md", "requirements.txt"]:
        if Path(file).exists():
            shutil.copy2(file, package_folder)
    
    # Copy sample data
    if Path("sample_data").exists():
        shutil.copytree("sample_data", package_folder / "sample_data", dirs_exist_ok=True)
    
    # Create installation instructions
    install_instructions = """# Work Order Duplicate Checker - Installation Instructions

## Quick Start

1. **GUI Version (Recommended)**:
   - Double-click `WorkOrderChecker.exe`
   - Drag and drop your work order files or use the "Add Files" button
   - Click "Check for Duplicates"

2. **Command Line Version**:
   - Open Command Prompt in this folder
   - Run: `WorkOrderChecker-Console.exe path\\to\\your\\workorders\\`

## Features

- Supports multiple file formats: TXT, CSV, JSON, PDF, HTML, XML, Excel, Word
- Drag-and-drop file support (GUI version)
- Location-specific duplicate detection
- Equipment ID matching
- Export results to text files

## Sample Data

The `sample_data` folder contains example work order files you can use to test the application.

## Support

For issues or questions, visit: https://github.com/built-by-Beck/work-order-checker
"""
    
    with open(package_folder / "INSTALLATION.md", "w", encoding="utf-8") as f:
        f.write(install_instructions)
    
    print(f"Package created in: {package_folder}")
    print("Ready for distribution!")

def main():
    """Main build process."""
    print("=" * 60)
    print("Work Order Duplicate Checker - Windows Build Script")
    print("=" * 60)
    
    try:
        # Step 1: Install requirements
        install_requirements()
        
        # Step 2: Build executable
        build_executable()
        
        # Step 3: Create installer package
        create_installer_package()
        
        print("\n" + "=" * 60)
        print("BUILD SUCCESSFUL!")
        print("Your Windows executable is ready in the 'dist' folder.")
        print("=" * 60)
        
    except subprocess.CalledProcessError as e:
        print(f"Error during build: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()