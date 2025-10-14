#!/usr/bin/env python3
"""
Work Order Duplicate Checker

This program analyzes work order files to identify duplicate tasks across different work orders.
Prevents double work by detecting when the same task appears in multiple work orders.

Usage: python main.py [work_order_files...]
"""

import sys
import os
from pathlib import Path
from work_order_checker import WorkOrderChecker


def main():
    """Main entry point for the work order duplicate checker."""
    if len(sys.argv) < 2:
        print("Usage: python main.py <work_order_file1> <work_order_file2> [...]")
        print("   or: python main.py <directory_with_work_orders>")
        return 1
    
    # Initialize the checker
    checker = WorkOrderChecker()
    
    # Collect work order files
    files_to_check = []
    
    for arg in sys.argv[1:]:
        path = Path(arg)
        if path.is_file():
            files_to_check.append(path)
        elif path.is_dir():
            # Add all supported file types in the directory
            supported_extensions = ['*.txt', '*.csv', '*.json', '*.pdf', '*.html', '*.htm', '*.HTM',
                                  '*.xml', '*.xls', '*.xlsx', '*.doc', '*.docx']
            for pattern in supported_extensions:
                files_to_check.extend(path.glob(pattern))
        else:
            print(f"Warning: {arg} is not a valid file or directory")
    
    if not files_to_check:
        print("No valid work order files found.")
        return 1
    
    print(f"Checking {len(files_to_check)} work order files for duplicates...")
    
    # Load and analyze work orders
    for file_path in files_to_check:
        try:
            checker.load_work_order(file_path)
            print(f"✓ Loaded: {file_path.name}")
        except Exception as e:
            print(f"✗ Error loading {file_path.name}: {e}")
    
    # Find and report duplicates
    duplicates = checker.find_duplicates()
    
    if duplicates:
        print(f"\n🚨 Found {len(duplicates)} duplicate tasks:")
        print("=" * 80)
        
        for i, duplicate in enumerate(duplicates, 1):
            print(f"\nDuplicate #{i}:")
            print(f"Task: {duplicate['task']}")
            print(f"Found in work orders:")
            for work_order in duplicate['work_orders']:
                print(f"  - {work_order}")
            print("-" * 40)
    else:
        print("\n✅ No duplicate tasks found!")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())