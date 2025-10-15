#!/usr/bin/env python3
"""
Desktop launcher for Work Order Duplicate Checker GUI
This script can be used to create desktop shortcuts
"""

import sys
import os
from pathlib import Path

# Add the script directory to Python path
script_dir = Path(__file__).parent
sys.path.insert(0, str(script_dir))

try:
    from gui import main
    main()
except ImportError as e:
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    messagebox.showerror(
        "Missing Dependencies",
        f"Required modules are not installed.\n\n"
        f"Error: {e}\n\n"
        f"Please run 'install_windows.bat' first to install dependencies."
    )
    
    sys.exit(1)
except Exception as e:
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    messagebox.showerror(
        "Application Error",
        f"An error occurred while starting the application:\n\n{e}"
    )
    
    sys.exit(1)