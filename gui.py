#!/usr/bin/env python3
"""
Work Order Duplicate Checker - GUI Version

A user-friendly graphical interface for checking duplicate tasks across work order files.
Supports drag-and-drop functionality and multiple file formats.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from pathlib import Path
import sys
import os
from typing import List
import webbrowser

# Add the current directory to the path so we can import our module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from work_order_checker import WorkOrderChecker

# Try to import tkinterdnd2 for drag-and-drop (optional)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False


class WorkOrderGUI:
    def __init__(self):
        # Use TkinterDnD if available, otherwise regular tkinter
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.setup_ui()
        self.checker = WorkOrderChecker()
        self.files_to_check = []
        
    def setup_ui(self):
        """Set up the user interface."""
        self.root.title("Work Order Duplicate Checker")
        self.root.geometry("800x600")
        self.root.minsize(600, 500)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Work Order Duplicate Checker", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection area
        files_frame = ttk.LabelFrame(main_frame, text="Work Order Files", padding="10")
        files_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(1, weight=1)
        
        # Instructions
        instructions = ("Select work order files to check for duplicates.\n"
                       "Supported formats: TXT, CSV, JSON, PDF, HTML, XML, XLS, XLSX, DOC, DOCX")
        if HAS_DND:
            instructions += "\nYou can drag and drop files here!"
        
        instruction_label = ttk.Label(files_frame, text=instructions, 
                                    font=('Arial', 9), foreground='gray')
        instruction_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # File list
        self.file_listbox = tk.Listbox(files_frame, height=6, selectmode=tk.EXTENDED)
        self.file_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Scrollbar for file list
        file_scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        file_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.file_listbox.configure(yscrollcommand=file_scrollbar.set)
        
        # Buttons frame
        buttons_frame = ttk.Frame(files_frame)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=(10, 0))
        
        # Add files button
        self.add_files_btn = ttk.Button(buttons_frame, text="Add Files", command=self.add_files)
        self.add_files_btn.grid(row=0, column=0, padx=(0, 5))
        
        # Add folder button
        self.add_folder_btn = ttk.Button(buttons_frame, text="Add Folder", command=self.add_folder)
        self.add_folder_btn.grid(row=0, column=1, padx=5)
        
        # Remove files button
        self.remove_files_btn = ttk.Button(buttons_frame, text="Remove Selected", command=self.remove_files)
        self.remove_files_btn.grid(row=0, column=2, padx=5)
        
        # Clear all button
        self.clear_files_btn = ttk.Button(buttons_frame, text="Clear All", command=self.clear_files)
        self.clear_files_btn.grid(row=0, column=3, padx=5)
        
        # Check duplicates button
        self.check_btn = ttk.Button(buttons_frame, text="Check for Duplicates", 
                                   command=self.check_duplicates, style='Accent.TButton')
        self.check_btn.grid(row=0, column=4, padx=(20, 0))
        
        # Results area
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="10")
        results_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Results text area
        self.results_text = scrolledtext.ScrolledText(results_frame, height=15, 
                                                     font=('Courier New', 10))
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Export button
        export_frame = ttk.Frame(results_frame)
        export_frame.grid(row=1, column=0, pady=(10, 0), sticky=tk.W)
        
        self.export_btn = ttk.Button(export_frame, text="Export Results", 
                                    command=self.export_results, state=tk.DISABLED)
        self.export_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Add work order files to check for duplicates")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Configure drag and drop if available
        if HAS_DND:
            self.file_listbox.drop_target_register(DND_FILES)
            self.file_listbox.dnd_bind('<<Drop>>', self.drop_files)
        
        # Menu bar
        self.create_menu()
        
    def create_menu(self):
        """Create the menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Add Files...", command=self.add_files, accelerator="Ctrl+O")
        file_menu.add_command(label="Add Folder...", command=self.add_folder, accelerator="Ctrl+Shift+O")
        file_menu.add_separator()
        file_menu.add_command(label="Export Results...", command=self.export_results, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit, accelerator="Ctrl+Q")
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Documentation", command=self.show_documentation)
        
        # Keyboard shortcuts
        self.root.bind('<Control-o>', lambda e: self.add_files())
        self.root.bind('<Control-O>', lambda e: self.add_folder())
        self.root.bind('<Control-s>', lambda e: self.export_results())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        
    def add_files(self):
        """Add individual files to the check list."""
        filetypes = [
            ("All Supported", "*.txt;*.csv;*.json;*.pdf;*.html;*.htm;*.HTM;*.xml;*.xls;*.xlsx;*.doc;*.docx"),
            ("Text files", "*.txt"),
            ("CSV files", "*.csv"),
            ("JSON files", "*.json"),
            ("PDF files", "*.pdf"),
            ("HTML files", "*.html;*.htm;*.HTM"),
            ("XML files", "*.xml"),
            ("Excel files", "*.xls;*.xlsx"),
            ("Word files", "*.doc;*.docx"),
            ("All files", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="Select Work Order Files",
            filetypes=filetypes
        )
        
        for file in files:
            if file not in self.files_to_check:
                self.files_to_check.append(file)
                self.file_listbox.insert(tk.END, Path(file).name)
        
        self.update_status(f"Added {len(files)} file(s). Total: {len(self.files_to_check)} files")
        
    def add_folder(self):
        """Add all supported files from a folder."""
        folder = filedialog.askdirectory(title="Select Folder with Work Order Files")
        if not folder:
            return
        
        folder_path = Path(folder)
        supported_extensions = ['.txt', '.csv', '.json', '.pdf', '.html', '.htm', '.HTM',
                              '.xml', '.xls', '.xlsx', '.doc', '.docx']
        
        added_count = 0
        for file_path in folder_path.iterdir():
            if file_path.is_file() and file_path.suffix in supported_extensions:
                file_str = str(file_path)
                if file_str not in self.files_to_check:
                    self.files_to_check.append(file_str)
                    self.file_listbox.insert(tk.END, file_path.name)
                    added_count += 1
        
        self.update_status(f"Added {added_count} file(s) from folder. Total: {len(self.files_to_check)} files")
        
    def remove_files(self):
        """Remove selected files from the list."""
        selected_indices = list(self.file_listbox.curselection())
        selected_indices.reverse()  # Remove from end to beginning to maintain indices
        
        for index in selected_indices:
            del self.files_to_check[index]
            self.file_listbox.delete(index)
        
        self.update_status(f"Removed {len(selected_indices)} file(s). Total: {len(self.files_to_check)} files")
        
    def clear_files(self):
        """Clear all files from the list."""
        self.files_to_check.clear()
        self.file_listbox.delete(0, tk.END)
        self.results_text.delete(1.0, tk.END)
        self.export_btn.config(state=tk.DISABLED)
        self.update_status("All files cleared")
        
    def drop_files(self, event):
        """Handle drag and drop files."""
        if not HAS_DND:
            return
        
        files = self.root.tk.splitlist(event.data)
        added_count = 0
        
        for file in files:
            file_path = Path(file)
            if file_path.is_file():
                file_str = str(file_path)
                if file_str not in self.files_to_check:
                    self.files_to_check.append(file_str)
                    self.file_listbox.insert(tk.END, file_path.name)
                    added_count += 1
        
        self.update_status(f"Dropped {added_count} file(s). Total: {len(self.files_to_check)} files")
        
    def check_duplicates(self):
        """Check for duplicates in the selected files."""
        if not self.files_to_check:
            messagebox.showwarning("No Files", "Please add some work order files to check.")
            return
        
        # Disable the check button and start progress
        self.check_btn.config(state=tk.DISABLED)
        self.progress.start()
        self.update_status("Checking for duplicates...")
        self.results_text.delete(1.0, tk.END)
        
        # Run the check in a separate thread to avoid freezing the UI
        thread = threading.Thread(target=self._check_duplicates_thread)
        thread.daemon = True
        thread.start()
        
    def _check_duplicates_thread(self):
        """Thread function to check duplicates without blocking the UI."""
        try:
            # Reset the checker
            self.checker = WorkOrderChecker()
            
            # Load all files
            loaded_count = 0
            failed_files = []
            
            for file_path in self.files_to_check:
                try:
                    self.checker.load_work_order(Path(file_path))
                    loaded_count += 1
                    # Update UI in main thread
                    self.root.after(0, self.update_status, 
                                  f"Loaded {loaded_count}/{len(self.files_to_check)} files...")
                except Exception as e:
                    failed_files.append((Path(file_path).name, str(e)))
            
            # Find duplicates
            duplicates = self.checker.find_duplicates()
            stats = self.checker.get_statistics()
            
            # Update UI in main thread
            self.root.after(0, self._display_results, duplicates, stats, loaded_count, failed_files)
            
        except Exception as e:
            self.root.after(0, self._display_error, str(e))
        finally:
            self.root.after(0, self._finish_check)
            
    def _display_results(self, duplicates, stats, loaded_count, failed_files):
        """Display the results in the UI."""
        results = []
        
        # Header
        results.append("=" * 60)
        results.append("WORK ORDER DUPLICATE CHECK RESULTS")
        results.append("=" * 60)
        results.append("")
        
        # Summary
        results.append("SUMMARY:")
        results.append(f"  Files processed: {loaded_count}")
        results.append(f"  Total work orders: {stats['work_orders_count']}")
        results.append(f"  Total tasks: {stats['total_tasks']}")
        results.append(f"  Unique tasks: {stats['unique_tasks']}")
        results.append(f"  Duplicate tasks found: {len(duplicates)}")
        if stats['unique_tasks'] > 0:
            results.append(f"  Duplication rate: {stats['duplication_rate']:.1f}%")
        results.append("")
        
        # Failed files
        if failed_files:
            results.append("FAILED TO LOAD:")
            for filename, error in failed_files:
                results.append(f"  ✗ {filename}: {error}")
            results.append("")
        
        # Duplicates
        if duplicates:
            results.append(f"🚨 DUPLICATE TASKS FOUND ({len(duplicates)}):")
            results.append("-" * 60)
            
            for i, duplicate in enumerate(duplicates, 1):
                results.append(f"\nDuplicate #{i}:")
                results.append(f"  Task: {duplicate['task']}")
                results.append(f"  Equipment ID: {duplicate['task_id']}")
                results.append(f"  Found in {duplicate['count']} work orders:")
                for wo_name in duplicate['work_orders']:
                    results.append(f"    - {wo_name}")
                results.append("-" * 40)
        else:
            results.append("✅ NO DUPLICATE TASKS FOUND!")
            results.append("All work orders contain unique tasks.")
        
        # Display results
        result_text = "\n".join(results)
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(1.0, result_text)
        
        # Enable export if we have results
        self.export_btn.config(state=tk.NORMAL if duplicates else tk.DISABLED)
        
        # Update status
        if duplicates:
            self.update_status(f"Check complete - Found {len(duplicates)} duplicate tasks")
        else:
            self.update_status("Check complete - No duplicates found")
            
    def _display_error(self, error_message):
        """Display an error message."""
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(1.0, f"ERROR: {error_message}")
        self.update_status("Check failed - See error message")
        messagebox.showerror("Error", f"An error occurred during duplicate checking:\n\n{error_message}")
        
    def _finish_check(self):
        """Finish the duplicate check process."""
        self.progress.stop()
        self.check_btn.config(state=tk.NORMAL)
        
    def export_results(self):
        """Export the results to a file."""
        if not self.results_text.get(1.0, tk.END).strip():
            messagebox.showwarning("No Results", "No results to export. Please run a duplicate check first.")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Results",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.results_text.get(1.0, tk.END))
                self.update_status(f"Results exported to {Path(filename).name}")
                messagebox.showinfo("Export Complete", f"Results exported successfully to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export results:\n{e}")
                
    def update_status(self, message):
        """Update the status bar."""
        self.status_var.set(message)
        
    def show_about(self):
        """Show about dialog."""
        about_text = """Work Order Duplicate Checker v1.0

A tool for identifying duplicate maintenance tasks across work order files.

Features:
• Multi-format support (TXT, CSV, JSON, PDF, HTML, XML, Excel, Word)
• Location-specific duplicate detection
• Equipment ID matching
• Drag-and-drop file support
• Batch processing

Built for maintenance professionals to prevent duplicate work and improve efficiency.

GitHub: https://github.com/built-by-Beck/work-order-checker"""
        
        messagebox.showinfo("About", about_text)
        
    def show_documentation(self):
        """Open documentation in web browser."""
        url = "https://github.com/built-by-Beck/work-order-checker"
        webbrowser.open(url)
        
    def run(self):
        """Start the GUI application."""
        self.root.mainloop()


def main():
    """Main entry point for the GUI application."""
    app = WorkOrderGUI()
    app.run()


if __name__ == "__main__":
    main()