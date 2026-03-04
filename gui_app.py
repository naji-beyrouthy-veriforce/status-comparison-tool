"""
GUI Interface for Dynamics 365 vs SafeContractor Status Comparison
Drag-and-drop file upload interface with manual workflow - Dark Mode Edition
"""

import tkinter as tk
from tkinter import messagebox, scrolledtext
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import ToolTip
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import logging
from pathlib import Path
import shutil
import sys
import warnings

# Suppress openpyxl style warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Import configuration and utilities
from config import INPUT_DIR, OUTPUT_DIR, DYNAMICS_DIR, REDASH_DIR, QUERY_IDS_DIR, D365_FILES, SC_FILES, Messages, setup_logging, get_dated_comparison_dir

# Import main processing functions
from main import extract_and_save_ids, generate_comparisons

# Setup logging for GUI
logger = setup_logging("comparison_tool_gui", console_output=False, file_output=True)


class ComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("D365 vs SafeContractor - Status Comparison Tool")
        self.root.geometry("1100x850")
        self.root.resizable(True, True)
        
        logger.info("GUI Application started")
        
        # Dark theme colors
        self.colors = {
            'bg_dark': '#1e1e1e',
            'bg_card': '#2d2d2d',
            'accent_blue': '#3b82f6',
            'accent_green': '#10b981',
            'accent_orange': '#f59e0b',
            'accent_purple': '#8b5cf6',
            'text_primary': '#e5e5e5',
            'text_secondary': '#9ca3af',
            'border': '#404040',
            'hover': '#374151'
        }

        # File storage
        self.uploaded_files = {
            "accreditation_d365": None,
            "wcb_d365": None,
            "client_d365": None,
            "accreditation_sc": None,
            "wcb_sc": None,
            "client_sc": None,
        }

        self.setup_ui()
        # Check for existing files after UI is fully initialized
        self.root.after(100, self.check_existing_files)
        
        # Bind window close event to cleanup
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        """Create the user interface"""
        # Modern header with gradient effect
        header_frame = tk.Frame(self.root, bg=self.colors['accent_blue'], height=90)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        # Title with status indicator
        title_container = tk.Frame(header_frame, bg=self.colors['accent_blue'])
        title_container.pack(expand=True)
        
        title_label = tk.Label(
            title_container,
            text="📊 D365 vs SafeContractor",
            font=("Segoe UI", 22, "bold"),
            fg="white",
            bg=self.colors['accent_blue'],
        )
        title_label.pack(pady=(10, 2))

        subtitle_label = tk.Label(
            title_container,
            text="Status Comparison & Reporting Tool",
            font=("Segoe UI", 11),
            fg="#e0f2fe",
            bg=self.colors['accent_blue'],
        )
        subtitle_label.pack()

        # Main container with tabs
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Create notebook with modern styling
        self.notebook = ttk.Notebook(container, bootstyle="dark")
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Upload D365 Files
        self.tab_d365 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_d365, text="  📁 1. Upload D365 Files  ")
        self.setup_d365_tab()

        # Tab 2: Extract IDs
        self.tab_extract = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_extract, text="  🔍 2. Extract IDs  ")
        self.setup_extract_tab()

        # Tab 3: Upload SafeContractor Files
        self.tab_sc = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_sc, text="  📊 3. Upload SC Files  ")
        self.setup_sc_tab()

        # Tab 4: Generate Comparisons
        self.tab_compare = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_compare, text="  🚀 4. Generate Reports  ")
        self.setup_compare_tab()

        # Modern status bar with indicator
        status_frame = tk.Frame(self.root, bg=self.colors['bg_card'], height=40)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        status_frame.pack_propagate(False)
        
        # Status indicator (green dot)
        self.status_indicator = tk.Canvas(status_frame, width=12, height=12, bg=self.colors['bg_card'], highlightthickness=0)
        self.status_indicator.pack(side=tk.LEFT, padx=(15, 8), pady=14)
        self.status_dot = self.status_indicator.create_oval(2, 2, 10, 10, fill=self.colors['accent_green'], outline="")
        
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Start by uploading D365 files")
        status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg=self.colors['bg_card'],
            fg=self.colors['text_primary'],
            font=("Segoe UI", 9),
            anchor=tk.W,
        )
        status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=10)

    def setup_d365_tab(self):
        """Setup D365 file upload tab"""
        # Info card
        info_frame = tk.Frame(self.tab_d365, bg=self.colors['bg_card'], relief=tk.SOLID, bd=1)
        info_frame.pack(fill=tk.X, padx=20, pady=20)

        tk.Label(
            info_frame,
            text="📁 Upload Dynamics 365 Exports",
            font=("Segoe UI", 14, "bold"),
            bg=self.colors['bg_card'],
            fg=self.colors['accent_blue'],
        ).pack(anchor=tk.W, padx=20, pady=(15, 5))

        instruction_text = "Drag & drop your D365 files (any combination of Accreditation, WCB, Client). Not all 3 are required."

        tk.Label(
            info_frame,
            text=instruction_text,
            bg=self.colors['bg_card'],
            fg=self.colors['text_secondary'],
            font=("Segoe UI", 10),
            wraplength=900,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, padx=20, pady=(0, 15))

        # Enhanced drop zone with border
        multi_drop_frame = ttk.Frame(self.tab_d365, padding=10)
        multi_drop_frame.pack(fill=tk.X, padx=20, pady=10)

        bulk_drop = tk.Frame(
            multi_drop_frame,
            bg="#2a3f5f",
            relief=tk.SOLID,
            bd=2,
            highlightbackground=self.colors['accent_blue'],
            highlightthickness=2,
            height=110,
        )
        bulk_drop.pack(fill=tk.X, pady=5)
        bulk_drop.pack_propagate(False)

        # Configure drag and drop
        bulk_drop.drop_target_register(DND_FILES)
        bulk_drop.dnd_bind("<<Drop>>", lambda e: self.handle_bulk_drop(e, "d365"))
        bulk_drop.dnd_bind("<<DragEnter>>", lambda e, f=bulk_drop: self.on_drag_enter(e, f, "d365"))
        bulk_drop.dnd_bind("<<DragLeave>>", lambda e, f=bulk_drop: self.on_drag_leave(e, f, "d365"))

        tk.Label(
            bulk_drop,
            text="🚀 DRAG & DROP D365 FILES HERE\n\nDrop one or more files — Accreditation, WCB, and/or Client",
            bg="#2a3f5f",
            fg="#93c5fd",
            font=("Segoe UI", 11, "bold"),
            justify=tk.CENTER,
        ).pack(expand=True)
        
        ToolTip(bulk_drop, text="Drop Excel files (.xlsx, .xls) containing D365 export data", bootstyle="info")

        # Status indicators for files
        status_frame = tk.Frame(self.tab_d365, bg=self.colors['bg_card'], bd=1, relief=tk.SOLID)
        status_frame.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(
            status_frame,
            text="File Upload Status:",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['bg_card'],
            fg=self.colors['text_primary']
        ).pack(anchor=tk.W, padx=15, pady=(10, 5))
        
        self.d365_status_labels = {}
        for report_type, display_name in [("accreditation", "Accreditation"), ("wcb", "WCB"), ("client", "Client Specific")]:
            row = tk.Frame(status_frame, bg=self.colors['bg_card'])
            row.pack(fill=tk.X, padx=15, pady=3)
            
            indicator = tk.Canvas(row, width=10, height=10, bg=self.colors['bg_card'], highlightthickness=0)
            indicator.pack(side=tk.LEFT, padx=(0, 8))
            indicator.create_oval(2, 2, 8, 8, fill="#6b7280", outline="")
            
            label = tk.Label(row, text=f"{display_name}: Not uploaded", bg=self.colors['bg_card'], fg=self.colors['text_secondary'], font=("Segoe UI", 9))
            label.pack(side=tk.LEFT)
            
            self.d365_status_labels[report_type] = {"indicator": indicator, "label": label}
        
        tk.Label(status_frame, text="", bg=self.colors['bg_card']).pack(pady=5)

        # Action button with modern styling
        btn_frame = ttk.Frame(self.tab_d365)
        btn_frame.pack(pady=25)

        self.btn_process_d365 = ttk.Button(
            btn_frame,
            text="✓ Save D365 Files & Continue",
            command=self.save_d365_files,
            bootstyle="primary",
            width=35,
            state=tk.DISABLED
        )
        self.btn_process_d365.pack()
        
        ToolTip(self.btn_process_d365, text="Save uploaded D365 files and proceed to ID extraction", bootstyle="primary")

    def setup_extract_tab(self):
        """Setup ID extraction tab"""
        # Info card
        info_frame = tk.Frame(self.tab_extract, bg=self.colors['bg_card'], relief=tk.SOLID, bd=1)
        info_frame.pack(fill=tk.X, padx=20, pady=20)

        tk.Label(
            info_frame,
            text="🔍 Extract IDs for Redash Queries",
            font=("Segoe UI", 14, "bold"),
            bg=self.colors['bg_card'],
            fg=self.colors['accent_orange'],
        ).pack(anchor=tk.W, padx=20, pady=(15, 5))

        tk.Label(
            info_frame,
            text="Extract and format Global Alcumus IDs from D365 files. The formatted IDs will be ready to copy into Redash queries.",
            bg=self.colors['bg_card'],
            fg=self.colors['text_secondary'],
            font=("Segoe UI", 10),
            wraplength=850,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, padx=20, pady=(5, 15))

        # Action button with progress bar
        btn_frame = ttk.Frame(self.tab_extract)
        btn_frame.pack(pady=20)

        self.btn_extract = ttk.Button(
            btn_frame,
            text="⚙️ Extract IDs",
            command=self.extract_ids,
            bootstyle="warning",
            width=35,
        )
        self.btn_extract.pack()
        
        ToolTip(self.btn_extract, text="Extract Global Alcumus IDs from uploaded D365 files", bootstyle="warning")
        
        # Progress bar (hidden by default)
        self.extract_progress = ttk.Progressbar(btn_frame, bootstyle="warning-striped", mode="indeterminate", length=300)
        # Don't pack it yet, will show when processing

        # Modern console output
        console_label = tk.Label(
            self.tab_extract, 
            text="Output:", 
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['bg_dark'],
            fg=self.colors['text_primary']
        )
        console_label.pack(anchor=tk.W, padx=20, pady=(20, 5))

        self.extract_console = scrolledtext.ScrolledText(
            self.tab_extract,
            height=20,
            bg="#0d1117",
            fg="#58a6ff",
            font=("Consolas", 9),
            wrap=tk.WORD,
            insertbackground="#58a6ff",
            selectbackground="#1f6feb",
            relief=tk.SOLID,
            bd=1,
        )
        self.extract_console.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))

    def setup_sc_tab(self):
        """Setup SafeContractor file upload tab"""
        # Info card
        info_frame = tk.Frame(self.tab_sc, bg=self.colors['bg_card'], relief=tk.SOLID, bd=1)
        info_frame.pack(fill=tk.X, padx=20, pady=20)

        tk.Label(
            info_frame,
            text="📊 Upload SafeContractor (Redash) Exports",
            font=("Segoe UI", 14, "bold"),
            bg=self.colors['bg_card'],
            fg=self.colors['accent_green'],
        ).pack(anchor=tk.W, padx=20, pady=(15, 5))

        tk.Label(
            info_frame,
            text="Drag & drop your SafeContractor files (any combination of Accreditation, WCB, Client). Not all 3 are required.",
            bg=self.colors['bg_card'],
            fg=self.colors['text_secondary'],
            font=("Segoe UI", 10),
            wraplength=900,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, padx=20, pady=(0, 15))

        # Enhanced drop zone
        multi_drop_frame = ttk.Frame(self.tab_sc, padding=10)
        multi_drop_frame.pack(fill=tk.X, padx=20, pady=10)

        bulk_drop = tk.Frame(
            multi_drop_frame, 
            bg="#1f3a2c", 
            relief=tk.SOLID, 
            bd=2,
            highlightbackground=self.colors['accent_green'],
            highlightthickness=2,
            height=110
        )
        bulk_drop.pack(fill=tk.X, pady=5)
        bulk_drop.pack_propagate(False)

        # Configure drag and drop
        bulk_drop.drop_target_register(DND_FILES)
        bulk_drop.dnd_bind("<<Drop>>", lambda e: self.handle_bulk_drop(e, "sc"))
        bulk_drop.dnd_bind("<<DragEnter>>", lambda e, f=bulk_drop: self.on_drag_enter(e, f, "sc"))
        bulk_drop.dnd_bind("<<DragLeave>>", lambda e, f=bulk_drop: self.on_drag_leave(e, f, "sc"))

        tk.Label(
            bulk_drop,
            text="🚀 DRAG & DROP SAFECONTRACTOR FILES HERE\n(Drop one or more files — Accreditation, WCB, and/or Client)",
            bg="#1f3a2c",
            fg="#6ee7b7",
            font=("Segoe UI", 11, "bold"),
            justify=tk.CENTER,
        ).pack(expand=True)
        
        ToolTip(bulk_drop, text="Drop Excel files from Redash query results", bootstyle="success")

        # Status indicators
        status_frame = tk.Frame(self.tab_sc, bg=self.colors['bg_card'], bd=1, relief=tk.SOLID)
        status_frame.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(
            status_frame,
            text="File Upload Status:",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['bg_card'],
            fg=self.colors['text_primary']
        ).pack(anchor=tk.W, padx=15, pady=(10, 5))
        
        self.sc_status_labels = {}
        for report_type, display_name in [("accreditation", "Accreditation"), ("wcb", "WCB"), ("client", "Client Specific")]:
            row = tk.Frame(status_frame, bg=self.colors['bg_card'])
            row.pack(fill=tk.X, padx=15, pady=3)
            
            indicator = tk.Canvas(row, width=10, height=10, bg=self.colors['bg_card'], highlightthickness=0)
            indicator.pack(side=tk.LEFT, padx=(0, 8))
            indicator.create_oval(2, 2, 8, 8, fill="#6b7280", outline="")
            
            label = tk.Label(row, text=f"{display_name}: Not uploaded", bg=self.colors['bg_card'], fg=self.colors['text_secondary'], font=("Segoe UI", 9))
            label.pack(side=tk.LEFT)
            
            self.sc_status_labels[report_type] = {"indicator": indicator, "label": label}
        
        tk.Label(status_frame, text="", bg=self.colors['bg_card']).pack(pady=5)

        # Action button
        btn_frame = ttk.Frame(self.tab_sc)
        btn_frame.pack(pady=25)

        self.btn_process_sc = ttk.Button(
            btn_frame,
            text="✅ Save SafeContractor Files & Proceed",
            command=self.save_sc_files,
            bootstyle="success",
            width=35,
            state=tk.DISABLED,
        )
        self.btn_process_sc.pack()
        
        ToolTip(self.btn_process_sc, text="Save uploaded SC files and proceed to comparison", bootstyle="success")

    def setup_compare_tab(self):
        """Setup comparison generation tab"""
        # Info card
        info_frame = tk.Frame(self.tab_compare, bg=self.colors['bg_card'], relief=tk.SOLID, bd=1)
        info_frame.pack(fill=tk.X, padx=20, pady=20)

        tk.Label(
            info_frame,
            text="📊 Generate Status Comparison Files",
            font=("Segoe UI", 14, "bold"),
            bg=self.colors['bg_card'],
            fg=self.colors['accent_purple'],
        ).pack(anchor=tk.W, padx=20, pady=(15, 5))

        tk.Label(
            info_frame,
            text="Click below to generate the final comparison Excel files with status matching and differences.",
            bg=self.colors['bg_card'],
            fg=self.colors['text_secondary'],
            font=("Segoe UI", 10),
            wraplength=850,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, padx=20, pady=(5, 15))

        # Action button with progress
        btn_frame = ttk.Frame(self.tab_compare)
        btn_frame.pack(pady=20)

        self.btn_compare = ttk.Button(
            btn_frame,
            text="🚀 Generate Comparisons",
            command=self.generate_comparison,
            bootstyle="info",
            width=35,
        )
        self.btn_compare.pack()
        
        ToolTip(self.btn_compare, text="Generate Excel comparison reports with status analysis", bootstyle="info")
        
        # Progress bar (hidden by default)
        self.compare_progress = ttk.Progressbar(btn_frame, bootstyle="info-striped", mode="indeterminate", length=300)

        # Header with copy button
        output_header_frame = ttk.Frame(self.tab_compare)
        output_header_frame.pack(fill=tk.X, padx=20, pady=(20, 5))
        
        output_label = tk.Label(
            output_header_frame, 
            text="📊 Output:", 
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['bg_dark'],
            fg=self.colors['text_primary']
        )
        output_label.pack(side=tk.LEFT)
        
        self.btn_copy_report = ttk.Button(
            output_header_frame,
            text="📋 Copy Email Report",
            command=self.copy_email_to_clipboard,
            bootstyle="success",
            width=20,
            state="disabled"
        )
        self.btn_copy_report.pack(side=tk.RIGHT)
        
        ToolTip(self.btn_copy_report, text="Copy email report to clipboard", bootstyle="success")

        # Unified output console
        self.unified_output = scrolledtext.ScrolledText(
            self.tab_compare,
            height=28,
            bg="#0d1117",
            fg="#58a6ff",
            font=("Consolas", 9),
            wrap=tk.WORD,
            insertbackground="#58a6ff",
            selectbackground="#1f6feb",
            relief=tk.SOLID,
            bd=1,
        )
        self.unified_output.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
        
        # Configure text tags for different sections
        self.unified_output.tag_config("header", foreground="#10b981", font=("Consolas", 10, "bold"))
        self.unified_output.tag_config("separator", foreground="#60a5fa")
        self.unified_output.tag_config("email", foreground="#e5e5e5", font=("Consolas", 9))
        self.unified_output.tag_config("success", foreground="#10b981")
        self.unified_output.tag_config("error", foreground="#ef4444")
        
        # Add placeholder text
        placeholder = "Waiting to generate comparisons...\n\n" \
                     "Click 'Generate Comparisons' to start:\n" \
                     "  • Generates Excel comparison files\n" \
                     "  • Creates comparison.zip archive\n" \
                     "  • Displays email report with statistics"
        self.unified_output.insert("1.0", placeholder)
        self.unified_output.config(state="disabled")
        
        # Store email report start position for clipboard copying
        self.email_report_start = None

    def classify_file(self, file_path, file_type_suffix):
        """
        Automatically classify a file based on its name
        Returns the key (e.g., 'accreditation_d365') or None if can't classify
        """
        from main import D365_PATTERNS, SC_PATTERNS

        filename_lower = Path(file_path).name.lower()
        patterns = D365_PATTERNS if file_type_suffix == "d365" else SC_PATTERNS

        # Try to match against each pattern
        for report_type, pattern_list in patterns.items():
            # Convert single pattern to list
            if isinstance(pattern_list, str):
                pattern_list = [pattern_list]

            # Check if any pattern matches
            for pattern in pattern_list:
                if pattern.lower() in filename_lower:
                    return f"{report_type}_{file_type_suffix}"

        return None

    def handle_bulk_drop(self, event, file_type):
        """Handle multiple files dropped at once"""
        try:
            # Parse dropped files
            file_paths = self.parse_dropped_files(event.data)

            if not file_paths:
                return

            classified = {}
            unclassified = []
            suffix = file_type  # 'd365' or 'sc'

            # Classify each file
            for file_path in file_paths:
                # Validate file type
                if not file_path.lower().endswith((".xlsx", ".xls", ".csv")):
                    continue

                # Try to classify
                key = self.classify_file(file_path, suffix)
                if key:
                    classified[key] = file_path
                else:
                    unclassified.append(Path(file_path).name)

            # Update uploaded files and UI
            for key, file_path in classified.items():
                self.uploaded_files[key] = file_path
                # Update status indicators
                report_type = key.replace(f"_{suffix}", "")
                self.update_file_status(report_type, suffix, True)

            # Show results
            if classified:
                report = f"✅ Automatically classified {len(classified)} file(s):\n"
                for key in classified:
                    report_name = key.replace(f"_{suffix}", "").replace("_", " ").title()
                    report += f"  • {report_name}\n"

                if unclassified:
                    report += f"\n⚠ Could not classify {len(unclassified)} file(s):\n"
                    for name in unclassified:
                        report += f"  • {name}\n"

                messagebox.showinfo("Files Classified", report)
                self.update_status_indicator("success")
            else:
                messagebox.showerror(
                    "Classification Failed",
                    f"Could not automatically classify the dropped files.\n\n"
                    f"Make sure filenames contain:\n"
                    f"  • 'accreditation' for Accreditation\n"
                    f"  • 'wcb' for WCB\n"
                    f"  • 'cs' or 'client' for Client Specific",
                )
                self.update_status_indicator("error")

            # Check if we can enable process buttons
            self.check_upload_status()
            self.status_var.set(f"Classified {len(classified)} file(s)")

        except Exception as e:
            import traceback

            traceback.print_exc()
            messagebox.showerror("Error", f"Error processing dropped files: {e}")
            self.update_status_indicator("error")

    def parse_dropped_files(self, data):
        """Parse file paths from drag-and-drop event data"""
        import re
        
        file_paths = []
        data = data.strip()

        # Handle mixed braced and non-braced files
        if "{" in data:
            # More robust parsing for mixed formats
            # Example: "file1 {file2} file3" or "{file1} {file2} {file3}"
            
            # First extract braced files
            braced_pattern = r'\{([^}]+)\}'
            braced_files = re.findall(braced_pattern, data)
            
            for bf in braced_files:
                if Path(bf).exists():
                    file_paths.append(bf)
            
            # Remove braced sections to find non-braced files
            remaining = re.sub(braced_pattern, '', data).strip()
            # Also remove the extra braces themselves
            remaining = remaining.replace('{}', '').strip()
            
            if remaining:
                # Split remaining by spaces (careful with paths that have spaces)
                parts = remaining.split()
                current_path = ""
                
                for part in parts:
                    if current_path:
                        test_path = current_path + " " + part
                        if Path(test_path).exists():
                            current_path = test_path
                        elif Path(current_path).exists():
                            file_paths.append(current_path)
                            current_path = part
                        else:
                            current_path = test_path
                    else:
                        current_path = part
                
                if current_path and Path(current_path).exists():
                    file_paths.append(current_path)

        else:
            # Files without braces - split by space and handle paths with spaces
            # Windows paths can have spaces, so we need to be smart about splitting
            parts = data.split()
            current_path = ""

            for part in parts:
                if current_path:
                    # Try appending to current path
                    test_path = current_path + " " + part
                    if Path(test_path).exists():
                        current_path = test_path
                    elif Path(current_path).exists():
                        # Current path is valid, save it and start new path
                        file_paths.append(current_path)
                        current_path = part
                    else:
                        # Keep building current path
                        current_path = test_path
                else:
                    current_path = part

            # Don't forget the last path
            if current_path and Path(current_path).exists():
                file_paths.append(current_path)
        
        return file_paths

    def handle_drop(self, event, file_key):
        """Handle file drop event"""
        try:
            # Get the dropped file path (tkinterdnd2 returns paths in curly braces)
            file_path = event.data.strip()

            # Remove curly braces if present
            if file_path.startswith("{") and file_path.endswith("}"):
                file_path = file_path[1:-1]

            # Handle multiple files (take first one)
            if " " in file_path and not Path(file_path).exists():
                # Might be multiple files, take the first valid path
                possible_paths = file_path.split("} {")
                for p in possible_paths:
                    p = p.strip("{}")
                    if Path(p).exists():
                        file_path = p
                        break

            # Validate file exists and type
            if not Path(file_path).exists():
                logger.warning(f"Dropped file not found: {file_path}")
                messagebox.showerror("Error", f"Could not find file: {file_path}")
                return

            if not file_path.lower().endswith((".xlsx", ".xls", ".csv")):
                logger.warning(f"Invalid file type dropped: {file_path}")
                messagebox.showerror(
                    "Error", "Please drop an Excel file (.xlsx, .xls, or .csv)"
                )
                return

            # Store the file
            self.uploaded_files[file_key] = file_path
            filename = Path(file_path).name
            logger.info(f"File dropped and stored: {filename} as {file_key}")

            # Check if we can enable process buttons
            self.check_upload_status()
            self.status_var.set(f"Dropped: {filename}")

        except Exception as e:
            logger.exception(f"Error processing dropped file: {str(e)}")
            import traceback

            traceback.print_exc()
            messagebox.showerror("Error", f"Error processing dropped file: {e}")

    def on_drag_enter(self, event, frame, file_type):
        """Visual feedback when dragging over drop zone - smooth animation"""
        current_bg = str(frame["bg"])
        
        if file_type == "d365":
            frame.config(bg="#3b5a7f")  # Brighter blue on hover
        elif file_type == "sc":
            frame.config(bg="#2d5a43")  # Brighter green on hover

    def on_drag_leave(self, event, frame, file_type):
        """Reset visual when leaving drop zone"""
        if file_type == "d365":
            frame.config(bg="#2a3f5f")  # Original D365 blue
        elif file_type == "sc":
            frame.config(bg="#1f3a2c")  # Original SC green

    def update_file_status(self, report_type, file_type, uploaded):
        """Update file status indicators with colored dots"""
        print(f"\n=== UPDATE_FILE_STATUS ===")
        print(f"Report Type: {report_type}")
        print(f"File Type: {file_type}")
        print(f"Uploaded: {uploaded}")
        
        status_dict = self.d365_status_labels if file_type == "d365" else self.sc_status_labels
        print(f"Status dict keys: {list(status_dict.keys())}")
        
        if report_type in status_dict:
            print(f"  \u2713 Report type found in status dict")
            indicator = status_dict[report_type]["indicator"]
            label = status_dict[report_type]["label"]
            
            if uploaded:
                # Green dot for uploaded
                indicator.delete("all")
                indicator.create_oval(2, 2, 8, 8, fill=self.colors['accent_green'], outline="")
                label.config(text=f"{report_type.replace('_', ' ').title()}: ✓ Uploaded", fg=self.colors['text_primary'])
                print(f"  Updated to GREEN (uploaded)")
            else:
                # Gray dot for not uploaded
                indicator.delete("all")
                indicator.create_oval(2, 2, 8, 8, fill="#6b7280", outline="")
                label.config(text=f"{report_type.replace('_', ' ').title()}: Not uploaded", fg=self.colors['text_secondary'])
                print(f"  Updated to GRAY (not uploaded)")
            
            # Force UI update
            indicator.update_idletasks()
            label.update_idletasks()
        else:
            print(f"  \u2717 WARNING: {report_type} NOT FOUND in status dict!")
    
    def update_status_indicator(self, status_type):
        """Update the main status bar indicator dot"""
        if status_type == "success":
            self.status_indicator.itemconfig(self.status_dot, fill=self.colors['accent_green'])
        elif status_type == "error":
            self.status_indicator.itemconfig(self.status_dot, fill="#ef4444")
        elif status_type == "warning":
            self.status_indicator.itemconfig(self.status_dot, fill=self.colors['accent_orange'])
        elif status_type == "processing":
            self.status_indicator.itemconfig(self.status_dot, fill=self.colors['accent_blue'])
        else:  # idle
            self.status_indicator.itemconfig(self.status_dot, fill="#6b7280")

    def check_upload_status(self):
        """Check if any files are uploaded and enable buttons"""
        # Check D365 files - enable button if ANY file is uploaded
        d365_any = any(
            self.uploaded_files[k] for k in ["accreditation_d365", "wcb_d365", "client_d365"]
        )
        if hasattr(self, "btn_process_d365"):
            self.btn_process_d365.config(state=tk.NORMAL if d365_any else tk.DISABLED)

        # Check SC files - enable button if ANY file is uploaded
        sc_any = any(
            self.uploaded_files[k] for k in ["accreditation_sc", "wcb_sc", "client_sc"]
        )
        if hasattr(self, "btn_process_sc"):
            self.btn_process_sc.config(state=tk.NORMAL if sc_any else tk.DISABLED)

    def check_existing_files(self):
        """Check for existing files in input folder and mark as uploaded"""
        DYNAMICS_DIR.mkdir(parents=True, exist_ok=True)
        REDASH_DIR.mkdir(parents=True, exist_ok=True)
        
        # Check for D365 files in dynamics directory
        for key, filename in D365_FILES.items():
            file_path = DYNAMICS_DIR / filename
            if file_path.exists():
                full_key = f"{key}_d365"
                self.uploaded_files[full_key] = str(file_path)
                self.update_file_status(key, "d365", True)
        
        # Check for SC files in redash directory
        for key, filename in SC_FILES.items():
            file_path = REDASH_DIR / filename
            if file_path.exists():
                full_key = f"{key}_sc"
                self.uploaded_files[full_key] = str(file_path)
                self.update_file_status(key, "sc", True)
        
        # Update button states after checking
        self.check_upload_status()
        
        # Update status bar with detected files
        d365_count = sum(1 for k, v in self.uploaded_files.items() if "_d365" in k and v is not None)
        sc_count = sum(1 for k, v in self.uploaded_files.items() if "_sc" in k and v is not None)
        if d365_count > 0 or sc_count > 0:
            self.status_var.set(f"Loaded {d365_count} D365 file(s) and {sc_count} SC file(s) from disk")

    def cleanup_files(self):
        """Delete all uploaded files from input directories"""
        logger.info("Cleanup: Removing uploaded files from input directories")
        try:
            deleted_count = 0
            
            # Delete D365 files
            if DYNAMICS_DIR.exists():
                for file in DYNAMICS_DIR.glob('*.xlsx'):
                    try:
                        file.unlink()
                        deleted_count += 1
                        logger.debug(f"Deleted: {file.name}")
                    except Exception as e:
                        logger.warning(f"Error deleting {file.name}: {e}")
                        print(f"Error deleting {file.name}: {e}")
            
            # Delete SC files
            if REDASH_DIR.exists():
                for file in REDASH_DIR.glob('*.xlsx'):
                    try:
                        file.unlink()
                        deleted_count += 1
                        logger.debug(f"Deleted: {file.name}")
                    except Exception as e:
                        logger.warning(f"Error deleting {file.name}: {e}")
                        print(f"Error deleting {file.name}: {e}")
            
            if deleted_count > 0:
                logger.info(f"Cleaned up {deleted_count} file(s)")
                print(f"Cleaned up {deleted_count} file(s)")
        except Exception as e:
            logger.exception(f"Error during cleanup: {str(e)}")
            print(f"Error during cleanup: {e}")

    def on_closing(self):
        """Handle window close event"""
        logger.info("Application closing")
        self.cleanup_files()
        self.root.destroy()

    def save_d365_files(self):
        """Copy uploaded D365 files to input folder (only files that were uploaded)"""
        try:
            logger.info("Starting D365 file save process")
            DYNAMICS_DIR.mkdir(parents=True, exist_ok=True)
            
            saved_files = []
            skipped_files = []
            for key in ["accreditation_d365", "wcb_d365", "client_d365"]:
                report_type = key.replace("_d365", "")
                if self.uploaded_files[key]:
                    source = Path(self.uploaded_files[key])

                    # Validate source exists
                    if not source.exists():
                        raise FileNotFoundError(f"Source file not found: {source}")

                    dest = DYNAMICS_DIR / D365_FILES[report_type]

                    # Ensure destination path is valid
                    dest = dest.resolve()

                    shutil.copy2(str(source), str(dest))
                    saved_files.append(report_type)
                    logger.info(f"Saved D365 {report_type} file: {source.name} -> {dest.name}")
                else:
                    skipped_files.append(report_type)
            
            logger.info(f"Successfully saved {len(saved_files)} D365 files: {', '.join(saved_files)}")
            if skipped_files:
                logger.info(f"Skipped (not uploaded): {', '.join(skipped_files)}")
            
            msg = f"{Messages.SUCCESS} Saved {len(saved_files)} D365 file(s): {', '.join(f.title() for f in saved_files)}"
            if skipped_files:
                msg += f"\n\nSkipped (not uploaded): {', '.join(f.title() for f in skipped_files)}"
            msg += "\n\nNext step: Go to 'Extract IDs' tab to generate ID lists for Redash."
            messagebox.showinfo("Success", msg)

            # Always go to Extract IDs tab in manual mode
            self.notebook.select(1)  # Switch to Extract IDs tab
            self.status_var.set(f"Saved {len(saved_files)} D365 file(s) - Ready to extract IDs")
            self.update_status_indicator("success")

        except Exception as e:
            logger.exception(f"Failed to save D365 files: {str(e)}")
            messagebox.showerror("Error", f"Failed to save files:\n{e}")
            self.update_status_indicator("error")

    def save_sc_files(self):
        """Copy uploaded SafeContractor files to input folder (only files that were uploaded)"""
        try:
            logger.info("Starting SC file save process")
            REDASH_DIR.mkdir(parents=True, exist_ok=True)
            
            saved_files = []
            skipped_files = []
            for key in ["accreditation_sc", "wcb_sc", "client_sc"]:
                report_type = key.replace("_sc", "")
                if self.uploaded_files[key]:
                    source = Path(self.uploaded_files[key])

                    # Validate source exists
                    if not source.exists():
                        raise FileNotFoundError(f"Source file not found: {source}")

                    dest = REDASH_DIR / SC_FILES[report_type]

                    # Ensure destination path is valid
                    dest = dest.resolve()

                    shutil.copy2(str(source), str(dest))
                    saved_files.append(report_type)
                    logger.info(f"Saved SC {report_type} file: {source.name} -> {dest.name}")
                else:
                    skipped_files.append(report_type)
            
            logger.info(f"Successfully saved {len(saved_files)} SC files: {', '.join(saved_files)}")
            if skipped_files:
                logger.info(f"Skipped (not uploaded): {', '.join(skipped_files)}")
            
            msg = f"{Messages.SUCCESS} Saved {len(saved_files)} SC file(s): {', '.join(f.title() for f in saved_files)}"
            if skipped_files:
                msg += f"\n\nSkipped (not uploaded): {', '.join(f.title() for f in skipped_files)}"
            msg += "\n\nNext step: Go to 'Generate Comparisons' tab to create comparison reports."
            messagebox.showinfo("Success", msg)
            self.notebook.select(3)  # Switch to Generate Comparisons tab
            self.status_var.set(f"Saved {len(saved_files)} SC file(s) - Ready to generate comparisons")
            self.update_status_indicator("success")

        except Exception as e:
            logger.exception(f"Failed to save SC files: {str(e)}")
            messagebox.showerror("Error", f"Failed to save files:\n{e}")
            self.update_status_indicator("error")

    def extract_ids(self):
        """Run ID extraction in background thread"""
        logger.info("Starting ID extraction process")
        self.btn_extract.config(state=tk.DISABLED, text="⏳ Extracting...")
        self.extract_console.delete(1.0, tk.END)
        self.status_var.set("Extracting IDs...")
        self.update_status_indicator("processing")
        
        # Show and start progress bar
        self.extract_progress.pack(pady=10)
        self.extract_progress.start()

        def run_extraction():
            # Redirect stdout to console
            import io
            import contextlib

            f = io.StringIO()
            with contextlib.redirect_stdout(f):
                try:
                    extract_and_save_ids()
                    output = f.getvalue()
                    logger.info("ID extraction completed successfully")
                except Exception as e:
                    logger.exception(f"ID extraction failed: {str(e)}")
                    output = f"Error: {e}\n{f.getvalue()}"

            # Update UI in main thread
            self.root.after(0, self.extraction_complete, output)

        thread = threading.Thread(target=run_extraction, daemon=True)
        thread.start()

    def extraction_complete(self, output):
        """Handle extraction completion"""
        self.extract_console.insert(tk.END, output)
        self.extract_console.see(tk.END)
        self.btn_extract.config(state=tk.NORMAL, text="⚙️ Extract IDs")
        
        # Hide and stop progress bar
        self.extract_progress.stop()
        self.extract_progress.pack_forget()

        if "Error" in output or "❌" in output:
            self.status_var.set("ID extraction failed - Check console for errors")
            self.update_status_indicator("error")
            messagebox.showerror(
                "Error", "ID extraction failed. Check the console output for details."
            )
        else:
            self.status_var.set("IDs extracted successfully!")
            self.update_status_indicator("success")
            
            # Automatically open the query_ids folder
            self.open_folder(QUERY_IDS_DIR)
            
            messagebox.showinfo(
                "Success",
                f"{Messages.SUCCESS} IDs extracted successfully!\n\n"
                "Next steps:\n"
                "1. Copy IDs from .sql.txt files (folder opened)\n"
                "2. Run Redash queries with those IDs\n"
                "3. Upload SC results in the next tab",
            )
            self.notebook.select(2)  # Switch to Upload SC tab

    def generate_comparison(self):
        """Run comparison generation in background thread"""
        logger.info("Starting comparison generation")
        # Check if SC files exist
        from config import SC_PATTERNS, REPORT_TYPES
        from utils import find_file_by_pattern

        sc_files_exist = any(
            find_file_by_pattern(REDASH_DIR, SC_PATTERNS[t]) is not None for t in REPORT_TYPES
        )

        if not sc_files_exist:
            logger.warning("Comparison generation aborted - SC files missing")
            messagebox.showerror(
                "SC Files Missing",
                "SafeContractor files are not found!\n\n"
                "Please:\n"
                "1. Extract IDs from D365 files (Tab 2)\n"
                "2. Run Redash queries with those IDs\n"
                "3. Upload the SC files from Redash (Tab 3)\n"
                "4. Then generate comparisons",
            )
            self.status_var.set("SC files missing - Cannot generate comparisons")
            return

        self.btn_compare.config(state=tk.DISABLED, text="⏳ Generating...")
        self.unified_output.config(state="normal")
        self.unified_output.delete(1.0, tk.END)
        self.unified_output.config(state="disabled")
        self.status_var.set("Generating comparisons...")

        def run_comparison():
            # Redirect stdout to console with real-time updates
            import sys
            from io import StringIO

            class StreamToConsole:
                def __init__(self, console_widget, root):
                    self.console = console_widget
                    self.root = root
                    self.buffer = StringIO()

                def write(self, text):
                    self.buffer.write(text)
                    if self.console:
                        self.root.after(0, lambda: self._update_console(text))

                def _update_console(self, text):
                    if self.console:
                        self.console.config(state="normal")
                        self.console.insert(tk.END, text)
                        self.console.see(tk.END)
                        self.console.update_idletasks()
                        self.console.config(state="disabled")

                def flush(self):
                    pass

                def getvalue(self):
                    return self.buffer.getvalue()

            stream = StreamToConsole(self.unified_output, self.root)

            try:
                old_stdout = sys.stdout
                sys.stdout = stream
                generate_comparisons()
                output = stream.getvalue()
                logger.info("Comparison generation completed successfully")
            except Exception as e:
                logger.exception(f"Comparison generation failed: {str(e)}")
                output = f"Error: {e}\n{stream.getvalue()}"
            finally:
                sys.stdout = old_stdout

            # Update UI in main thread
            self.root.after(0, self.comparison_complete, output)

        thread = threading.Thread(target=run_comparison, daemon=True)
        thread.start()

    def comparison_complete(self, output):
        """Handle comparison completion"""
        self.unified_output.config(state="normal")
        self.unified_output.insert(tk.END, output)
        self.unified_output.see(tk.END)
        self.unified_output.config(state="disabled")
        
        self.btn_compare.config(state=tk.NORMAL, text="🚀 Generate Comparisons")
        
        # Hide and stop progress bar
        self.compare_progress.stop()
        self.compare_progress.pack_forget()

        if "Error" in output or "❌" in output:
            self.status_var.set("Comparison generation failed - Check output for errors")
            self.update_status_indicator("error")
            messagebox.showerror(
                "Error", "Comparison generation failed. Check the output for details."
            )
        else:
            self.status_var.set("Comparisons generated successfully!")
            self.update_status_indicator("success")
            
            # Generate and display email report in the same output area
            self.auto_generate_email_report()
            
            # Automatically open the output folder
            self.open_folder(OUTPUT_DIR)
            
            messagebox.showinfo(
                "Success",
                f"{Messages.SUCCESS} Comparison files generated successfully!\n\n"
                "The Excel files are ready in the output folder (opened automatically).\n"
                "Email report displayed below - click 'Copy Email Report' to use it.",
            )

    def auto_generate_email_report(self):
        """Automatically generate email report from comparison files after successful comparison"""
        logger.info("Auto-generating email report after comparison completion")
        
        def run_email_generation():
            try:
                # Import the email report generator
                from generate_email_report import (
                    read_comparison_file,
                    analyze_sc_sheet,
                    analyze_d365_sheet,
                    format_status_name
                )
                
                # Get the dated comparison directory
                comparison_dir = get_dated_comparison_dir()
                
                # Define comparison types and their file paths
                comparisons = {
                    "Client": comparison_dir / "Client_Comparison.xlsx",
                    "WCB": comparison_dir / "WCB_Comparison.xlsx",
                    "Accreditation": comparison_dir / "Accreditation_Comparison.xlsx"
                }
                
                # Check which files exist
                available_comparisons = {}
                
                for name, path in comparisons.items():
                    if path.exists():
                        available_comparisons[name] = path
                
                if not available_comparisons:
                    logger.warning("No comparison files found for email report")
                    return
                
                # Analyze each comparison
                results = {}
                for name, file_path in available_comparisons.items():
                    # Read the file
                    sc_df, d365_df = read_comparison_file(file_path)
                    
                    if sc_df is None or d365_df is None:
                        continue
                    
                    # Analyze both sheets (pass report type for correct column detection)
                    sc_stats = analyze_sc_sheet(sc_df, d365_df, report_type=name)
                    d365_stats = analyze_d365_sheet(d365_df, sc_df, report_type=name)
                    
                    results[name] = {
                        "sc": sc_stats,
                        "d365": d365_stats
                    }
                
                # Generate email text
                email_lines = []
                
                # Process in the order: Client, WCB, Accreditation
                order = ["Client", "WCB", "Accreditation"]
                
                for name in order:
                    if name not in results:
                        continue
                    
                    data = results[name]
                    
                    # Section header
                    if name == "Client":
                        email_lines.append("Client specific:\n")
                        email_lines.append("SC:")
                    else:
                        email_lines.append(f"\n{name}:\n")
                        email_lines.append("SC:")
                    
                    # SC statistics
                    sc_diff = data["sc"]["differences"]
                    sc_not_found = data["sc"]["not_found"]
                    
                    if sc_not_found > 0:
                        email_lines.append(f"{sc_diff} differences between dynamics and SafeContractor, {sc_not_found} Not found")
                    else:
                        email_lines.append(f"{sc_diff} differences between dynamics and SafeContractor")
                    
                    # D365 statistics
                    email_lines.append("\nD365:")
                    
                    total_not_found = data["d365"]["total_not_found"]
                    email_lines.append(f"{total_not_found} not found in SafeContractor:")
                    
                    # Sort status breakdown alphabetically for consistency
                    status_breakdown = data["d365"]["status_breakdown"]
                    if status_breakdown:
                        sorted_statuses = sorted(status_breakdown.items(), key=lambda x: x[0])
                        for status, count in sorted_statuses:
                            formatted_status = format_status_name(status)
                            email_lines.append(f"{count} {formatted_status}")
                
                # Join all lines
                email_text = "\n".join(email_lines)
                
                # Save to file
                output_file = OUTPUT_DIR / "email_report.txt"
                try:
                    with open(output_file, "w", encoding="utf-8") as f:
                        f.write(email_text)
                    logger.info(f"Email report saved to {output_file}")
                except Exception as save_error:
                    logger.error(f"Failed to save email report to file: {save_error}")
                
                # Update UI in main thread
                self.root.after(0, lambda: self.display_email_report(email_text))
                
            except Exception as e:
                logger.exception(f"Email report generation failed: {str(e)}")
                error_msg = f"\nNote: Email report generation failed: {str(e)}\n"
                self.root.after(0, lambda msg=error_msg: self._insert_to_unified_output(msg, "error"))
        
        thread = threading.Thread(target=run_email_generation, daemon=True)
        thread.start()
    
    def display_email_report(self, email_text):
        """Display the generated email report in the unified output"""
        self.unified_output.config(state="normal")
        
        # Add separator
        separator = "\n" + "="*70 + "\n"
        self.unified_output.insert(tk.END, separator, "separator")
        
        # Add header
        header = "📧 EMAIL REPORT\n"
        self.unified_output.insert(tk.END, header, "header")
        
        separator2 = "="*70 + "\n\n"
        self.unified_output.insert(tk.END, separator2, "separator")
        
        # Store position where email report starts (for clipboard copying)
        self.email_report_start = self.unified_output.index(tk.INSERT)
        
        # Insert email report
        self.unified_output.insert(tk.END, email_text, "email")
        
        # Scroll to show email report
        self.unified_output.see(tk.END)
        self.unified_output.config(state="disabled")
        
        # Enable copy button
        self.btn_copy_report.config(state=tk.NORMAL)
        
        logger.info("Email report displayed successfully")
    
    def _insert_to_unified_output(self, text, tag=None):
        """Helper method to insert text to unified output"""
        self.unified_output.config(state="normal")
        if tag:
            self.unified_output.insert(tk.END, text, tag)
        else:
            self.unified_output.insert(tk.END, text)
        self.unified_output.see(tk.END)
        self.unified_output.config(state="disabled")
    
    def copy_email_to_clipboard(self):
        """Copy email report text to clipboard"""
        try:
            # Get email report content only (from stored position to end)
            if self.email_report_start:
                email_content = self.unified_output.get(self.email_report_start, tk.END).strip()
            else:
                # Fallback: try to get from file if position not stored
                email_file = OUTPUT_DIR / "email_report.txt"
                if email_file.exists():
                    with open(email_file, "r", encoding="utf-8") as f:
                        email_content = f.read().strip()
                else:
                    email_content = ""
            
            if not email_content:
                messagebox.showwarning("Warning", "No email report to copy!")
                return
            
            # Copy to clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(email_content)
            self.root.update()  # Required to finalize clipboard operation
            
            self.status_var.set("Email report copied to clipboard!")
            messagebox.showinfo("Success", "✓ Email report copied to clipboard!\n\nYou can now paste it into your email.")
            logger.info("Email report copied to clipboard")
            
        except Exception as e:
            logger.exception(f"Failed to copy to clipboard: {str(e)}")
            messagebox.showerror("Error", f"Failed to copy to clipboard: {str(e)}")

    def open_folder(self, folder_path):
        """Open folder in file explorer"""
        import os
        import subprocess

        folder_path.mkdir(parents=True, exist_ok=True)

        if sys.platform == "win32":
            os.startfile(folder_path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder_path])
        else:
            subprocess.Popen(["xdg-open", folder_path])


def main():
    """Launch the GUI application"""
    # Create TkinterDnD root first, then apply ttkbootstrap theme
    root = TkinterDnD.Tk()
    
    # Apply ttkbootstrap dark theme to existing window
    style = ttk.Style("darkly")
    root.title("D365 vs SafeContractor - Status Comparison Tool")
    root.geometry("1100x850")
    root.resizable(True, True)
    
    app = ComparisonApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
