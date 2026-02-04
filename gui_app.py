"""
GUI Interface for Dynamics 365 vs SafeContractor Status Comparison
Drag-and-drop file upload interface with manual workflow
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
from pathlib import Path
import shutil
import sys
import warnings

# Suppress openpyxl style warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Import configuration and utilities
from config import INPUT_DIR, OUTPUT_DIR, DYNAMICS_DIR, REDASH_DIR, D365_FILES, SC_FILES, Messages

# Import main processing functions
from automate_comparison import extract_and_save_ids, generate_comparisons


class ComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("D365 vs SafeContractor - Status Comparison Tool")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f5f5")

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
        self.check_existing_files()

    def setup_ui(self):
        """Create the user interface"""
        # Header frame with gradient-like appearance
        header_frame = tk.Frame(self.root, bg="#1976D2", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        # Title
        title_label = tk.Label(
            header_frame,
            text="📊 D365 vs SafeContractor",
            font=("Segoe UI", 20, "bold"),
            fg="white",
            bg="#1976D2",
        )
        title_label.pack(pady=8)

        subtitle_label = tk.Label(
            header_frame,
            text="Status Comparison & Reporting Tool",
            font=("Segoe UI", 11),
            fg="#E3F2FD",
            bg="#1976D2",
        )
        subtitle_label.pack()

        # Main container with tabs
        container = tk.Frame(self.root, bg="#f5f5f5")
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # Style for notebook
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background="#f5f5f5", borderwidth=0)
        style.configure("TNotebook.Tab", padding=[20, 10], font=("Segoe UI", 10, "bold"))
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#1976D2")],
            foreground=[("selected", "white")],
        )

        self.notebook = ttk.Notebook(container)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Upload D365 Files
        self.tab_d365 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_d365, text="1. Upload D365 Files")
        self.setup_d365_tab()

        # Tab 2: Extract IDs
        self.tab_extract = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_extract, text="2. Extract IDs")
        self.setup_extract_tab()

        # Tab 3: Upload SafeContractor Files
        self.tab_sc = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_sc, text="3. Upload SafeContractor Files")
        self.setup_sc_tab()

        # Tab 4: Generate Comparisons
        self.tab_compare = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_compare, text="4. Generate Comparisons")
        self.setup_compare_tab()

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("✓ Ready - Start by uploading D365 files")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            bd=0,
            relief=tk.FLAT,
            anchor=tk.W,
            bg="#263238",
            fg="white",
            font=("Segoe UI", 9),
            pady=8,
            padx=15,
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_d365_tab(self):
        """Setup D365 file upload tab"""
        self.tab_d365.configure(style="TFrame")

        info_frame = tk.Frame(self.tab_d365, bg="#E3F2FD", relief=tk.FLAT)
        info_frame.pack(fill=tk.X, padx=20, pady=20)

        tk.Label(
            info_frame,
            text="📁 Upload Dynamics 365 Exports",
            font=("Segoe UI", 14, "bold"),
            bg="#E3F2FD",
            fg="#0D47A1",
        ).pack(anchor=tk.W, padx=20, pady=(15, 5))

        instruction_text = "Drag & drop all 3 D365 files at once. The system will automatically detect Accreditation, WCB, and Client files."

        tk.Label(
            info_frame,
            text=instruction_text,
            bg="#E3F2FD",
            fg="#424242",
            font=("Segoe UI", 10),
            wraplength=900,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, padx=20, pady=(0, 15))

        # Multi-file drop zone
        multi_drop_frame = tk.Frame(self.tab_d365, bg="#f5f5f5")
        multi_drop_frame.pack(fill=tk.X, padx=20, pady=10)

        bulk_drop = tk.Frame(
            multi_drop_frame,
            bg="#FFF8E1",
            relief=tk.SOLID,
            bd=2,
            highlightbackground="#FFA726",
            highlightthickness=2,
            height=100,
        )
        bulk_drop.pack(fill=tk.X, pady=5)
        bulk_drop.pack_propagate(False)

        # Configure drag and drop for multiple files
        bulk_drop.drop_target_register(DND_FILES)
        bulk_drop.dnd_bind("<<Drop>>", lambda e: self.handle_bulk_drop(e, "d365"))
        bulk_drop.dnd_bind("<<DragEnter>>", lambda e, f=bulk_drop: self.on_drag_enter(e, f))
        bulk_drop.dnd_bind("<<DragLeave>>", lambda e, f=bulk_drop: self.on_drag_leave(e, f))

        tk.Label(
            bulk_drop,
            text="🚀 DRAG & DROP ALL 3 D365 FILES HERE\n\nSystem will automatically identify Accreditation, WCB, and Client files",
            bg="#FFF8E1",
            fg="#E65100",
            font=("Segoe UI", 11, "bold"),
            justify=tk.CENTER,
        ).pack(expand=True)

        # Process button
        btn_frame = tk.Frame(self.tab_d365, bg="#f5f5f5")
        btn_frame.pack(pady=30)

        self.btn_process_d365 = tk.Button(
            btn_frame,
            text="✓ Save D365 Files & Continue",
            command=self.save_d365_files,
            bg="#1976D2",
            fg="white",
            font=("Segoe UI", 12, "bold"),
            width=28,
            height=2,
            cursor="hand2",
            relief=tk.FLAT,
            activebackground="#1565C0",
            activeforeground="white",
            state=tk.DISABLED,
            disabledforeground="#BDBDBD",
        )
        self.btn_process_d365.pack()

    def setup_extract_tab(self):
        """Setup ID extraction tab"""
        info_frame = tk.Frame(self.tab_extract, bg="#FFF9C4", relief=tk.FLAT)
        info_frame.pack(fill=tk.X, padx=20, pady=20)

        tk.Label(
            info_frame,
            text="🔍 Extract IDs for Redash Queries",
            font=("Segoe UI", 14, "bold"),
            bg="#FFF9C4",
            fg="#F57F17",
        ).pack(anchor=tk.W, padx=20, pady=(15, 5))

        tk.Label(
            info_frame,
            text="Extract and format Global Alcumus IDs from D365 files. The formatted IDs will be ready to copy into Redash queries.",
            bg="#FFF9C4",
            fg="#424242",
            font=("Segoe UI", 10),
            wraplength=850,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(5, 0))

        # Extract button
        btn_frame = tk.Frame(self.tab_extract)
        btn_frame.pack(pady=20)

        self.btn_extract = tk.Button(
            btn_frame,
            text="⚙️ Extract IDs",
            command=self.extract_ids,
            bg="#FF9800",
            fg="white",
            font=("Arial", 11, "bold"),
            width=30,
            height=2,
            cursor="hand2",
        )
        self.btn_extract.pack()

        # Output console
        console_label = tk.Label(self.tab_extract, text="Output:", font=("Arial", 10, "bold"))
        console_label.pack(anchor=tk.W, padx=15, pady=(20, 5))

        self.extract_console = scrolledtext.ScrolledText(
            self.tab_extract,
            height=20,
            bg="#1e1e1e",
            fg="#00ff00",
            font=("Consolas", 9),
            wrap=tk.WORD,
        )
        self.extract_console.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

        # Open output folder button
        tk.Button(
            self.tab_extract,
            text="📁 Open Output Folder",
            command=lambda: self.open_folder(OUTPUT_DIR),
            bg="#607D8B",
            fg="white",
            font=("Arial", 10),
            cursor="hand2",
        ).pack(pady=10)

    def setup_sc_tab(self):
        """Setup SafeContractor file upload tab"""
        info_frame = tk.Frame(self.tab_sc, bg="#E8F5E9", relief=tk.FLAT)
        info_frame.pack(fill=tk.X, padx=20, pady=20)

        tk.Label(
            info_frame,
            text="📊 Upload SafeContractor (Redash) Exports",
            font=("Segoe UI", 14, "bold"),
            bg="#E8F5E9",
            fg="#1B5E20",
        ).pack(anchor=tk.W, padx=20, pady=(15, 5))

        tk.Label(
            info_frame,
            text="Drag & drop all 3 SafeContractor files at once. The system will automatically detect Accreditation, WCB, and Client files.",
            bg="#E8F5E9",
            fg="#424242",
            font=("Segoe UI", 10),
            wraplength=900,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, padx=20, pady=(0, 15))

        # Multi-file drop zone
        multi_drop_frame = tk.Frame(self.tab_sc, padx=15)
        multi_drop_frame.pack(fill=tk.X, pady=(10, 5))

        bulk_drop = tk.Frame(multi_drop_frame, bg="#fff3cd", relief=tk.RIDGE, bd=3, height=80)
        bulk_drop.pack(fill=tk.X, pady=5)
        bulk_drop.pack_propagate(False)

        # Configure drag and drop for multiple files
        bulk_drop.drop_target_register(DND_FILES)
        bulk_drop.dnd_bind("<<Drop>>", lambda e: self.handle_bulk_drop(e, "sc"))
        bulk_drop.dnd_bind("<<DragEnter>>", lambda e, f=bulk_drop: self.on_drag_enter(e, f))
        bulk_drop.dnd_bind("<<DragLeave>>", lambda e, f=bulk_drop: self.on_drag_leave(e, f))

        tk.Label(
            bulk_drop,
            text="🚀 QUICK: Drag & Drop All 3 SafeContractor Files Here\n(System will auto-detect Accreditation, WCB, and Client files)",
            bg="#fff3cd",
            fg="#856404",
            font=("Arial", 10, "bold"),
            justify=tk.CENTER,
        ).pack(expand=True)

        # Process button
        btn_frame = tk.Frame(self.tab_sc)
        btn_frame.pack(pady=20)

        self.btn_process_sc = tk.Button(
            btn_frame,
            text="✅ Save SafeContractor Files & Proceed",
            command=self.save_sc_files,
            bg="#2196F3",
            fg="white",
            font=("Arial", 11, "bold"),
            width=30,
            height=2,
            cursor="hand2",
            state=tk.DISABLED,
        )
        self.btn_process_sc.pack()

    def setup_compare_tab(self):
        """Setup comparison generation tab"""
        info_frame = tk.Frame(self.tab_compare, bg="#f3e5f5", padx=15, pady=15)
        info_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(
            info_frame,
            text="📊 Generate Status Comparison Files",
            font=("Arial", 12, "bold"),
            bg="#f3e5f5",
        ).pack(anchor=tk.W)

        tk.Label(
            info_frame,
            text="Click below to generate the final comparison Excel files with status matching and differences.",
            bg="#f3e5f5",
            wraplength=850,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(5, 0))

        # Generate button
        btn_frame = tk.Frame(self.tab_compare)
        btn_frame.pack(pady=20)

        self.btn_compare = tk.Button(
            btn_frame,
            text="🚀 Generate Comparisons",
            command=self.generate_comparison,
            bg="#9C27B0",
            fg="white",
            font=("Arial", 11, "bold"),
            width=30,
            height=2,
            cursor="hand2",
        )
        self.btn_compare.pack()

        # Output console
        console_label = tk.Label(self.tab_compare, text="Output:", font=("Arial", 10, "bold"))
        console_label.pack(anchor=tk.W, padx=15, pady=(20, 5))

        self.compare_console = scrolledtext.ScrolledText(
            self.tab_compare,
            height=20,
            bg="#1e1e1e",
            fg="#00ff00",
            font=("Consolas", 9),
            wrap=tk.WORD,
        )
        self.compare_console.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

        # Open output folder button
        tk.Button(
            self.tab_compare,
            text="📁 Open Output Folder",
            command=lambda: self.open_folder(OUTPUT_DIR),
            bg="#607D8B",
            fg="white",
            font=("Arial", 10),
            cursor="hand2",
        ).pack(pady=10)

    def classify_file(self, file_path, file_type_suffix):
        """
        Automatically classify a file based on its name
        Returns the key (e.g., 'accreditation_d365') or None if can't classify
        """
        from automate_comparison import D365_PATTERNS, SC_PATTERNS

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
            else:
                messagebox.showerror(
                    "Classification Failed",
                    f"Could not automatically classify the dropped files.\n\n"
                    f"Make sure filenames contain:\n"
                    f"  • 'accreditation' for Accreditation\n"
                    f"  • 'wcb' for WCB\n"
                    f"  • 'cs' or 'client' for Client Specific",
                )

            # Check if we can enable process buttons
            self.check_upload_status()
            self.status_var.set(f"Classified {len(classified)} file(s)")

        except Exception as e:
            import traceback

            traceback.print_exc()
            messagebox.showerror("Error", f"Error processing dropped files: {e}")

    def parse_dropped_files(self, data):
        """Parse file paths from drag-and-drop event data"""
        file_paths = []
        data = data.strip()

        # Handle multiple files (separated by spaces, wrapped in {})
        if "{" in data:
            # Split by '} {' to handle multiple files with braces
            parts = data.split("} {")
            for part in parts:
                part = part.strip("{}").strip()
                if part and Path(part).exists():
                    file_paths.append(part)
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
                messagebox.showerror("Error", f"Could not find file: {file_path}")
                return

            if not file_path.lower().endswith((".xlsx", ".xls", ".csv")):
                messagebox.showerror(
                    "Error", "Please drop an Excel file (.xlsx, .xls, or .csv)"
                )
                return

            # Store the file
            self.uploaded_files[file_key] = file_path
            filename = Path(file_path).name

            # Check if we can enable process buttons
            self.check_upload_status()
            self.status_var.set(f"Dropped: {filename}")

        except Exception as e:
            import traceback

            traceback.print_exc()
            messagebox.showerror("Error", f"Error processing dropped file: {e}")

    def on_drag_enter(self, event, frame):
        """Visual feedback when dragging over drop zone"""
        # Get current background to determine type
        current_bg = str(frame["bg"])
        if current_bg in ("#e8f5e9", "#c8e6c9"):  # D365 colors
            frame.config(bg="#81c784")
        elif current_bg in ("#e3f2fd", "#bbdefb"):  # SC colors
            frame.config(bg="#64b5f6")
        elif current_bg == "#fff3cd":  # Bulk drop zone
            frame.config(bg="#ffe082")

    def on_drag_leave(self, event, frame):
        """Reset visual when leaving drop zone"""
        # Restore original color based on current state
        current_bg = str(frame["bg"])
        if current_bg == "#81c784":  # Was D365 hover
            frame.config(bg="#e8f5e9")
        elif current_bg == "#64b5f6":  # Was SC hover
            frame.config(bg="#e3f2fd")
        elif current_bg == "#ffe082":  # Was bulk hover
            frame.config(bg="#fff3cd")

    def check_upload_status(self):
        """Check if all required files are uploaded and enable buttons"""
        # Check D365 files
        d365_complete = all(
            self.uploaded_files[k] for k in ["accreditation_d365", "wcb_d365", "client_d365"]
        )
        if hasattr(self, "btn_process_d365"):
            self.btn_process_d365.config(state=tk.NORMAL if d365_complete else tk.DISABLED)

        # Check SC files
        sc_complete = all(
            self.uploaded_files[k] for k in ["accreditation_sc", "wcb_sc", "client_sc"]
        )
        if hasattr(self, "btn_process_sc"):
            self.btn_process_sc.config(state=tk.NORMAL if sc_complete else tk.DISABLED)

    def check_existing_files(self):
        """Check for existing files in input folder and mark as uploaded"""
        INPUT_DIR.mkdir(exist_ok=True)

        for key, filename in {**D365_FILES, **SC_FILES}.items():
            file_path = INPUT_DIR / filename
            if file_path.exists():
                self.uploaded_files[key] = str(file_path)

        # Update button states after checking
        self.check_upload_status()

    def save_d365_files(self):
        """Copy D365 files to input folder"""
        try:
            DYNAMICS_DIR.mkdir(parents=True, exist_ok=True)

            for key in ["accreditation_d365", "wcb_d365", "client_d365"]:
                if self.uploaded_files[key]:
                    source = Path(self.uploaded_files[key])

                    # Validate source exists
                    if not source.exists():
                        raise FileNotFoundError(f"Source file not found: {source}")

                    # Strip _d365 suffix to get the report type
                    report_type = key.replace("_d365", "")
                    dest = DYNAMICS_DIR / D365_FILES[report_type]

                    # Ensure destination path is valid
                    dest = dest.resolve()

                    shutil.copy2(str(source), str(dest))

            messagebox.showinfo(
                "Success",
                f"{Messages.SUCCESS} D365 files saved successfully!\n\nNext step: Go to 'Extract IDs' tab to generate ID lists for Redash.",
            )

            # Always go to Extract IDs tab in manual mode
            self.notebook.select(1)  # Switch to Extract IDs tab
            self.status_var.set("D365 files saved - Ready to extract IDs")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save files:\n{e}")

    def save_sc_files(self):
        """Copy SafeContractor files to input folder"""
        try:
            REDASH_DIR.mkdir(parents=True, exist_ok=True)

            for key in ["accreditation_sc", "wcb_sc", "client_sc"]:
                if self.uploaded_files[key]:
                    source = Path(self.uploaded_files[key])

                    # Validate source exists
                    if not source.exists():
                        raise FileNotFoundError(f"Source file not found: {source}")

                    # Strip _sc suffix to get the report type
                    report_type = key.replace("_sc", "")
                    dest = REDASH_DIR / SC_FILES[report_type]

                    # Ensure destination path is valid
                    dest = dest.resolve()

                    shutil.copy2(str(source), str(dest))

            messagebox.showinfo(
                "Success",
                f"{Messages.SUCCESS} SC files saved successfully!\n\nNext step: Go to 'Generate Comparisons' tab to create comparison reports.",
            )
            self.notebook.select(3)  # Switch to Generate Comparisons tab
            self.status_var.set("SC files saved - Ready to generate comparisons")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save files:\n{e}")

    def extract_ids(self):
        """Run ID extraction in background thread"""
        self.btn_extract.config(state=tk.DISABLED, text="⏳ Extracting...")
        self.extract_console.delete(1.0, tk.END)
        self.status_var.set("Extracting IDs...")

        def run_extraction():
            # Redirect stdout to console
            import io
            import contextlib

            f = io.StringIO()
            with contextlib.redirect_stdout(f):
                try:
                    extract_and_save_ids()
                    output = f.getvalue()
                except Exception as e:
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

        if "Error" in output or "❌" in output:
            self.status_var.set("ID extraction failed - Check console for errors")
            messagebox.showerror(
                "Error", "ID extraction failed. Check the console output for details."
            )
        else:
            self.status_var.set("IDs extracted successfully!")
            messagebox.showinfo(
                "Success",
                f"{Messages.SUCCESS} IDs extracted successfully!\n\n"
                "Next steps:\n"
                "1. Open the output folder\n"
                "2. Copy IDs from .sql.txt files\n"
                "3. Run Redash queries\n"
                "4. Upload SC results in the next tab",
            )
            self.notebook.select(2)  # Switch to Upload SC tab

    def generate_comparison(self):
        """Run comparison generation in background thread"""
        # Check if SC files exist
        from config import SC_PATTERNS, REPORT_TYPES
        from utils import find_file_by_pattern

        sc_files_exist = any(
            find_file_by_pattern(REDASH_DIR, SC_PATTERNS[t]) is not None for t in REPORT_TYPES
        )

        if not sc_files_exist:
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
        self.compare_console.delete(1.0, tk.END)
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
                        self.console.insert(tk.END, text)
                        self.console.see(tk.END)
                        self.console.update_idletasks()

                def flush(self):
                    pass

                def getvalue(self):
                    return self.buffer.getvalue()

            stream = StreamToConsole(self.compare_console, self.root)

            try:
                old_stdout = sys.stdout
                sys.stdout = stream
                generate_comparisons()
                output = stream.getvalue()
            except Exception as e:
                output = f"Error: {e}\n{stream.getvalue()}"
            finally:
                sys.stdout = old_stdout

            # Update UI in main thread
            self.root.after(0, self.comparison_complete, output)

        thread = threading.Thread(target=run_comparison, daemon=True)
        thread.start()

    def comparison_complete(self, output):
        """Handle comparison completion"""
        self.compare_console.insert(tk.END, output)
        self.compare_console.see(tk.END)
        self.btn_compare.config(state=tk.NORMAL, text="🚀 Generate Comparisons")

        if "Error" in output or "❌" in output:
            self.status_var.set("Comparison generation failed - Check console for errors")
            messagebox.showerror(
                "Error", "Comparison generation failed. Check the console output for details."
            )
        else:
            self.status_var.set("Comparisons generated successfully!")
            messagebox.showinfo(
                "Success",
                f"{Messages.SUCCESS} Comparison files generated successfully!\n\n"
                "The Excel files are ready in the output folder.\n"
                "Click 'Open Output Folder' to view them.",
            )

    def open_folder(self, folder_path):
        """Open folder in file explorer"""
        import os
        import subprocess

        folder_path.mkdir(exist_ok=True)

        if sys.platform == "win32":
            os.startfile(folder_path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder_path])
        else:
            subprocess.Popen(["xdg-open", folder_path])


def main():
    """Launch the GUI application"""
    root = TkinterDnD.Tk()  # Use TkinterDnD root instead of regular Tk
    app = ComparisonApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
