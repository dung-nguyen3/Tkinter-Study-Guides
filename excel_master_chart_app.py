#!/usr/bin/env python3
"""
Excel Master Chart Creator v2.5 - Desktop Application
A tksheet-based GUI application for creating formatted Excel master charts
with auto-color assignment, professional formatting, and multi-format export.

Version 2.5 Features:
- Excel-like grid interface with tksheet
- Right-click context menu
- Auto-save and crash recovery
- Live color preview
- 3-shade color system
- Two export formats: Master Chart (single sheet) and Comprehensive (4-tab)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import platform
import os
import json
import threading
import time
from pathlib import Path
from datetime import datetime
from tksheet import Sheet
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================================
# COLOR CONSTANTS - 3-Shade System for Professional Gradients
# ============================================================================

# 10-color sets with 3 shades each (header/main/row_label) - IN ORDER
# 1. Ice Blue - General topics
# 2. Seafoam - Normal findings
# 3. Light Orchid - Special topics
# 4. Champagne - Warnings/cautions
# 5. Sky Blue - Diagnostic
# 6. Pale Azure - Alternative blue
# 7. Blush Pink - Important alerts
# 8. Soft Lilac - Alternative purple
# 9. Soft Tangerine - Highlights
# 10. Powder Blue - Alternative blue
COLOR_SETS = [
    {'header': 'B3D4ED', 'main': 'E3F2FD', 'row_label': 'CBE7FA'},  # 1. Ice Blue - General topics
    {'header': 'A8CCA8', 'main': 'C8E6C9', 'row_label': 'B8D9B9'},  # 2. Seafoam - Normal findings
    {'header': 'B8A4D0', 'main': 'D1C4E9', 'row_label': 'C4B4DC'},  # 3. Light Orchid - Special topics
    {'header': 'E0D0B0', 'main': 'F7E7CE', 'row_label': 'EBDBBF'},  # 4. Champagne - Warnings/cautions
    {'header': '9DC3E6', 'main': 'BDD7EE', 'row_label': 'AECDEA'},  # 5. Sky Blue - Diagnostic
    {'header': 'D0E8FF', 'main': 'F0F8FF', 'row_label': 'E0F0FF'},  # 6. Pale Azure - Alternative blue
    {'header': 'E8C4CC', 'main': 'FCE4EC', 'row_label': 'F2D4DC'},  # 7. Blush Pink - Important alerts
    {'header': 'D0C8DC', 'main': 'EDE7F6', 'row_label': 'DED7E9'},  # 8. Soft Lilac - Alternative purple
    {'header': 'E0C8B0', 'main': 'FFE8D6', 'row_label': 'EFD8C3'},  # 9. Soft Tangerine - Highlights
    {'header': 'A0C4E8', 'main': 'BBDEFB', 'row_label': 'ADD1F1'},  # 10. Powder Blue - Alternative blue
]

# Special purpose colors
MNEMONIC_BG = 'E6F3FF'        # Light blue for mnemonics
CLINICAL_PEARL_BG = 'E8F5E9'   # Light green for clinical pearls
ANALOGY_BOX_BG = 'FFF9E6'      # Light yellow for analogies
MAIN_TITLE_COLOR = '4472C4'    # Dark blue for sheet titles

# Header formatting constants (backward compatible)
HEADER_BG_COLOR = "#4472C4"
HEADER_FONT_COLOR = "#FFFFFF"
DATA_FONT_COLOR = "#000000"
BORDER_COLOR = "#FFFFFF"

# ============================================================================
# COLUMN PRESETS
# ============================================================================
COLUMN_PRESETS = {
    "Drug Chart (11 columns)": [
        "Drug Class",
        "Drug Name (Brand)",
        "Route",
        "Mechanism",
        "Uses",
        "Adverse Effects",
        "Contraindications",
        "Resistance",
        "Drug Interactions",
        "Drug Combinations",
        "Special Considerations"
    ],
    "Condition Chart (7 columns)": [
        "Condition",
        "Epidemiology",
        "Risk Factors",
        "Clinical Presentation",
        "Diagnostics",
        "Labs",
        "Treatment"
    ],
    "Lab Values (5 columns)": [
        "Test Name",
        "Normal Range",
        "Increased In",
        "Decreased In",
        "Clinical Significance"
    ],
    "Custom": []
}

# Column width mapping
COLUMN_WIDTH_MAP = {
    "short": 12,
    "medium": 22,
    "name": 28,
    "long": 35
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_color_set(index):
    """Get color set by index (rotates through 10 sets)"""
    return COLOR_SETS[index % len(COLOR_SETS)]


def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple for tksheet"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


# ============================================================================
# MAIN APPLICATION CLASS
# ============================================================================
class ExcelMasterChartApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Master Chart Creator v2.5")
        self.root.geometry("1400x800")
        self.root.minsize(800, 600)  # Set minimum window size (reduced for better flexibility)

        # Data storage
        self.current_columns = []
        self.sheet = None
        self.current_preset = tk.StringVar(value="Drug Chart (11 columns)")
        self.output_filename = tk.StringVar(value="Master_Chart.xlsx")
        self.output_directory = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.export_format = tk.StringVar(value="master_chart")
        self.live_preview_var = tk.BooleanVar(value=False)

        # Auto-save settings
        self.autosave_path = Path.home() / ".excel_master_chart_autosave.json"
        self.autosave_thread = None
        self.autosave_running = True
        self.unsaved_changes = False
        self.last_save_time = None

        # Recent files
        self.recent_files_path = Path.home() / ".excel_master_chart_recent.json"
        self.recent_files = self.load_recent_files()

        # Setup UI
        self.create_menu_bar()
        self.setup_ui()
        self.load_preset()

        # Start auto-save thread
        self.start_autosave()

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Check for crash recovery
        self.check_crash_recovery()

    # ========================================================================
    # MENU BAR
    # ========================================================================

    def create_menu_bar(self):
        """Create application menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", command=self.new_file, accelerator="Cmd+N")
        file_menu.add_command(label="Open...", command=self.load_data_json, accelerator="Cmd+O")
        file_menu.add_command(label="Save", command=self.save_file, accelerator="Cmd+S")
        file_menu.add_command(label="Save As...", command=self.save_data_json)
        file_menu.add_separator()

        # Recent files submenu
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Recent Files", menu=self.recent_menu)
        self.update_recent_files_menu()

        file_menu.add_separator()
        file_menu.add_command(label="Export to Excel...", command=self.export_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Quit", command=self.on_closing, accelerator="Cmd+Q")

        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Undo", command=self.undo, accelerator="Cmd+Z")
        edit_menu.add_command(label="Redo", command=self.redo, accelerator="Cmd+Shift+Z")
        edit_menu.add_separator()
        edit_menu.add_command(label="Cut", accelerator="Cmd+X")
        edit_menu.add_command(label="Copy", accelerator="Cmd+C")
        edit_menu.add_command(label="Paste", accelerator="Cmd+V")
        edit_menu.add_separator()
        edit_menu.add_command(label="Select All", accelerator="Cmd+A")
        edit_menu.add_separator()
        edit_menu.add_command(label="Edit Column Headers...", command=self.edit_column_headers)
        edit_menu.add_command(label="Clear All Data", command=self.clear_all_data)

        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_checkbutton(label="Live Color Preview", variable=self.live_preview_var,
                                  command=self.toggle_color_preview)
        view_menu.add_command(label="Preview Colors", command=self.preview_colors)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Quick Start Guide", command=self.show_quick_start)
        help_menu.add_command(label="About", command=self.show_about)

    # ========================================================================
    # UI SETUP
    # ========================================================================

    def setup_ui(self):
        """Create the main user interface - Ribbon style"""
        # Configure root window
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=0)  # Ribbon - fixed
        self.root.rowconfigure(1, weight=1)  # Data grid - expandable
        self.root.rowconfigure(2, weight=0)  # Status bar - fixed

        # Create ribbon toolbar
        self.create_ribbon()

        # Create data grid (main area)
        self.create_data_grid_ribbon_style()

        # Create status bar at bottom
        self.create_status_bar_bottom()

    def create_ribbon(self):
        """Create Excel-style ribbon toolbar with tabs"""
        ribbon_container = ttk.Frame(self.root, relief=tk.RAISED, borderwidth=1)
        ribbon_container.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=0, pady=0)

        # Create notebook for tabbed ribbon
        self.ribbon_notebook = ttk.Notebook(ribbon_container)
        self.ribbon_notebook.pack(fill=tk.BOTH, expand=True)

        # HOME TAB
        home_tab = ttk.Frame(self.ribbon_notebook, padding="10")
        self.ribbon_notebook.add(home_tab, text="Home")

        # Column Setup Group
        col_group = ttk.LabelFrame(home_tab, text="Column Setup", padding="10")
        col_group.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        ttk.Label(col_group, text="Preset:").grid(row=0, column=0, sticky=tk.W, pady=2)
        preset_combo = ttk.Combobox(
            col_group,
            textvariable=self.current_preset,
            values=list(COLUMN_PRESETS.keys()),
            state="readonly",
            width=20
        )
        preset_combo.grid(row=0, column=1, columnspan=2, pady=2, padx=5)
        preset_combo.bind("<<ComboboxSelected>>", self.on_preset_change)

        ttk.Button(col_group, text="Custom Columns", command=self.define_custom_columns, width=15).grid(row=1, column=0, columnspan=2, pady=2, padx=2)
        ttk.Button(col_group, text="Edit Headers", command=self.edit_column_headers, width=15).grid(row=1, column=2, pady=2, padx=2)

        # Rows Group
        rows_group = ttk.LabelFrame(home_tab, text="Rows", padding="10")
        rows_group.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        ttk.Label(rows_group, text="Quick Add:").grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=2)
        ttk.Button(rows_group, text="+10", command=lambda: self.add_quick_rows(10), width=6).grid(row=1, column=0, padx=2, pady=2)
        ttk.Button(rows_group, text="+50", command=lambda: self.add_quick_rows(50), width=6).grid(row=1, column=1, padx=2, pady=2)
        ttk.Button(rows_group, text="+100", command=lambda: self.add_quick_rows(100), width=6).grid(row=1, column=2, padx=2, pady=2)

        ttk.Label(rows_group, text="Manage:").grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(10,2))
        ttk.Button(rows_group, text="Delete Empty", command=self.delete_empty_rows, width=12).grid(row=3, column=0, columnspan=2, pady=2, padx=2)
        ttk.Button(rows_group, text="Clear All", command=self.clear_all_data, width=12).grid(row=3, column=2, pady=2, padx=2)

        # DATA TAB
        data_tab = ttk.Frame(self.ribbon_notebook, padding="10")
        self.ribbon_notebook.add(data_tab, text="Data")

        # Save/Load Group
        file_group = ttk.LabelFrame(data_tab, text="File Operations", padding="10")
        file_group.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        ttk.Button(file_group, text="Save Data", command=self.save_data_json, width=15).grid(row=0, column=0, pady=5, padx=5)
        ttk.Button(file_group, text="Load Data", command=self.load_data_json, width=15).grid(row=1, column=0, pady=5, padx=5)

        # EXPORT TAB
        export_tab = ttk.Frame(self.ribbon_notebook, padding="10")
        self.ribbon_notebook.add(export_tab, text="Export")

        # Format Group
        format_group = ttk.LabelFrame(export_tab, text="Export Format", padding="10")
        format_group.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        ttk.Radiobutton(
            format_group,
            text="Master Chart\n(Single Sheet)",
            variable=self.export_format,
            value="master_chart"
        ).pack(anchor=tk.W, pady=5, padx=5)

        ttk.Radiobutton(
            format_group,
            text="Comprehensive\n(4 Tabs)",
            variable=self.export_format,
            value="comprehensive"
        ).pack(anchor=tk.W, pady=5, padx=5)

        # Export Action Group
        action_group = ttk.LabelFrame(export_tab, text="Action", padding="10")
        action_group.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        export_btn = ttk.Button(
            action_group,
            text="EXPORT TO EXCEL",
            command=self.export_to_excel,
            width=20
        )
        export_btn.pack(pady=20, padx=10)

    def create_data_grid_ribbon_style(self):
        """Create maximized data grid for ribbon interface"""
        # Main grid container
        grid_frame = ttk.Frame(self.root, padding="10")
        grid_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.rowconfigure(0, weight=1)

        # Sheet container (takes all available space)
        sheet_container = ttk.Frame(grid_frame)
        sheet_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        sheet_container.columnconfigure(0, weight=1)
        sheet_container.rowconfigure(0, weight=1)

        # This will be replaced when preset is loaded
        self.sheet_container = sheet_container

    def create_status_bar_bottom(self):
        """Create status bar at bottom of window"""
        status_frame = ttk.Frame(self.root, relief=tk.SUNKEN, borderwidth=1)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))

        # Left side: Status indicator
        self.status_label = ttk.Label(status_frame, text="● Ready", foreground="green")
        self.status_label.pack(side=tk.LEFT, padx=10)

        ttk.Label(status_frame, text="|").pack(side=tk.LEFT, padx=5)

        # Save status
        self.save_status_label = ttk.Label(status_frame, text="Not saved")
        self.save_status_label.pack(side=tk.LEFT, padx=5)

        ttk.Label(status_frame, text="|").pack(side=tk.LEFT, padx=5)

        # Row count
        self.row_count_label = ttk.Label(status_frame, text="Rows: 0/50")
        self.row_count_label.pack(side=tk.RIGHT, padx=10)

    def create_status_bar(self, parent):
        """Create status bar showing save status (DEPRECATED - kept for compatibility)"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        self.status_label = ttk.Label(status_frame, text="● Ready", foreground="green")
        self.status_label.pack(side=tk.LEFT)

        ttk.Label(status_frame, text=" | ").pack(side=tk.LEFT)

        self.save_status_label = ttk.Label(status_frame, text="Not saved")
        self.save_status_label.pack(side=tk.LEFT)

    def create_column_setup_section(self, parent):
        """Create Section 1: Column Setup"""
        setup_frame = ttk.LabelFrame(parent, text="Section 1: Column Setup", padding="10")
        setup_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        # Preset selector
        preset_frame = ttk.Frame(setup_frame)
        preset_frame.pack(fill=tk.X)

        ttk.Label(preset_frame, text="Select Preset:").pack(side=tk.LEFT, pady=5)
        preset_combo = ttk.Combobox(
            preset_frame,
            textvariable=self.current_preset,
            values=list(COLUMN_PRESETS.keys()),
            state="readonly",
            width=30
        )
        preset_combo.pack(side=tk.LEFT, padx=(10, 20), pady=5)
        preset_combo.bind("<<ComboboxSelected>>", self.on_preset_change)

        self.custom_btn = ttk.Button(
            preset_frame,
            text="Define Custom Columns...",
            command=self.define_custom_columns
        )
        self.custom_btn.pack(side=tk.LEFT, padx=10, pady=5)

        # Edit column headers button
        self.edit_headers_btn = ttk.Button(
            preset_frame,
            text="Edit Column Headers...",
            command=self.edit_column_headers
        )
        self.edit_headers_btn.pack(side=tk.LEFT, padx=10, pady=5)

    def create_data_grid_section(self, parent):
        """Create Section 2: Excel-Style Data Grid"""
        grid_frame = ttk.LabelFrame(parent, text="Section 2: Data Grid", padding="10")
        grid_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.rowconfigure(1, weight=1)

        # Control panel
        control_panel = ttk.Frame(grid_frame)
        control_panel.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        # Row management
        ttk.Label(control_panel, text="Add Rows:").pack(side=tk.LEFT, padx=5)
        self.row_count_var = tk.StringVar(value="50")
        ttk.Entry(control_panel, textvariable=self.row_count_var, width=8).pack(side=tk.LEFT)
        ttk.Button(control_panel, text="Add", command=self.add_custom_rows).pack(side=tk.LEFT, padx=5)

        ttk.Separator(control_panel, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Label(control_panel, text="Quick:").pack(side=tk.LEFT, padx=5)
        ttk.Button(control_panel, text="+10", command=lambda: self.add_quick_rows(10)).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_panel, text="+50", command=lambda: self.add_quick_rows(50)).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_panel, text="+100", command=lambda: self.add_quick_rows(100)).pack(side=tk.LEFT, padx=2)

        ttk.Separator(control_panel, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(control_panel, text="Delete Empty Rows", command=self.delete_empty_rows).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_panel, text="Clear All", command=self.clear_all_data).pack(side=tk.LEFT, padx=5)

        ttk.Separator(control_panel, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Checkbutton(control_panel, text="Live Color Preview", variable=self.live_preview_var,
                       command=self.toggle_color_preview).pack(side=tk.LEFT, padx=5)

        # Row count
        self.row_count_label = ttk.Label(control_panel, text="Rows: 0/50")
        self.row_count_label.pack(side=tk.RIGHT, padx=10)

        # Sheet container
        sheet_container = ttk.Frame(grid_frame)
        sheet_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        sheet_container.columnconfigure(0, weight=1)
        sheet_container.rowconfigure(0, weight=1)

        # This will be replaced when preset is loaded
        self.sheet_container = sheet_container

    def create_export_section(self, parent):
        """Create Section 3: Export Controls"""
        export_frame = ttk.LabelFrame(parent, text="Section 3: Export", padding="10")
        export_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        export_frame.columnconfigure(0, weight=1)

        # Export format selection
        format_frame = ttk.Frame(export_frame)
        format_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Radiobutton(
            format_frame,
            text="Excel Master Chart (Single Sheet)",
            variable=self.export_format,
            value="master_chart"
        ).pack(anchor=tk.W, padx=10, pady=2)

        ttk.Radiobutton(
            format_frame,
            text="Excel Comprehensive Chart (4 Tabs)",
            variable=self.export_format,
            value="comprehensive"
        ).pack(anchor=tk.W, padx=10, pady=2)

        # Action buttons
        button_frame = ttk.Frame(export_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="Save Data (JSON)", command=self.save_data_json).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Load Data (JSON)", command=self.load_data_json).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Preview Colors", command=self.preview_colors).pack(side=tk.LEFT, padx=5)

        # Export button (prominent)
        export_btn = ttk.Button(
            button_frame,
            text="Export to Excel",
            command=self.export_to_excel
        )
        export_btn.pack(side=tk.RIGHT, padx=5)

    # ========================================================================
    # TKSHEET SETUP
    # ========================================================================

    def setup_sheet(self):
        """Setup tksheet widget with current columns"""
        # Clear existing sheet if any
        if self.sheet:
            self.sheet.destroy()

        # Create initial data (50 empty rows)
        initial_data = [[""] * len(self.current_columns) for _ in range(50)]

        # Create sheet (no fixed height - will expand with window)
        self.sheet = Sheet(
            self.sheet_container,
            data=initial_data,
            headers=self.current_columns,
            theme="light blue",
            font=("Calibri", 11, "normal"),
            header_font=("Calibri", 11, "bold"),
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            empty_horizontal=0,
            empty_vertical=0
        )

        # Enable all bindings including header editing
        self.sheet.enable_bindings(
            "single_select", "drag_select", "column_width_resize",
            "double_click_column_resize", "row_height_resize",
            "column_select", "row_select", "edit_cell", "edit_index", "edit_header",
            "copy", "paste", "delete", "undo", "redo", "right_click_popup_menu"
        )

        self.sheet.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Bind right-click for context menu
        self.sheet.bind("<Button-2>", self.show_context_menu)  # macOS right-click
        self.sheet.bind("<Button-3>", self.show_context_menu)  # Windows/Linux right-click

        # Bind cell changes for unsaved indicator
        self.sheet.bind("<<SheetModified>>", self.on_sheet_modified)

        # Bind header changes to sync column names
        self.sheet.bind("<<SheetModified>>", self.sync_column_headers, add=True)

        # Update row count
        self.update_row_count()

    # ========================================================================
    # COLUMN MANAGEMENT
    # ========================================================================

    def load_preset(self):
        """Load the selected preset columns"""
        preset_name = self.current_preset.get()
        self.current_columns = COLUMN_PRESETS[preset_name].copy()
        self.update_columns_display()
        self.setup_sheet()

    def on_preset_change(self, event=None):
        """Handle preset selection change"""
        preset_name = self.current_preset.get()
        if preset_name == "Custom" and not COLUMN_PRESETS["Custom"]:
            self.define_custom_columns()
        else:
            self.load_preset()

    def define_custom_columns(self):
        """Open dialog to define custom columns"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Define Custom Columns")
        dialog.geometry("450x350")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Enter column names (one per line):").pack(pady=10, padx=10)

        text = tk.Text(dialog, height=15, width=50)
        text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        if COLUMN_PRESETS["Custom"]:
            text.insert("1.0", "\n".join(COLUMN_PRESETS["Custom"]))

        def save_custom():
            content = text.get("1.0", tk.END).strip()
            if content:
                columns = [line.strip() for line in content.split("\n") if line.strip()]
                if columns:
                    COLUMN_PRESETS["Custom"] = columns
                    self.current_preset.set("Custom")
                    self.load_preset()
                    dialog.destroy()
                else:
                    messagebox.showwarning("Invalid Input", "Please enter at least one column name.")
            else:
                messagebox.showwarning("Invalid Input", "Please enter at least one column name.")

        ttk.Button(dialog, text="Save", command=save_custom).pack(pady=10)

    def edit_column_headers(self):
        """Open dialog to edit current column headers"""
        if not self.current_columns:
            messagebox.showwarning("No Columns", "Please select a preset or define custom columns first.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Column Headers")
        dialog.geometry("450x400")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Edit column names (one per line):", font=("", 10, "bold")).pack(pady=10, padx=10)
        ttk.Label(dialog, text="Changes will be applied to the current data grid.", font=("", 9, "italic")).pack(padx=10)

        text = tk.Text(dialog, height=18, width=50)
        text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Pre-populate with current column names
        text.insert("1.0", "\n".join(self.current_columns))

        def save_edited_headers():
            content = text.get("1.0", tk.END).strip()
            if content:
                new_columns = [line.strip() for line in content.split("\n") if line.strip()]
                if new_columns:
                    # Check if number of columns changed
                    if len(new_columns) != len(self.current_columns):
                        result = messagebox.askyesno(
                            "Column Count Changed",
                            f"You changed the number of columns from {len(self.current_columns)} to {len(new_columns)}.\n\n"
                            "This may result in data loss if you reduced the number of columns.\n\n"
                            "Do you want to continue?"
                        )
                        if not result:
                            return

                    # Get current data
                    current_data = self.sheet.get_sheet_data() if self.sheet else []

                    # Update columns
                    old_col_count = len(self.current_columns)
                    self.current_columns = new_columns
                    self.update_columns_display()

                    # Adjust data to match new column count
                    if len(new_columns) > old_col_count:
                        # Add empty columns
                        current_data = [row + [""] * (len(new_columns) - len(row)) for row in current_data]
                    elif len(new_columns) < old_col_count:
                        # Truncate columns
                        current_data = [row[:len(new_columns)] for row in current_data]

                    # Recreate sheet with new headers
                    self.setup_sheet()

                    # Restore data
                    if current_data:
                        self.sheet.set_sheet_data(current_data)

                    self.mark_unsaved()
                    dialog.destroy()
                    messagebox.showinfo("Success", "Column headers updated successfully!")
                else:
                    messagebox.showwarning("Invalid Input", "Please enter at least one column name.")
            else:
                messagebox.showwarning("Invalid Input", "Please enter at least one column name.")

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Save Changes", command=save_edited_headers).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def update_columns_display(self):
        """Update the column display (no-op since we removed the display widget)"""
        pass

    def sync_column_headers(self, event=None):
        """Sync column headers from sheet to internal list"""
        try:
            headers = self.sheet.headers()
            if headers:
                self.current_columns = list(headers)
        except:
            pass

    # ========================================================================
    # ROW MANAGEMENT
    # ========================================================================

    def add_custom_rows(self):
        """Add user-specified number of rows"""
        try:
            count = int(self.row_count_var.get())
            if count > 0:
                self.add_quick_rows(count)
            else:
                messagebox.showwarning("Invalid Input", "Please enter a positive number.")
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid number.")

    def add_quick_rows(self, count):
        """Add specified number of empty rows"""
        current_data = self.sheet.get_sheet_data()
        new_rows = [[""] * len(self.current_columns) for _ in range(count)]
        self.sheet.set_sheet_data(current_data + new_rows)
        self.update_row_count()
        self.mark_unsaved()

    def delete_empty_rows(self):
        """Delete all rows that are completely empty"""
        data = self.sheet.get_sheet_data()
        non_empty = [row for row in data if any(cell.strip() if isinstance(cell, str) else cell for cell in row)]

        if len(non_empty) < len(data):
            self.sheet.set_sheet_data(non_empty if non_empty else [[""] * len(self.current_columns)])
            self.update_row_count()
            self.mark_unsaved()
            messagebox.showinfo("Rows Deleted", f"Deleted {len(data) - len(non_empty)} empty rows.")
        else:
            messagebox.showinfo("No Empty Rows", "There are no empty rows to delete.")

    def clear_all_data(self):
        """Clear all data from the sheet"""
        if messagebox.askyesno("Clear All Data", "Are you sure you want to clear all data? This cannot be undone."):
            self.sheet.set_sheet_data([[""] * len(self.current_columns) for _ in range(50)])
            self.update_row_count()
            self.mark_unsaved()

    def update_row_count(self):
        """Update the row count label"""
        data = self.sheet.get_sheet_data()
        non_empty = sum(1 for row in data if any(cell.strip() if isinstance(cell, str) else cell for cell in row))
        self.row_count_label.config(text=f"Rows with data: {non_empty}/{len(data)}")

    # ========================================================================
    # CONTEXT MENU
    # ========================================================================

    def show_context_menu(self, event):
        """Show right-click context menu"""
        menu = tk.Menu(self.root, tearoff=0)

        menu.add_command(label="Cut", accelerator="Cmd+X", command=lambda: self.sheet.cut())
        menu.add_command(label="Copy", accelerator="Cmd+C", command=lambda: self.sheet.copy())
        menu.add_command(label="Paste", accelerator="Cmd+V", command=lambda: self.sheet.paste())
        menu.add_separator()

        menu.add_command(label="Insert Row Above", command=self.insert_row_above)
        menu.add_command(label="Insert Row Below", command=self.insert_row_below)
        menu.add_command(label="Insert 10 Rows Below", command=lambda: self.insert_rows_below(10))
        menu.add_separator()

        menu.add_command(label="Insert Column Before", command=self.insert_column_before)
        menu.add_command(label="Insert Column After", command=self.insert_column_after)
        menu.add_separator()

        menu.add_command(label="Delete Selected Row(s)", command=self.delete_selected_rows)
        menu.add_command(label="Delete Selected Column(s)", command=self.delete_selected_columns)
        menu.add_command(label="Delete All Empty Rows", command=self.delete_empty_rows)
        menu.add_separator()

        menu.add_command(label="Clear Selected Cells", command=lambda: self.sheet.delete_key())
        menu.add_command(label="Fill Down", command=self.fill_down)
        menu.add_command(label="Fill Right", command=self.fill_right)

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def insert_row_above(self):
        """Insert row above selected"""
        selected = self.sheet.get_currently_selected()
        if selected:
            row = selected.row
            data = self.sheet.get_sheet_data()
            data.insert(row, [""] * len(self.current_columns))
            self.sheet.set_sheet_data(data)
            self.update_row_count()
            self.mark_unsaved()

    def insert_row_below(self):
        """Insert row below selected"""
        selected = self.sheet.get_currently_selected()
        if selected:
            row = selected.row + 1
            data = self.sheet.get_sheet_data()
            data.insert(row, [""] * len(self.current_columns))
            self.sheet.set_sheet_data(data)
            self.update_row_count()
            self.mark_unsaved()
        else:
            # Insert at end if nothing selected
            self.add_quick_rows(1)

    def insert_rows_below(self, count):
        """Insert multiple rows below selected"""
        selected = self.sheet.get_currently_selected()
        if selected:
            row = selected.row + 1
            data = self.sheet.get_sheet_data()
            for _ in range(count):
                data.insert(row, [""] * len(self.current_columns))
            self.sheet.set_sheet_data(data)
            self.update_row_count()
            self.mark_unsaved()
        else:
            self.add_quick_rows(count)

    def delete_selected_rows(self):
        """Delete currently selected rows"""
        selected = self.sheet.get_all_selection_boxes()
        if selected:
            data = self.sheet.get_sheet_data()
            rows_to_delete = set()

            for box in selected:
                for row in range(box.from_r, box.upto_r):
                    rows_to_delete.add(row)

            # Delete rows in reverse order
            for row in sorted(rows_to_delete, reverse=True):
                if row < len(data):
                    del data[row]

            self.sheet.set_sheet_data(data if data else [[""] * len(self.current_columns)])
            self.update_row_count()
            self.mark_unsaved()

    def fill_down(self):
        """Fill down from selected cell"""
        selected = self.sheet.get_currently_selected()
        if selected and selected.row is not None and selected.column is not None:
            data = self.sheet.get_sheet_data()
            value = data[selected.row][selected.column]

            # Fill down to next 10 rows or end
            for i in range(selected.row + 1, min(selected.row + 11, len(data))):
                data[i][selected.column] = value

            self.sheet.set_sheet_data(data)
            self.mark_unsaved()

    def fill_right(self):
        """Fill right from selected cell"""
        selected = self.sheet.get_currently_selected()
        if selected and selected.row is not None and selected.column is not None:
            data = self.sheet.get_sheet_data()
            value = data[selected.row][selected.column]

            # Fill right to end of row
            for j in range(selected.column + 1, len(self.current_columns)):
                data[selected.row][j] = value

            self.sheet.set_sheet_data(data)
            self.mark_unsaved()

    def insert_column_before(self):
        """Insert column before selected column"""
        selected = self.sheet.get_currently_selected()
        if selected and selected.column is not None:
            col_idx = selected.column

            # Insert into column headers
            new_col_name = f"Column {len(self.current_columns) + 1}"
            self.current_columns.insert(col_idx, new_col_name)
            self.update_columns_display()

            # Insert into data
            data = self.sheet.get_sheet_data()
            for row in data:
                row.insert(col_idx, "")

            # Recreate sheet
            self.setup_sheet()
            self.sheet.set_sheet_data(data)
            self.mark_unsaved()
        else:
            messagebox.showinfo("No Column Selected", "Please select a column first.")

    def insert_column_after(self):
        """Insert column after selected column"""
        selected = self.sheet.get_currently_selected()
        if selected and selected.column is not None:
            col_idx = selected.column + 1

            # Insert into column headers
            new_col_name = f"Column {len(self.current_columns) + 1}"
            self.current_columns.insert(col_idx, new_col_name)
            self.update_columns_display()

            # Insert into data
            data = self.sheet.get_sheet_data()
            for row in data:
                row.insert(col_idx, "")

            # Recreate sheet
            self.setup_sheet()
            self.sheet.set_sheet_data(data)
            self.mark_unsaved()
        else:
            messagebox.showinfo("No Column Selected", "Please select a column first.")

    def delete_selected_columns(self):
        """Delete currently selected columns"""
        selected = self.sheet.get_all_selection_boxes()
        if selected:
            if len(self.current_columns) <= 1:
                messagebox.showwarning("Cannot Delete", "Cannot delete all columns. At least one column must remain.")
                return

            columns_to_delete = set()

            for box in selected:
                for col in range(box.from_c, box.upto_c):
                    columns_to_delete.add(col)

            if len(columns_to_delete) >= len(self.current_columns):
                messagebox.showwarning("Cannot Delete", "Cannot delete all columns. At least one column must remain.")
                return

            # Confirm deletion
            if not messagebox.askyesno("Delete Columns",
                                      f"Are you sure you want to delete {len(columns_to_delete)} column(s)?"):
                return

            # Delete columns from headers
            for col_idx in sorted(columns_to_delete, reverse=True):
                if col_idx < len(self.current_columns):
                    del self.current_columns[col_idx]

            self.update_columns_display()

            # Delete columns from data
            data = self.sheet.get_sheet_data()
            for row in data:
                for col_idx in sorted(columns_to_delete, reverse=True):
                    if col_idx < len(row):
                        del row[col_idx]

            # Recreate sheet
            self.setup_sheet()
            self.sheet.set_sheet_data(data)
            self.mark_unsaved()
        else:
            messagebox.showinfo("No Selection", "Please select column(s) to delete.")

    # ========================================================================
    # UNDO/REDO
    # ========================================================================

    def undo(self):
        """Undo last action"""
        try:
            self.sheet.undo()
        except:
            pass

    def redo(self):
        """Redo last undone action"""
        try:
            self.sheet.redo()
        except:
            pass

    # ========================================================================
    # COLOR PREVIEW
    # ========================================================================

    def toggle_color_preview(self):
        """Toggle live color preview on/off"""
        if self.live_preview_var.get():
            self.apply_live_colors()
        else:
            self.clear_live_colors()

    def apply_live_colors(self):
        """Apply live color preview to rows"""
        data = self.sheet.get_sheet_data()
        color_map = self.calculate_color_assignments_from_data(data)

        # Clear existing highlights first
        try:
            self.sheet.dehighlight_all()
        except:
            pass

        # Apply colors to ALL rows in each drug class
        last_group = None
        current_color = None

        for row_idx, row_data in enumerate(data):
            if row_data and row_data[0]:  # Has value in first column
                group_name = str(row_data[0]).strip()
                if group_name:
                    # Update current color when we encounter a new group
                    if group_name != last_group:
                        last_group = group_name
                        current_color = color_map.get(group_name)

                    # Apply color to this row
                    if current_color:
                        try:
                            self.sheet.highlight_rows([row_idx], bg=f"#{current_color['main']}")
                        except:
                            pass

    def clear_live_colors(self):
        """Clear live color preview"""
        try:
            self.sheet.dehighlight_all()
        except:
            pass

    def preview_colors(self):
        """Show color assignment preview dialog"""
        data = self.sheet.get_sheet_data()
        non_empty = [row for row in data if row and row[0]]

        if not non_empty:
            messagebox.showinfo("No Data", "Please add some data first to preview colors.")
            return

        color_map = self.calculate_color_assignments_from_data(data)

        # Create preview window
        preview = tk.Toplevel(self.root)
        preview.title("Color Preview")
        preview.geometry("500x600")

        canvas = tk.Canvas(preview)
        scrollbar = ttk.Scrollbar(preview, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Display color assignments
        ttk.Label(scrollable_frame, text="Color Assignments:", font=("", 14, "bold")).pack(pady=10)
        ttk.Label(scrollable_frame, text="(Colors as they will appear in Excel)", font=("", 10, "italic")).pack(pady=5)

        for group_name, color_set in color_map.items():
            frame = tk.Frame(scrollable_frame, bg=f"#{color_set['main']}", relief=tk.RAISED, borderwidth=2)
            frame.pack(fill=tk.X, padx=20, pady=3)

            label = tk.Label(frame, text=f"  {group_name}  ", bg=f"#{color_set['main']}",
                           font=("Calibri", 11, "bold"), anchor=tk.W, padx=15, pady=8)
            label.pack(fill=tk.X)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def calculate_color_assignments_from_data(self, data):
        """Calculate color assignments for data"""
        color_map = {}
        color_index = 0
        seen_groups = []

        for row in data:
            if row and row[0]:
                group_name = str(row[0]).strip()
                if group_name and group_name not in seen_groups:
                    color_map[group_name] = get_color_set(color_index)
                    seen_groups.append(group_name)
                    color_index += 1

        return color_map

    # ========================================================================
    # AUTO-SAVE & FILE MANAGEMENT
    # ========================================================================

    def start_autosave(self):
        """Start background auto-save thread"""
        self.autosave_thread = threading.Thread(target=self.autosave_loop, daemon=True)
        self.autosave_thread.start()

    def autosave_loop(self):
        """Auto-save every 2 minutes"""
        while self.autosave_running:
            time.sleep(120)  # 2 minutes
            if self.unsaved_changes:
                try:
                    self.save_to_json(self.autosave_path, silent=True)
                    self.last_save_time = datetime.now()
                except:
                    pass

    def mark_unsaved(self):
        """Mark document as having unsaved changes"""
        self.unsaved_changes = True
        self.status_label.config(text="● Unsaved changes", foreground="orange")

    def mark_saved(self):
        """Mark document as saved"""
        self.unsaved_changes = False
        self.status_label.config(text="● Saved", foreground="green")
        now = datetime.now().strftime("%I:%M %p")
        self.save_status_label.config(text=f"Last saved: {now}")

    def on_sheet_modified(self, event=None):
        """Handle sheet modifications"""
        self.mark_unsaved()
        self.update_row_count()

        # Always auto-apply colors after cell modification (real-time color update)
        self.apply_live_colors()

    def check_crash_recovery(self):
        """Check for auto-save file and offer recovery"""
        if self.autosave_path.exists():
            if messagebox.askyesno("Crash Recovery",
                                  "Found an auto-saved file from a previous session.\nWould you like to recover it?"):
                try:
                    self.load_from_json(self.autosave_path)
                    messagebox.showinfo("Recovery Successful", "Your previous work has been recovered.")
                except:
                    messagebox.showerror("Recovery Failed", "Could not recover the auto-saved file.")
            else:
                # Delete auto-save file
                self.autosave_path.unlink()

    def new_file(self):
        """Create new file"""
        if self.unsaved_changes:
            if not messagebox.askyesno("Unsaved Changes", "You have unsaved changes. Create new file anyway?"):
                return

        self.sheet.set_sheet_data([[""] * len(self.current_columns) for _ in range(50)])
        self.update_row_count()
        self.unsaved_changes = False
        self.mark_saved()

    def save_file(self):
        """Quick save to last location"""
        # For now, just save as JSON
        self.save_data_json()

    def browse_directory(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(initialdir=self.output_directory.get())
        if directory:
            self.output_directory.set(directory)

    def save_data_json(self):
        """Save current data to JSON file"""
        data = self.sheet.get_sheet_data()
        non_empty = [row for row in data if any(cell.strip() if isinstance(cell, str) else cell for cell in row)]

        if not non_empty:
            messagebox.showinfo("No Data", "There is no data to save.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=self.output_directory.get()
        )

        if filename:
            self.save_to_json(Path(filename))
            self.add_to_recent_files(filename)
            messagebox.showinfo("Success", f"Data saved to:\n{filename}")

    def save_to_json(self, filepath, silent=False):
        """Save data to specified JSON file"""
        data = self.sheet.get_sheet_data()

        save_data = {
            "preset": self.current_preset.get(),
            "columns": self.current_columns,
            "rows": data
        }

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(save_data, f, indent=2, ensure_ascii=False)

        if not silent:
            self.mark_saved()

    def load_data_json(self):
        """Load data from JSON file"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=self.output_directory.get()
        )

        if filename:
            try:
                self.load_from_json(Path(filename))
                self.add_to_recent_files(filename)
                messagebox.showinfo("Success", f"Data loaded from:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load data:\n{str(e)}")

    def load_from_json(self, filepath):
        """Load data from specified JSON file"""
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Load preset
        preset = data.get("preset", "Custom")
        if preset == "Custom":
            COLUMN_PRESETS["Custom"] = data["columns"]

        self.current_preset.set(preset)
        self.current_columns = data["columns"]

        # Update UI
        self.update_columns_display()
        self.setup_sheet()

        # Load data
        self.sheet.set_sheet_data(data["rows"])
        self.update_row_count()
        self.mark_saved()

    def load_recent_files(self):
        """Load recent files list"""
        if self.recent_files_path.exists():
            try:
                with open(self.recent_files_path, 'r') as f:
                    return json.load(f)
            except:
                return []
        return []

    def save_recent_files(self):
        """Save recent files list"""
        with open(self.recent_files_path, 'w') as f:
            json.dump(self.recent_files, f)

    def add_to_recent_files(self, filepath):
        """Add file to recent files list"""
        filepath = str(Path(filepath).resolve())
        if filepath in self.recent_files:
            self.recent_files.remove(filepath)
        self.recent_files.insert(0, filepath)
        self.recent_files = self.recent_files[:10]  # Keep only 10 most recent
        self.save_recent_files()
        self.update_recent_files_menu()

    def update_recent_files_menu(self):
        """Update recent files menu"""
        self.recent_menu.delete(0, tk.END)

        if self.recent_files:
            for filepath in self.recent_files:
                filename = Path(filepath).name
                self.recent_menu.add_command(
                    label=filename,
                    command=lambda f=filepath: self.load_from_json(Path(f))
                )
        else:
            self.recent_menu.add_command(label="(No recent files)", state=tk.DISABLED)

    # ========================================================================
    # EXPORT TO EXCEL
    # ========================================================================

    def export_to_excel(self):
        """Export data to Excel (route to appropriate format)"""
        format_type = self.export_format.get()

        if format_type == "master_chart":
            self.export_to_excel_master_chart()
        elif format_type == "comprehensive":
            self.export_to_excel_comprehensive()

    def export_to_excel_master_chart(self):
        """Export to Excel Master Chart (single sheet with enhanced colors)"""
        # Get data
        data = self.sheet.get_sheet_data()
        non_empty = [row for row in data if any(cell.strip() if isinstance(cell, str) else cell for cell in row)]

        # Validation
        if not non_empty:
            messagebox.showwarning("No Data", "Please add some data before exporting.")
            return

        # Show file save dialog
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self.output_directory.get(),
            initialfile=self.output_filename.get()
        )

        if not output_path:
            return  # User cancelled

        output_path = Path(output_path)

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Master Chart"

            # Calculate color assignments
            color_map = self.calculate_color_assignments_from_data(non_empty)

            # Header row
            ws.append(self.current_columns)

            header_font = Font(name="Calibri", size=12, bold=True, color=HEADER_FONT_COLOR[1:])
            header_fill = PatternFill(start_color=HEADER_BG_COLOR[1:], end_color=HEADER_BG_COLOR[1:], fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for col_idx in range(1, len(self.current_columns) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment

            ws.row_dimensions[1].height = 25

            # Data rows with auto-coloring
            current_color_set = None
            last_group = None

            for row_idx, row_data in enumerate(non_empty, 2):
                ws.append(row_data)

                # Determine color
                group_name = str(row_data[0]).strip() if row_data[0] else ""
                if group_name and group_name != last_group:
                    current_color_set = color_map.get(group_name)
                    last_group = group_name

                if current_color_set:
                    # Apply 3-shade system
                    data_fill = PatternFill(
                        start_color=current_color_set['main'],
                        end_color=current_color_set['main'],
                        fill_type="solid"
                    )

                    data_alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                    white_border = Border(
                        left=Side(style="thin", color=BORDER_COLOR[1:]),
                        right=Side(style="thin", color=BORDER_COLOR[1:]),
                        top=Side(style="thin", color=BORDER_COLOR[1:]),
                        bottom=Side(style="thin", color=BORDER_COLOR[1:])
                    )

                    for col_idx in range(1, len(self.current_columns) + 1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.fill = data_fill
                        cell.alignment = data_alignment
                        cell.border = white_border

                        # First column bold and size 12
                        if col_idx == 1:
                            cell.font = Font(name="Calibri", size=12, bold=True, color=DATA_FONT_COLOR[1:])
                        else:
                            cell.font = Font(name="Calibri", size=10, color=DATA_FONT_COLOR[1:])

            # Column widths
            for col_idx, col_name in enumerate(self.current_columns, 1):
                width = self.get_excel_column_width(col_name)
                ws.column_dimensions[get_column_letter(col_idx)].width = width

            # Freeze panes
            ws.freeze_panes = "A2"

            # Save
            wb.save(output_path)

            messagebox.showinfo(
                "Export Successful",
                f"Excel Master Chart created successfully!\n\nLocation:\n{output_path}\n\nRows exported: {len(non_empty)}"
            )

            self.open_file(output_path)
            self.mark_saved()

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export Excel file:\n\n{str(e)}")

    def export_to_excel_comprehensive(self):
        """Export to Excel Comprehensive Drug Chart (4 tabs)"""
        # Get data
        data = self.sheet.get_sheet_data()
        non_empty = [row for row in data if any(cell.strip() if isinstance(cell, str) else cell for cell in row)]

        # Validation
        if not non_empty:
            messagebox.showwarning("No Data", "Please add some data before exporting.")
            return

        # Show file save dialog
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self.output_directory.get(),
            initialfile=self.output_filename.get()
        )

        if not output_path:
            return  # User cancelled

        output_path = Path(output_path)

        try:
            # Calculate color assignments
            color_map = self.calculate_color_assignments_from_data(non_empty)

            # Group by drug class
            grouped_data = self.group_by_drug_class(non_empty)

            # Create workbook
            wb = Workbook()
            wb.remove(wb.active)  # Remove default sheet

            # Create 4 tabs
            self.create_drug_details_tab(wb, grouped_data, color_map)
            self.create_key_comparisons_tab(wb, grouped_data, color_map)
            self.create_master_chart_tab(wb, non_empty, color_map)
            self.create_high_yield_tab(wb, grouped_data, color_map)

            # Save
            wb.save(output_path)

            messagebox.showinfo(
                "Export Successful",
                f"Excel Comprehensive Drug Chart created!\n\nLocation:\n{output_path}\n\n4 tabs created:\n- Drug Details\n- Key Comparisons\n- Master Chart\n- High-Yield & Pearls"
            )

            self.open_file(output_path)
            self.mark_saved()

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export comprehensive chart:\n\n{str(e)}")

    def group_by_drug_class(self, data):
        """Group rows by first column (Drug Class)"""
        grouped = {}
        for row in data:
            if row and row[0]:
                drug_class = str(row[0]).strip()
                if drug_class not in grouped:
                    grouped[drug_class] = []
                grouped[drug_class].append(row[1:])  # Skip drug class column
        return grouped

    def create_drug_details_tab(self, wb, grouped_data, color_map):
        """Create Tab 1: Drug Details (transposed tables by class)"""
        ws = wb.create_sheet("Drug Details")

        current_row = 1

        for drug_class, drugs in grouped_data.items():
            color_set = color_map.get(drug_class, get_color_set(0))
            num_drugs = len(drugs)

            # Drug class header (merged)
            end_col = get_column_letter(num_drugs + 1)
            ws.merge_cells(f'A{current_row}:{end_col}{current_row}')

            cell = ws[f'A{current_row}']
            cell.value = drug_class.upper()
            cell.font = Font(bold=True, size=16, color='FFFFFF')
            cell.fill = PatternFill(start_color=color_set['header'], end_color=color_set['header'], fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin', color='FFFFFF'),
                right=Side(style='thin', color='FFFFFF'),
                top=Side(style='thin', color='FFFFFF'),
                bottom=Side(style='thin', color='FFFFFF')
            )
            ws.row_dimensions[current_row].height = 25

            current_row += 1

            # Drug names as column headers
            ws['A' + str(current_row)] = ""  # Empty corner cell
            ws['A' + str(current_row)].fill = PatternFill(start_color=color_set['row_label'],
                                                           end_color=color_set['row_label'], fill_type='solid')
            ws['A' + str(current_row)].border = Border(
                left=Side(style='thin', color='FFFFFF'),
                right=Side(style='thin', color='FFFFFF'),
                top=Side(style='thin', color='FFFFFF'),
                bottom=Side(style='thin', color='FFFFFF')
            )

            for idx, drug in enumerate(drugs, 1):
                col_letter = get_column_letter(idx + 1)
                cell = ws[f'{col_letter}{current_row}']
                cell.value = drug[0] if drug else ""  # Drug name
                cell.font = Font(bold=True, size=14, color='000000')
                cell.fill = PatternFill(start_color=color_set['main'], end_color=color_set['main'], fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = Border(
                    left=Side(style='thin', color='FFFFFF'),
                    right=Side(style='thin', color='FFFFFF'),
                    top=Side(style='thin', color='FFFFFF'),
                    bottom=Side(style='thin', color='FFFFFF')
                )
                ws.column_dimensions[col_letter].width = 25

            current_row += 1

            # Properties as rows (starting from column 2 since we skip Drug Name)
            if len(self.current_columns) > 2:
                properties = self.current_columns[2:]  # Skip Drug Class and Drug Name

                for prop_idx, prop_name in enumerate(properties):
                    cell = ws[f'A{current_row}']
                    cell.value = prop_name
                    cell.font = Font(bold=True, size=11, color='000000')
                    cell.fill = PatternFill(start_color=color_set['row_label'], end_color=color_set['row_label'], fill_type='solid')
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                    cell.border = Border(
                        left=Side(style='thin', color='FFFFFF'),
                        right=Side(style='thin', color='FFFFFF'),
                        top=Side(style='thin', color='FFFFFF'),
                        bottom=Side(style='thin', color='FFFFFF')
                    )

                    # Data for each drug
                    for drug_idx, drug in enumerate(drugs, 1):
                        col_letter = get_column_letter(drug_idx + 1)
                        cell = ws[f'{col_letter}{current_row}']

                        # Get data (offset by 1 since we removed drug class, +prop_idx for property)
                        if len(drug) > prop_idx + 1:
                            cell.value = drug[prop_idx + 1]
                        else:
                            cell.value = ""

                        cell.font = Font(size=10, color='000000')
                        cell.fill = PatternFill(start_color=color_set['main'], end_color=color_set['main'], fill_type='solid')
                        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                        cell.border = Border(
                            left=Side(style='thin', color='FFFFFF'),
                            right=Side(style='thin', color='FFFFFF'),
                            top=Side(style='thin', color='FFFFFF'),
                            bottom=Side(style='thin', color='FFFFFF')
                        )

                    ws.row_dimensions[current_row].height = 30
                    current_row += 1

            # Mnemonic row
            ws.merge_cells(f'A{current_row}:{end_col}{current_row}')
            cell = ws[f'A{current_row}']
            cell.value = f"MEMORY TRICKS & MNEMONICS\n[Add {drug_class} mnemonics here]"
            cell.font = Font(size=10, italic=True, color='000000')
            cell.fill = PatternFill(start_color=MNEMONIC_BG, end_color=MNEMONIC_BG, fill_type='solid')
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            cell.border = Border(
                left=Side(style='thin', color='FFFFFF'),
                right=Side(style='thin', color='FFFFFF'),
                top=Side(style='thin', color='FFFFFF'),
                bottom=Side(style='thin', color='FFFFFF')
            )
            ws.row_dimensions[current_row].height = 40

            current_row += 3  # Blank rows

        # Set column A width
        ws.column_dimensions['A'].width = 25

    def create_key_comparisons_tab(self, wb, grouped_data, color_map):
        """Create Tab 2: Key Comparisons"""
        ws = wb.create_sheet("Key Comparisons")

        # Title
        ws.merge_cells('A1:F1')
        ws['A1'] = "KEY COMPARISONS ACROSS DRUG CLASSES"
        ws['A1'].font = Font(bold=True, size=16, color='FFFFFF')
        ws['A1'].fill = PatternFill(start_color=MAIN_TITLE_COLOR, end_color=MAIN_TITLE_COLOR, fill_type='solid')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        ws['A3'] = "Side-by-side comparison tables for key drug properties"
        ws['A3'].font = Font(size=11, italic=True)

        current_row = 5

        # Create comparison for each property
        if len(self.current_columns) > 2:
            for prop_idx in range(2, min(5, len(self.current_columns))):  # Compare first few properties
                prop_name = self.current_columns[prop_idx]

                # Property header
                ws.merge_cells(f'A{current_row}:F{current_row}')
                ws[f'A{current_row}'] = f"COMPARISON: {prop_name.upper()}"
                ws[f'A{current_row}'].font = Font(bold=True, size=13, color='FFFFFF')
                ws[f'A{current_row}'].fill = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
                ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
                ws.row_dimensions[current_row].height = 20

                current_row += 1

                # Create comparison table
                ws[f'A{current_row}'] = "Drug Class"
                ws[f'A{current_row}'].font = Font(bold=True, size=11)
                ws[f'A{current_row}'].fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')

                ws[f'B{current_row}'] = prop_name
                ws[f'B{current_row}'].font = Font(bold=True, size=11)
                ws[f'B{current_row}'].fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')
                ws.column_dimensions['B'].width = 60

                current_row += 1

                # Add each drug class
                for drug_class, drugs in list(grouped_data.items())[:10]:  # Limit to 10 classes
                    color_set = color_map.get(drug_class, get_color_set(0))

                    ws[f'A{current_row}'] = drug_class
                    ws[f'A{current_row}'].font = Font(bold=True, size=10)
                    ws[f'A{current_row}'].fill = PatternFill(start_color=color_set['row_label'],
                                                              end_color=color_set['row_label'], fill_type='solid')

                    # Aggregate property values from all drugs in class
                    values = []
                    for drug in drugs:
                        if len(drug) > prop_idx:
                            val = drug[prop_idx]
                            if val:
                                values.append(str(val))

                    ws[f'B{current_row}'] = "\n".join(values) if values else "(Not specified)"
                    ws[f'B{current_row}'].font = Font(size=10)
                    ws[f'B{current_row}'].fill = PatternFill(start_color=color_set['main'],
                                                              end_color=color_set['main'], fill_type='solid')
                    ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                    ws.row_dimensions[current_row].height = 40

                    current_row += 1

                current_row += 2  # Space between comparisons

        ws.column_dimensions['A'].width = 25

    def create_master_chart_tab(self, wb, data, color_map):
        """Create Tab 3: Master Chart (same as single-sheet export)"""
        ws = wb.create_sheet("Master Chart")

        # Header row
        ws.append(self.current_columns)

        header_font = Font(name="Calibri", size=12, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color=HEADER_BG_COLOR[1:], end_color=HEADER_BG_COLOR[1:], fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col_idx in range(1, len(self.current_columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        ws.row_dimensions[1].height = 25

        # Data rows
        last_group = None
        for row_idx, row_data in enumerate(data, 2):
            ws.append(row_data)

            group_name = str(row_data[0]).strip() if row_data[0] else ""
            if group_name and group_name != last_group:
                current_color_set = color_map.get(group_name)
                last_group = group_name

            if current_color_set:
                data_fill = PatternFill(start_color=current_color_set['main'],
                                       end_color=current_color_set['main'], fill_type="solid")

                for col_idx in range(1, len(self.current_columns) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.fill = data_fill
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

                    if col_idx == 1:
                        cell.font = Font(name="Calibri", size=10, bold=True, color='000000')
                    else:
                        cell.font = Font(name="Calibri", size=10, color='000000')

        # Column widths
        for col_idx, col_name in enumerate(self.current_columns, 1):
            width = self.get_excel_column_width(col_name)
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Freeze panes
        ws.freeze_panes = "A2"

    def create_high_yield_tab(self, wb, grouped_data, color_map):
        """Create Tab 4: High-Yield & Pearls"""
        ws = wb.create_sheet("High-Yield & Pearls")

        # Title
        ws.merge_cells('A1:F1')
        ws['A1'] = "HIGH-YIELD POINTS & CLINICAL PEARLS"
        ws['A1'].font = Font(bold=True, size=16, color='FFFFFF')
        ws['A1'].fill = PatternFill(start_color=MAIN_TITLE_COLOR, end_color=MAIN_TITLE_COLOR, fill_type='solid')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        current_row = 3

        for drug_class in grouped_data.keys():
            color_set = color_map.get(drug_class, get_color_set(0))

            # Class header
            ws.merge_cells(f'A{current_row}:F{current_row}')
            ws[f'A{current_row}'] = f"{drug_class.upper()} - KEY POINTS"
            ws[f'A{current_row}'].font = Font(bold=True, size=14, color='FFFFFF')
            ws[f'A{current_row}'].fill = PatternFill(start_color=color_set['header'],
                                                      end_color=color_set['header'], fill_type='solid')
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[current_row].height = 25

            current_row += 2

            # Clinical pearls box
            ws.merge_cells(f'A{current_row}:F{current_row + 2}')
            ws[f'A{current_row}'] = f"CLINICAL PEARLS - {drug_class}\n\n[Add key clinical tips and must-know facts here]"
            ws[f'A{current_row}'].font = Font(size=11, color='000000')
            ws[f'A{current_row}'].fill = PatternFill(start_color=CLINICAL_PEARL_BG,
                                                      end_color=CLINICAL_PEARL_BG, fill_type='solid')
            ws[f'A{current_row}'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            ws.row_dimensions[current_row].height = 60

            current_row += 3

            # Mnemonic box
            ws.merge_cells(f'A{current_row}:F{current_row + 1}')
            ws[f'A{current_row}'] = f"MNEMONICS\n\n[Add {drug_class} mnemonics here]"
            ws[f'A{current_row}'].font = Font(size=11, italic=True, color='000000')
            ws[f'A{current_row}'].fill = PatternFill(start_color=MNEMONIC_BG,
                                                      end_color=MNEMONIC_BG, fill_type='solid')
            ws[f'A{current_row}'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            ws.row_dimensions[current_row].height = 40

            current_row += 4  # Space before next class

        ws.column_dimensions['A'].width = 80

    def get_excel_column_width(self, col_name):
        """Get Excel column width based on column name"""
        col_lower = col_name.lower()

        if any(word in col_lower for word in ["route", "status", "type"]):
            return COLUMN_WIDTH_MAP["short"]
        elif any(word in col_lower for word in ["class", "condition", "category", "test"]):
            return COLUMN_WIDTH_MAP["medium"]
        elif any(word in col_lower for word in ["name", "brand"]):
            return COLUMN_WIDTH_MAP["name"]
        else:
            return COLUMN_WIDTH_MAP["long"]

    def open_file(self, filepath):
        """Open file with default application"""
        try:
            if platform.system() == "Darwin":
                subprocess.run(["open", str(filepath)], check=True)
            elif platform.system() == "Windows":
                os.startfile(str(filepath))
            else:
                subprocess.run(["xdg-open", str(filepath)], check=True)
        except Exception as e:
            print(f"Could not auto-open file: {e}")

    # ========================================================================
    # HELP & ABOUT
    # ========================================================================

    def show_quick_start(self):
        """Show quick start guide"""
        guide = """Excel Master Chart Creator v2.5 - Quick Start

1. SELECT PRESET
   Choose from Drug Chart, Condition Chart, Lab Values, or Custom

2. ENTER DATA
   - Click any cell to start typing
   - Press Tab to move right, Enter to move down
   - Right-click for more options

3. ADD ROWS
   - Use +10, +50, +100 buttons for quick adding
   - Or enter custom number and click Add

4. PREVIEW COLORS
   - Enable "Live Color Preview" to see colors as you type
   - Or click "Preview Colors" button

5. EXPORT
   - Choose format: Master Chart (1 sheet) or Comprehensive (4 tabs)
   - Click "Export to Excel"
   - File opens automatically!

TIP: Use Cmd+S to save your work as JSON
TIP: Auto-save runs every 2 minutes
TIP: Right-click for quick actions"""

        messagebox.showinfo("Quick Start Guide", guide)

    def show_about(self):
        """Show about dialog"""
        about = """Excel Master Chart Creator v2.5

A powerful desktop application for creating professional Excel charts
with auto-color assignment and multi-format export.

Features:
✓ Excel-like grid interface
✓ Auto-color by drug class
✓ Two export formats (Master Chart & Comprehensive)
✓ Auto-save and crash recovery
✓ Right-click context menu
✓ Live color preview

Created with Python, tkinter, tksheet, and openpyxl

© 2025"""

        messagebox.showinfo("About", about)

    # ========================================================================
    # WINDOW CLOSING
    # ========================================================================

    def on_closing(self):
        """Handle window close event"""
        if self.unsaved_changes:
            response = messagebox.askyesnocancel(
                "Unsaved Changes",
                "You have unsaved changes. Do you want to save before closing?"
            )

            if response is None:  # Cancel
                return
            elif response:  # Yes - save
                self.save_data_json()

        # Stop auto-save thread
        self.autosave_running = False

        # Delete auto-save file
        if self.autosave_path.exists():
            self.autosave_path.unlink()

        self.root.destroy()

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    root = tk.Tk()
    app = ExcelMasterChartApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
