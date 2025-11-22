#!/usr/bin/env python3
"""
Study Guide Creator v3.0 - Desktop Application
A tksheet-based GUI application for creating formatted Excel master charts
AND converting Markdown study guides to styled Word documents.

Version 3.0 Features:
- Excel-like grid interface with tksheet
- Right-click context menu
- Auto-save and crash recovery
- Live color preview
- 3-shade color system
- Two Excel export formats: Master Chart (single sheet) and Comprehensive (4-tab)
- Data validation dropdowns for common medical fields (Route, Contraindications, etc.)
- NEW: Markdown to Word conversion with styled tables and clinical pearls
- NEW: Word tab for importing .md files and exporting to .docx
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import platform
import os
import json
import csv
import re
import threading
import time
from pathlib import Path
from datetime import datetime
from tksheet import Sheet
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# python-docx for Word document generation
try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Twips
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    print("Warning: python-docx not installed. Word export will be disabled.")
    print("Install with: pip install python-docx")

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

# Word document color themes (from LO Word template)
WORD_COLOR_THEMES = {
    'Purple - General Topics': {
        'header': 'D1C4E9',
        'header_text': (74, 20, 140),
        'light': 'EDE7F6',
        'name': 'purple'
    },
    'Blue - Diagnostic': {
        'header': 'B3E5FC',
        'header_text': (1, 87, 155),
        'light': 'E1F5FE',
        'name': 'blue'
    },
    'Green - Normal Findings': {
        'header': 'C8E6C9',
        'header_text': (27, 94, 32),
        'light': 'E8F5E9',
        'name': 'green'
    },
    'Red - Pathology': {
        'header': 'FFCDD2',
        'header_text': (183, 28, 28),
        'light': 'FFEBEE',
        'name': 'red'
    },
    'Auto (Rotate Colors)': {
        'header': None,  # Will use COLOR_SETS rotation
        'header_text': (0, 0, 0),
        'light': None,
        'name': 'auto'
    }
}

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
# DATA VALIDATION DROPDOWNS - Medical Study Guide Fields
# ============================================================================
# Dropdown options for common medical/pharmaceutical fields
DROPDOWN_OPTIONS = {
    # Drug administration routes
    "Route": [
        "Oral (PO)",
        "Intravenous (IV)",
        "Intramuscular (IM)",
        "Subcutaneous (SC)",
        "Sublingual (SL)",
        "Buccal",
        "Transdermal",
        "Inhalation",
        "Nasal",
        "Ophthalmic",
        "Otic",
        "Rectal (PR)",
        "Vaginal",
        "Topical",
        "Intrathecal",
        "Epidural",
        "Intra-articular",
        "Multiple routes"
    ],

    # Contraindication severity levels
    "Contraindications": [
        "Absolute: ",
        "Relative: ",
        "Caution in: ",
        "Avoid in: ",
        "None known"
    ],

    # Common drug interaction severity
    "Drug Interactions": [
        "Major: ",
        "Moderate: ",
        "Minor: ",
        "None known",
        "Theoretical: "
    ],

    # Resistance patterns (for antimicrobials)
    "Resistance": [
        "Common",
        "Increasing",
        "Rare",
        "Variable",
        "Not applicable",
        "Mechanisms: "
    ],

    # Clinical presentation urgency
    "Clinical Presentation": [
        "Acute: ",
        "Subacute: ",
        "Chronic: ",
        "Acute on chronic: ",
        "Emergent: ",
        "Variable: "
    ],

    # Lab value interpretation
    "Labs": [
        "Elevated in: ",
        "Decreased in: ",
        "Normal in: ",
        "Variable in: ",
        "Not typically ordered"
    ],

    # Common risk factor categories
    "Risk Factors": [
        "Modifiable: ",
        "Non-modifiable: ",
        "Environmental: ",
        "Genetic: ",
        "Lifestyle: ",
        "Iatrogenic: "
    ],

    # Treatment urgency/type
    "Treatment": [
        "Emergent: ",
        "Urgent: ",
        "Elective: ",
        "Supportive: ",
        "Definitive: ",
        "Palliative: ",
        "First-line: ",
        "Second-line: ",
        "Adjunctive: "
    ],

    # Diagnostic test types
    "Diagnostics": [
        "Laboratory: ",
        "Imaging: ",
        "Biopsy: ",
        "Functional test: ",
        "Physical exam: ",
        "Clinical diagnosis",
        "Gold standard: "
    ]
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
        self.root.title("Study Guide Creator v3.0")
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

        # Search and filter state
        self.original_data = None  # Store original data when filtering
        self.is_filtered = False
        self.search_results = []  # List of matching row indices
        self.current_search_index = 0

        # Ribbon state
        self.ribbon_expanded = tk.BooleanVar(value=True)  # Start expanded

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
        file_menu.add_separator()

        # Import submenu
        import_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Import", menu=import_menu)
        import_menu.add_command(label="From CSV...", command=self.import_from_csv)
        import_menu.add_command(label="From Excel...", command=self.import_from_excel)

        file_menu.add_separator()
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

        # Word/Markdown submenu
        word_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Markdown/Word", menu=word_menu)
        word_menu.add_command(label="Import Markdown File...", command=self.browse_markdown_file)
        word_menu.add_command(label="Preview Markdown...", command=self.preview_markdown)
        word_menu.add_command(label="Convert to Word...", command=self.convert_markdown_to_word)

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
        edit_menu.add_separator()
        edit_menu.add_command(label="Find...", command=self.show_find_dialog, accelerator="Cmd+F")
        edit_menu.add_separator()
        edit_menu.add_command(label="Validate Data Quality...", command=self.show_data_validation)

        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_checkbutton(label="Live Color Preview", variable=self.live_preview_var,
                                  command=self.toggle_color_preview)
        view_menu.add_command(label="Preview Colors", command=self.preview_colors)
        view_menu.add_separator()
        view_menu.add_command(label="Clear Filter", command=self.clear_filter)
        view_menu.add_separator()
        view_menu.add_command(label="Refresh Auto-Complete", command=self.refresh_autocomplete)

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
        # Main ribbon container
        ribbon_wrapper = ttk.Frame(self.root)
        ribbon_wrapper.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=0, pady=0)

        # Toggle button bar
        toggle_bar = ttk.Frame(ribbon_wrapper, relief=tk.RAISED, borderwidth=1)
        toggle_bar.pack(fill=tk.X)

        ttk.Button(
            toggle_bar,
            text="▼ Hide Ribbon" if self.ribbon_expanded.get() else "▶ Show Ribbon",
            command=self.toggle_ribbon,
            width=15
        ).pack(side=tk.LEFT, padx=5, pady=2)

        self.toggle_button = toggle_bar.winfo_children()[0]  # Store reference

        # Ribbon content container
        self.ribbon_container = ttk.Frame(ribbon_wrapper, relief=tk.RAISED, borderwidth=1)
        self.ribbon_container.pack(fill=tk.BOTH, expand=True)

        # Create notebook for tabbed ribbon
        self.ribbon_notebook = ttk.Notebook(self.ribbon_container)
        self.ribbon_notebook.pack(fill=tk.BOTH, expand=True)
        # Bind tab change event to manage sheet visibility
        self.ribbon_notebook.bind("<<NotebookTabChanged>>", self._on_ribbon_tab_changed)

        # HOME TAB
        home_tab = ttk.Frame(self.ribbon_notebook, padding="5")
        self.ribbon_notebook.add(home_tab, text="Home")

        # === CLIPBOARD DROPDOWN ===
        clipboard_menu = tk.Menu(self.root, tearoff=0)
        clipboard_menu.add_command(label="Cut", command=lambda: self.sheet.cut())
        clipboard_menu.add_command(label="Copy", command=lambda: self.sheet.copy())
        clipboard_menu.add_command(label="Paste", command=lambda: self.sheet.paste())

        clipboard_btn = ttk.Menubutton(home_tab, text="Clipboard ▼", menu=clipboard_menu)
        clipboard_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === EDIT DROPDOWN ===
        edit_menu = tk.Menu(self.root, tearoff=0)
        edit_menu.add_command(label="⟲ Undo", command=self.undo_action)
        edit_menu.add_command(label="⟳ Redo", command=self.redo_action)
        edit_menu.add_separator()
        edit_menu.add_command(label="Find", command=self.show_find_dialog)

        edit_btn = ttk.Menubutton(home_tab, text="Edit ▼", menu=edit_menu)
        edit_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === DELETE DROPDOWN ===
        delete_menu = tk.Menu(self.root, tearoff=0)
        delete_menu.add_command(label="Delete Rows", command=self.delete_selected_rows)
        delete_menu.add_command(label="Delete Columns", command=self.delete_selected_columns)
        delete_menu.add_command(label="Clear Cells", command=self.clear_selected_cells)

        delete_btn = ttk.Menubutton(home_tab, text="Delete ▼", menu=delete_menu)
        delete_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === CELLS DROPDOWN ===
        cells_menu = tk.Menu(self.root, tearoff=0)
        cells_menu.add_command(label="Insert Row", command=self.insert_row_below)
        cells_menu.add_command(label="Insert Column", command=self.insert_column_after)

        cells_btn = ttk.Menubutton(home_tab, text="Cells ▼", menu=cells_menu)
        cells_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === DATA TAB ===
        data_tab = ttk.Frame(self.ribbon_notebook, padding="5")
        self.ribbon_notebook.add(data_tab, text="Data")

        # === SETUP DROPDOWN ===
        # Keep preset combobox visible, put other options in dropdown
        setup_frame = ttk.Frame(data_tab)
        setup_frame.pack(side=tk.LEFT, padx=2, pady=2)

        ttk.Label(setup_frame, text="Preset:").pack(side=tk.LEFT, padx=2)
        preset_combo = ttk.Combobox(
            setup_frame,
            textvariable=self.current_preset,
            values=list(COLUMN_PRESETS.keys()),
            state="readonly",
            width=18
        )
        preset_combo.pack(side=tk.LEFT, padx=2)
        preset_combo.bind("<<ComboboxSelected>>", self.on_preset_change)

        setup_menu = tk.Menu(self.root, tearoff=0)
        setup_menu.add_command(label="Custom", command=self.define_custom_columns)
        setup_menu.add_command(label="Edit Headers", command=self.edit_column_headers)

        setup_btn = ttk.Menubutton(setup_frame, text="Setup ▼", menu=setup_menu)
        setup_btn.pack(side=tk.LEFT, padx=2)

        # === FILE DROPDOWN ===
        file_menu = tk.Menu(self.root, tearoff=0)
        file_menu.add_command(label="Save", command=self.save_data_json)
        file_menu.add_command(label="Load", command=self.load_data_json)
        file_menu.add_separator()
        file_menu.add_command(label="Import CSV", command=self.import_from_csv)

        file_btn = ttk.Menubutton(data_tab, text="File ▼", menu=file_menu)
        file_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === ROWS DROPDOWN ===
        rows_menu = tk.Menu(self.root, tearoff=0)
        rows_menu.add_command(label="Add 10 Rows", command=lambda: self.add_quick_rows(10))
        rows_menu.add_command(label="Add 50 Rows", command=lambda: self.add_quick_rows(50))
        rows_menu.add_command(label="Add 100 Rows", command=lambda: self.add_quick_rows(100))
        rows_menu.add_separator()
        rows_menu.add_command(label="Delete Empty Rows", command=self.delete_empty_rows)

        rows_btn = ttk.Menubutton(data_tab, text="Rows ▼", menu=rows_menu)
        rows_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === VIEW DROPDOWN ===
        view_menu = tk.Menu(self.root, tearoff=0)
        view_menu.add_checkbutton(label="Live Colors", variable=self.live_preview_var,
                                   command=self.toggle_color_preview)
        view_menu.add_separator()
        view_menu.add_command(label="Preview", command=self.preview_colors)
        view_menu.add_command(label="Clear Filter", command=self.clear_filter)

        view_btn = ttk.Menubutton(data_tab, text="View ▼", menu=view_menu)
        view_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === EXPORT TAB ===
        export_tab = ttk.Frame(self.ribbon_notebook, padding="5")
        self.ribbon_notebook.add(export_tab, text="Export")

        # === FORMAT DROPDOWN ===
        format_menu = tk.Menu(self.root, tearoff=0)
        format_menu.add_radiobutton(label="Master Chart", variable=self.export_format,
                                     value="master_chart")
        format_menu.add_radiobutton(label="Comprehensive (4 Tabs)", variable=self.export_format,
                                     value="comprehensive")

        format_btn = ttk.Menubutton(export_tab, text="Format ▼", menu=format_menu)
        format_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === EXPORT DROPDOWN ===
        export_menu = tk.Menu(self.root, tearoff=0)
        export_menu.add_command(label="Export to Excel", command=self.export_to_excel)
        export_menu.add_command(label="Export to CSV", command=self.export_to_csv)
        export_menu.add_separator()
        export_menu.add_command(label="Validate Data", command=self.show_data_validation)

        export_btn = ttk.Menubutton(export_tab, text="Export ▼", menu=export_menu)
        export_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # === WORD TAB (Markdown to Word conversion) ===
        # Controls only in ribbon - preview is in separate container below
        word_tab = ttk.Frame(self.ribbon_notebook, padding="2")
        self.ribbon_notebook.add(word_tab, text="Word")

        # Single row control panel (compact)
        control_panel = ttk.Frame(word_tab)
        control_panel.pack(fill=tk.X, padx=0, pady=0)

        # Markdown file selection
        md_frame = ttk.Frame(control_panel)
        md_frame.pack(side=tk.LEFT, padx=2, pady=2)

        ttk.Label(md_frame, text="Markdown:").pack(side=tk.LEFT, padx=2)
        self.markdown_path_var = tk.StringVar(value="No file selected")
        self.md_path_label = ttk.Label(md_frame, textvariable=self.markdown_path_var, width=30)
        self.md_path_label.pack(side=tk.LEFT, padx=2)

        ttk.Button(md_frame, text="Browse...", command=self.browse_markdown_file).pack(side=tk.LEFT, padx=2)

        # Color theme dropdown
        theme_frame = ttk.Frame(control_panel)
        theme_frame.pack(side=tk.LEFT, padx=10, pady=2)

        ttk.Label(theme_frame, text="Theme:").pack(side=tk.LEFT, padx=2)
        self.word_theme_var = tk.StringVar(value="Purple - General Topics")
        theme_combo = ttk.Combobox(
            theme_frame,
            textvariable=self.word_theme_var,
            values=list(WORD_COLOR_THEMES.keys()),
            state="readonly",
            width=20
        )
        theme_combo.pack(side=tk.LEFT, padx=2)

        # Export button (renamed from "Convert to Word")
        ttk.Button(control_panel, text="Export", command=self.convert_markdown_to_word).pack(side=tk.LEFT, padx=5, pady=2)

        # Store markdown content
        self.markdown_content = ""
        self.markdown_file_path = None

        # Create Word preview container (separate from ribbon, shown when Word tab is active)
        self._create_word_preview_container()

    def toggle_ribbon(self):
        """Toggle ribbon visibility"""
        if self.ribbon_expanded.get():
            # Collapse ribbon
            self.ribbon_container.pack_forget()
            self.toggle_button.config(text="▶ Show Ribbon")
            self.ribbon_expanded.set(False)
        else:
            # Expand ribbon
            self.ribbon_container.pack(fill=tk.BOTH, expand=True)
            self.toggle_button.config(text="▼ Hide Ribbon")
            self.ribbon_expanded.set(True)

        # Force window to update layout
        self.root.update_idletasks()

    def _create_word_preview_container(self):
        """Create the Word preview container (separate from ribbon, shown when Word tab is active)"""
        # Container frame for Word preview (placed in main content area)
        self.word_preview_container = ttk.Frame(self.root, padding="10")
        # Initially hidden - will be shown when Word tab is selected

        # Preview label
        ttk.Label(self.word_preview_container, text="Document Preview:",
                  font=("Calibri", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))

        # Create scrollable preview with styled text
        preview_frame = ttk.Frame(self.word_preview_container)
        preview_frame.pack(fill=tk.BOTH, expand=True)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(preview_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create text widget for styled preview
        self.word_preview_text = tk.Text(
            preview_frame,
            wrap=tk.WORD,
            font=("Calibri", 11),
            yscrollcommand=scrollbar.set,
            padx=10,
            pady=10,
            bg="white"
        )
        scrollbar.config(command=self.word_preview_text.yview)
        self.word_preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure text tags for styled preview (mimics Word document colors)
        self.word_preview_text.tag_config("title",
            font=("Calibri", 16, "bold"),
            foreground="#1a1a1a",
            spacing3=10)
        self.word_preview_text.tag_config("heading2",
            font=("Calibri", 14, "bold"),
            foreground="#2c3e50",
            spacing1=15,
            spacing3=5)
        self.word_preview_text.tag_config("heading3",
            font=("Calibri", 12, "bold"),
            foreground="#34495e",
            spacing1=10,
            spacing3=5)
        self.word_preview_text.tag_config("clinical_pearl",
            background="#E8F5E9",
            font=("Calibri", 10),
            lmargin1=20,
            lmargin2=20,
            rmargin=20,
            spacing1=5,
            spacing3=5)
        self.word_preview_text.tag_config("clinical_header",
            background="#E8F5E9",
            font=("Calibri", 11, "bold"),
            foreground="#2E7D32",
            lmargin1=20,
            lmargin2=20)
        self.word_preview_text.tag_config("memory_trick",
            background="#E6F3FF",
            font=("Calibri", 10),
            lmargin1=20,
            lmargin2=20,
            rmargin=20,
            spacing1=5,
            spacing3=5)
        self.word_preview_text.tag_config("memory_header",
            background="#E6F3FF",
            font=("Calibri", 11, "bold"),
            foreground="#1565C0",
            lmargin1=20,
            lmargin2=20)
        self.word_preview_text.tag_config("analogy",
            background="#FFF3E0",
            font=("Calibri", 10),
            lmargin1=20,
            lmargin2=20,
            rmargin=20,
            spacing1=5,
            spacing3=5)
        self.word_preview_text.tag_config("analogy_header",
            background="#FFF3E0",
            font=("Calibri", 11, "bold"),
            foreground="#E65100",
            lmargin1=20,
            lmargin2=20)
        self.word_preview_text.tag_config("table_header",
            background="#D1C4E9",
            font=("Calibri", 10, "bold"),
            foreground="#4A148C")
        self.word_preview_text.tag_config("table_cell",
            background="#F3E5F5",
            font=("Calibri", 10))
        self.word_preview_text.tag_config("normal",
            font=("Calibri", 10))
        self.word_preview_text.tag_config("list_item",
            font=("Calibri", 10),
            lmargin1=30,
            lmargin2=40)

    def _on_ribbon_tab_changed(self, event=None):
        """Handle ribbon tab changes - show/hide sheet based on active tab"""
        selected_tab_index = self.ribbon_notebook.index(self.ribbon_notebook.select())
        tab_names = ["Home", "Data", "Export", "Word"]

        if selected_tab_index < len(tab_names):
            selected_tab = tab_names[selected_tab_index]

            # Handle Word tab - show word preview, hide sheet
            if selected_tab == "Word":
                if self.sheet_container and self.sheet_container.winfo_exists():
                    self.sheet_container.grid_forget()
                # Show Word preview container
                if hasattr(self, 'word_preview_container') and self.word_preview_container.winfo_exists():
                    self.word_preview_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            # Show sheet container for all Excel tabs
            else:
                # Hide Word preview container
                if hasattr(self, 'word_preview_container') and self.word_preview_container.winfo_exists():
                    self.word_preview_container.grid_forget()
                if self.sheet_container and self.sheet_container.winfo_exists():
                    self.sheet_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

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

        # Enable all bindings including drag and drop
        self.sheet.enable_bindings(
            "single_select", "drag_select", "column_width_resize",
            "double_click_column_resize", "row_height_resize",
            "column_select", "row_select", "edit_cell", "edit_index", "edit_header",
            "copy", "paste", "delete", "undo", "redo",
            "column_drag_and_drop", "row_drag_and_drop"  # Enable drag and drop
            # Note: right_click_popup_menu disabled to use custom context menu
        )

        self.sheet.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Bind right-click for context menu
        self.sheet.bind("<Button-2>", self.show_context_menu)  # macOS Control+Click
        self.sheet.bind("<Button-3>", self.show_context_menu)  # macOS/Windows/Linux right-click

        # Bind keyboard shortcuts for delete operations
        # Mac: Command+Backspace to delete selected rows
        self.sheet.bind("<Command-BackSpace>", lambda e: self.delete_selected_rows())
        # Windows/Linux: Ctrl+Delete to delete selected rows
        self.sheet.bind("<Control-Delete>", lambda e: self.delete_selected_rows())

        # Bind cell changes for unsaved indicator
        self.sheet.bind("<<SheetModified>>", self.on_sheet_modified)

        # Bind header changes to sync column names
        self.sheet.bind("<<SheetModified>>", self.sync_column_headers, add=True)

        # NOTE: Dropdown validation disabled - tksheet auto-fills cells with dropdown values
        # Users can still manually enter values and use auto-complete suggestions
        # self.configure_dropdown_validations()

        # Update row count
        self.update_row_count()

    def get_column_unique_values(self, col_idx, max_values=50):
        """Get unique non-empty values from a column for auto-complete

        Args:
            col_idx: Column index
            max_values: Maximum number of unique values to return (prevents huge dropdowns)

        Returns:
            List of unique values sorted alphabetically
        """
        if not self.sheet:
            return []

        data = self.sheet.get_sheet_data()
        unique_values = set()

        for row in data:
            if col_idx < len(row):
                value = row[col_idx]
                if value and isinstance(value, str):
                    value = value.strip()
                    if value:  # Non-empty after stripping
                        unique_values.add(value)

        # Sort and limit
        sorted_values = sorted(list(unique_values))
        return sorted_values[:max_values]

    def configure_dropdown_validations(self):
        """Configure dropdown validation with both predefined options and auto-complete from previous entries"""
        if not self.sheet or not self.current_columns:
            return

        # Map columns to their indices and check if they have dropdown options
        for col_idx, col_name in enumerate(self.current_columns):
            dropdown_values = []

            # Get predefined dropdown options if available
            if col_name in DROPDOWN_OPTIONS:
                dropdown_values = DROPDOWN_OPTIONS[col_name].copy()

            # Get unique values from previous entries (auto-complete)
            previous_values = self.get_column_unique_values(col_idx, max_values=30)

            # Merge: Add previous values that aren't already in predefined options
            # This prevents duplicates while preserving the order of predefined options
            if previous_values:
                for prev_val in previous_values:
                    if prev_val not in dropdown_values:
                        dropdown_values.append(prev_val)

            # Only configure dropdown if we have values
            if dropdown_values:
                # Set dropdown for this entire column
                try:
                    self.sheet.dropdown(
                        col_idx,
                        values=dropdown_values,
                        set_value=None,  # Don't pre-fill
                        state="normal",  # Allow custom text too
                        redraw=True
                    )
                except Exception as e:
                    # Some versions of tksheet might have different API
                    # Try alternative method
                    try:
                        for row_idx in range(self.sheet.get_total_rows()):
                            self.sheet.dropdown(
                                row_idx,
                                col_idx,
                                values=dropdown_values,
                                set_value=None,
                                state="normal",
                                redraw=False
                            )
                        self.sheet.refresh()
                    except:
                        pass  # Skip if dropdown not supported

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
        # NOTE: Dropdown validation disabled (auto-fills cells)
        # self.configure_dropdown_validations()
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

    def show_data_validation(self):
        """Show data validation results in a dialog"""
        warnings = self.validate_data_quality()

        if not warnings:
            messagebox.showinfo(
                "Data Quality Check",
                "✓ No data quality issues detected!\n\nYour data looks good and is ready for export."
            )
        else:
            warning_text = "Data quality issues detected:\n\n" + "\n\n".join(warnings)
            warning_text += "\n\nReview and fix these issues before exporting for best results."

            # Create custom dialog with scrollable text
            dialog = tk.Toplevel(self.root)
            dialog.title("Data Quality Validation")
            dialog.geometry("600x400")
            dialog.transient(self.root)

            # Text widget with scrollbar
            text_frame = ttk.Frame(dialog)
            text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            scrollbar = ttk.Scrollbar(text_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            text_widget = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=text_widget.yview)

            # Insert warning text
            text_widget.insert("1.0", warning_text)
            text_widget.config(state=tk.DISABLED)  # Make read-only

            # OK button
            ttk.Button(dialog, text="OK", command=dialog.destroy).pack(pady=10)

    # ========================================================================
    # SEARCH AND FILTER
    # ========================================================================

    def show_find_dialog(self):
        """Show find/search dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Find")
        dialog.geometry("500x250")
        dialog.transient(self.root)
        dialog.grab_set()

        # Search input
        ttk.Label(dialog, text="Find what:").pack(anchor=tk.W, padx=10, pady=(10, 0))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(dialog, textvariable=search_var, width=50)
        search_entry.pack(padx=10, pady=5, fill=tk.X)
        search_entry.focus()

        # Search options
        options_frame = ttk.LabelFrame(dialog, text="Options", padding="10")
        options_frame.pack(padx=10, pady=10, fill=tk.BOTH)

        case_sensitive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Case sensitive",
                       variable=case_sensitive_var).pack(anchor=tk.W)

        whole_word_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Match whole word",
                       variable=whole_word_var).pack(anchor=tk.W)

        # Search scope
        scope_var = tk.StringVar(value="all")
        ttk.Radiobutton(options_frame, text="Search all columns",
                       variable=scope_var, value="all").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(options_frame, text="Search first column only",
                       variable=scope_var, value="first").pack(anchor=tk.W, pady=2)

        # Button frame
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        def do_find():
            search_text = search_var.get()
            if not search_text:
                messagebox.showwarning("No Search Text", "Please enter text to search for.")
                return

            matches = self.find_in_data(
                search_text,
                case_sensitive=case_sensitive_var.get(),
                whole_word=whole_word_var.get(),
                scope=scope_var.get()
            )

            if matches:
                dialog.destroy()
                self.show_search_results(search_text, matches)
            else:
                messagebox.showinfo("No Matches", f"No matches found for '{search_text}'.")

        def do_filter():
            search_text = search_var.get()
            if not search_text:
                messagebox.showwarning("No Search Text", "Please enter text to filter by.")
                return

            matches = self.find_in_data(
                search_text,
                case_sensitive=case_sensitive_var.get(),
                whole_word=whole_word_var.get(),
                scope=scope_var.get()
            )

            if matches:
                dialog.destroy()
                self.apply_filter(search_text, matches)
            else:
                messagebox.showinfo("No Matches", f"No matches found for '{search_text}'.")

        ttk.Button(button_frame, text="Find All", command=do_find).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Filter", command=do_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def find_in_data(self, search_text, case_sensitive=False, whole_word=False, scope="all"):
        """Find all rows containing the search text

        Args:
            search_text: Text to search for
            case_sensitive: Whether search is case sensitive
            whole_word: Whether to match whole words only
            scope: "all" for all columns, "first" for first column only

        Returns:
            List of (row_index, column_index, cell_value) tuples for matches
        """
        data = self.sheet.get_sheet_data()
        matches = []

        # Prepare search text
        if not case_sensitive:
            search_text = search_text.lower()

        for row_idx, row in enumerate(data):
            # Skip completely empty rows
            if not any(cell.strip() if isinstance(cell, str) else cell for cell in row):
                continue

            # Determine which columns to search
            columns_to_search = [0] if scope == "first" else range(len(row))

            for col_idx in columns_to_search:
                if col_idx >= len(row):
                    continue

                cell_value = str(row[col_idx]) if row[col_idx] else ""
                if not cell_value:
                    continue

                # Prepare cell value
                compare_value = cell_value if case_sensitive else cell_value.lower()

                # Check for match
                if whole_word:
                    # Match whole words only
                    pattern = r'\b' + re.escape(search_text) + r'\b'
                    if re.search(pattern, compare_value):
                        matches.append((row_idx, col_idx, cell_value))
                else:
                    # Substring match
                    if search_text in compare_value:
                        matches.append((row_idx, col_idx, cell_value))

        return matches

    def show_search_results(self, search_text, matches):
        """Display search results in a dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Search Results for '{search_text}'")
        dialog.geometry("600x400")
        dialog.transient(self.root)

        ttk.Label(dialog, text=f"Found {len(matches)} matches:",
                 font=("", 12, "bold")).pack(pady=10)

        # Create scrollable listbox
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, font=("Courier", 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        # Populate results
        for row_idx, col_idx, cell_value in matches:
            col_name = self.current_columns[col_idx] if col_idx < len(self.current_columns) else f"Col{col_idx}"
            # Truncate long cell values
            display_value = cell_value[:50] + "..." if len(cell_value) > 50 else cell_value
            listbox.insert(tk.END, f"Row {row_idx + 1}, {col_name}: {display_value}")

        def jump_to_match(event=None):
            selection = listbox.curselection()
            if selection:
                idx = selection[0]
                row_idx, col_idx, _ = matches[idx]
                # Select the cell in the sheet
                try:
                    self.sheet.see(row_idx, col_idx)
                    self.sheet.select_cell(row_idx, col_idx)
                    dialog.destroy()
                except:
                    pass

        listbox.bind("<Double-Button-1>", jump_to_match)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Go to Selected", command=jump_to_match).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def apply_filter(self, search_text, matches):
        """Filter the data to show only matching rows"""
        if self.is_filtered:
            # Already filtered, restore first
            self.clear_filter()

        # Store original data
        self.original_data = self.sheet.get_sheet_data()
        self.is_filtered = True

        # Get unique row indices from matches
        matching_row_indices = sorted(set(m[0] for m in matches))

        # Filter data
        filtered_data = [self.original_data[idx] for idx in matching_row_indices]

        # Update sheet
        self.sheet.set_sheet_data(filtered_data)
        self.update_row_count()

        # Update status
        self.status_label.config(text=f"Filtered: showing {len(filtered_data)} rows matching '{search_text}'  |  Click View > Clear Filter to show all")

        # Add Clear Filter to View menu temporarily
        messagebox.showinfo("Filter Applied",
                          f"Showing {len(filtered_data)} rows matching '{search_text}'.\n\n"
                          "To show all data again, use View > Clear Filter.")

    def clear_filter(self):
        """Clear any active filter and restore all data"""
        if not self.is_filtered:
            messagebox.showinfo("No Filter Active", "There is no active filter to clear.")
            return

        # Restore original data
        if self.original_data:
            self.sheet.set_sheet_data(self.original_data)
            self.update_row_count()

        self.original_data = None
        self.is_filtered = False

        # Update status
        self.status_label.config(text="All rows displayed")

        messagebox.showinfo("Filter Cleared", "All data is now visible.")

    # ========================================================================
    # CONTEXT MENU
    # ========================================================================

    def show_context_menu(self, event):
        """Show right-click context menu"""
        # On Mac, ensure we're responding to right-click
        # Button-2 is typically Control+Click, Button-3 is right-click
        menu = tk.Menu(self.root, tearoff=0)

        # Check if anything is selected
        selected = self.sheet.get_all_selection_boxes()
        has_selection = bool(selected)

        menu.add_command(label="Cut", accelerator="Cmd+X", command=lambda: self.sheet.cut(),
                        state=tk.NORMAL if has_selection else tk.DISABLED)
        menu.add_command(label="Copy", accelerator="Cmd+C", command=lambda: self.sheet.copy(),
                        state=tk.NORMAL if has_selection else tk.DISABLED)
        menu.add_command(label="Paste", accelerator="Cmd+V", command=lambda: self.sheet.paste())
        menu.add_separator()

        menu.add_command(label="Insert Row Above", command=self.insert_row_above)
        menu.add_command(label="Insert Row Below", command=self.insert_row_below)
        menu.add_command(label="Insert 10 Rows Below", command=lambda: self.insert_rows_below(10))
        menu.add_separator()

        menu.add_command(label="Insert Column Before", command=self.insert_column_before)
        menu.add_command(label="Insert Column After", command=self.insert_column_after)
        menu.add_separator()

        menu.add_command(label="Delete Selected Row(s)", command=self.delete_selected_rows,
                        state=tk.NORMAL if has_selection else tk.DISABLED)
        menu.add_command(label="Delete Selected Column(s)", command=self.delete_selected_columns,
                        state=tk.NORMAL if has_selection else tk.DISABLED)
        menu.add_command(label="Delete All Empty Rows", command=self.delete_empty_rows)
        menu.add_separator()

        menu.add_command(label="Clear Selected Cells", command=lambda: self.clear_selected_cells(),
                        state=tk.NORMAL if has_selection else tk.DISABLED)
        menu.add_command(label="Fill Down", command=self.fill_down,
                        state=tk.NORMAL if has_selection else tk.DISABLED)
        menu.add_command(label="Fill Right", command=self.fill_right,
                        state=tk.NORMAL if has_selection else tk.DISABLED)
        menu.add_separator()
        menu.add_command(label="Refresh Auto-Complete", command=self.refresh_autocomplete)

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def clear_selected_cells(self):
        """Clear content from selected cells"""
        try:
            self.sheet.delete_key()
            self.mark_unsaved()
        except:
            pass

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

    def undo_action(self):
        """Undo action with user feedback"""
        try:
            self.sheet.undo()
            self.status_label.config(text="● Undo successful")
        except:
            self.status_label.config(text="● Nothing to undo")

    def redo_action(self):
        """Redo action with user feedback"""
        try:
            self.sheet.redo()
            self.status_label.config(text="● Redo successful")
        except:
            self.status_label.config(text="● Nothing to redo")

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

    def refresh_autocomplete(self):
        """Refresh auto-complete suggestions based on current data"""
        # NOTE: Dropdown validation disabled (auto-fills cells)
        # self.configure_dropdown_validations()
        messagebox.showinfo("Auto-Complete Note",
                          "Dropdown auto-complete is currently disabled due to tksheet auto-fill issues.\n\n"
                          "You can still manually type values - the app learns from your entries.")

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

        # NOTE: Dropdown validation disabled (auto-fills cells)
        # self.configure_dropdown_validations()

        self.update_row_count()
        self.mark_saved()

    # ========================================================================
    # IMPORT FROM CSV/EXCEL
    # ========================================================================

    def import_from_csv(self):
        """Import data from CSV file with optional color metadata restoration"""
        filename = filedialog.askopenfilename(
            title="Import from CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=self.output_directory.get()
        )

        if not filename:
            return

        try:
            # Try to open UTF-8 with BOM first (Excel compatibility), fallback to regular UTF-8
            try:
                with open(filename, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    rows = list(reader)
            except:
                with open(filename, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    rows = list(reader)

            if not rows:
                messagebox.showwarning("Empty File", "The CSV file is empty.")
                return

            # First row is headers
            headers = rows[0]
            data_rows = rows[1:]

            # Check for companion metadata file
            metadata_path = Path(filename).with_suffix('.csv.meta')
            metadata = None
            if metadata_path.exists():
                try:
                    with open(metadata_path, 'r', encoding='utf-8') as f:
                        metadata = json.load(f)
                except:
                    metadata = None

            self._import_data(headers, data_rows, filename, metadata=metadata)

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import CSV:\n{str(e)}")

    def import_from_excel(self):
        """Import data from Excel file with multi-sheet support"""
        filename = filedialog.askopenfilename(
            title="Import from Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=self.output_directory.get()
        )

        if not filename:
            return

        try:
            wb = load_workbook(filename, data_only=True)

            # If multiple sheets, ask user which to import
            if len(wb.sheetnames) > 1:
                selected_sheet = self._show_sheet_selector_dialog(wb.sheetnames)
                if not selected_sheet:
                    return
                ws = wb[selected_sheet]
            else:
                ws = wb.active

            # Get all rows
            rows = list(ws.values)

            if not rows:
                messagebox.showwarning("Empty File", "The Excel file is empty.")
                return

            # First row is headers
            headers = [str(cell) if cell is not None else "" for cell in rows[0]]
            # Preserve data types when converting to strings
            data_rows = []
            for row in rows[1:]:
                converted_row = []
                for cell in row:
                    if cell is None:
                        converted_row.append("")
                    else:
                        converted_row.append(str(cell))
                data_rows.append(converted_row)

            self._import_data(headers, data_rows, filename)

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import Excel:\n{str(e)}")

    def _show_sheet_selector_dialog(self, sheet_names):
        """Show dialog to select which sheet to import from multi-sheet workbook"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Sheet")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="This workbook has multiple sheets.\nSelect which sheet to import:",
                 font=("", 11)).pack(pady=15)

        # Listbox for sheet selection
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, height=10)
        scrollbar.config(command=listbox.yview)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for sheet_name in sheet_names:
            listbox.insert(tk.END, sheet_name)

        # Select first sheet by default
        listbox.selection_set(0)
        listbox.activate(0)

        result = {"selected": None}

        def confirm_selection():
            selection = listbox.curselection()
            if selection:
                result["selected"] = sheet_names[selection[0]]
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=15)
        ttk.Button(button_frame, text="Import", command=confirm_selection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

        dialog.wait_window()
        return result.get("selected")

    def _import_data(self, import_headers, import_data, filename, metadata=None):
        """Process imported data with column mapping and append/replace options

        Args:
            import_headers: List of column names from imported file
            import_data: List of data rows
            filename: Source filename for display
            metadata: Optional metadata dict with color map from CSV export
        """
        # Ask user: Append or Replace?
        dialog = tk.Toplevel(self.root)
        dialog.title("Import Options")
        dialog.geometry("500x300")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text=f"Importing from:\n{Path(filename).name}",
                 font=("", 12, "bold")).pack(pady=10)

        ttk.Label(dialog, text=f"Found {len(import_data)} data rows with {len(import_headers)} columns").pack()

        # Show preview of headers
        preview_text = "Columns: " + ", ".join(import_headers[:5])
        if len(import_headers) > 5:
            preview_text += f", ... ({len(import_headers)} total)"
        ttk.Label(dialog, text=preview_text, wraplength=450).pack(pady=10)

        # Import mode selection
        mode_var = tk.StringVar(value="append")
        ttk.Radiobutton(dialog, text="Append to existing data", variable=mode_var,
                       value="append").pack(anchor=tk.W, padx=30, pady=5)
        ttk.Radiobutton(dialog, text="Replace all existing data", variable=mode_var,
                       value="replace").pack(anchor=tk.W, padx=30, pady=5)

        result = {"confirmed": False}

        def confirm_import():
            result["confirmed"] = True
            result["mode"] = mode_var.get()
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        ttk.Button(button_frame, text="Import", command=confirm_import).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

        dialog.wait_window()

        if not result.get("confirmed"):
            return

        # Map imported columns to current columns
        mapped_data = self._map_columns(import_headers, import_data)

        # Apply import
        if result["mode"] == "replace":
            self.sheet.set_sheet_data(mapped_data)
        else:  # append
            current_data = self.sheet.get_sheet_data()
            # Remove trailing empty rows from current data
            non_empty_current = [row for row in current_data if any(cell.strip() if isinstance(cell, str) else cell for cell in row)]
            self.sheet.set_sheet_data(non_empty_current + mapped_data)

        # NOTE: Dropdown validation disabled (auto-fills cells)
        # self.configure_dropdown_validations()

        self.update_row_count()
        self.mark_unsaved()

        # Automatically apply color coding if metadata exists (from CSV export)
        if metadata and 'color_map' in metadata:
            # Colors were saved with this CSV - restore them
            self.live_preview_var.set(True)
            self.apply_live_colors()
            messagebox.showinfo(
                "Colors Restored",
                f"Successfully imported {len(mapped_data)} rows with saved color coding!\n\n"
                "Color coding has been automatically restored.\n"
                "Live Color Preview is now enabled."
            )
        # Automatically apply color coding if Live Preview is enabled
        elif self.live_preview_var.get():
            self.apply_live_colors()
        else:
            # Ask user if they want to apply colors
            apply_colors = messagebox.askyesno(
                "Import Complete",
                f"Successfully imported {len(mapped_data)} rows!\n\n"
                "Would you like to apply color coding to drug classes now?"
            )

            if apply_colors:
                # Enable live preview and apply colors
                self.live_preview_var.set(True)
                self.apply_live_colors()
                messagebox.showinfo("Colors Applied",
                                  "Color coding has been applied to all drug classes.\n\n"
                                  "Live Color Preview is now enabled.")

    def _map_columns(self, import_headers, import_data):
        """Map imported columns to current column structure

        Args:
            import_headers: List of imported column names
            import_data: List of imported data rows

        Returns:
            List of data rows mapped to current column structure
        """
        # Create column mapping (case-insensitive match)
        import_headers_lower = [h.lower().strip() for h in import_headers]
        current_headers_lower = [h.lower().strip() for h in self.current_columns]

        column_map = {}  # Maps import column index to current column index
        for imp_idx, imp_header in enumerate(import_headers_lower):
            if imp_header in current_headers_lower:
                curr_idx = current_headers_lower.index(imp_header)
                column_map[imp_idx] = curr_idx

        # Map data rows
        mapped_rows = []
        for row in import_data:
            # Skip completely empty rows
            if not any(cell.strip() if isinstance(cell, str) else cell for cell in row):
                continue

            new_row = [""] * len(self.current_columns)
            for imp_idx, curr_idx in column_map.items():
                if imp_idx < len(row):
                    new_row[curr_idx] = row[imp_idx]

            mapped_rows.append(new_row)

        return mapped_rows

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

    def validate_data_quality(self):
        """Validate data quality and return list of warnings

        Returns:
            List of warning strings, empty if no issues found
        """
        warnings = []
        data = self.sheet.get_sheet_data()

        # Track statistics
        empty_first_column = []
        incomplete_rows = []
        duplicate_entries = {}
        very_sparse_rows = []

        for row_idx, row in enumerate(data, start=1):
            # Skip completely empty rows
            if not any(cell.strip() if isinstance(cell, str) else cell for cell in row):
                continue

            # Check for empty first column (drug class / main identifier)
            if not row[0] or (isinstance(row[0], str) and not row[0].strip()):
                # But has data in other columns
                if any(cell.strip() if isinstance(cell, str) else cell for cell in row[1:]):
                    empty_first_column.append(row_idx)

            # Check for incomplete rows (some columns filled, many empty)
            filled_cells = sum(1 for cell in row if cell and (not isinstance(cell, str) or cell.strip()))
            total_cells = len(row)

            if filled_cells > 0:
                # If less than 30% of cells are filled (but not completely empty)
                if filled_cells / total_cells < 0.3 and filled_cells < total_cells:
                    very_sparse_rows.append((row_idx, filled_cells, total_cells))

                # Check for duplicate entries in first column (for Drug Chart)
                if self.current_preset.get() == "Drug Chart (11 columns)" and row[0]:
                    key = str(row[0]).strip().lower()
                    if key:
                        if key in duplicate_entries:
                            duplicate_entries[key].append(row_idx)
                        else:
                            duplicate_entries[key] = [row_idx]

        # Generate warnings
        if empty_first_column:
            if len(empty_first_column) <= 5:
                rows_str = ", ".join(map(str, empty_first_column))
            else:
                rows_str = ", ".join(map(str, empty_first_column[:5])) + f", ... ({len(empty_first_column)} total)"
            warnings.append(f"⚠ Empty first column in rows: {rows_str}\n  (First column should contain the main identifier)")

        if very_sparse_rows:
            if len(very_sparse_rows) <= 3:
                rows_str = ", ".join(f"{r} ({f}/{t} filled)" for r, f, t in very_sparse_rows)
            else:
                rows_str = ", ".join(f"{r} ({f}/{t})" for r, f, t in very_sparse_rows[:3]) + f"... ({len(very_sparse_rows)} total)"
            warnings.append(f"⚠ Incomplete rows detected: {rows_str}\n  (Less than 30% of cells filled)")

        # Check for duplicates (only if there are duplicates)
        actual_duplicates = {k: v for k, v in duplicate_entries.items() if len(v) > 1}
        if actual_duplicates:
            if len(actual_duplicates) <= 3:
                dup_str = "\n  ".join(f"{k}: rows {', '.join(map(str, v))}" for k, v in list(actual_duplicates.items())[:3])
            else:
                dup_str = "\n  ".join(f"{k}: rows {', '.join(map(str, v))}" for k, v in list(actual_duplicates.items())[:3])
                dup_str += f"\n  ... ({len(actual_duplicates)} duplicate entries total)"
            warnings.append(f"⚠ Duplicate entries found:\n  {dup_str}")

        return warnings

    def export_to_excel(self):
        """Export data to Excel (route to appropriate format)"""
        # Validate data quality first
        warnings = self.validate_data_quality()

        if warnings:
            warning_text = "Data quality issues detected:\n\n" + "\n\n".join(warnings)
            warning_text += "\n\nDo you want to continue with export anyway?"

            response = messagebox.askyesno(
                "Data Quality Warnings",
                warning_text,
                icon='warning'
            )

            if not response:
                return  # User chose not to export

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

    def export_to_csv(self):
        """Export grid data to CSV with color metadata"""
        # Get data
        data = self.sheet.get_sheet_data()
        non_empty = [row for row in data if any(cell.strip() if isinstance(cell, str) else cell for cell in row)]

        # Validation
        if not non_empty:
            messagebox.showwarning("No Data", "Please add some data before exporting.")
            return

        # Show file save dialog
        output_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=self.output_directory.get(),
            initialfile=self.output_filename.get().replace(".xlsx", ".csv")
        )

        if not output_path:
            return  # User cancelled

        output_path = Path(output_path)
        metadata_path = output_path.with_suffix('.csv.meta')

        try:
            # Export CSV with UTF-8 BOM encoding (Excel compatible)
            with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
                # Write header row
                writer.writerow(self.current_columns)
                # Write data rows
                for row in non_empty:
                    cleaned_row = ['' if cell is None else str(cell) for cell in row]
                    writer.writerow(cleaned_row)

            # Save color metadata to companion file
            try:
                color_map = self.calculate_color_assignments_from_data(non_empty)
                metadata = {
                    'headers': self.current_columns,
                    'color_map': color_map,
                    'row_count': len(non_empty),
                    'preset': self.current_preset.get(),
                    'export_date': datetime.now().isoformat()
                }

                with open(metadata_path, 'w', encoding='utf-8') as f:
                    json.dump(metadata, f, indent=2)
            except Exception as e:
                # Metadata save failed, but CSV is still created
                pass

            messagebox.showinfo(
                "Export Successful",
                f"CSV exported successfully!\n\nLocation:\n{output_path}\n\nRows exported: {len(non_empty)}\n\nColor metadata saved to: {metadata_path.name}"
            )

            self.mark_saved()

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export CSV file:\n\n{str(e)}")

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

            current_row += 2  # Blank rows between drug classes

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

                    ws[f'B{current_row}'] = "\n".join(values) if values else ""
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
        guide = """Excel Master Chart Creator v2.6 - Quick Start

1. SELECT PRESET
   Choose from Drug Chart, Condition Chart, Lab Values, or Custom

2. ENTER DATA
   - Click any cell to start typing
   - Press Tab to move right, Enter to move down
   - Right-click for more options
   - Use Edit > Find to search your data

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
        about = """Excel Master Chart Creator v2.6

A powerful desktop application for creating professional Excel charts
with auto-color assignment and multi-format export.

Features:
✓ Excel-like grid interface
✓ Auto-color by drug class
✓ Two export formats (Master Chart & Comprehensive)
✓ Auto-save and crash recovery
✓ Right-click context menu
✓ Live color preview
✓ Data quality validation
✓ CSV/Excel import
✓ Search and filter

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

    # ========================================================================
    # WORD TAB - MARKDOWN TO WORD CONVERSION
    # ========================================================================

    def browse_markdown_file(self):
        """Open file dialog to select a markdown file and auto-display preview"""
        file_path = filedialog.askopenfilename(
            title="Select Markdown File",
            filetypes=[
                ("Markdown files", "*.md"),
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )

        if file_path:
            self.markdown_file_path = Path(file_path)
            # Show truncated path in label
            display_name = self.markdown_file_path.name
            if len(display_name) > 28:
                display_name = display_name[:25] + "..."
            self.markdown_path_var.set(display_name)

            # Read the file
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.markdown_content = f.read()
                self.status_label.config(text=f"● Loaded: {self.markdown_file_path.name}", foreground="green")

                # Auto-display preview in the Word tab
                self._display_word_preview()
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file:\n{str(e)}")
                self.markdown_content = ""

    def _display_word_preview(self):
        """Display styled Word preview with colored backgrounds (like actual Word doc)"""
        if not self.markdown_content:
            return

        try:
            # Parse the markdown
            parsed = self.parse_markdown(self.markdown_content)

            # Clear and enable the text widget
            self.word_preview_text.config(state=tk.NORMAL)
            self.word_preview_text.delete("1.0", tk.END)

            # Add title
            if parsed['title']:
                self.word_preview_text.insert(tk.END, f"{parsed['title']}\n", "title")
                self.word_preview_text.insert(tk.END, "\n")

            # Track rendered elements to avoid duplicates
            rendered_tables = set()
            rendered_blockquotes = set()

            # Process sections
            for section in parsed['sections']:
                # Add section heading with appropriate style
                if section['level'] == 2:
                    self.word_preview_text.insert(tk.END, f"\n{section['title']}\n", "heading2")
                else:
                    self.word_preview_text.insert(tk.END, f"\n{section['title']}\n", "heading3")

                # Add section content
                for content_type, content in section['content']:
                    if content_type == 'text':
                        self.word_preview_text.insert(tk.END, f"{content}\n", "normal")
                    elif content_type == 'list':
                        self.word_preview_text.insert(tk.END, f"• {content}\n", "list_item")

                # Add tables in this section
                for i, table_data in enumerate(parsed['tables']):
                    if table_data.get('section') == section['title'] and i not in rendered_tables:
                        rendered_tables.add(i)
                        self.word_preview_text.insert(tk.END, "\n")
                        # Table headers
                        headers_str = " │ ".join(table_data['headers'])
                        self.word_preview_text.insert(tk.END, f"  {headers_str}\n", "table_header")
                        # Table rows
                        for row in table_data['rows']:
                            row_str = " │ ".join(str(cell).replace('\n', ' ') for cell in row)
                            self.word_preview_text.insert(tk.END, f"  {row_str}\n", "table_cell")
                        self.word_preview_text.insert(tk.END, "\n")

                # Add blockquotes in this section (with proper colored backgrounds)
                for i, bq in enumerate(parsed['blockquotes']):
                    if bq.get('section') == section['title'] and i not in rendered_blockquotes:
                        rendered_blockquotes.add(i)
                        bq_type = bq.get('type', 'clinical')
                        content = bq['content']

                        self.word_preview_text.insert(tk.END, "\n")

                        # Add header and content with appropriate styling
                        if bq_type == 'memory':
                            self.word_preview_text.insert(tk.END, "💡 MEMORY TRICKS & MNEMONICS\n", "memory_header")
                            self.word_preview_text.insert(tk.END, f"{content}\n", "memory_trick")
                        elif bq_type == 'analogy':
                            self.word_preview_text.insert(tk.END, "🔗 ANALOGY\n", "analogy_header")
                            self.word_preview_text.insert(tk.END, f"{content}\n", "analogy")
                        else:
                            self.word_preview_text.insert(tk.END, "📋 CLINICAL PEARLS & HIGH-YIELD\n", "clinical_header")
                            self.word_preview_text.insert(tk.END, f"{content}\n", "clinical_pearl")

                        self.word_preview_text.insert(tk.END, "\n")

            # Leave editable for user modifications
            # self.word_preview_text.config(state=tk.DISABLED)

        except Exception as e:
            # Show error in preview
            self.word_preview_text.config(state=tk.NORMAL)
            self.word_preview_text.delete("1.0", tk.END)
            self.word_preview_text.insert("1.0", f"Error displaying preview:\n{str(e)}")
            self.word_preview_text.config(state=tk.DISABLED)

    def preview_markdown(self):
        """Show a preview of the parsed markdown content"""
        if not self.markdown_content:
            messagebox.showwarning("No File", "Please select a markdown file first.")
            return

        # Create preview window
        preview_win = tk.Toplevel(self.root)
        preview_win.title("Markdown Preview")
        preview_win.geometry("800x600")
        preview_win.transient(self.root)

        # Create text widget with scrollbar
        frame = ttk.Frame(preview_win)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        text = tk.Text(frame, wrap=tk.WORD, font=("Courier", 11))
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Parse and display structured content
        parsed = self.parse_markdown(self.markdown_content)
        preview_text = self.format_parsed_preview(parsed)
        text.insert("1.0", preview_text)
        text.config(state=tk.DISABLED)

        # Close button
        ttk.Button(preview_win, text="Close", command=preview_win.destroy).pack(pady=10)

    def preprocess_markdown_tables(self, content):
        """Pre-process markdown to join multi-line table cells before parsing.

        This fixes issues where table cells span multiple lines, like:
        | **HDL** | >40 mg/dL (male)
        >50 mg/dL (female) | - |

        These get joined into single lines for proper parsing.
        """
        lines = content.split('\n')
        result = []
        i = 0
        in_table = False

        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            # Check if this is a table row (starts with |)
            if stripped.startswith('|'):
                in_table = True
                # Check if this row is complete (ends with |) or continues on next line
                combined = line

                # Look ahead for continuation lines (don't start with |, not empty, not a new section)
                while i + 1 < len(lines):
                    next_line = lines[i + 1]
                    next_stripped = next_line.strip()

                    # Stop conditions: empty line, new table row, section header, separator row
                    if not next_stripped:
                        break
                    if next_stripped.startswith('|'):
                        break
                    if next_stripped.startswith('#'):
                        break
                    if next_stripped.startswith('>') and not any(c.isalnum() for c in next_stripped[1:3]):
                        # Likely a blockquote, not a ">" in table content
                        break
                    if all(c in '|-: ' for c in next_stripped):
                        break

                    # This is a continuation of the table cell
                    # Join with a line break marker that we'll preserve
                    combined = combined.rstrip() + '\n' + next_line
                    i += 1

                result.append(combined)
                in_table = True
            else:
                # Not a table row
                if stripped == '' or stripped.startswith('#'):
                    in_table = False
                result.append(line)

            i += 1

        return '\n'.join(result)

    def split_blockquotes_by_type(self, blockquote_lines, section):
        """Split a collected blockquote into separate typed entries.

        Detects headers like **Clinical Pearls:**, **Memory Tricks:**, **Analogy:**
        and creates separate blockquote entries for each type.
        """
        if not blockquote_lines:
            return []

        groups = []
        current_type = 'clinical'  # Default type
        current_lines = []

        # Patterns that indicate a new blockquote section
        header_patterns = {
            'clinical': ['**clinical pearls', '**high-yield', '**epidemiology', '**critical'],
            'memory': ['**memory tricks', '**mnemonics', '**mnemonic'],
            'analogy': ['**analogy']
        }

        for line in blockquote_lines:
            line_lower = line.lower()

            # Check if this line is a header for a new type
            new_type = None
            for bq_type, patterns in header_patterns.items():
                if any(pattern in line_lower for pattern in patterns):
                    new_type = bq_type
                    break

            if new_type:
                # Save current group if it has content
                if current_lines:
                    # Clean up empty lines at the end
                    while current_lines and current_lines[-1] == '':
                        current_lines.pop()
                    if current_lines:
                        groups.append({
                            'type': current_type,
                            'content': '\n'.join(current_lines),
                            'section': section
                        })
                # Start new group
                current_type = new_type
                current_lines = [line]
            else:
                current_lines.append(line)

        # Don't forget the last group
        if current_lines:
            while current_lines and current_lines[-1] == '':
                current_lines.pop()
            if current_lines:
                groups.append({
                    'type': current_type,
                    'content': '\n'.join(current_lines),
                    'section': section
                })

        return groups

    def parse_markdown(self, content):
        """Parse markdown content into structured data with proper table and blockquote handling"""
        # Pre-process to handle multi-line table cells
        content = self.preprocess_markdown_tables(content)

        lines = content.split('\n')
        parsed = {
            'title': '',
            'sections': [],
            'tables': [],
            'blockquotes': [],
            'lists': [],
            'toc': []  # Table of Contents
        }

        # Track section hierarchy (stack of sections)
        section_stack = []
        current_section_title = None

        # For collecting blockquote lines (preserves internal blank lines)
        blockquote_lines = []
        blockquote_start_idx = -1

        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            # Handle H1 (Title)
            if stripped.startswith('# ') and not stripped.startswith('## '):
                parsed['title'] = stripped[2:].strip()
                current_section_title = parsed['title']
                section_stack = []
                i += 1
                continue

            # Handle H2 (Major sections)
            if stripped.startswith('## '):
                section_title = stripped[3:].strip()

                # Check if this is Table of Contents (special handling)
                if section_title.lower() in ['table of contents', 'toc', 'contents']:
                    # Collect TOC lines until next ## or ### section
                    toc_lines = []
                    i += 1
                    while i < len(lines):
                        toc_line = lines[i].strip()
                        if toc_line.startswith('## ') or toc_line.startswith('### '):
                            break
                        if toc_line:
                            toc_lines.append(toc_line)
                        i += 1
                    parsed['toc'] = toc_lines
                    continue

                # Add new section
                parsed['sections'].append({
                    'level': 2,
                    'title': section_title,
                    'content': []
                })
                current_section_title = section_title
                section_stack = [section_title]
                i += 1
                continue

            # Handle H3 (Subsections)
            if stripped.startswith('### '):
                section_title = stripped[4:].strip()
                parsed['sections'].append({
                    'level': 3,
                    'title': section_title,
                    'content': []
                })
                current_section_title = section_title
                section_stack = [parsed['sections'][-2]['title'] if len(parsed['sections']) > 1 else '', section_title]
                i += 1
                continue

            # Handle blockquotes (consecutive lines starting with >)
            if stripped.startswith('>'):
                # Collect all consecutive blockquote lines (including internal blank lines)
                blockquote_lines = []
                blockquote_start_idx = i
                while i < len(lines):
                    current_line = lines[i].strip()
                    if current_line.startswith('>'):
                        # Add the content after the '>'
                        quote_content = current_line[1:].strip()
                        blockquote_lines.append(quote_content)
                        i += 1
                    elif not current_line:
                        # Blank line - check if next non-empty line is also a blockquote
                        peek_idx = i + 1
                        while peek_idx < len(lines) and not lines[peek_idx].strip():
                            peek_idx += 1

                        if peek_idx < len(lines) and lines[peek_idx].strip().startswith('>'):
                            # It's a continuation of blockquote, keep internal blank line
                            blockquote_lines.append('')
                            i += 1
                        else:
                            # End of blockquote
                            break
                    else:
                        # Non-blockquote content, end blockquote
                        break

                # Save the complete blockquote - split by type (Clinical Pearls, Memory Tricks, etc.)
                if blockquote_lines:
                    # Clean up trailing empty lines
                    while blockquote_lines and blockquote_lines[-1] == '':
                        blockquote_lines.pop()

                    # Split into typed groups (Clinical Pearls, Memory Tricks, Analogy)
                    typed_groups = self.split_blockquotes_by_type(blockquote_lines, current_section_title)
                    parsed['blockquotes'].extend(typed_groups)
                continue

            # Handle tables (markdown tables: header, separator, rows)
            if stripped.startswith('|'):
                # Validate this is a proper markdown table
                # Need: header row, separator row, at least one data row
                table_lines = [stripped]
                header_row = stripped

                # Look for separator row (must be next non-empty line)
                sep_idx = i + 1
                while sep_idx < len(lines) and not lines[sep_idx].strip():
                    sep_idx += 1

                if sep_idx < len(lines):
                    sep_line = lines[sep_idx].strip()
                    # Valid separator: contains | and - and nothing else
                    is_valid_separator = (sep_line.startswith('|') and
                                        '|' in sep_line and
                                        all(c in '|-: ' for c in sep_line))

                    if is_valid_separator:
                        # This is a valid table, collect all table rows
                        current_table = {
                            'headers': [c.strip() for c in header_row.split('|')[1:-1]],
                            'rows': [],
                            'section': current_section_title
                        }

                        # Skip to after separator
                        i = sep_idx + 1

                        # Collect data rows
                        while i < len(lines):
                            row_line = lines[i].strip()
                            if row_line.startswith('|') and any(c not in '|- ' for c in row_line):
                                # This is a data row
                                cells = [c.strip() for c in row_line.split('|')[1:-1]]
                                current_table['rows'].append(cells)
                                i += 1
                            elif not row_line:
                                # Empty line, might be end of table
                                i += 1
                                break
                            else:
                                # Non-table content
                                break

                        # Add valid table to list
                        if current_table['headers'] and current_table['rows']:
                            parsed['tables'].append(current_table)
                        continue

            # Handle empty lines
            if not stripped:
                i += 1
                continue

            # Handle list items
            if stripped.startswith('- ') or stripped.startswith('* '):
                list_item = stripped[2:].strip()
                if parsed['sections']:
                    parsed['sections'][-1]['content'].append(('list', list_item))
                i += 1
                continue

            # Handle regular text
            if stripped and parsed['sections']:
                parsed['sections'][-1]['content'].append(('text', stripped))

            i += 1

        return parsed

    def format_parsed_preview(self, parsed):
        """Format parsed markdown for preview display"""
        lines = []
        lines.append("=" * 60)
        lines.append(f"TITLE: {parsed['title']}")
        lines.append("=" * 60)
        lines.append("")

        # Show Table of Contents if present
        if parsed.get('toc'):
            lines.append(f"TABLE OF CONTENTS ({len(parsed['toc'])} items):")
            for toc_item in parsed['toc'][:5]:
                lines.append(f"  • {toc_item}")
            if len(parsed['toc']) > 5:
                lines.append(f"  ... and {len(parsed['toc']) - 5} more items")
            lines.append("")

        lines.append(f"SECTIONS FOUND: {len(parsed['sections'])}")
        for sec in parsed['sections'][:10]:
            prefix = "  " if sec['level'] == 2 else "    "
            lines.append(f"{prefix}[H{sec['level']}] {sec['title']}")
        if len(parsed['sections']) > 10:
            lines.append(f"  ... and {len(parsed['sections']) - 10} more sections")

        lines.append("")
        lines.append(f"TABLES FOUND: {len(parsed['tables'])}")
        for i, table in enumerate(parsed['tables'][:5]):
            lines.append(f"  Table {i+1}: {len(table['headers'])} columns, {len(table['rows'])} rows (Section: {table['section'][:40]}...)")
            lines.append(f"    Headers: {', '.join(table['headers'][:3])}...")
        if len(parsed['tables']) > 5:
            lines.append(f"  ... and {len(parsed['tables']) - 5} more tables")

        lines.append("")
        lines.append(f"BLOCKQUOTES (Clinical Pearls): {len(parsed['blockquotes'])}")
        for i, bq in enumerate(parsed['blockquotes'][:5]):
            preview = bq['content'][:50] + "..." if len(bq['content']) > 50 else bq['content']
            lines.append(f"  [{i+1}] Section: {bq['section'][:30]}... | {preview}")
        if len(parsed['blockquotes']) > 5:
            lines.append(f"  ... and {len(parsed['blockquotes']) - 5} more blockquotes")

        lines.append("")
        lines.append("=" * 60)
        lines.append("✓ All elements properly associated with their sections")
        lines.append("✓ Ready to convert to Word document")
        lines.append("=" * 60)

        return '\n'.join(lines)

    def convert_markdown_to_word(self):
        """Convert loaded markdown to a styled Word document"""
        if not DOCX_AVAILABLE:
            messagebox.showerror(
                "Missing Dependency",
                "python-docx is not installed.\n\nInstall it with:\npip install python-docx"
            )
            return

        if not self.markdown_content:
            messagebox.showwarning("No File", "Please select a markdown file first.")
            return

        # Parse the markdown
        parsed = self.parse_markdown(self.markdown_content)

        # Get output file path
        default_name = "Study_Guide.docx"
        if self.markdown_file_path:
            default_name = self.markdown_file_path.stem + ".docx"

        output_path = filedialog.asksaveasfilename(
            title="Save Word Document",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("All files", "*.*")],
            initialfile=default_name
        )

        if not output_path:
            return

        try:
            # Create Word document
            doc = Document()

            # Set margins
            for section in doc.sections:
                section.top_margin = Inches(0.8)
                section.bottom_margin = Inches(0.8)
                section.left_margin = Inches(0.8)
                section.right_margin = Inches(0.8)

            # Get selected theme
            theme_name = self.word_theme_var.get()
            theme = WORD_COLOR_THEMES.get(theme_name, WORD_COLOR_THEMES['Purple - General Topics'])

            # Add title
            if parsed['title']:
                title = doc.add_heading(parsed['title'], 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in title.runs:
                    run.font.color.rgb = RGBColor(*theme['header_text'])

            # Add Table of Contents if it exists (at the beginning, right after title)
            if parsed['toc']:
                doc.add_heading('Table of Contents', level=2)
                for toc_line in parsed['toc']:
                    doc.add_paragraph(toc_line, style='List Bullet')
                doc.add_paragraph()  # Blank line after TOC

            # Process sections with integrated tables
            color_index = 0
            for section in parsed['sections']:
                # Add section heading
                level = min(section['level'], 3)
                heading = doc.add_heading(section['title'], level)
                for run in heading.runs:
                    run.font.color.rgb = RGBColor(*theme['header_text'])

                # Add section content
                for content_type, content in section['content']:
                    if content_type == 'text':
                        doc.add_paragraph(content)
                    elif content_type == 'list':
                        doc.add_paragraph(content, style='List Bullet')

                # Add tables that belong to this section (INLINE, not at the end)
                for table_data in parsed['tables']:
                    # Check if this table belongs to the current section
                    if table_data.get('section') == section['title']:
                        doc.add_paragraph()  # Space before table

                        # Determine colors
                        if theme['name'] == 'auto':
                            color_set = COLOR_SETS[color_index % len(COLOR_SETS)]
                            header_color = color_set['header']
                            data_color = color_set['main']
                            color_index += 1
                        else:
                            header_color = theme['header']
                            data_color = theme['light']

                        # Create table
                        num_cols = len(table_data['headers'])
                        num_rows = len(table_data['rows']) + 1  # +1 for header
                        table = doc.add_table(rows=num_rows, cols=num_cols)
                        table.style = 'Table Grid'
                        table.alignment = WD_TABLE_ALIGNMENT.LEFT

                        # Header row
                        for col_idx, header_text in enumerate(table_data['headers']):
                            cell = table.rows[0].cells[col_idx]
                            cell.text = header_text
                            self._set_cell_shading(cell, header_color)
                            for para in cell.paragraphs:
                                for run in para.runs:
                                    run.font.bold = True
                                    run.font.size = Pt(11)

                        # Data rows
                        for row_idx, row_data in enumerate(table_data['rows'], 1):
                            for col_idx, cell_text in enumerate(row_data):
                                if col_idx < num_cols:
                                    cell = table.rows[row_idx].cells[col_idx]
                                    cell.text = cell_text
                                    self._set_cell_shading(cell, data_color)
                                    for para in cell.paragraphs:
                                        for run in para.runs:
                                            run.font.size = Pt(10)

                # Add blockquotes that belong to this section (INLINE)
                for bq in parsed['blockquotes']:
                    # Check if this blockquote belongs to the current section
                    if bq.get('section') == section['title']:
                        doc.add_paragraph()

                        # Use the explicit type field from parsing, with fallback to content detection
                        content = bq['content']
                        bq_type = bq.get('type', 'clinical')

                        # Create a single-cell table for the box effect
                        box_table = doc.add_table(rows=1, cols=1)
                        box_table.style = 'Table Grid'
                        cell = box_table.rows[0].cells[0]

                        # Format based on type
                        if bq_type == 'memory':
                            cell.text = "💡 MEMORY TRICKS & MNEMONICS\n\n" + content
                            self._set_cell_shading(cell, MNEMONIC_BG)
                        elif bq_type == 'analogy':
                            cell.text = "🔗 ANALOGY\n\n" + content
                            self._set_cell_shading(cell, ANALOGY_BOX_BG)
                        else:
                            # Default to clinical pearls
                            cell.text = "📋 CLINICAL PEARLS & HIGH-YIELD\n\n" + content
                            self._set_cell_shading(cell, CLINICAL_PEARL_BG)

                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.font.size = Pt(10)

            # Save document
            doc.save(output_path)

            messagebox.showinfo(
                "Success",
                f"Word document created successfully!\n\nLocation:\n{output_path}"
            )

            # Open the file
            self.open_file(output_path)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to create Word document:\n\n{str(e)}")

    def _set_cell_shading(self, cell, hex_color):
        """Set cell background color in Word table"""
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), hex_color)
        cell._element.get_or_add_tcPr().append(shading_elm)


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    root = tk.Tk()
    app = ExcelMasterChartApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
