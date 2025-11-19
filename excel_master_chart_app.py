#!/usr/bin/env python3
"""
Excel Master Chart Creator - Desktop Application
A tkinter-based GUI application for creating formatted Excel master charts
with auto-color assignment and professional formatting.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import platform
import os
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================================
# COLOR CONSTANTS - 10-Color Rotation for Auto-Assignment
# ============================================================================
PASTEL_COLORS = [
    "#D9E2F3",  # Color 0: Ice Blue
    "#C8E6C9",  # Color 1: Seafoam
    "#D1C4E9",  # Color 2: Light Orchid
    "#F7E7CE",  # Color 3: Champagne
    "#BDD7EE",  # Color 4: Sky Blue
    "#F0F8FF",  # Color 5: Pale Azure
    "#FCE4EC",  # Color 6: Blush Pink
    "#EDE7F6",  # Color 7: Soft Lilac
    "#FFE8D6",  # Color 8: Soft Tangerine
    "#BBDEFB",  # Color 9: Powder Blue
]

# Header formatting constants
HEADER_BG_COLOR = "#4472C4"  # Dark blue
HEADER_FONT_COLOR = "#FFFFFF"  # White
DATA_FONT_COLOR = "#000000"  # Black
BORDER_COLOR = "#FFFFFF"  # White (invisible borders)

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
    "Custom": []  # Will be filled by user
}

# Column width mapping based on column name patterns
COLUMN_WIDTH_MAP = {
    "short": 12,    # Route, Status, Type
    "medium": 22,   # Drug Class, Condition, Category, Test Name
    "name": 28,     # Drug Name, Brand Name
    "long": 35      # Mechanism, Uses, Adverse Effects, etc.
}

# ============================================================================
# MAIN APPLICATION CLASS
# ============================================================================
class ExcelMasterChartApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Master Chart Creator")
        self.root.geometry("1200x700")

        # Data storage
        self.current_columns = []
        self.data_rows = []
        self.current_preset = tk.StringVar(value="Drug Chart (11 columns)")
        self.output_filename = tk.StringVar(value="Master_Chart.xlsx")
        self.output_directory = tk.StringVar(value=str(Path.home() / "Desktop"))

        # Setup UI
        self.setup_ui()
        self.load_preset()

    def setup_ui(self):
        """Create the main user interface"""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # ====================================================================
        # SECTION 1: Column Setup
        # ====================================================================
        setup_frame = ttk.LabelFrame(main_frame, text="Section 1: Column Setup", padding="10")
        setup_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Preset selector
        ttk.Label(setup_frame, text="Select Preset:").grid(row=0, column=0, sticky=tk.W, pady=5)
        preset_combo = ttk.Combobox(
            setup_frame,
            textvariable=self.current_preset,
            values=list(COLUMN_PRESETS.keys()),
            state="readonly",
            width=30
        )
        preset_combo.grid(row=0, column=1, sticky=tk.W, padx=(10, 20), pady=5)
        preset_combo.bind("<<ComboboxSelected>>", self.on_preset_change)

        # Custom columns button
        self.custom_btn = ttk.Button(
            setup_frame,
            text="Define Custom Columns...",
            command=self.define_custom_columns
        )
        self.custom_btn.grid(row=0, column=2, sticky=tk.W, padx=10, pady=5)

        # Column display
        ttk.Label(setup_frame, text="Current Columns:").grid(row=1, column=0, sticky=(tk.W, tk.N), pady=5)
        self.columns_text = tk.Text(setup_frame, height=3, width=80, wrap=tk.WORD)
        self.columns_text.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        self.columns_text.config(state=tk.DISABLED)

        # ====================================================================
        # SECTION 2: Data Entry Grid
        # ====================================================================
        grid_frame = ttk.LabelFrame(main_frame, text="Section 2: Data Entry Grid", padding="10")
        grid_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.rowconfigure(0, weight=1)

        # Create treeview with scrollbars
        tree_container = ttk.Frame(grid_frame)
        tree_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_container, orient="vertical")
        hsb = ttk.Scrollbar(tree_container, orient="horizontal")

        # Treeview
        self.tree = ttk.Treeview(
            tree_container,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            selectmode="browse",
            height=15
        )

        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Double-click to edit
        self.tree.bind("<Double-1>", self.on_double_click)

        # Button panel
        button_panel = ttk.Frame(grid_frame)
        button_panel.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

        ttk.Button(button_panel, text="Add Row", command=self.add_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_panel, text="Delete Selected Row", command=self.delete_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_panel, text="Clear All Data", command=self.clear_all_data).pack(side=tk.LEFT, padx=5)

        # Row count label
        self.row_count_label = ttk.Label(button_panel, text="Rows: 0")
        self.row_count_label.pack(side=tk.RIGHT, padx=10)

        # Color preview button
        ttk.Button(button_panel, text="Preview Colors", command=self.preview_colors).pack(side=tk.RIGHT, padx=5)

        # Save/Load buttons
        ttk.Separator(button_panel, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(button_panel, text="Save Data (JSON)", command=self.save_data_json).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_panel, text="Load Data (JSON)", command=self.load_data_json).pack(side=tk.LEFT, padx=5)

        # ====================================================================
        # SECTION 3: Export Controls
        # ====================================================================
        export_frame = ttk.LabelFrame(main_frame, text="Section 3: Export Controls", padding="10")
        export_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))

        # Filename
        ttk.Label(export_frame, text="Output Filename:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(export_frame, textvariable=self.output_filename, width=40).grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)

        # Directory
        ttk.Label(export_frame, text="Save Location:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(export_frame, textvariable=self.output_directory, width=50).grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Button(export_frame, text="Browse...", command=self.browse_directory).grid(row=1, column=2, sticky=tk.W, pady=5)

        # Export button (large and prominent)
        export_btn = ttk.Button(
            export_frame,
            text="🚀 Export to Excel",
            command=self.export_to_excel
        )
        export_btn.grid(row=2, column=0, columnspan=3, pady=15)

    def load_preset(self):
        """Load the selected preset columns"""
        preset_name = self.current_preset.get()
        self.current_columns = COLUMN_PRESETS[preset_name].copy()
        self.update_columns_display()
        self.setup_treeview()

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
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Enter column names (one per line):").pack(pady=10, padx=10)

        text = tk.Text(dialog, height=12, width=40)
        text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Pre-fill if custom columns already exist
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

    def update_columns_display(self):
        """Update the column display text"""
        self.columns_text.config(state=tk.NORMAL)
        self.columns_text.delete("1.0", tk.END)
        self.columns_text.insert("1.0", ", ".join(self.current_columns))
        self.columns_text.config(state=tk.DISABLED)

    def setup_treeview(self):
        """Setup the treeview with current columns"""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Configure columns
        self.tree["columns"] = self.current_columns
        self.tree["show"] = "tree headings"  # Show both tree column and headings

        # Configure tree column (row numbers)
        self.tree.column("#0", width=50, minwidth=50, anchor=tk.CENTER)
        self.tree.heading("#0", text="#")

        # Configure data columns
        for col in self.current_columns:
            width = self.calculate_column_width(col)
            self.tree.column(col, width=width, minwidth=50, anchor=tk.W)
            self.tree.heading(col, text=col)

        # Clear data rows
        self.data_rows = []
        self.update_row_count()

    def calculate_column_width(self, col_name):
        """Calculate column width based on name"""
        col_lower = col_name.lower()

        # Short columns
        if any(word in col_lower for word in ["route", "status", "type"]):
            return 80
        # Medium columns
        elif any(word in col_lower for word in ["class", "condition", "category", "test"]):
            return 150
        # Name columns
        elif any(word in col_lower for word in ["name", "brand"]):
            return 180
        # Long text columns
        else:
            return 250

    def add_row(self):
        """Add a new empty row"""
        row_data = [""] * len(self.current_columns)
        self.data_rows.append(row_data)
        row_num = len(self.data_rows)
        self.tree.insert("", tk.END, text=str(row_num), values=row_data)
        self.update_row_count()

    def delete_row(self):
        """Delete the selected row"""
        selected = self.tree.selection()
        if selected:
            item = selected[0]
            index = self.tree.index(item)
            self.tree.delete(item)
            del self.data_rows[index]

            # Re-number rows
            for idx, item in enumerate(self.tree.get_children(), 1):
                self.tree.item(item, text=str(idx))

            self.update_row_count()
        else:
            messagebox.showinfo("No Selection", "Please select a row to delete.")

    def clear_all_data(self):
        """Clear all data rows"""
        if self.data_rows:
            if messagebox.askyesno("Clear All Data", "Are you sure you want to clear all data?"):
                for item in self.tree.get_children():
                    self.tree.delete(item)
                self.data_rows = []
                self.update_row_count()
        else:
            messagebox.showinfo("No Data", "There is no data to clear.")

    def on_double_click(self, event):
        """Handle double-click to edit cell"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        # Get clicked item and column
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)

        if not item or column == "#0":  # Don't edit row number
            return

        # Get column index
        col_index = int(column.replace("#", "")) - 1

        # Get current value
        values = list(self.tree.item(item, "values"))
        current_value = values[col_index] if col_index < len(values) else ""

        # Get cell bbox
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return

        # Create entry widget
        entry = tk.Entry(self.tree)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()

        # Position entry
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])

        def save_edit(event=None):
            new_value = entry.get()
            values[col_index] = new_value
            self.tree.item(item, values=values)

            # Update data storage
            row_index = self.tree.index(item)
            self.data_rows[row_index] = values

            entry.destroy()

        def cancel_edit(event=None):
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<Escape>", cancel_edit)
        entry.bind("<FocusOut>", save_edit)

    def update_row_count(self):
        """Update the row count label"""
        count = len(self.data_rows)
        self.row_count_label.config(text=f"Rows: {count}")

    def preview_colors(self):
        """Show a preview of color assignments"""
        if not self.data_rows:
            messagebox.showinfo("No Data", "Please add some data first to preview colors.")
            return

        # Calculate color assignments
        color_map = self.calculate_color_assignments()

        # Create preview window
        preview = tk.Toplevel(self.root)
        preview.title("Color Preview")
        preview.geometry("400x400")

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
        ttk.Label(scrollable_frame, text="Color Assignments:", font=("", 12, "bold")).pack(pady=10)

        for group_name, color in sorted(color_map.items()):
            frame = tk.Frame(scrollable_frame, bg=color, relief=tk.RAISED, borderwidth=2)
            frame.pack(fill=tk.X, padx=10, pady=2)

            label = tk.Label(frame, text=f"{group_name}", bg=color, anchor=tk.W, padx=10)
            label.pack(fill=tk.X, pady=5)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def calculate_color_assignments(self):
        """Calculate color assignments for each group"""
        if not self.data_rows or not self.current_columns:
            return {}

        color_map = {}
        color_index = 0
        last_group = None

        for row in self.data_rows:
            if row:  # Skip empty rows
                group_name = row[0]  # First column is the grouping column
                if group_name and group_name != last_group:
                    if group_name not in color_map:
                        color_map[group_name] = PASTEL_COLORS[color_index % len(PASTEL_COLORS)]
                        color_index += 1
                    last_group = group_name

        return color_map

    def browse_directory(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(initialdir=self.output_directory.get())
        if directory:
            self.output_directory.set(directory)

    def save_data_json(self):
        """Save current data to JSON file"""
        if not self.data_rows:
            messagebox.showinfo("No Data", "There is no data to save.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=self.output_directory.get()
        )

        if filename:
            data = {
                "preset": self.current_preset.get(),
                "columns": self.current_columns,
                "rows": self.data_rows
            }

            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

            messagebox.showinfo("Success", f"Data saved to:\n{filename}")

    def load_data_json(self):
        """Load data from JSON file"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=self.output_directory.get()
        )

        if filename:
            try:
                with open(filename, "r", encoding="utf-8") as f:
                    data = json.load(f)

                # Load preset
                preset = data.get("preset", "Custom")
                if preset == "Custom":
                    COLUMN_PRESETS["Custom"] = data["columns"]

                self.current_preset.set(preset)
                self.current_columns = data["columns"]
                self.data_rows = data["rows"]

                # Update UI
                self.update_columns_display()
                self.setup_treeview()

                # Populate treeview
                for idx, row_data in enumerate(self.data_rows, 1):
                    self.tree.insert("", tk.END, text=str(idx), values=row_data)

                self.update_row_count()
                messagebox.showinfo("Success", f"Data loaded from:\n{filename}")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to load data:\n{str(e)}")

    def export_to_excel(self):
        """Export data to formatted Excel file"""
        # Validation
        if not self.data_rows:
            messagebox.showwarning("No Data", "Please add some data before exporting.")
            return

        if not self.output_filename.get():
            messagebox.showwarning("No Filename", "Please enter an output filename.")
            return

        # Ensure .xlsx extension
        filename = self.output_filename.get()
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"

        output_path = Path(self.output_directory.get()) / filename

        try:
            # Create workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Master Chart"

            # Calculate color assignments
            color_map = self.calculate_color_assignments()

            # ================================================================
            # HEADER ROW
            # ================================================================
            ws.append(self.current_columns)

            # Header styling
            header_font = Font(name="Calibri", size=12, bold=True, color=HEADER_FONT_COLOR[1:])
            header_fill = PatternFill(start_color=HEADER_BG_COLOR[1:], end_color=HEADER_BG_COLOR[1:], fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for col_idx, col_name in enumerate(self.current_columns, 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment

            # Set header row height
            ws.row_dimensions[1].height = 25

            # ================================================================
            # DATA ROWS
            # ================================================================
            current_color = None
            last_group = None

            for row_idx, row_data in enumerate(self.data_rows, 2):
                ws.append(row_data)

                # Determine color for this row
                group_name = row_data[0] if row_data else ""
                if group_name and group_name != last_group:
                    current_color = color_map.get(group_name, PASTEL_COLORS[0])
                    last_group = group_name

                # Apply styling
                data_fill = PatternFill(start_color=current_color[1:] if current_color else PASTEL_COLORS[0][1:],
                                       end_color=current_color[1:] if current_color else PASTEL_COLORS[0][1:],
                                       fill_type="solid")
                data_alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                white_border = Border(
                    left=Side(style="thin", color=BORDER_COLOR[1:]),
                    right=Side(style="thin", color=BORDER_COLOR[1:]),
                    top=Side(style="thin", color=BORDER_COLOR[1:]),
                    bottom=Side(style="thin", color=BORDER_COLOR[1:])
                )

                for col_idx, col_name in enumerate(self.current_columns, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.fill = data_fill
                    cell.alignment = data_alignment
                    cell.border = white_border

                    # First column bold
                    if col_idx == 1:
                        cell.font = Font(name="Calibri", size=10, bold=True, color=DATA_FONT_COLOR[1:])
                    else:
                        cell.font = Font(name="Calibri", size=10, color=DATA_FONT_COLOR[1:])

            # ================================================================
            # COLUMN WIDTHS
            # ================================================================
            for col_idx, col_name in enumerate(self.current_columns, 1):
                width = self.get_excel_column_width(col_name)
                ws.column_dimensions[self.get_column_letter(col_idx)].width = width

            # ================================================================
            # FREEZE PANES
            # ================================================================
            ws.freeze_panes = "A2"  # Freeze header row

            # Save workbook
            wb.save(output_path)

            # Success message
            messagebox.showinfo(
                "Export Successful",
                f"Excel file created successfully!\n\nLocation:\n{output_path}\n\nRows exported: {len(self.data_rows)}"
            )

            # Auto-open the file
            self.open_file(output_path)

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export Excel file:\n\n{str(e)}")

    def get_excel_column_width(self, col_name):
        """Get Excel column width based on column name"""
        col_lower = col_name.lower()

        # Short columns
        if any(word in col_lower for word in ["route", "status", "type"]):
            return COLUMN_WIDTH_MAP["short"]
        # Medium columns
        elif any(word in col_lower for word in ["class", "condition", "category", "test"]):
            return COLUMN_WIDTH_MAP["medium"]
        # Name columns
        elif any(word in col_lower for word in ["name", "brand"]):
            return COLUMN_WIDTH_MAP["name"]
        # Long text columns
        else:
            return COLUMN_WIDTH_MAP["long"]

    def get_column_letter(self, col_idx):
        """Convert column index to Excel column letter"""
        result = ""
        while col_idx > 0:
            col_idx -= 1
            result = chr(col_idx % 26 + 65) + result
            col_idx //= 26
        return result

    def open_file(self, filepath):
        """Open file with default application"""
        try:
            if platform.system() == "Darwin":  # macOS
                subprocess.run(["open", str(filepath)], check=True)
            elif platform.system() == "Windows":
                os.startfile(str(filepath))
            else:  # Linux
                subprocess.run(["xdg-open", str(filepath)], check=True)
        except Exception as e:
            print(f"Could not auto-open file: {e}")

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def main():
    root = tk.Tk()
    app = ExcelMasterChartApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
