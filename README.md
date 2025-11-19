# Excel Master Chart Creator

A powerful desktop application for macOS that generates beautifully formatted Excel master charts with automatic color-coding and professional styling. No more manually editing Python files - just enter your data visually and export!

![Application Type: Desktop GUI](https://img.shields.io/badge/Type-Desktop%20GUI-blue)
![Platform: macOS](https://img.shields.io/badge/Platform-macOS-lightgrey)
![Python: 3.7+](https://img.shields.io/badge/Python-3.7%2B-green)

## Features

### ✨ Core Features
- **Visual Data Entry**: Spreadsheet-like interface - no code editing required
- **Auto-Color Assignment**: Automatically assigns colors when grouping changes
- **Pre-set Templates**: Drug Chart (11 columns), Condition Chart (7 columns), Lab Values (5 columns)
- **Custom Columns**: Define your own column layouts
- **Professional Formatting**: Dark blue headers, pastel-colored rows, frozen panes
- **Auto-Open**: Exported Excel files open automatically
- **Save/Load Data**: Save your work as JSON and reload later

### 🎨 Auto-Color System
The application automatically assigns one of 10 beautiful pastel colors based on the first column (Drug Class, Condition, etc.):

1. Ice Blue (#D9E2F3)
2. Seafoam (#C8E6C9)
3. Light Orchid (#D1C4E9)
4. Champagne (#F7E7CE)
5. Sky Blue (#BDD7EE)
6. Pale Azure (#F0F8FF)
7. Blush Pink (#FCE4EC)
8. Soft Lilac (#EDE7F6)
9. Soft Tangerine (#FFE8D6)
10. Powder Blue (#BBDEFB)

Colors change automatically when the grouping column value changes!

### 📊 Excel Output Features
- **Header Row**: Dark blue background (#4472C4), white bold text, centered, frozen
- **Data Rows**: Auto-colored backgrounds, black text, word wrap enabled
- **First Column**: Bold formatting
- **Borders**: Invisible white borders for clean appearance
- **Column Widths**: Automatically sized based on content type
- **Freeze Panes**: Header row stays visible when scrolling

## Installation

### Prerequisites
- **Python 3.7 or higher** (macOS usually comes with Python 3)
- **pip** (Python package installer)

### Quick Setup

1. **Clone or download this repository**
   ```bash
   cd ~/Desktop
   git clone <repository-url>
   cd Tkinter-Study-Guides
   ```

2. **Install required packages**
   ```bash
   pip3 install openpyxl
   ```

3. **Run the application**
   ```bash
   python3 excel_master_chart_app.py
   ```

That's it! The application should open.

## Usage Guide

### 1. Select Your Column Layout

**Choose a preset:**
- **Drug Chart (11 columns)**: Drug Class, Drug Name (Brand), Route, Mechanism, Uses, Adverse Effects, Contraindications, Resistance, Drug Interactions, Drug Combinations, Special Considerations
- **Condition Chart (7 columns)**: Condition, Epidemiology, Risk Factors, Clinical Presentation, Diagnostics, Labs, Treatment
- **Lab Values (5 columns)**: Test Name, Normal Range, Increased In, Decreased In, Clinical Significance
- **Custom**: Define your own columns

### 2. Enter Your Data

- **Double-click any cell** to edit it
- **Add Row**: Adds a new empty row
- **Delete Selected Row**: Removes the currently selected row
- **Clear All Data**: Removes all data (with confirmation)

**Tips:**
- The first column is the "grouping" column - colors change when this value changes
- You can add 50+ rows with smooth scrolling
- Data is saved in real-time as you edit

### 3. Preview Colors (Optional)

Click **"Preview Colors"** to see how your data will be color-coded in Excel.

### 4. Save Your Work (Optional)

- Click **"Save Data (JSON)"** to save your current work
- Click **"Load Data (JSON)"** to restore previously saved data
- JSON files are portable and can be shared

### 5. Export to Excel

1. Enter a filename (default: `Master_Chart.xlsx`)
2. Choose save location (default: Desktop)
3. Click **"🚀 Export to Excel"**
4. The Excel file will be created and automatically opened!

### Example Workflow

```
1. Select "Drug Chart (11 columns)"
2. Add rows for different drug classes (e.g., "Beta Blockers", "ACE Inhibitors")
3. Fill in drug information
4. Click "Preview Colors" to verify color assignments
5. Click "Export to Excel"
6. Excel opens with beautifully formatted chart!
```

## Keyboard Shortcuts

- **Double-click**: Edit cell
- **Return/Enter**: Save cell edit
- **Escape**: Cancel cell edit

## Creating a macOS .app Bundle (Optional)

You can package this application as a native macOS .app that you can double-click to launch!

### Method 1: Using py2app (Recommended)

1. **Install py2app**
   ```bash
   pip3 install py2app
   ```

2. **Create setup file**
   ```bash
   py2applet --make-setup excel_master_chart_app.py
   ```

3. **Build the .app**
   ```bash
   python3 setup.py py2app
   ```

4. **Find your app**
   The application will be in `dist/excel_master_chart_app.app`

5. **Move to Applications folder**
   ```bash
   cp -r dist/excel_master_chart_app.app /Applications/
   ```

6. **Launch from Applications!**
   Now you can double-click the app from your Applications folder or Launchpad.

### Method 2: Using PyInstaller

1. **Install PyInstaller**
   ```bash
   pip3 install pyinstaller
   ```

2. **Build the app**
   ```bash
   pyinstaller --onefile --windowed --name="Excel Master Chart Creator" excel_master_chart_app.py
   ```

3. **Find your app**
   The application will be in `dist/Excel Master Chart Creator.app`

### Troubleshooting .app Creation

**Issue**: "App can't be opened because it is from an unidentified developer"

**Solution**:
```bash
# Right-click the app → Open (instead of double-clicking)
# Or remove quarantine attribute:
xattr -cr /Applications/excel_master_chart_app.app
```

**Issue**: Module not found errors

**Solution**: Make sure openpyxl is installed in the same Python environment:
```bash
pip3 install --upgrade openpyxl
```

## Advanced Configuration

### Customizing Colors

Edit the `PASTEL_COLORS` list in `excel_master_chart_app.py`:

```python
PASTEL_COLORS = [
    "#D9E2F3",  # Your color 1
    "#C8E6C9",  # Your color 2
    # ... add more colors
]
```

### Customizing Column Widths

Edit the `COLUMN_WIDTH_MAP` in `excel_master_chart_app.py`:

```python
COLUMN_WIDTH_MAP = {
    "short": 12,    # Adjust as needed
    "medium": 22,
    "name": 28,
    "long": 35
}
```

## File Structure

```
Tkinter-Study-Guides/
├── excel_master_chart_app.py    # Main application
├── README.md                     # This file
└── example_data/                 # (Optional) Sample JSON files
```

## Requirements

- **Python**: 3.7+
- **openpyxl**: Excel file generation
- **tkinter**: GUI framework (included with Python)

## Platform Compatibility

| Platform | Status | Notes |
|----------|--------|-------|
| macOS | ✅ Fully Supported | Primary platform |
| Windows | ⚠️ Should work | File opening may differ |
| Linux | ⚠️ Should work | File opening may differ |

## Troubleshooting

### Application won't start
```bash
# Check Python version (should be 3.7+)
python3 --version

# Check if tkinter is available
python3 -m tkinter

# Reinstall openpyxl
pip3 install --upgrade openpyxl
```

### Excel file won't open automatically
- The file is still created successfully
- Navigate to the save location and open manually
- On macOS: Check if Microsoft Excel or Numbers is installed

### Colors not appearing correctly
- Make sure you have data in the first column (grouping column)
- Use "Preview Colors" to verify assignments
- Different groups should have different values in the first column

### Can't edit cells
- Make sure to **double-click** the cell (single-click just selects)
- Press Return to save, Escape to cancel

## Tips & Best Practices

1. **Save frequently**: Use "Save Data (JSON)" to preserve your work
2. **Preview before export**: Use "Preview Colors" to check color assignments
3. **Group logically**: Ensure items in the same group have identical first-column values
4. **Use descriptive names**: Clear column names help with readability
5. **Test with small data first**: Try exporting 5-10 rows before adding 50+

## Examples

### Example 1: Drug Chart
```
Drug Class: Beta Blockers
Drug Name: Metoprolol
Route: Oral
Mechanism: Blocks beta-1 receptors
Uses: Hypertension, heart failure
...
```

### Example 2: Condition Chart
```
Condition: Diabetes Mellitus Type 2
Epidemiology: 10% of US adults
Risk Factors: Obesity, sedentary lifestyle
...
```

## Contributing

Feel free to submit issues or pull requests!

## License

MIT License - feel free to use and modify as needed.

## Support

If you encounter issues:
1. Check the Troubleshooting section above
2. Verify all prerequisites are installed
3. Try running with `python3 -v excel_master_chart_app.py` for verbose output

## Acknowledgments

- Built with Python's tkinter framework
- Excel generation powered by openpyxl
- Designed for medical/pharmacy education workflows

---

**Enjoy creating beautiful Excel master charts without touching code!** 🚀📊
