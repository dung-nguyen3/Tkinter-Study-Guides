# Quick Start Guide - Excel Master Chart Creator

## 5-Minute Setup

### 1. Install Dependencies (30 seconds)
```bash
pip3 install openpyxl
```

### 2. Run the App (10 seconds)
```bash
python3 excel_master_chart_app.py
```

### 3. Try the Example Data (2 minutes)
1. Click **"Load Data (JSON)"**
2. Navigate to `example_data/sample_drug_chart.json`
3. Click **"Preview Colors"** to see the auto-color assignments
4. Click **"🚀 Export to Excel"**
5. Excel opens automatically with your formatted chart!

### 4. Create Your Own Chart (2 minutes)
1. Select a preset from the dropdown (e.g., "Condition Chart")
2. Click **"Add Row"** several times
3. **Double-click** cells to edit them
4. Fill in your data
5. Export!

---

## Common Questions

**Q: How do I edit a cell?**
A: **Double-click** the cell, type your content, then press **Return**

**Q: How does auto-coloring work?**
A: Colors change when the first column value changes. All rows with "Beta Blockers" get one color, all rows with "ACE Inhibitors" get another color, etc.

**Q: Can I add more than 11 columns?**
A: Yes! Choose "Custom" preset and define as many columns as you want.

**Q: Where is my Excel file saved?**
A: By default, your Desktop. You can change this in Section 3 before exporting.

**Q: How do I save my work?**
A: Click "Save Data (JSON)" - you can reload this file later with "Load Data (JSON)"

**Q: Can I use this on Windows/Linux?**
A: The app should work, but it's optimized for macOS. File auto-opening may behave differently.

---

## Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Edit cell | Double-click |
| Save edit | Return/Enter |
| Cancel edit | Escape |
| Close edit | Click elsewhere |

---

## Tips for Best Results

1. **Keep first column consistent**: If using "Beta Blockers", don't also use "Beta-Blockers" or "beta blockers" - colors won't group correctly
2. **Use Preview Colors**: Always check color assignments before exporting
3. **Save frequently**: Use JSON save feature to preserve your work
4. **Start small**: Test with 5-10 rows before adding 50+
5. **Descriptive column names**: Clear names make your chart more readable

---

## Example Workflow: Creating a Drug Chart

```
Step 1: Launch app
  → python3 excel_master_chart_app.py

Step 2: Select preset
  → Choose "Drug Chart (11 columns)"

Step 3: Add data
  → Click "Add Row" 5 times
  → Double-click first cell of row 1
  → Type "Beta Blockers" → Return
  → Fill in remaining cells for that drug
  → Repeat for other drugs in different classes

Step 4: Preview
  → Click "Preview Colors"
  → Verify each drug class has unique color

Step 5: Export
  → Enter filename: "My_Cardiology_Drugs.xlsx"
  → Click "Export to Excel"
  → Excel opens automatically!
```

---

## Troubleshooting Quick Fixes

**App won't start**
```bash
# Check Python version
python3 --version  # Should be 3.7+

# Reinstall openpyxl
pip3 install --upgrade openpyxl
```

**Can't edit cells**
- Make sure you're **double-clicking** (not single-click)
- Try clicking outside the cell first, then double-click again

**Excel file is blank**
- Make sure you added data (click "Add Row")
- Check that cells aren't empty (double-click to add content)

**Colors are wrong**
- Check that first column values are exactly the same for items in the same group
- Use "Preview Colors" to see assignments before exporting

---

## Next Steps

1. ✅ Try the example data in `example_data/`
2. ✅ Create your first custom chart
3. ✅ Learn to use JSON save/load feature
4. ✅ Package as macOS .app (see README.md)
5. ✅ Customize colors and widths (advanced - edit .py file)

---

**Need more help?** See the full README.md for detailed documentation.

**Enjoy creating beautiful charts!** 🚀
