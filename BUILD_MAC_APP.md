# Building the macOS Application

## Quick Build

To create a macOS application bundle that you can run without opening Terminal:

```bash
./build_app.sh
```

This will:
1. Build the app using py2app
2. Install it to `/Applications/Excel Master Chart Creator.app`

## Launch the App

After building, you can launch the app in several ways:

1. **Spotlight**: Press `Cmd + Space`, type "Excel Master Chart", press Enter
2. **Applications folder**: Open Finder → Applications → Double-click "Excel Master Chart Creator"
3. **Launchpad**: Click the Launchpad icon and find "Excel Master Chart Creator"

The app will run without showing Python or Terminal!

## Manual Build Instructions

If you prefer to build manually:

```bash
# Clean previous builds
rm -rf build dist

# Build the app (use -A flag for alias mode - faster and works best)
python3 setup.py py2app -A

# Copy to Applications
cp -R "dist/Excel Master Chart Creator.app" /Applications/
```

## Troubleshooting

### "Launch error" message

If you see a launch error, try:

1. Make sure all dependencies are installed:
   ```bash
   pip3 install tksheet openpyxl
   ```

2. Rebuild the app:
   ```bash
   ./build_app.sh
   ```

### App won't open (Gatekeeper)

If macOS says the app is "damaged":

```bash
xattr -cr "/Applications/Excel Master Chart Creator.app"
```

Then try opening it again.

### After updating the Python code

Whenever you modify `excel_master_chart_app.py`, rebuild the app:

```bash
./build_app.sh
```

## About Alias Mode (-A flag)

The build uses py2app's "alias" mode which creates a lightweight app that links to your Python installation. This means:

- **Pros**: Fast builds, smaller app size, changes to Python code require only a quick rebuild
- **Cons**: Requires Python and dependencies to remain installed on your Mac

If you need a standalone app that bundles everything (to share with others), use:

```bash
python3 setup.py py2app
```

This creates a fully self-contained app but takes longer to build and may encounter code-signing issues on some systems.
