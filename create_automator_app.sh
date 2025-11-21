#!/bin/bash
# Script to create an Automator application for Excel Master Chart Creator

APP_NAME="Excel Master Chart Creator"
APP_PATH="/Applications/$APP_NAME.app"
PYTHON_SCRIPT="/Users/kimnguyen/Documents/Github/Tkinter Study Guides/excel_master_chart_app.py"
ICON_PATH="/Users/kimnguyen/Documents/Github/Tkinter Study Guides/app_icon.icns"

echo "Creating Automator app: $APP_NAME"

# Remove old app if it exists
if [ -d "$APP_PATH" ]; then
    echo "Removing old app..."
    rm -rf "$APP_PATH"
fi

# Create the app bundle structure
mkdir -p "$APP_PATH/Contents/MacOS"
mkdir -p "$APP_PATH/Contents/Resources"

# Create the Info.plist
cat > "$APP_PATH/Contents/Info.plist" << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>Application Stub</string>
    <key>CFBundleIconFile</key>
    <string>app_icon.icns</string>
    <key>CFBundleIdentifier</key>
    <string>com.yourname.excelmasterchart</string>
    <key>CFBundleInfoDictionaryVersion</key>
    <string>6.0</string>
    <key>CFBundleName</key>
    <string>Excel Master Chart Creator</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>2.6.0</string>
    <key>CFBundleVersion</key>
    <string>2.6.0</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13</string>
    <key>LSUIElement</key>
    <false/>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
EOF

# Create the executable script
cat > "$APP_PATH/Contents/MacOS/Application Stub" << 'EOFSCRIPT'
#!/bin/bash
# Launch the Python application

PYTHON_SCRIPT="/Users/kimnguyen/Documents/Github/Tkinter Study Guides/excel_master_chart_app.py"

# Launch Python script
/Library/Frameworks/Python.framework/Versions/3.13/bin/python3 "$PYTHON_SCRIPT" &

EOFSCRIPT

# Make the script executable
chmod +x "$APP_PATH/Contents/MacOS/Application Stub"

# Copy the icon
if [ -f "$ICON_PATH" ]; then
    cp "$ICON_PATH" "$APP_PATH/Contents/Resources/app_icon.icns"
    echo "Icon copied"
else
    echo "Warning: Icon file not found at $ICON_PATH"
fi

# Set the icon on the app bundle
if [ -f "$ICON_PATH" ]; then
    # Convert icns to iconset for setting
    sips -i "$ICON_PATH" >/dev/null 2>&1
    DeRez -only icns "$ICON_PATH" > /tmp/tmpicns.rsrc
    Rez -append /tmp/tmpicns.rsrc -o "$APP_PATH"$'/Contents/Resources/app_icon.icns\r'
    SetFile -a C "$APP_PATH"
    rm /tmp/tmpicns.rsrc 2>/dev/null
fi

# Touch the app to update timestamp
touch "$APP_PATH"

echo ""
echo "✓ App created at: $APP_PATH"
echo ""
echo "Launch it from:"
echo "  - Applications folder"
echo "  - Spotlight (Cmd+Space, type 'Excel Master Chart')"
echo ""
