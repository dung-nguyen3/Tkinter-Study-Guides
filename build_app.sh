#!/bin/bash
# Build script for Excel Master Chart Creator macOS app

echo "Building Excel Master Chart Creator..."

# Clean previous builds
rm -rf build dist

# Build the app in alias mode (lightweight, fast)
python3 setup.py py2app -A

if [ $? -eq 0 ]; then
    echo "Build successful!"

    # Copy to Applications folder
    echo "Copying to /Applications..."
    cp -R "dist/Excel Master Chart Creator.app" /Applications/

    # Touch the app to update its timestamp
    touch "/Applications/Excel Master Chart Creator.app"

    # Refresh Dock to pick up new icon
    echo "Refreshing Dock..."
    killall Dock 2>/dev/null

    echo ""
    echo "✓ App installed to /Applications/Excel Master Chart Creator.app"
    echo ""
    echo "You can now launch it from:"
    echo "  - Spotlight (Cmd+Space, type 'Excel Master Chart')"
    echo "  - Applications folder"
    echo ""
    echo "Note: If the Dock shows the old Python icon when running,"
    echo "      close the app and relaunch it. The icon should update."
    echo ""
else
    echo "Build failed!"
    exit 1
fi
