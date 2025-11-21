#!/bin/bash
# Launcher script for Excel Master Chart Creator

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Change to the app directory
cd "$SCRIPT_DIR"

# Launch the Python app without showing Python in Dock
/usr/bin/python3 "$SCRIPT_DIR/excel_master_chart_app.py" &

# Exit this script so it doesn't stay in the Dock
exit 0
