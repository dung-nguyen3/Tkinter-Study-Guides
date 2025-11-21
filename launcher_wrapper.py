#!/usr/bin/env python3
"""
Wrapper to launch the app without showing Python icon in Dock
"""
import subprocess
import sys
import os

# Get the actual app path
app_path = "/Users/kimnguyen/Documents/Github/Tkinter Study Guides/excel_master_chart_app.py"

# Set environment to prevent Python framework from showing in Dock
env = os.environ.copy()
env['PYTHONPATH'] = os.path.dirname(app_path)

# Use subprocess to launch the app
# The key is to use the Python from the framework but not show it
subprocess.Popen(
    [sys.executable, app_path],
    env=env,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
    start_new_session=True
)

# Exit immediately
sys.exit(0)
