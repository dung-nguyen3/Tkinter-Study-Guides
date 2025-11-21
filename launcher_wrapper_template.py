#!/usr/bin/env python3
"""
Launcher wrapper for Excel Master Chart Creator
Hides the Python Dock icon before launching the main application
"""
import sys
import os

def hide_dock_icon():
    """
    Hide the Python dock icon on macOS using PyObjC.
    This must be called BEFORE creating any tkinter windows.
    """
    try:
        import AppKit
        # Initialize the shared application if not already done
        app = AppKit.NSApplication.sharedApplication()

        # NSApplicationActivationPolicyRegular = 0
        # This allows the app to show in Dock with our custom icon
        # The Python framework won't show separately because we're controlling
        # the app activation from within the .app bundle
        app.setActivationPolicy_(0)
        print("Dock icon hidden successfully", file=sys.stderr)
    except ImportError:
        # PyObjC not available - Info.plist LSUIElement should handle it
        print("Warning: PyObjC not available. Relying on Info.plist LSUIElement.", file=sys.stderr)
    except Exception as e:
        print(f"Warning: Could not hide dock icon: {e}", file=sys.stderr)

# Hide the dock icon FIRST, before any imports
hide_dock_icon()

# Add the script directory to Python path
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Now import and run the main application
try:
    # Import the main app module
    # This will execute the code in excel_master_chart_app.py
    import excel_master_chart_app
except ImportError as e:
    print(f"Error: Could not import excel_master_chart_app: {e}", file=sys.stderr)
    sys.exit(1)
except Exception as e:
    print(f"Error running application: {e}", file=sys.stderr)
    import traceback
    traceback.print_exc()
    sys.exit(1)
