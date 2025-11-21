from setuptools import setup

APP = ['excel_master_chart_app.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'packages': ['tkinter', 'tksheet', 'openpyxl'],
    'includes': ['openpyxl.styles', 'openpyxl.utils', 'openpyxl.cell', 'openpyxl.cell.cell', 'tksheet'],
    'iconfile': 'app_icon.icns',
    'site_packages': True,
    'excludes': ['numpy', 'matplotlib', 'scipy', 'pkg_resources'],  # Exclude pkg_resources
    'plist': {
        'CFBundleName': 'Excel Master Chart Creator',
        'CFBundleDisplayName': 'Excel Master Chart Creator',
        'CFBundleGetInfoString': "Medical Study Guide Creator",
        'CFBundleIdentifier': "com.yourname.excelmasterchart",
        'CFBundleVersion': "2.6.0",
        'CFBundleShortVersionString': "2.6.0",
        'LSUIElement': False,  # Show in Dock
        'NSHighResolutionCapable': True,  # Support Retina displays
    }
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)


