from setuptools import setup

APP = ['search_excel.py']
DATA_FILES = ['logo.jpg']    # Logo landet später in Contents/Resources
OPTIONS = {
    'argv_emulation': True,
    'packages': ['pandas','openpyxl','PIL'],
    # 'iconfile': 'appicon.icns',    # optional, nur wenn Du ein Icon hast
    'plist': {
        'CFBundleName': 'ExcelSearcher',
        'CFBundleIdentifier': 'com.example.ExcelSearcher',
        'CFBundleVersion': '1.0.0',
    },
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
