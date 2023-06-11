import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    'includes': [],
    'excludes': [],
    'packages': [],
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable('blattfunktion.py', base=base, icon='C:\\Users\\MSI\\Desktop\\ico.ico')
]

setup(
    name="Sheet Function",
    version="1.0",
    description="Sheet Function Application",
    options={'build_exe': build_exe_options},
    executables=executables
)
