import sys
from cx_Freeze import setup, Executable

# Options for building the executable
build_exe_options = {
    "packages": ["tkinter", "openpyxl", "datetime"],  # include necessary packages
    "include_files": ["app_icon.ico"]  # add additional files if needed
}

# Base is set to "Win32GUI" for Windows GUI applications to hide the console window.
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Analiza Stanu Magazynu MAD Moda",
    version="1.0",
    description="Aplikacja do analizy stanu magazynu",
    options={"build_exe": build_exe_options},
    executables=[Executable("madmodamagazyn_run.py",
                            base=base,
                            icon="app_icon.ico",
                            target_name="madmodamagazyn.exe")]
)
