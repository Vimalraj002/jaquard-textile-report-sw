from cx_Freeze import setup, Executable
import sys
import os

# Include additional files like .py scripts, icons, or data files
include_files = [
    "dataEntry.py", 
    "reportGenerateCustomer.py", 
    "reportGenerateLacing.py", 
    "managePaymentsCustomer.py", 
    "managePaymentsLacing.py"
]

# Add necessary packages
packages = [
    "tkinter", 
    "subprocess", 
    "tkcalendar", 
    "openpyxl", 
    "fpdf", 
    "datetime", 
    "os", 
    "collections"
]

# Base selection based on the platform
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Use "Win32GUI" to suppress the console window

# Main executable
executables = [Executable("main.py", base=base, target_name="ShanmugarajApp.exe")]

# Setup function
setup(
    name="ShanmugarajApp",
    version="1.0",
    description="A GUI app for managing customer and lacing data.",
    options={
        "build_exe": {
            "packages": packages,
            "include_files": include_files,
            "include_msvcr": True,  # Include MSVC runtime if on Windows
        }
    },
    executables=executables,
)
