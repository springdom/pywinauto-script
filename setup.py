import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
#build_exe_options = {"packages": ["os"], "excludes": ["tkinter"]}
buildOptions = dict(include_files = ['excel_orgchart/'])
# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(  name = "IA Auto",
        version = "0.1",
        description = "IA AutoMation",
        options = dict(build_exe = buildOptions),
        executables = [Executable("CreateUser.py", base=None)])
