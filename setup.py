import os
import sys

from cx_Freeze import Executable, setup

includes = []
include_files = [r'c:\Anaconda3\DLLs\tcl86t.dll',
                 r'c:\Anaconda3\DLLs\tk86t.dll']
os.environ["TCL_LIBRARY"] = 'c:\\Anaconda3\\tcl\\tcl8.6'
os.environ["TK_LIBRARY"] = 'c:\\Anaconda3\\tcl\\tk8.6'

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [Executable("application.py", base=base)]

packages = ["numpy"]
options = {
    'build_exe': {
        "includes":      includes,
        "include_files": include_files,

        'packages':      packages,
    },

}

setup(
    name="run.exe",
    options=options,
    version="0.0",
    description='测试版',
    executables=executables
)
