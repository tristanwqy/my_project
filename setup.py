import sys
import os
from cx_Freeze import Executable, setup

os.environ["TCL_LIBRARY"] = 'c:\\Anaconda3\\tcl\\tcl8.6'
os.environ["TK_LIBRARY"] = 'c:\\Anaconda3\\tcl\\tk8.6'

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [Executable("application.py", base=base)]

packages = ["idna"]
options = {
    'build_exe': {

        'packages': packages,
    },

}

setup(
    name="run.exe",
    options=options,
    version="0.0",
    description='测试版',
    executables=executables
)
