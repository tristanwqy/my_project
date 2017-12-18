from cx_Freeze import setup, Executable

base = None

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
