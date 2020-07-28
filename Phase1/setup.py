import sys
import os
from cx_Freeze import setup, Executable
python_dir = os.path.dirname(sys.executable)  # directory of your python installation

include_files = [os.path.join(python_dir, "python3.dll"), os.path.join(python_dir, "vcruntime140.dll"), 'autorun.inf']
build_exe_options = {"packages": ["os"],
                     "includes": ["PyQt5"],
                     "include_files": include_files,
                     "excludes": ["tkinter"]}
base = None

if sys.platform == "win32":
    base = "win32GUI"

setup(name='Stucco Phase Analyzer',
      version=1.0,
      description='Program to calculate stucco phase content from weigh gain and loss measurements',
      options={'build_exe': build_exe_options},
      executables=[Executable('main.py', base=base)])