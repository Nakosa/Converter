import os
import sys
import tkinter
from glob import glob

from cx_Freeze import Executable, setup

root = tkinter.Tk()
os.environ['TCL_LIBRARY'] = root.tk.exprstring('$tcl_library')
os.environ['TK_LIBRARY'] = root.tk.exprstring('$tk_library')
del root


# # Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "packages": ["tkinter", 'numpy', 'pandas'],
    "excludes": [
        #'encodings',
        #'importlib',
        #'pandas',
        #'tkinter',
        #'dateutil',
        #'openpyxl',
        #'collections',
        #'numpy',
        #'StyleFrame',
        #'pytz',
        #'unittest',
        #'distutils',
        #'http',
        #'ctypes',
        #'urllib',
        #'urllib3',
        #'xml',
        #'html',
        #'json',
        #'xlrd',
        #'email',
        #'logging',
        #'numpy.linalg'
        'scipy',
        'PyQt5',
        'matplotlib',
        'multiprocessing',
        'ipython_genutils',
        'asyncio',
        'parso',
        'notebook',
        'jsonschema',
        'testpath',
        'dill',
        'requests',
        'cryptography',
        'tornado',
        'nbconvert',
        'sqlalchemy',
        'pygments',
        'concurrent',
        'idna',
        'pydoc_data',
        'jupyter_core',
        'html5lib',
        'gi',
        'PIL',
        'curses',
        'apport',
        'markupsafe',
        'prompt_toolkit',
        'nbformat',
        'jinja2',
        'wcwidth',
        'zmq',
        'ipykernel',
        'certifi',
        'pexpect',
        'chardet',
        'dbm',
        'jedi',
        'apt',
        'sqlite3',
        'traitlets',
        'IPython',
        'asn1crypto',
        'ptyprocess',
        'lib2to3',
        'simplejson',
        'jupyter_client',
        'pkg_resources'
    ],
    "include_files": [],
    "optimize": 2
}

dll_dir = os.path.join(os.path.dirname(sys.executable), "DLLs")
for dll in glob(os.path.join(dll_dir, "tcl*.dll")):
    build_exe_options["include_files"].append(dll)

for dll in glob(os.path.join(dll_dir, "tk*.dll")):
    build_exe_options["include_files"].append(dll)

# GUI applications require a different base on Windows (the default is for a
# console application).
# base = "Win32GUI"
base = None

setup(name="ui",
      version="0.1",
      description="",
      options={"build_exe": build_exe_options},
      executables=[Executable("F:/Job/7_converter/gett/gett-converter/ui.py")])
