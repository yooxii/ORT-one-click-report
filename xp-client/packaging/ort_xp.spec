# -*- mode: python -*-
# ORT lab management system - XP lite client - PyInstaller spec
#
# Build (PyInstaller 3.3.1-3.4 + Python 3.4.x 32-bit only):
#     python -m PyInstaller packaging\ort_xp.spec --noconfirm --clean --distpath dist --workpath build
#
# KEEP THIS FILE ASCII-ONLY.
# PyInstaller 3.3.1 reads the spec with the *locale* encoding (cp950 on a zh-TW/zh-CN
# machine), so any non-ASCII byte here aborts the build with UnicodeDecodeError.
# Chinese notes belong in packaging/README.md, not in this file.
#
# This spec targets the PyInstaller 3.x spec API: COLLECT needs a.zipfiles, and the
# SPEC variable does not exist yet (only SPECPATH). Modern PyInstaller 4+ cannot use it.
#
# onedir on purpose: onefile unpacks to %TEMP% on every start (slow on XP, AV-prone).

import os

try:
    _spec_file = SPEC  # noqa: F821  (PyInstaller 4+)
except NameError:
    _spec_file = os.path.join(SPECPATH, "ort_xp.spec")  # noqa: F821  (PyInstaller 3.x)

spec_dir = os.path.dirname(os.path.abspath(_spec_file))
project_dir = os.path.dirname(spec_dir)

block_cipher = None

a = Analysis(
    [os.path.join(project_dir, "packaging", "entry_xp.py")],
    pathex=[project_dir],
    binaries=[],
    datas=[],
    hiddenimports=[
        "tkinter",
        "tkinter.ttk",
        "tkinter.messagebox",
        "sqlite3",
        "smtplib",
        "ssl",
        "email.mime.text",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        # Nothing outside the standard library is used; drop the heavy optional stuff
        # so the bundle stays small and starts fast on XP.
        "numpy",
        "pandas",
        "matplotlib",
        "PIL",
        "PyQt5",
        "PySide2",
        "unittest",
        "pydoc",
        "doctest",
        "distutils",
        "setuptools",
        "pip",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name="ORT-XP",
    debug=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name="ORT-XP",
)
