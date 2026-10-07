# -*- mode: python -*-
# ORT 实验室管理系统 · XP 精简客户端 —— PyInstaller 配置
#
# 必须用 PyInstaller 3.3.1–3.4（XP 可用的最后一批启动器）与 Python 3.4.10（32 位）执行：
#     python -m PyInstaller packaging\ort_xp.spec --noconfirm --clean --distpath dist --workpath build
#
# 刻意用 onedir（不是 onefile）：onefile 每次启动都解压到 %TEMP%，
# 在 XP 的老磁盘与老杀软上既慢又容易被拦。

import os

spec_dir = os.path.dirname(os.path.abspath(SPEC))
project_dir = os.path.dirname(spec_dir)

block_cipher = None

a = Analysis(
    [os.path.join(project_dir, "ort_xp", "__main__.py")],
    pathex=[project_dir],
    binaries=[],
    datas=[],
    hiddenimports=["tkinter", "tkinter.ttk", "tkinter.messagebox", "sqlite3", "smtplib", "ssl", "email.mime.text"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        # 本范围用不到的重量级模块，排除后包更小、启动更快
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
    bootloader_ignore_signals=False,
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
