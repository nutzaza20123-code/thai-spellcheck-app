# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller build spec สำหรับ "โปรแกรมตรวจอักษรไทย-อังกฤษ"
ต้องรันบน Windows เท่านั้น (PyInstaller ไม่ข้ามแพลตฟอร์ม) ปกติแล้ว GitHub Actions
(ดู .github/workflows/build-windows.yml) จะรันไฟล์นี้ให้อัตโนมัติ ไม่ต้องรันเอง
เว้นแต่ต้องการ build ด้วยเครื่อง Windows ของตัวเอง:

    pip install -r requirements.txt pyinstaller
    pyinstaller build.spec

ผลลัพธ์จะอยู่ที่ dist/ThaiSpellCheckApp.exe (ไฟล์เดียว รันได้ทันทีแบบออฟไลน์)
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = []
hiddenimports = []

for pkg in ("pythainlp", "docx", "spellchecker"):
    datas += collect_data_files(pkg)

hiddenimports += collect_submodules("pythainlp")

block_cipher = None

a = Analysis(
    ["gui_app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # แพ็กเกจเสริมหนัก ๆ ของ pythainlp ที่เราไม่ได้ใช้ (ไม่ได้ติดตั้งอยู่แล้ว
        # แต่ระบุกันไว้กันพลาดในอนาคตถ้ามีคนติดตั้งเพิ่ม)
        "torch", "tensorflow", "transformers", "spacy", "sentence_transformers",
        "matplotlib", "scipy", "sklearn", "IPython", "notebook", "pandas",
    ],
    noarchive=False,
    optimize=0,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ThaiSpellCheckApp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # ไม่ต้องมีหน้าต่างดำ (คอนโซล) โผล่มา
    windowed=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    onefile=True,
)
