# -*- mode: python ; coding: utf-8 -*-
"""Fisier de compilare PyInstaller pentru Dozare Lingouri Titan.

De ce exista acest fisier: o comanda simpla "pyinstaller main.py" NU
include folderul dozare_titan/assets (unde e logo.png) in exe — de-aia
documentele generate din .exe arata doar textul de rezerva "ZIROM
TITANIUM" in loc de logo. Acest fisier .spec include explicit folderul,
o singura data, ca sa nu mai fie nevoie sa tii minte flag-ul --add-data
(si separatorul lui, care difera intre Windows ";" si Linux/Mac ":")
de fiecare data cand compilezi.

Compilare (din radacina proiectului, langa main.py):
    pyinstaller --noconfirm dozare_titan.spec
    (sau ruleaza build_windows.bat, care face exact asta)

Rezultatul apare in dist/dozare_titan/dozare_titan.exe — MUTA/COPIAZA
tot folderul "dist/dozare_titan" acolo unde vrei sa foloseasti aplicatia
(nu doar exe-ul singur), ca sa ramana langa el fisierele lui interne.
"""

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    # Folderul assets (logo.png + README.md) — cheia intregii probleme:
    # fara acest rand, exe-ul nu are de unde sa afle unde e logo-ul.
    datas=[('dozare_titan/assets', 'dozare_titan/assets')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='dozare_titan',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # aplicatie grafica (PySide6) — nu deschide o consola
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='dozare_titan',
)
