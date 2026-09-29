# -*- mode: python ; coding: utf-8 -*-
# Paketleme: python -m PyInstaller Donusturucu.spec --distpath build_cikti/dist --workpath build_cikti/build --noconfirm
import os

from PyInstaller.utils.hooks import collect_all

SIMGE = os.path.join(SPECPATH, 'icon.ico')  # noqa: F821 (SPECPATH PyInstaller'dan gelir)

datas = [(SIMGE, '.')]
binaries = []
hiddenimports = ['win32timezone', 'win32com.client']
# Yerel DLL/ikili/veri dosyaları otomatik algılanmayan paketler:
# pillow_heif (libheif), fitz (PyMuPDF), imageio_ffmpeg (ffmpeg.exe),
# sv_ttk (tema dosyaları), tkinterdnd2 (tkdnd sürükle-bırak kütüphanesi)
for paket in ('pillow_heif', 'fitz', 'imageio_ffmpeg', 'sv_ttk', 'tkinterdnd2'):
    tmp_ret = collect_all(paket)
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['donusturucu.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', '_pytest', 'pip_audit'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Donusturucu',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[SIMGE],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Donusturucu',
)
