# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['PIL._tkinter_finder']
for package in ('treepoem', 'zxingcpp'):
    tmp_ret = collect_all(package)
    datas += tmp_ret[0]
    binaries += tmp_ret[1]
    hiddenimports += tmp_ret[2]

gs_root = Path(r'C:\Program Files\gs\gs10.07.1')
if not (gs_root / 'bin' / 'gswin64c.exe').is_file():
    raise FileNotFoundError(f'Ghostscript runtime not found: {gs_root}')
for source in gs_root.rglob('*'):
    if source.is_file():
        datas.append((str(source), str(Path('ghostscript') / source.relative_to(gs_root).parent)))

a = Analysis(['app.py'], pathex=[], binaries=binaries, datas=datas,
             hiddenimports=hiddenimports, hookspath=[], hooksconfig={},
             runtime_hooks=[], excludes=[], noarchive=False, optimize=0)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [],
          name='PROJECTZ GS1 DataMatrix 4.0.7.6',
          version='version_info.txt', debug=False,
          bootloader_ignore_signals=False, strip=False, upx=True,
          upx_exclude=[], runtime_tmpdir=None, console=False,
          disable_windowed_traceback=False, argv_emulation=False,
          target_arch=None, codesign_identity=None, entitlements_file=None)
