# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['main', 'paso1_limpieza', 'paso2_clusterizacion', 'paso3_modelos', 'paso4_proyeccion', 'paso5_orden_produccion', 'sklearn', 'sklearn.ensemble', 'sklearn.linear_model', 'sklearn.neural_network', 'sklearn.preprocessing', 'sklearn.metrics', 'xgboost', 'openpyxl', 'pandas', 'numpy', 'matplotlib', 'matplotlib.backends.backend_pdf', 'matplotlib.backends.backend_tkagg', 'holidays']
hiddenimports += collect_submodules('sklearn')
tmp_ret = collect_all('xgboost')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('holidays')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['codigo\\app.py'],
    pathex=['codigo', 'codigo\\pasos'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SeedPack Planner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
