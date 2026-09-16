# -*- mode: python ; coding: utf-8 -*-

shared_datas = [
    ('README-CheckVocal.txt', '.'),
    ('cv.ico', 'icons'),
    ('cf.ico', 'icons'),
    ('a2t.ico', 'icons'),
]

cv_a = Analysis(
    ['CheckVocal.pyw'],
    pathex=[],
    binaries=[],
    datas=shared_datas,
    hiddenimports=['importlib.resources'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

at_a = Analysis(
    ['azk2txt.pyw'],
    pathex=[],
    binaries=[],
    datas=shared_datas,
    hiddenimports=['importlib.resources'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

splash = Splash('cv_loading.png',
                binaries=cv_a.binaries,
                datas=cv_a.datas,
                text_pos=None)
				
cv_pyz = PYZ(cv_a.pure)

cv_exe = EXE(
    cv_pyz,
    cv_a.scripts,
	splash,
	splash.binaries,
    [],
    exclude_binaries=True,
    name='CheckVocal',
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
	version='cv_file_version_info.txt',
    icon=['cv.ico'],
)

cf_exe = EXE(
    cv_pyz,
    cv_a.scripts,
	splash,
	splash.binaries,
    [],
    exclude_binaries=True,
    name='CheckFiles',
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
	version='cf_file_version_info.txt',
    icon=['cf.ico'],
)

at_pyz = PYZ(at_a.pure)

at_exe = EXE(
    at_pyz,
    at_a.scripts,
    [],
    exclude_binaries=True,
    name='azk2txt',
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
	version='at_file_version_info.txt',
    icon=['a2t.ico'],
)

coll = COLLECT(
    cv_exe,
    cf_exe,
    at_exe,
    cv_a.binaries,
    at_a.binaries,
    cv_a.datas,
    at_a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CheckVocal',
)
