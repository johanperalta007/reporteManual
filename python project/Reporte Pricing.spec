# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\Johan\\Documents\\reporteManual\\python project\\app_pricing.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Johan\\Documents\\reporteManual\\python project\\images', 'images')],
    hiddenimports=['customtkinter', 'PIL', 'PIL._tkinter_finder', 'pandas', 'openpyxl', 'selenium', 'selenium.webdriver', 'selenium.webdriver.chrome', 'selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chrome.service', 'selenium.webdriver.chrome.options', 'selenium.webdriver.common', 'selenium.webdriver.common.by', 'selenium.webdriver.support', 'selenium.webdriver.support.ui', 'selenium.webdriver.support.expected_conditions', 'webdriver_manager', 'webdriver_manager.chrome', 'webdriver_manager.core', 'urllib3', 'requests'],
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
    [],
    exclude_binaries=True,
    name='Reporte Pricing',
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
    icon=['C:\\Users\\Johan\\Documents\\reporteManual\\python project\\images\\logoWhite.png'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Reporte Pricing',
)
