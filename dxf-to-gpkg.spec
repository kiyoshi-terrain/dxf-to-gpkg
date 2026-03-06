# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for DXF → GeoPackage Converter"""

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect data files for packages that need them
datas = [
    ('templates', 'templates'),
    ('static', 'static'),
]
datas += collect_data_files('fiona')
datas += collect_data_files('pyproj')
datas += collect_data_files('geopandas')

# Hidden imports that PyInstaller misses
hiddenimports = (
    collect_submodules('fiona') +
    collect_submodules('pyproj') +
    collect_submodules('geopandas') +
    collect_submodules('shapely') +
    collect_submodules('webview') +
    ['ezdxf', 'numpy', 'flask', 'werkzeug', 'jinja2', 'markupsafe']
)

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'PIL', 'IPython', 'notebook'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DXF-to-GeoPackage',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DXF-to-GeoPackage',
)
