# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('style/conflux-dark-red.json', 'style'),        # ← moved here
        ('style/Conflux-Logo.ico',       'style'),
        ('style/conflux-logo.png',       'style'),
    ],
    hiddenimports=['tkinterdnd2'],
    hookspath=['.'],
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='conflux',
    console=False,
    icon='style/Conflux-Logo.ico',
)

coll = COLLECT(
    exe, a.binaries, a.datas,
    name='conflux',
)
