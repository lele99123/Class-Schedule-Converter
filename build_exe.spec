# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['schedule_converter.py'],
             pathex=[],
             binaries=[],
             datas=[],
             hiddenimports=[
                 'pandas',
                 'numpy',  # pandas 依赖
                 'PyQt5',
                 'PyQt5.QtWidgets',
                 'PyQt5.QtCore',
                 'openpyxl',
                 'datetime',
                 'openpyxl.styles',
                 'openpyxl.utils'
             ],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='课程表转换器',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,)
