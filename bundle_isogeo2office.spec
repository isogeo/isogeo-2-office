# -*- mode: python -*-

block_cipher = None

# -- Include ------------------------------------------------------------------
added_files = [('i18n', 'i18n'),
               ('resources', 'resources'),
               ('templates', 'templates'),
               ('thumbnails/thumbnails.xlsx', 'thumbnails'),
              ]

# -- PyInstaller process ------------------------------------------------------
a = Analysis(['__main__.py'],
             pathex=[],
             binaries=[],
             datas=added_files,
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

pyz = PYZ(a.pure,
          a.zipped_data,
          cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='isogeo2office',
          debug=True,
          strip=False,
          upx=False,
          console=True,
          icon='img\\favicon.ico',
          windowed=True,
          version='bundle_version.txt')

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=False,
               name='isogeo2office',
               icon='img\\logo_isogeo.gif'
               )
