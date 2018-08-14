# -*- mode: python -*-

block_cipher = None

added_files = [('i18n', 'i18n'),
               ('resources', 'resources'),
              ]

a = Analysis(['__main__.py'],
             pathex=['C:\\Users\\julien.moura\\Documents\\GitHub\\Isogeo\\isogeo-2-office'],
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
          debug=False,
          strip=False,
          upx=False,
          console=False,
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
