# -*- mode: python -*-

block_cipher = None

# -- Include ------------------------------------------------------------------
added_files = [('i18n', 'i18n'),
               ('resources', 'resources'),
               ('_templates/template_Isogeo.docx', '_templates'),
               ('_thumbnails/thumbnails.xlsx', '_thumbnails'),
              ]

# -- PyInstaller process ------------------------------------------------------
a = Analysis(['IsogeoToOffice.py'],
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
          debug=False,
          strip=False,
          upx=False,
          console=False,
          icon='resources\\favicon.ico',
          windowed=True,
          version='bundle_version.txt')

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=False,
               name='isogeo2office',
               icon='resources\\logo_isogeo.gif'
               )
