# -*- mode: python -*-

block_cipher = None

from configparser import SafeConfigParser
from os import path

# ------------ Initial settings file -----------------------------------------
config = SafeConfigParser()
config.read(path.realpath("settings_TPL.ini"))
#config.set("global", "excel", "word", "xml", "proxy")
with open(path.realpath("build\\settings.ini"), "w") as configfile:
    config.write(configfile)

# ----------------------------------------------------------------------------

added_files = [('build\\settings.ini', '.'),
               ('LICENSE', '.'),
               ('README.md', '.'),
               ('templates\\template_Isogeo.docx', 'templates'),
               ('output\\README.md', 'output'),
               ('img\\favicon.ico', 'img')
              ]


a = Analysis(['isogeo2office.py'],
             pathex=['C:\\Users\\julien.moura\\Documents\\GitHub\\Isogeo\\isogeo-2-office'],
             binaries=None,
             datas=added_files,
             hiddenimports=[],
             hookspath=["hooks"],
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
          upx=True,
          console=False,
          icon='img\\favicon.ico',
          windowed=True,
          version='bundle_version.txt')

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='isogeo2office',
               icon='img\\logo_isogeo.gif'
               )
