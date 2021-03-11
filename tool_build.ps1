###############################################
#
# Script to package Isogeo To Office (Python 3)
#
###############################################

# ENV VARS
$env:PYTHONOPTIMIZE = 2;

# EXECUTION
"-- STEP -- Creating temp virtualenv to perform dependencies packaging"
py -3 -m venv .venv_packaging
./.venv_packaging/Scripts/activate

"`n-- STEP -- Install dependencies within the virtualenv and get UPX"
python -m pip install --upgrade pip
python -m pip install --upgrade setuptools wheel
python -m pip install --upgrade -r ./requirements_dev.txt

"`n-- STEP -- Update and compile UI"
.\tool_ui_compile.ps1

"`n-- STEP -- Build and bundle forcing clean"
# mkdir dist
New-Item -Name "dist" -ItemType "directory" -Force
# rm -r dist\*
Get-ChildItem -Path ".\dist\" -Recurse | Remove-Item -force -recurse

python -m PyInstaller -y bundle_isogeo2office.spec --upx-dir .\build\upx-3.95-win64

"`n-- STEP -- Zipping"
Add-Type -assembly "system.io.compression.filesystem"
$source=Join-Path (pwd) dist\isogeo2office
$dest=Join-Path (pwd) dist\IsogeoToOffice_GenericBundle.zip
[io.compression.zipfile]::CreateFromDirectory($Source, $dest)

"`n-- STEP -- Get out the virtualenv and cleanup"
deactivate
#rm -r .venv_packaging

