###############################################
#
# Script to package Isogeo To Office (Python 3)
#
###############################################

# EXECUTION
"-- STEP -- Creating temp virtualenv to perform dependencies packaging"
py -3 -m venv env3_packaging
./env3_packaging/Scripts/activate

"`n-- STEP -- Install dependencies within the virtualenv and get UPX"
python -m pip install -U pip
pip install --upgrade -r ./requirements.txt
pip install --upgrade -r ./requirements_dev.txt

"`n-- STEP -- Update and compile UI"
.\ui_compile.ps1

"`n-- STEP -- Build and bundle forcing clean"
rm -r dist\*
pyinstaller -y --clean bundle_isogeo2office.spec

"`n-- STEP -- Add required empty folders"
new-item -Name "dist/isogeo2office/_auth" -ItemType directory 
new-item -Name "dist/isogeo2office/_logs" -ItemType directory 

"`n-- STEP -- Zipping"
Add-Type -assembly "system.io.compression.filesystem"
$source=Join-Path (pwd) dist\isogeo2office
$dest=Join-Path (pwd) dist\i2o.zip
[io.compression.zipfile]::CreateFromDirectory($Source, $dest)

"`n-- STEP -- Get out the virtualenv and cleanup"
deactivate
rm -r env3_packaging
