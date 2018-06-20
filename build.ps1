###############################################
#
# Script to package Isogeo To Office (Python 3)
#
#   To use UPX:
#       1. download if from https://github.com/upx/upx/releases
#       2. then extrat it in 'lib/upx' directory
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

"`n-- STEP -- Build and bundle forcing clean"
rm -r dist\*
pyinstaller -y --clean bundle_isogeo2office.spec
# UPX? comment previous line, uncomment the next one and adjust UPX path if needed
#pyinstaller -y --clean --upx-dir=lib/upx/ bundle_isogeo2office.spec

"`n-- STEP -- Zipping"
Add-Type -assembly "system.io.compression.filesystem"
$source=Join-Path (pwd) dist\isogeo2office
$dest=Join-Path (pwd) dist\i2o.zip
[io.compression.zipfile]::CreateFromDirectory($Source, $dest)

"`n-- STEP -- Get out the virtualenv and cleanup"
deactivate
rm -r env3_packaging
