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

"`n-- STEP -- Update UI translations"
pylupdate5 -noobsolete -verbose .\isogeo2office.pro

"`n-- STEP -- Compile UI resources"
pyrcc5 ".\resources.qrc" -o ".\resources_rc.py"

"`n-- STEP -- Update UI elements"
pyuic5 -x .\modules\ui\auth\ui_authentication.ui -o .\modules\ui\auth\ui_authentication.py
pyuic5 -x .\modules\ui\credits\ui_credits.ui -o .\modules\ui\credits\ui_credits.py
pyuic5 -x .\modules\ui\main\ui_IsogeoToOffice.ui -o .\modules\ui\main\ui_IsogeoToOffice.py

"`n-- STEP -- Build and bundle forcing clean"
rm -r dist\*
pyinstaller -y --clean bundle_isogeo2office.spec

"`n-- STEP -- Zipping"
# Add-Type -assembly "system.io.compression.filesystem"
# $source=Join-Path (pwd) dist\isogeo2office
# $dest=Join-Path (pwd) dist\i2o.zip
# [io.compression.zipfile]::CreateFromDirectory($Source, $dest)

"`n-- STEP -- Get out the virtualenv and cleanup"
deactivate
#rm -r env3_packaging
