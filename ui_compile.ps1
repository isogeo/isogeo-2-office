###############################################
#
# Script to package Isogeo To Office (Python 3)
#
###############################################

# EXECUTION
"-- STEP -- Creating temp virtualenv to perform dependencies packaging"
./env3_packaging/Scripts/activate

"`n-- STEP -- Update UI translations"
pylupdate5 -noobsolete -verbose .\isogeo2office.pro

"`n-- STEP -- Compile UI resources"
pyrcc5 ".\resources.qrc" -o ".\resources_rc.py"

"`n-- STEP -- Update UI elements"
pyuic5 -x .\modules\ui\auth\ui_authentication.ui -o .\modules\ui\auth\ui_authentication.py
pyuic5 -x .\modules\ui\credits\ui_credits.ui -o .\modules\ui\credits\ui_credits.py
pyuic5 -x .\modules\ui\main\ui_win_IsogeoToOffice.ui -o .\modules\ui\main\ui_win_IsogeoToOffice.py
