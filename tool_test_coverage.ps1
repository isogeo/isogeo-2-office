#################################################################
#
# Script to package and upload Isogeo Python package.
#
#################################################################

"-- STEP -- Creating virtualenv"
py -3 -m venv env3_tests
./env3_tests/Scripts/activate

"-- STEP -- Install and display dependencies within the virtualenv"
python -m pip install -U pip
pip install --upgrade setuptools wheel
pip install --upgrade -r .\requirements.txt
pip install --upgrade -r .\tests\requirements_test.txt

"-- STEP -- Python code style"
# main
pycodestyle .\IsogeoToOffice.py --ignore="E265,E501" --statistics --show-source
# threads
pycodestyle modules/threads.py --ignore="E265,E501" --statistics --show-source
# export modules
pycodestyle modules/export/isogeo2docx.py --ignore="E265,E501" --statistics --show-source
pycodestyle modules/export/isogeo2xlsx.py --ignore="E265,E501" --statistics --show-source
pycodestyle modules/export/isogeo_stats.py --ignore="E265,E501" --statistics --show-source
pycodestyle modules/export/formatter.py --ignore="E265,E501" --statistics --show-source
# utils
pycodestyle modules/utils/api.py --ignore="E265,E501" --statistics --show-source
pycodestyle modules/utils/utils.py --ignore="E265,E501" --statistics --show-source

"-- STEP -- Fixturing"
python .\tests\fixturing.py

  "-- STEP -- Run coverage"
coverage run -m unittest discover -s tests/

"-- STEP -- Build and open coverage report"
coverage html
Invoke-Item htmlcov/index.html

"-- STEP -- Exit virtualenv"
deactivate
