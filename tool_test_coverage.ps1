#################################################################
#
# Script to package and upload Isogeo Python package.
#
#################################################################

"-- STEP -- Creating virtualenv"
py -3 -m venv .venv_tests
./.venv_tests/Scripts/activate

"-- STEP -- Install and display dependencies within the virtualenv"
python -m pip install -U pip
pip install --upgrade setuptools wheel
pip install --upgrade -r .\requirements.txt
pip install --upgrade -r .\tests\requirements_test.txt

"-- STEP -- Fixturing"
python .\tests\fixturing.py

  "-- STEP -- Run coverage"
coverage run -m unittest discover -s tests/

"-- STEP -- Build and open coverage report"
coverage html
Invoke-Item htmlcov/index.html

"-- STEP -- Exit virtualenv"
deactivate
