import sys
from PyQt5 import QtWidgets
from PyQt5 import QtCore

app = QtWidgets.QApplication(sys.argv)

model = QtCore.QStringListModel()
model.setStringList(
    ["some", "words", "in", "my", "dictionary", "ploup", "dinausaurus", "chokobons"]
)

completer = QtWidgets.QCompleter()
completer.setModel(model)

lineedit = QtWidgets.QLineEdit()
lineedit.setCompleter(completer)
lineedit.show()

sys.exit(app.exec_())
