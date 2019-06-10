import time
from PyQt5 import QtGui, QtCore, QtWidgets


class MyCustomWidget(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(MyCustomWidget, self).__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)

        # Create a progress bar and a button and add them to the main layout
        self.progressBar = QtWidgets.QProgressBar(self)
        self.progressBar.setRange(0, 1)
        layout.addWidget(self.progressBar)
        button = QtWidgets.QPushButton("Start", self)
        layout.addWidget(button)

        button.clicked.connect(self.onStart)

        self.myLongTask = TaskThread()
        self.myLongTask.taskFinished.connect(self.onFinished)

    def onStart(self):
        self.progressBar.setRange(0, 0)
        self.myLongTask.start()

    def onFinished(self):
        # Stop the pulsation
        self.progressBar.setRange(0, 1)


class TaskThread(QtCore.QThread):
    taskFinished = QtCore.pyqtSignal()

    def run(self):
        time.sleep(3)
        self.taskFinished.emit()


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    window = MyCustomWidget()
    window.resize(640, 480)
    window.show()
    sys.exit(app.exec_())
