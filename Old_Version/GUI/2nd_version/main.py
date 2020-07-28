import sys
from PyQt5 import QtWidgets as qtw
from PyQt5 import QtCore as qtc
from PyQt5 import QtGui as qtg

class MainWindow(qtw.QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.show()

if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    w = MainWindow()
    sys.exit(app.exec_())