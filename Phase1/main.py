import sys
from PyQt5 import QtWidgets as qtw
from PyQt5 import QtCore as qtc
from PyQt5 import QtGui as qtg

from DataEntry import *
from Compute import *
from utils import *

class MyGroupBox(qtw.QWidget):
    def __init__(self):
        super().__init__()

        self.sample_num = qtw.QLabel()
        self.sample_num.setText(str(0))
        self.compute_num = qtw.QLabel()
        self.compute_num.setText(str(0))
        self.data_entry_button = qtw.QPushButton('Enter Measurement')
        self.excel_import_button = qtw.QPushButton('Import Excel')
        self.compute_button = qtw.QPushButton('Compute')
        self.save_button = qtw.QPushButton('Save')

        self.setLayout(qtw.QVBoxLayout())

        messages = qtw.QWidget()
        messages.setLayout(qtw.QFormLayout())
        messages.setFont(qtg.QFont('Arial', 12))
        messages.layout().addRow('Number of samples loaded:', self.sample_num)
        messages.layout().addRow('Number of samples analyzed:', self.compute_num)

        buttons = qtw.QWidget()
        buttons.setLayout(qtw.QGridLayout())
        buttons.setFont(qtg.QFont('Arial', 12))
        buttons.layout().addWidget(self.data_entry_button, 0, 0)
        buttons.layout().addWidget(self.excel_import_button, 0, 1)
        buttons.layout().addWidget(self.compute_button, 1, 0)
        buttons.layout().addWidget(self.save_button, 1, 1)

        self.layout().addWidget(messages)
        self.layout().addWidget(buttons)

class MainWindow(qtw.QMainWindow):

    def __init__(self):
        """MainWindow constructor."""
        super().__init__()
        self.title = 'Main Window'
        self.left = 100
        self.top = 100
        self.width = 640
        self.height = 480
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # menu bar
        self.menuBar = self.menuBar()
        self.fileMenu = self.menuBar.addMenu("File")
        self.helpMenu = self.menuBar.addMenu("Help")

        # import from Excel file action
        importAction = qtw.QAction("Import Measurements From Excel", self)
        importAction.setShortcut("Ctrl+E")
        importAction.triggered.connect(self.open_excel_import)

        # export to Excel file action
        exportAction = qtw.QAction("Export Analyses to Excel", self)
        exportAction.setShortcut("Ctrl+X")
        exportAction.triggered.connect(self.open_save)

        # exit action
        exitAction = qtw.QAction("Exit", self)
        exitAction.setShortcut("Ctrl+Q")
        exitAction.triggered.connect(lambda: self.close())

        self.fileMenu.addAction(importAction)
        self.fileMenu.addAction(exportAction)
        self.fileMenu.addAction(exitAction)

        # layout
        self.centralWidget = MyGroupBox()
        self.setCentralWidget(self.centralWidget)
        self.samples = {}
        self.phases = {'TGA':{}, 'GPA':{}, 'GPA_ORG':{}, 'GPA_HUM':{}, 'GPA_HYD':{}}
        self.childWindow1 = DataEntryWin(self.samples)
        self.childWindow2 = ComputeWin(self.phases)
        self.centralWidget.data_entry_button.clicked.connect(self.open_data_entry)
        self.centralWidget.excel_import_button.clicked.connect(self.open_excel_import)
        self.centralWidget.compute_button.clicked.connect(self.open_compute)
        self.centralWidget.save_button.clicked.connect(self.open_save)
        self.childWindow1.submitted.connect(self.update_messages)

    def open_data_entry(self):
        self.childWindow1.show()
        self.samples = self.childWindow1.samples

    def update_messages(self):
        # updating the number of loaded samples in the main window
        self.centralWidget.sample_num.setText(str(len(self.samples)))
        l = 0
        for p in self.phases:
            if len(self.phases[p]) > l:
                l = len(self.phases[p])
        self.centralWidget.compute_num.setText(str(l))

    def open_excel_import(self):
        fileopen = qtw.QFileDialog.getOpenFileName(self, "Open File", filter="Excel files (*.xlsx *.xls)")
        filename = fileopen[0]
        if filename:
            temp, err = load_excel(filename)

            if err == 0:
                self.samples = temp
                self.childWindow1 = DataEntryWin(self.samples)
                self.update_messages()
            elif err == 1:
                msgBox = qtw.QMessageBox()
                msgBox.setIcon(qtw.QMessageBox.Critical)
                msgBox.setText("The Label column in different sheets of the Excel file do not match!")
                msgBox.setStandardButtons(qtw.QMessageBox.Ok)
                returnValue = msgBox.exec()
            elif err == 2:
                msgBox = qtw.QMessageBox()
                msgBox.setIcon(qtw.QMessageBox.Critical)
                msgBox.setText("No measurements were found in the Excel file!")
                msgBox.setStandardButtons(qtw.QMessageBox.Ok)
                returnValue = msgBox.exec()

    def open_compute(self):
        tga_phases = tga_solver(self.samples)
        gpa_phases, gpa_phases_org, gpa_phases_hum, gpa_phases_hyd = gpa_solver(self.samples)
        self.phases.update({'TGA': tga_phases})
        self.phases.update({'GPA': gpa_phases})
        self.phases.update({'GPA_ORG': gpa_phases_org})
        self.phases.update({'GPA_HUM': gpa_phases_hum})
        self.phases.update({'GPA_HYD': gpa_phases_hyd})
        self.childWindow2 = ComputeWin(self.phases)
        self.childWindow2.show()
        self.update_messages()

    def open_save(self):
        fileopen = qtw.QFileDialog.getSaveFileName(self, "Save File", filter="Excel files (*.xlsx *.xls)")
        filename = fileopen[0]
        save_excel(filename, self.phases)

if __name__ == '__main__':
    app = qtw.QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())