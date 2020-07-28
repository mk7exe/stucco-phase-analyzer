from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
import numpy as np
import Main_Window, Excel_load_window, Compute_window, Save_window
import xlrd
import xlwt

def isdigit(str0):
    str1 = str0.replace('.', '', 1)
    str2 = str1.replace('-', '', 1)
    str3 = str2.replace('+', '', 1)
    return str3.isdigit()

class Errors:
    def __init__(self, msg):
        self.msg = msg
        self.err_msg()

    def err_msg(self):
        msg = QMessageBox()
        msg.setWindowTitle('Error')
        msg.setText(self.msg)
        msg.setIcon(QMessageBox.Critical)
        temp = msg.exec_()

class DW:
    w0 = 136.14
    dhwl = 36 / (w0 + 36)
    hhwl = 9 / (w0 + 9)
    hhwg = 1.5 * 18 / (w0 + 9)
    a3wq1 = 9 / w0
    a3wq2 = 36 / w0

class PaGPA:
    def __init__(self, wl_id, dwl, dwa, dwh):
        self.dwl = dwl
        self.wl_id = wl_id
        self.dwa = dwa
        self.dwh = dwh

    def phase(self):
        x = np.zeros((5, 1))
        if self.dwa >= 0:
            x[2] = self.dwa/DW.a3wq1
            x[1] = (self.dwh-DW.a3wq2*x[2])/DW.hhwg
            x[3] = 0
            if self.wl_id == 1:
                x[0] = (self.dwl-DW.hhwl*x[1])/DW.dhwl
            if self.wl_id == 2:
                x[0] = (self.dwl*(1+0.01*DW.a3wq1*x[2])-DW.hhwl*(x[1]+(1+DW.a3wq1)*x[2]))/DW.dhwl
            if self.wl_id == 3:
                x[0] = (self.dwl*(1+0.01*DW.hhwg*x[1]+0.01*DW.a3wq2*x[2])-DW.dhwl*((1+DW.hhwg)*x[1]+(1+DW.a3wq2)*x[2]))/DW.dhwl
            x[4] = 100-(x[0]+x[1]+x[2]+x[3])
        else:
            x[3] = -self.dwa
            x[2] = 0
            x[1] = (self.dwh+x[3])/DW.hhwg
            if self.wl_id == 1:
                x[0] = (self.dwl-DW.hhwl*x[1]-x[3])/DW.dhwl
            if self.wl_id == 2:
                x[0] = (self.dwl*(1-0.01*x[3])-DW.hhwl*x[1])/DW.dhwl
            if self.wl_id == 3:
                x[0] = (self.dwl*(1+0.01*DW.hhwg*x[1]-0.01*x[3])-DW.dhwl*(1+DW.hhwg)*x[1])/DW.dhwl
            x[4] = 100-(x[0]+x[1]+x[2]+x[3])
        return x

class PaTGA:
    def __init__(self, dwo, dwa, dwh):
        self.dwo = dwo
        self.dwa = dwa
        self.dwh = dwh

    def invA(self):
        A = np.zeros((4, 4))
        if self.dwa >= self.dwo:
            A[0, 0] = 1
            A[0, 1] = 1
            A[0, 2] = 1
            A[0, 3] = 1
            A[1, 0] = DW.dhwl
            A[1, 1] = DW.hhwl
            A[1, 2] = 0
            A[1, 3] = 0
            A[2, 0] = 0
            A[2, 1] = DW.hhwg - 0.01 * DW.hhwg / DW.dhwl * self.dwh
            A[2, 2] = DW.a3wq2 - 0.01 * DW.a3wq2 / DW.dhwl * self.dwh
            A[2, 3] = -1
            A[3, 0] = DW.dhwl / DW.hhwl - 1
            A[3, 1] = 0
            A[3, 2] = DW.a3wq1 - 0.01 * DW.a3wq1 / DW.hhwl * self.dwa
            A[3, 3] = -1
        else:
            A[0, 0] = 1
            A[0, 1] = 1
            A[0, 2] = 1
            A[0, 3] = 1
            A[1, 0] = DW.dhwl
            A[1, 1] = DW.hhwl
            A[1, 2] = 1
            A[1, 3] = 0
            A[2, 0] = 0
            A[2, 1] = DW.hhwg - 0.01 * DW.hhwg / DW.dhwl * self.dwh
            A[2, 2] = 0.01 * self.dwh / DW.dhwl - 1
            A[2, 3] = -1
            A[3, 0] = DW.dhwl/DW.hhwl - 1
            A[3, 1] = 0
            A[3, 2] = 0.01*self.dwa / DW.hhwl - 1
            A[3, 3] = -1
        A_inv = np.linalg.inv(A)
        return A_inv

    def b(self):
        b = np.zeros((4, 1))
        if self.dwa >= self.dwo:
            b[0] = 100
            b[1] = self.dwo
            b[2] = 1 / DW.dhwl * self.dwh - 100
            b[3] = 1 / DW.hhwl * self.dwa - 100
        else:
            b[0] = 100
            b[1] = self.dwo
            b[2] = 1 / DW.dhwl * self.dwh - 100
            b[3] = 1 / DW.hhwl * self.dwa - 100
        return b

    def phase(self):
        x = np.zeros((5, 1))
        temp = np.matmul(self.invA(), self.b())
        if self.dwa >= self.dwo:
            x[0] = temp[0]
            x[1] = temp[1]
            x[2] = temp[2]
            x[3] = 0
            x[4] = temp[3]
        else:
            x[0] = temp[0]
            x[1] = temp[1]
            x[2] = 0
            x[3] = temp[2]
            x[4] = temp[3]
        return x

class PaTGA_NoAIII:
    def __init__(self, dwo, dwa, dwh, dgh):
        self.dwo = dwo
        self.dwa = dwa
        self.dwh = dwh
        self.dgh = dgh

    def invA(self):
        A = np.zeros((4, 4))
        if self.dwa >= self.dwo:
            A[0, 0] = 1
            A[0, 1] = 1
            A[0, 2] = 1
            A[0, 3] = 1
            A[1, 0] = DW.dhwl
            A[1, 1] = DW.hhwl
            A[1, 2] = 0
            A[1, 3] = 0
            A[2, 0] = 0
            A[2, 1] = DW.hhwg - 0.01 * DW.hhwg / DW.dhwl * self.dwh
            A[2, 2] = DW.a3wq2 - 0.01 * DW.a3wq2 / DW.dhwl * self.dwh
            A[2, 3] = -1
            A[3, 0] = 0
            A[3, 1] = DW.hhwg
            A[3, 2] = DW.a3wq2
            A[3, 3] = 0
        else:
            A[0, 0] = 1
            A[0, 1] = 1
            A[0, 2] = 1
            A[0, 3] = 1
            A[1, 0] = DW.dhwl
            A[1, 1] = DW.hhwl
            A[1, 2] = 1
            A[1, 3] = 0
            A[2, 0] = 0
            A[2, 1] = DW.hhwg - 0.01 * DW.hhwg / DW.dhwl * self.dwh
            A[2, 2] = 0.01 * self.dwh / DW.dhwl - 1
            A[2, 3] = -1
            A[3, 0] = 0
            A[3, 1] = DW.hhwg
            A[3, 2] = -1
            A[3, 3] = 0
        A_inv = np.linalg.inv(A)
        return A_inv

    def b(self):
        b = np.zeros((4, 1))
        if self.dwa >= self.dwo:
            b[0] = 100
            b[1] = self.dwo
            b[2] = 1 / DW.dhwl * self.dwh - 100
            b[3] = self.dgh
        else:
            b[0] = 100
            b[1] = self.dwo
            b[2] = 1 / DW.dhwl * self.dwh - 100
            b[3] = self.dgh
        return b

    def phase(self):
        x = np.zeros((5, 1))
        temp = np.matmul(self.invA(), self.b())
        if self.dwa >= self.dwo:
            x[0] = temp[0]
            x[1] = temp[1]
            x[2] = temp[2]
            x[3] = 0
            x[4] = temp[3]
        else:
            x[0] = temp[0]
            x[1] = temp[1]
            x[2] = 0
            x[3] = temp[2]
            x[4] = temp[3]
        return x

class WindowMain(QtWidgets.QMainWindow, Main_Window.Ui_MainWindow):
    def __init__(self, parent=None):
        super(WindowMain, self).__init__(parent)
        self.setupUi(self)

        self.childWindow1 = WindowExcel(self)
        self.childWindow2 = WindowCompute(self)
        self.childWindow3 = WindowSave(self)
        self.pushButton_excel.clicked.connect(self.load_excel)
        self.actionLoad.triggered.connect(self.load_excel)
        self.pushButton_calc.clicked.connect(self.load_compute)
        self.actionCalculate.triggered.connect(self.load_compute)
        self.pushButton_save.clicked.connect(self.load_save)
        self.actionExit.triggered.connect(self.exit)

        self.childWindow2.checkBoxTGA.setEnabled(False)
        self.childWindow2.checkBoxCOM.setEnabled(False)
        self.childWindow1.LoadTGA.setEnabled(False)
        self.childWindow1.LoadCOM.setEnabled(False)
        self.childWindow1.LoadSCA.setEnabled(False)
        self.childWindow3.checkBoxTGA.setEnabled(False)
        self.childWindow3.checkBoxCOM.setEnabled(False)
        self.childWindow3.checkBoxGPA_TGA_org.setEnabled(False)
        self.childWindow3.checkBoxGPA_TGA_a3.setEnabled(False)
        self.childWindow3.checkBoxGPA_TGA_hyd.setEnabled(False)
        self.childWindow3.checkBoxGPA_COM_org.setEnabled(False)
        self.childWindow3.checkBoxGPA_COM_a3.setEnabled(False)
        self.childWindow3.checkBoxGPA_COM_hyd.setEnabled(False)

    def exit(self):
        self.close()

    def load_excel(self):
        if self.checkBoxTGA.isChecked():
            self.childWindow1.LoadTGA.setEnabled(True)
        if self.checkBoxCOM.isChecked():
            self.childWindow1.LoadCOM.setEnabled(True)
        if self.checkBoxSCA.isChecked():
            self.childWindow1.LoadSCA.setEnabled(True)
        self.childWindow1.show()

    def load_compute(self):
        if self.checkBoxTGA.isChecked() and self.label_TGA_load.text() != '0':
            self.childWindow2.checkBoxTGA.setEnabled(True)
        if self.checkBoxCOM.isChecked() and self.label_COM_load.text() != '0':
            self.childWindow2.checkBoxCOM.setEnabled(True)
        self.childWindow2.pushButton_3.setEnabled(False)
        self.childWindow2.pushButton_4.setEnabled(False)
        self.childWindow2.show()

    def load_save(self):
        temp = self.childWindow2.get_output()
        checks = temp[0]
        chbox = [self.childWindow3.checkBoxTGA, self.childWindow3.checkBoxCOM, self.childWindow3.checkBoxGPA_TGA_org,
                 self.childWindow3.checkBoxGPA_TGA_a3, self.childWindow3.checkBoxGPA_TGA_hyd,
                 self.childWindow3.checkBoxGPA_COM_org, self.childWindow3.checkBoxGPA_COM_a3,
                 self.childWindow3.checkBoxGPA_COM_hyd]
        for i in range(8):
            if checks[i]:
                chbox[i].setEnabled(True)
        self.childWindow3.pushButton_save_excel.setEnabled(False)
        self.childWindow3.pushButton_save_ascii.setEnabled(False)
        self.childWindow3.show()


class WindowExcel(QtWidgets.QMainWindow, Excel_load_window.Ui_MainWindow):
    def __init__(self, parent=None):
        super(WindowExcel, self).__init__(parent)
        self.setupUi(self)

        self.TGA = []
        self.COM = []
        self.GPA = []

        self.LoadTGA.clicked.connect(self.load_TGA_clicked)
        self.LoadCOM.clicked.connect(self.load_COM_clicked)
        self.LoadSCA.clicked.connect(self.load_GPA_clicked)
        self.pushBotton_done.clicked.connect(self.done)

        self.main = parent  # otherwise, recursion

    def closeEvent(self, event):
        self.LoadTGA.setEnabled(False)
        self.LoadCOM.setEnabled(False)
        self.LoadSCA.setEnabled(False)

    def done(self):
        self.close()

    def load_TGA_clicked(self):
        filename = QFileDialog.getOpenFileName(filter="Excel files (*.xlsx *.xls)")
        if filename[0]:
            self.loadTGA = True
            file = filename[0]
            wb = xlrd.open_workbook(file)
            sheet = wb.sheet_by_index(0)
            label = []
            sample_count = sheet.nrows
            for i in range(1, sample_count, 1):
                label.append(sheet.cell_value(i, 0))

            sheet = wb.sheet_by_index(1)
            org = [[] for i in range(len(label))]
            for i in range(1, sample_count, 1):
                for j in range(1, sheet.ncols):
                    org[i-1].append(sheet.cell_value(i, j))

            sheet = wb.sheet_by_index(3)
            hyd = [[] for i in range(len(label))]
            for i in range(1, sample_count, 1):
                for j in range(1, sheet.ncols):
                    hyd[i-1].append(sheet.cell_value(i, j))

            sheet = wb.sheet_by_index(2)
            a3t = [[] for i in range(len(label))]
            for i in range(1, sample_count, 1):
                for j in range(1, sheet.ncols):
                    a3t[i-1].append(sheet.cell_value(i, j))
            self.TGA = label, org, a3t, hyd
            self.label_TGA_load.setText(str(len(label)))
            self.main.label_TGA_load.setText(str(len(label)))
            self.main.childWindow2.label_3.setText(str(len(label)))

    def load_COM_clicked(self):
        filename = QFileDialog.getOpenFileName(filter="Excel files (*.xlsx *.xls)")
        if filename[0]:
            file = filename[0]
            wb = xlrd.open_workbook(file)
            sheet = wb.sheet_by_index(0)
            label = []
            sample_count = sheet.nrows
            for i in range(1, sample_count, 1):
                label.append(sheet.cell_value(i, 0))

            sheet = wb.sheet_by_index(1)
            org = [[] for i in range(len(label))]
            for i in range(1, sample_count, 1):
                for j in range(1, sheet.ncols):
                    org[i-1].append(sheet.cell_value(i, j))

            sheet = wb.sheet_by_index(3)
            hyd = [[] for i in range(len(label))]
            for i in range(1, sample_count, 1):
                for j in range(1, sheet.ncols):
                    hyd[i-1].append(sheet.cell_value(i, j))

            sheet = wb.sheet_by_index(2)
            a3t = [[] for i in range(len(label))]
            for i in range(1, sample_count, 1):
                for j in range(1, sheet.ncols):
                    a3t[i-1].append(sheet.cell_value(i, j))
            self.COM = label, org, a3t, hyd
            self.label_COM_load.setText(str(len(label)))
            self.main.label_COM_load.setText(str(len(label)))
            self.main.childWindow2.label_4.setText(str(len(label)))
            if len(label) == 0:
                self.main.childWindow2.checkBoxCOM.setEnabled(False)

    def load_GPA_clicked(self):
        filename = QFileDialog.getOpenFileName(filter="Excel files (*.xlsx *.xls)")
        if filename[0]:
            file = filename[0]
            wb = xlrd.open_workbook(file)
            sheet = wb.sheet_by_index(0)
            label = []
            sample_count = sheet.nrows
            for i in range(1, sample_count, 1):
                label.append(sheet.cell_value(i, 0))

            sheet = wb.sheet_by_index(3)
            hyd = [[] for i in range(len(label))]
            for i in range(1, sample_count, 1):
                for j in range(1, sheet.ncols):
                    hyd[i-1].append(sheet.cell_value(i, j))

            sheet = wb.sheet_by_index(2)
            a3t = [[] for i in range(len(label))]
            for i in range(1, sample_count, 1):
                for j in range(1, sheet.ncols):
                    a3t[i-1].append(sheet.cell_value(i, j))
            self.GPA = label, [], a3t, hyd
            self.label_SCA_load.setText(str(len(label)))
            self.main.label_SCA_load.setText(str(len(label)))

    def get_output(self):
        return self.TGA, self.COM, self.GPA


class WindowCompute(QtWidgets.QMainWindow, Compute_window.Ui_MainWindow):
    def __init__(self, parent=None):
        super(WindowCompute, self).__init__(parent)
        self.setupUi(self)
        self.main = parent

        self.tests = np.zeros(8)
        self.output = []

        self.checkBoxTGA.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxCOM.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA1.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA2.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA3.stateChanged.connect(self.checkBoxChangedAction)

        self.pushButton_3.clicked.connect(self.calc_tga)
        self.pushButton_4.clicked.connect(self.calc_gpa)
        self.pushBotton_done.clicked.connect(self.done)
        self.pushButton_help.clicked.connect(self.help)


    def closeEvent(self, event):
        self.checkBoxTGA.setEnabled(False)
        self.checkBoxCOM.setEnabled(False)

    def done(self):
        self.close()

    def help(self):
        msg = QMessageBox()
        msg.setWindowTitle('Weight loss for GPA')
        Text = 'GPA phase analysis needs one weight loss measurement data.\n'
        Text = Text + 'This can come from original, AIII, or hydrated sample.\n'
        Text = Text + 'If you choose more than one option, multiple phase analysis will be performed\n'
        Text = Text + 'using each weight loss measurement selected.'
        msg.setText(Text)
        msg.setIcon(QMessageBox.Information)
        temp = msg.exec_()

    def checkBoxChangedAction(self, state):
        if self.checkBoxTGA.isChecked() or self.checkBoxCOM.isChecked():
            self.pushButton_3.setEnabled(True)
        if self.main.checkBoxSCA.isChecked() and self.main.label_SCA_load.text() != '0':
            if self.checkBoxGPA1.isChecked() or self.checkBoxGPA2.isChecked() or self.checkBoxGPA3.isChecked():
                self.pushButton_4.setEnabled(True)
        if not self.checkBoxTGA.isChecked() and not self.checkBoxCOM.isChecked():
            self.pushButton_3.setEnabled(False)
            self.pushButton_4.setEnabled(False)

    def calc_tga(self):
        msg = QMessageBox()
        msg.setWindowTitle('TGA Phase Analysis')
        self.TGA, self.COM, self.SCA = self.main.childWindow1.get_output()
        if self.checkBoxTGA.isChecked():
            self.tests[0] = 1
            label = self.TGA[0]
            org = self.TGA[1]
            a3t = self.TGA[2]
            hyd = self.TGA[3]
            dh = [[] for i in range(len(label))]
            hh = [[] for i in range(len(label))]
            a3 = [[] for i in range(len(label))]
            fm = [[] for i in range(len(label))]
            ine = [[] for i in range(len(label))]
            for i1 in range(len(label)):
                if label[i1]:
                    for i2 in range(len(org[i1])):
                        if str(org[i1][i2]).replace('.', '', 1).isdigit():
                            for i3 in range(len(a3t[i1])):
                                if str(a3t[i1][i3]).replace('.', '', 1).isdigit():
                                    for i4 in range(len(hyd[i1])):
                                        if str(hyd[i1][i4]).replace('.', '', 1).isdigit():
                                            t = PaTGA(org[i1][i2], a3t[i1][i3], hyd[i1][i4])
                                            tt = t.phase()
                                            dh[i1].append(tt[0])
                                            hh[i1].append(tt[1])
                                            a3[i1].append(tt[2])
                                            fm[i1].append(tt[3])
                                            ine[i1].append(tt[4])
            self.output.append([label,dh,hh,a3,fm,ine])
            Text = 'TGA Phase Analysis with TGA weight loss data is done!\n'
            self.main.childWindow3.checkBoxTGA.setEnabled(True)
        else:
            self.output.append([[], [], [], [], [], []])

        if self.checkBoxCOM.isChecked():
            self.tests[1] = 1
            label = self.COM[0]
            org = self.COM[1]
            a3t = self.COM[2]
            hyd = self.COM[3]
            dh = [[] for i in range(len(label))]
            hh = [[] for i in range(len(label))]
            a3 = [[] for i in range(len(label))]
            fm = [[] for i in range(len(label))]
            ine = [[] for i in range(len(label))]
            for i1 in range(len(label)):
                if label[i1]:
                    for i2 in range(len(org[i1])):
                        if str(org[i1][i2]).replace('.', '', 1).isdigit():
                            for i3 in range(len(a3t[i1])):
                                if str(a3t[i1][i3]).replace('.', '', 1).isdigit():
                                    for i4 in range(len(hyd[i1])):
                                        if str(hyd[i1][i4]).replace('.', '', 1).isdigit():
                                            t = PaTGA(org[i1][i2], a3t[i1][i3], hyd[i1][i4])
                                            tt = t.phase()
                                            dh[i1].append(tt[0])
                                            hh[i1].append(tt[1])
                                            a3[i1].append(tt[2])
                                            fm[i1].append(tt[3])
                                            ine[i1].append(tt[4])
            self.output.append([label,dh,hh,a3,fm,ine])
            Text = Text + 'TGA Phase Analysis with COMPUTRAC weight loss data is done!'
            self.main.childWindow3.checkBoxCOM.setEnabled(True)
        else:
            self.output.append([[], [], [], [], [], []])
        msg.setText(Text)
        msg.setIcon(QMessageBox.Information)
        temp = msg.exec_()

    def calc_gpa(self):
        msg = QMessageBox()
        msg.setWindowTitle('TGA Phase Analysis')
        text = ''
        self.TGA, self.COM, self.SCA = self.main.childWindow1.get_output()
        if self.checkBoxTGA.isChecked():
            err = 0
            for c, v in enumerate([self.checkBoxGPA1, self.checkBoxGPA2, self.checkBoxGPA3], 1):
                if v.isChecked():
                    label = self.SCA[0]
                    wl = self.TGA[c]
                    wl_id = c
                    a3t = self.SCA[2]
                    hyd = self.SCA[3]
                    dh = [[] for i in range(len(label))]
                    hh = [[] for i in range(len(label))]
                    a3 = [[] for i in range(len(label))]
                    fm = [[] for i in range(len(label))]
                    ine = [[] for i in range(len(label))]
                    for i1 in range(len(label)):
                        if label[i1] != self.TGA[0][i1]:
                            if not err:
                                text = 'Sample ids in TGA and Digital Scale datasets are not compatible\n'
                                text = text + 'Digital Scale: ' + label[i1] + '\n'
                                text = text + 'TGA: ' + self.TGA[0][i1]
                                err = Errors(text)
                                text = ''
                                self.done()
                            self.output.append([[], [], [], [], [], []])
                            break
                        if label[i1]:
                            for i2 in range(len(wl[i1])):
                                if isdigit(str(wl[i1][i2])):
                                    for i3 in range(len(a3t[i1])):
                                        if isdigit(str(a3t[i1][i3])):
                                            for i4 in range(len(hyd[i1])):
                                                if isdigit(str(hyd[i1][i4])):
                                                    t = PaGPA(wl_id, wl[i1][i2], a3t[i1][i3], hyd[i1][i4])
                                                    tt = t.phase()
                                                    dh[i1].append(tt[0])
                                                    hh[i1].append(tt[1])
                                                    a3[i1].append(tt[2])
                                                    fm[i1].append(tt[3])
                                                    ine[i1].append(tt[4])
                    self.output.append([label, dh, hh, a3, fm, ine])
                    if c == 1 and not err:
                        text = text + 'GPA Phase Analysis using TGA Original Sample weight loss data is done!\n'
                        self.main.childWindow3.checkBoxGPA_TGA_org.setEnabled(True)
                    if c == 2 and not err:
                        text = text + 'GPA Phase Analysis using TGA AIII Sample weight loss data is done!\n'
                        self.main.childWindow3.checkBoxGPA_TGA_a3.setEnabled(True)
                    if c == 3 and not err:
                        text = text + 'GPA Phase Analysis using TGA Hydrated Sample weight loss data is done!\n'
                        self.main.childWindow3.checkBoxGPA_TGA_hyd.setEnabled(True)
                    if not err:
                        self.tests[c+1] = 1
                else:
                    self.output.append([[], [], [], [], [], []])
        if self.checkBoxCOM.isChecked():
            err = 0
            for c, v in enumerate([self.checkBoxGPA1, self.checkBoxGPA2, self.checkBoxGPA3], 1):
                if v.isChecked():
                    label = self.SCA[0]
                    wl = self.COM[c]
                    wl_id = c
                    a3t = self.SCA[2]
                    hyd = self.SCA[3]
                    dh = [[] for i in range(len(label))]
                    hh = [[] for i in range(len(label))]
                    a3 = [[] for i in range(len(label))]
                    fm = [[] for i in range(len(label))]
                    ine = [[] for i in range(len(label))]
                    for i1 in range(len(label)):
                        if label[i1] != self.TGA[0][i1]:
                            if not err:
                                err = Errors('Sample ids in COMPUTRAC and Digital Scale datasets are not compatible')
                            self.output.append([[], [], [], [], [], []])
                            break
                        if label[i1]:
                            for i2 in range(len(wl[i1])):
                                if isdigit(str(wl[i1][i2])):
                                    for i3 in range(len(a3t[i1])):
                                        if isdigit(str(a3t[i1][i3])):
                                            for i4 in range(len(hyd[i1])):
                                                if isdigit(str(hyd[i1][i4])):
                                                    t = PaGPA(wl_id, wl[i1][i2], a3t[i1][i3], hyd[i1][i4])
                                                    tt = t.phase()
                                                    dh[i1].append(tt[0])
                                                    hh[i1].append(tt[1])
                                                    a3[i1].append(tt[2])
                                                    fm[i1].append(tt[3])
                                                    ine[i1].append(tt[4])
                    self.output.append([label, dh, hh, a3, fm, ine])
                    if c == 1 and not err:
                        text = text + 'GPA Phase Analysis using TGA Original Sample weight loss data is done!\n'
                        self.main.childWindow3.checkBoxGPA_COM_org.setEnabled(True)
                    if c == 2 and not err:
                        text = text + 'GPA Phase Analysis using TGA AIII Sample weight loss data is done!\n'
                        self.main.childWindow3.checkBoxGPA_COM_a3.setEnabled(True)
                    if c == 3 and not err:
                        text = text + 'GPA Phase Analysis using TGA Hydrated Sample weight loss data is done!\n'
                        self.main.childWindow3.checkBoxGPA_COM_hyd.setEnabled(True)
                    if not err:
                        self.tests[c+4] = 1
                else:
                    self.output.append([[], [], [], [], [], []])

        msg.setText(text)
        msg.setIcon(QMessageBox.Information)
        if text:
            temp = msg.exec_()

    def get_output(self):
        return [self.tests, self.output]


class WindowSave(QtWidgets.QMainWindow, Save_window.Ui_MainWindow):
    def __init__(self, parent=None):
        super(WindowSave, self).__init__(parent)
        self.setupUi(self)

        self.main = parent  # otherwise, recursion

        self.pushButton_save_excel.clicked.connect(self.save_excel_clicked)
        self.pushButton_save_ascii.clicked.connect(self.save_ascii_clicked)

        self.checkBoxTGA.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxCOM.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA_TGA_org.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA_TGA_a3.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA_TGA_hyd.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA_COM_org.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA_COM_a3.stateChanged.connect(self.checkBoxChangedAction)
        self.checkBoxGPA_COM_hyd.stateChanged.connect(self.checkBoxChangedAction)

        self.pushButton_done.clicked.connect(self.done)

    def closeEvent(self, event):
        self.checkBoxTGA.setChecked(False)
        self.checkBoxCOM.setChecked(False)
        self.checkBoxGPA_TGA_org.setChecked(False)
        self.checkBoxGPA_TGA_a3.setChecked(False)
        self.checkBoxGPA_TGA_hyd.setChecked(False)
        self.checkBoxGPA_COM_org.setChecked(False)
        self.checkBoxGPA_COM_a3.setChecked(False)
        self.checkBoxGPA_COM_hyd.setChecked(False)
        self.checkBoxTGA.setEnabled(False)
        self.checkBoxCOM.setEnabled(False)
        self.checkBoxGPA_TGA_org.setEnabled(False)
        self.checkBoxGPA_TGA_a3.setEnabled(False)
        self.checkBoxGPA_TGA_hyd.setEnabled(False)
        self.checkBoxGPA_COM_org.setEnabled(False)
        self.checkBoxGPA_COM_a3.setEnabled(False)
        self.checkBoxGPA_COM_hyd.setEnabled(False)

    def done(self):
        self.close()

    def checkBoxChangedAction(self, state):
        if any([self.checkBoxTGA.isChecked(), self.checkBoxCOM.isChecked(), self.checkBoxGPA_TGA_org.isChecked(),
               self.checkBoxGPA_TGA_a3.isChecked(), self.checkBoxGPA_TGA_hyd.isChecked(),
               self.checkBoxGPA_COM_org.isChecked(),  self.checkBoxGPA_COM_a3.isChecked(),
               self.checkBoxGPA_COM_hyd.isChecked()]):
            self.pushButton_save_excel.setEnabled(True)
            self.pushButton_save_ascii.setEnabled(False)
        else:
            self.pushButton_save_excel.setEnabled(False)
            self.pushButton_save_ascii.setEnabled(False)

    def save_excel_clicked(self):
        msg = QMessageBox()
        msg.setWindowTitle('Save to Excel')
        text = 'You need to select a destination to save Excel files at.\n'
        text = text + 'One Excel file will be made for each option you selected.'
        msg.setText(text)
        msg.setIcon(QMessageBox.Information)
        temp = msg.exec_()
        temp = self.main.childWindow2.get_output()
        checks = temp[0]
        calcs = temp[1]
        folder = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        tests = ['TGA-TGA', 'TGA-COMPUTRAC', 'GPA-TGA-original', 'GPA-TGA-AIII', 'GPA-TGA-Hydrated',
                 'GPA-COMPUTRAC-original', 'GPA-COMPUTRAC-AIII', 'GPA-COMPUTRAC-Hydrated']
        color = ['light_blue','light_green', 'light_orange', 'blue', 'green']
        text = ''
        chbox = [self.checkBoxTGA, self.checkBoxCOM, self.checkBoxGPA_TGA_org, self.checkBoxGPA_TGA_a3,
                 self.checkBoxGPA_TGA_hyd, self.checkBoxGPA_COM_org, self.checkBoxGPA_COM_a3, self.checkBoxGPA_COM_hyd]
        for i in range(8):
            if checks[i]:
                chbox[i].setEnabled(True)
                filename = folder + '/' + tests[i] + '.xls'
                text = text + tests[i] + '.xls is saved\n'
                workbook = xlwt.Workbook()
                style = xlwt.XFStyle()
                style.alignment.wrap = 1
                label = calcs[i][0]
                dh = calcs[i][1]
                hh = calcs[i][2]
                a3 = calcs[i][3]
                fm = calcs[i][4]
                ine = calcs[i][5]
                sheet0 = workbook.add_sheet('Overall')
                pattern = "pattern: pattern solid, fore_color " + color[
                    0] + "; font: color black, bold on; align: horiz center"
                sheet0.write_merge(0, 0, 1, 4, 'DH', xlwt.easyxf(pattern))
                pattern = "pattern: pattern solid, fore_color " + color[
                    1] + "; font: color black, bold on; align: horiz center"
                sheet0.write_merge(0, 0, 5, 8, 'HH', xlwt.easyxf(pattern))
                pattern = "pattern: pattern solid, fore_color " + color[
                    2] + "; font: color black, bold on; align: horiz center"
                sheet0.write_merge(0, 0, 9, 12, 'AIII', xlwt.easyxf(pattern))
                pattern = "pattern: pattern solid, fore_color " + color[
                    3] + "; font: color black, bold on; align: horiz center"
                sheet0.write_merge(0, 0, 13, 16, 'FM', xlwt.easyxf(pattern))
                pattern = "pattern: pattern solid, fore_color " + color[
                4] + "; font: color black, bold on; align: horiz center"
                sheet0.write_merge(0, 0, 17, 20, 'IN', xlwt.easyxf(pattern))
                for ii in range(5):
                    pattern = "pattern: pattern solid, fore_color " + color[
                        ii] + "; font: color white; align: horiz center"
                    sheet0.write(1, ii * 4 + 1, 'Mean', xlwt.easyxf(pattern))
                    sheet0.write(1, ii * 4 + 2, 'STDEV', xlwt.easyxf(pattern))
                    sheet0.write(1, ii * 4 + 3, '1st Q', xlwt.easyxf(pattern))
                    sheet0.write(1, ii * 4 + 4, '3st Q', xlwt.easyxf(pattern))
                length = 0
                for i1 in range(len(label)):
                    if label[i1]:
                        if len(str(label[i1])) > length:
                            length = len(str(label[i1]))
                        sheet0.write(2+i1,0,str(label[i1]))
                        for i2 in range(5):
                            p: object = calcs[i][i2 + 1]
                            if p[i1]:
                                pattern = "pattern: pattern solid, fore_color " + color[
                                    i2] + "; font: color black; align: horiz left"

                                if any(v != 0 for v in p[i1]):
                                    mean = np.nanmean(np.where(np.isclose(p[i1],0), np.nan, p[i1]))
                                    std = np.nanstd(np.where(np.isclose(p[i1],0), np.nan, p[i1]))
                                    q1 = np.nanquantile(np.where(np.isclose(p[i1],0), np.nan, p[i1]), 0.25)
                                    q3 = np.nanquantile(np.where(np.isclose(p[i1],0), np.nan, p[i1]), 0.75)
                                else:
                                    mean = 0
                                    std = 0
                                    q1 = 0
                                    q3 = 0

                                sheet0.write(2 + i1, i2*4 + 1, mean, xlwt.easyxf(pattern))
                                sheet0.write(2 + i1, i2*4 + 2, std, xlwt.easyxf(pattern))
                                sheet0.write(2 + i1, i2*4 + 3, q1, xlwt.easyxf(pattern))
                                sheet0.write(2 + i1, i2*4 + 4, q3, xlwt.easyxf(pattern))
                        first_col = sheet0.col(0)
                        first_col.width = 256 * length
                        name = "stucco" + str(i1 + 1)
                        sheet = workbook.add_sheet(name)
                        sheet.write_merge(0, 0, 0, 4, str(label[i1]))
                        sheet.write(1, 0, 'DH')
                        sheet.write(1, 1, 'HH')
                        sheet.write(1, 2, 'AIII')
                        sheet.write(1, 3, 'Free Moisture', style)
                        sheet.write(1, 4, 'Inert')
                        for i2 in range(len(dh[i1])):
                            sheet.write(i2 + 2, 0, float(dh[i1][i2]))
                            sheet.write(i2 + 2, 1, float(hh[i1][i2]))
                            sheet.write(i2 + 2, 2, float(a3[i1][i2]))
                            sheet.write(i2 + 2, 3, float(fm[i1][i2]))
                            sheet.write(i2 + 2, 4, float(ine[i1][i2]))
                workbook.save(filename)
        msg = QMessageBox()
        msg.setWindowTitle('Save to Excel')
        msg.setText(text)
        msg.setIcon(QMessageBox.Information)
        temp = msg.exec_()


    def save_ascii_clicked(self):
        pass


def main():
    import sys
    app = QtWidgets.QApplication(sys.argv)
    main_w = WindowMain()
    main_w.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
