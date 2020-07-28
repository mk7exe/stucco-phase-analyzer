import numpy as np
from PyQt5 import QtWidgets as qtw
from PyQt5 import QtCore as qtc


class ReadOnlyDelegate(qtw.QStyledItemDelegate):

    def createEditor(self, parent, option, index):
        return


class ComputeTable(qtw.QTableWidget):

    def __init__(self, row, col):
        super().__init__(row, col)
        self.setHorizontalHeaderLabels(['Label', 'DH', 'HH', 'AIII', 'FM', 'IN'])
        self.verticalHeader().setDefaultSectionSize(25)
        self.horizontalHeader().setDefaultSectionSize(150)
        self.horizontalHeader().setSectionResizeMode(qtw.QHeaderView.Fixed)
        stylesheet = "QHeaderView::section{border-radius: 14px;}"
        self.setStyleSheet(stylesheet)
        self.setColumnWidth(0, 400)

        # first column is read only and is updated from tab 1
        delegate = ReadOnlyDelegate(self)
        self.setItemDelegate(delegate)


class ComputeTabWidget(qtw.QWidget):

    def __init__(self, phases):
        super(qtw.QWidget, self).__init__()

        self.phases = phases
        self.layout = qtw.QVBoxLayout(self)

        # Initializing tabs
        self.tabs = qtw.QTabWidget()
        self.tab = []
        tab_names = ['TGA', 'GPA', 'GPA_ORG', 'GPA_HUM', 'GPA_HYD']
        i = 0
        for p in phases:
            self.tab.append(qtw.QWidget())
            self.tabs.addTab(self.tab[i], tab_names[i])
            self.tab[i].layout = qtw.QVBoxLayout(self)
            self.tab[i].table = ComputeTable(len(self.phases[tab_names[i]]), 6)
            ii = 0
            for s in self.phases[tab_names[i]]:
                self.tab[i].table.setItem(ii, 0, qtw.QTableWidgetItem(s))
                mean = phases[p][s]['mean']
                std = phases[p][s]['std']
                for k in range(5):
                    entry = str(np.around(mean[k], 2)) + ' (' + '\N{PLUS-MINUS SIGN}' + str(np.around(std[k], 2))\
                                + ')'
                    self.tab[i].table.setItem(ii, k + 1, qtw.QTableWidgetItem(entry))
                ii += 1
            self.tab[i].layout.addWidget(self.tab[i].table)
            self.tab[i].setLayout(self.tab[i].layout)
            i += 1

        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)


class ComputeWin(qtw.QMainWindow):

    def __init__(self, phases):
        super().__init__()
        self.setWindowModality(qtc.Qt.ApplicationModal)
        self.phases = phases
        self.ini_samples = phases
        self.title = 'Results Window'
        self.left = 100
        self.top = 100
        self.width = 1250
        l = 0
        for p in phases:
            if len(phases[p]) > l:
                l = len(phases[p])

        height = l*50 + 40
        if height > 800:
            height = 800

        self.height = height
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.w = ComputeTabWidget(self.phases)
        self.setCentralWidget(self.w)


# if __name__ == '__main__':
#     app = qtw.QApplication(sys.argv)
#     phases = {'TGA': {'DH 98% (old) MKBR9467V': {'mean': [99.99999999999999, 0.023702675928099628, 1.3141654294982622e-11, 0.0, 2.67864207549419e-17], 'std': [0.0, 0.004439166532685476, 1.811164465190758e-11, 0.0, 3.7317888448492955e-17]}, 'HH 97% SZBE2670V': {'mean': [2.637292053614291, 78.89382082518857, 0.0, 0.24395779112091276, 18.224929330076147], 'std': [1.3491201749515238, 1.417282262501928, 0.0, 0.16937935999592213, 0.23754144754643747]}, 'Drierite regular K13Z025 Powder': {'mean': [10.859609024424737, 1.912145114691061e-15, 14.99198187949763, 0.0, 74.14840909607763], 'std': [0.04685646794532872, 8.411592700734593e-17, 0.26969471754094076, 0.0, 0.22283824959561116]}, 'Drierite Pan': {'mean': [0.5055975088779824, 7.716967662150756, 68.77164607333266, 0.0, 23.005788755638644], 'std': [0.5055975084950278, 1.0718677658604516, 0.6627524399223859, 0.0, 0.09648218255691532]}, 'New HH 97%': {'mean': [1.9836491788674337, 88.26742418473505, 0.0, 0.1374358032603998, 9.611490833137134], 'std': [0.12155573545004043, 0.2224702531745521, 0.0, 0.017385933609438384, 0.11240401650096829]}}, 'GPA': {'DH 98% (old) MKBR9467V': {'mean': [99.87595692759139, 2.464769117248675e-14, 0.13849401720312648, 0.0, 6.149023262572711e-19], 'std': [0.08784085176058908, 3.4857004548318515e-14, 0.08662565147146234, 0.0, 8.616535344633927e-19]}, 'HH 97% SZBE2670V': {'mean': [4.894562812731589, 71.3080012491451, 4.98060895869797, 0.01348494001898833, 18.803342039406385], 'std': [2.0972422749701476, 7.11596484058244, 4.98060895869797, 0.01348494001898833, 1.7940437119117925]}, 'Drierite regular K13Z025 Powder': {'mean': [2.1385900884259046, 3.62419017574608e-11, 67.059865816382, 0.0, 30.80154409515586], 'std': [1.5685514057283272, 6.036407232210724e-11, 11.499660901276993, 0.0, 13.015311639930525]}, 'Drierite Pan': {'mean': [0.3069447885536849, 7.90060098461899, 69.53963901070293, 0.0, 22.25281521612441], 'std': [0.45371125766252757, 2.4953516597202103, 2.1632811099617224, 0.0, 0.3701817647775166]}, 'New HH 97%': {'mean': [1.2678910106408354, 89.08066221244093, 0.0, 0.2562768509779711, 9.395169925940273], 'std': [0.3130651444899398, 0.04783608076131164, 0.0, 0.013922377561119649, 0.315995844543531]}}, 'GPA_ORG': {'DH 98% (old) MKBR9467V': {'mean': [99.95789501962865, 7.394293974835896e-14, 0.05710216965664084, 0.0, 2.702452310086053e-24], 'std': [0.0, 0.0, 0.0, 0.0, 0.0]}, 'HH 97% SZBE2670V': {'mean': [5.988125306792239, 71.30800124914519, 4.980608958697886, 0.013484940018990699, 17.709779545345754], 'std': [2.055171367061651, 7.115964840582322, 4.980608958697886, 0.013484940018990699, 0.09366945484173783]}, 'Drierite regular K13Z025 Powder': {'mean': [2.7937261564304974, 1.0871466615696194e-10, 75.18615275489657, 0.0, 22.020121088564224], 'std': [0.18803236853910588, 5.525406088039157e-11, 0.4900648900833957, 0.0, 0.6780972585672504]}, 'Drierite Pan': {'mean': [0.6710966843762056, 8.027928684972675, 69.36740644606373, 0.0, 21.933568184587415], 'std': [0.6710966843762056, 2.4990156570202826, 2.1062577695325757, 0.0, 0.2783387968884785]}, 'New HH 97%': {'mean': [1.1746113482837792, 89.08066221244097, 0.0, 0.2562768509779829, 9.488449588297343], 'std': [0.1034050143865812, 0.0478360807613025, 0.0, 0.013922377561124558, 0.12780318029442256]}}, 'GPA_HUM': {'DH 98% (old) MKBR9467V': {'mean': [99.91584970607607, 4.807987310464473e-24, 0.09989599036215079, 0.0, 1.833447116558422e-18], 'std': [0.0, 0.0, 0.0, 0.0, 0.0]}, 'HH 97% SZBE2670V': {'mean': [3.395341318948512, 71.30800124914519, 4.980608958697925, 0.013484940019001357, 20.302563533189396], 'std': [0.3041757039277935, 7.115964840582379, 4.980608958697925, 0.013484940019001357, 2.453016525831277]}, 'Drierite regular K13Z025 Powder': {'mean': [3.6220441088472164, 1.1039115412601704e-14, 75.18615275496853, 0.0, 21.19180313618424], 'std': [0.37722367502625964, 2.7680912573004203e-15, 0.4900648900468312, 0.0, 0.11284121502057332]}, 'Drierite Pan': {'mean': [0.31254360689434574, 8.11223581304435, 69.45025298289049, 0.0, 22.12496759717084], 'std': [0.306897710132659, 2.8058010203820296, 2.3698886254441422, 0.0, 0.18546815890922302]}, 'New HH 97%': {'mean': [1.6728809852818862, 89.08066221244094, 0.0, 0.2562768509779521, 8.990179951299213], 'std': [0.0657479482404848, 0.04783608076129595, 0.0, 0.013922377561106935, 0.03911184243807809]}}, 'GPA_HYD': {'DH 98% (old) MKBR9467V': {'mean': [99.75412605706941, 1.337642932977518e-19, 0.2584838915905878, 0.0, 1.1257159761081472e-20], 'std': [0.0, 0.0, 0.0, 0.0, 0.0]}, 'HH 97% SZBE2670V': {'mean': [5.300221812454015, 71.30800124914488, 4.980608958698099, 0.013484940018972935, 18.397683039684004], 'std': [2.2958484342444625, 7.11596484058262, 4.980608958698099, 0.013484940018972935, 0.1470076123409907]}, 'Drierite regular K13Z025 Powder': {'mean': [7.089002473064803e-22, 7.863062580655441e-24, 50.80729193928089, 0.0, 49.19270806071911], 'std': [5.684217148402481e-22, 5.5010741909044126e-24, 0.16443507589076845, 0.0, 0.16443507589076845]}, 'Drierite Pan': {'mean': [0.05857803966467696, 7.604081022624506, 69.74384674827486, 0.0, 22.593494189435983], 'std': [0.06624635083893114, 2.1036442813528375, 1.958581230761621, 0.0, 0.2811151379178123]}, 'New HH 97%': {'mean': [0.9561806983568406, 89.08066221244091, 0.0, 0.2562768509779782, 9.706880238224258], 'std': [0.0955650073615191, 0.04783608076133644, 0.0, 0.013922377561127457, 0.10869884286722811]}}}
#     ex = ComputeWin(phases)
#     ex.show()
#     # ex.tab_widget.tab1.table.setItem(0, 0, qtw.QTableWidgetItem('test'))
#     sys.exit(app.exec_())
