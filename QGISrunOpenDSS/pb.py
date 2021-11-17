from __future__ import print_function
from builtins import range
import time
from qgis.PyQt import QtCore, QtGui


class SleepProgress(QtCore.QThread):
    procDone = QtCore.pyqtSignal(bool)
    partDone = QtCore.pyqtSignal(int)

    def run(self):
        # fix_print_with_import
        print('proc started')

        for a in range(1, 1 + 35):
            self.partDone.emit(float(a) / 35.0 * 100)
            # fix_print_with_import
            print('sleep', a)
            time.sleep(0.13)

        self.procDone.emit(True)
        # fix_print_with_import
        print('proc ended')


class AddProgresWin(QtGui.QWidget):
    def __init__(self, parent=None):
        super(AddProgresWin, self).__init__(parent)

        self.thread = SleepProgress()

        self.nameLabel = QtGui.QLabel("0.0%")
        self.nameLine = QtGui.QLineEdit()

        self.progressbar = QtGui.QProgressBar()
        self.progressbar.setMinimum(1)
        self.progressbar.setMaximum(100)

        mainLayout = QtGui.QGridLayout()
        mainLayout.addWidget(self.progressbar, 0, 0)
        mainLayout.addWidget(self.nameLabel, 0, 1)

        self.setLayout(mainLayout)
        self.setWindowTitle("Processing")

        self.thread.partDone.connect(self.updatePBar)
        self.thread.procDone.connect(self.fin)

        self.thread.start()

    def updatePBar(self, val):
        self.progressbar.setValue(val)

        perct = "{0}%".format(val)
        self.nameLabel.setText(perct)

    def fin(self):
        sys.exit()
        ##self.hide()


if __name__ == '__main__':
    import sys

    app = QtGui.QApplication(sys.path)

    pbarwin = AddProgresWin()
    pbarwin.show()

    sys.exit(app.exec_())
