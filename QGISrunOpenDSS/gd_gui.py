# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'gd_gui.ui'
#
# Created: Mon Mar 21 17:23:48 2016
#      by: PyQt4 UI code generator 4.10.2
#
# WARNING! All changes made in this file will be lost!

import os

#from qgis.PyQt import QtGui, uic
from PyQt5 import QtWidgets, uic, QtCore, QtGui
from PyQt5.QtWidgets import QDialog, QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMessageBox
from PyQt5.QtCore import QSettings, QTranslator, qVersion, QCoreApplication, QObject, Qt
import sys


FORM_CLASS, _ = uic.loadUiType(os.path.join(
os.path.dirname(__file__), 'gd_gui.ui'))

class Ui_Gd(QDialog, FORM_CLASS):

   def __init__(self, parent=None):
	   
      """Constructor."""
      super(Ui_Gd, self).__init__(parent)
      self.setupUi(self)
      
class GUI_Gd(object):
    """QGIS Plugin Implementation."""

    def __init__(self):
        self.dlg = Ui_Gd()
        self.pv_allocation_enabler()
        
        self.dlg.comboBoxPVTipoEstudio.addItem(QCoreApplication.translate('dialog', 'Aleatorio'))
        self.dlg.comboBoxPVTipoEstudio.addItem(QCoreApplication.translate('dialog', "Mayor consumo"))
        self.dlg.comboBoxPVTipoEstudio.addItem(QCoreApplication.translate('dialog', "Por probabilidad"))
        
        self.dlg.pushButton_pv_info.clicked.connect(self.select_pv_info)
        self.dlg.pushButton_pv_info_gen.clicked.connect(self.select_pv_info_gen)
        self.dlg.pushButton_pv_info_TMT.clicked.connect(self.select_pv_info_TMT)
        
        # PV allocation
        self.dlg.checkBox_PV_res.toggled.connect(self.pv_allocation_enabler)
        self.dlg.checkBox_PV_gen.toggled.connect(self.pv_allocation_enabler)
        self.dlg.checkBox_PV_TMT.toggled.connect(self.pv_allocation_enabler)
    
    def pv_allocation_enabler(self):
        self.dlg.lineEdit_pv_capacity.setEnabled(False)
        self.dlg.lineEdit_pv_capacity_gen.setEnabled(False)
        self.dlg.lineEdit_pv_capacity_TMT.setEnabled(False)
        self.dlg.lineEdit_pv_info.setEnabled(False)
        self.dlg.lineEdit_pv_info_gen.setEnabled(False)
        self.dlg.lineEdit_pv_info_TMT.setEnabled(False)
        self.dlg.pushButton_pv_info.setEnabled(False)
        self.dlg.pushButton_pv_info_gen.setEnabled(False)
        self.dlg.pushButton_pv_info_TMT.setEnabled(False)

        if self.dlg.checkBox_PV_res.isChecked():
            self.dlg.lineEdit_pv_capacity.setEnabled(True)
            self.dlg.lineEdit_pv_info.setEnabled(True)
            self.dlg.pushButton_pv_info.setEnabled(True)

        if self.dlg.checkBox_PV_gen.isChecked():
            self.dlg.lineEdit_pv_capacity_gen.setEnabled(True)
            self.dlg.lineEdit_pv_info_gen.setEnabled(True)
            self.dlg.pushButton_pv_info_gen.setEnabled(True)

        if self.dlg.checkBox_PV_TMT.isChecked():
            self.dlg.lineEdit_pv_capacity_TMT.setEnabled(True)
            self.dlg.lineEdit_pv_info_TMT.setEnabled(True)
            self.dlg.pushButton_pv_info_TMT.setEnabled(True)
            
    def select_pv_info_gen(self):
        load_pv_info = QFileDialog.getOpenFileName(self.dlg, QCoreApplication.translate('dialog',
                                                                                        u'Seleccione el archivo .CSV para asignar la información de los sistemas fotovoltaicos'),
                                                   "", "*.csv")
        load_pv_info = load_pv_info[0]
        self.dlg.lineEdit_pv_info_gen.setText(load_pv_info)

    def select_pv_info_TMT(self):
        load_pv_info = QFileDialog.getOpenFileName(self.dlg, QCoreApplication.translate('dialog',
                                                                                        u'Seleccione el archivo .CSV para asignar la información de los sistemas fotovoltaicos'),
                                                   "", "*.csv")
        load_pv_info = load_pv_info[0]
        self.dlg.lineEdit_pv_info_TMT.setText(load_pv_info)

    def select_pv_info(self):
        load_pv_info = QFileDialog.getOpenFileName(self.dlg, QCoreApplication.translate('dialog',
                                                                                        u'Seleccione el archivo .CSV para asignar la información de los sistemas fotovoltaicos'),
                                                   "", "*.csv")        

        load_pv_info = load_pv_info[0]
        self.dlg.lineEdit_pv_info.setText(load_pv_info)


      
    

      
if __name__ == "__main__":
	app = QtWidgets.QApplication(sys.argv)
	hola = GUI_Gd()
	hola.dlg.show()
	result = hola.dlg.exec_()
	
	
