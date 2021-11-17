# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ves_gui.ui'
#
# Created: Mon Mar 21 17:23:48 2016
#      by: PyQt4 UI code generator 4.10.2
#
# WARNING! All changes made in this file will be lost!

import os

from PyQt5 import QtWidgets, uic, QtCore, QtGui
from PyQt5.QtWidgets import QDialog, QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMessageBox
from PyQt5.QtCore import QSettings, QTranslator, qVersion, QCoreApplication, QObject, Qt
import sys


FORM_CLASS, _ = uic.loadUiType(os.path.join(
os.path.dirname(__file__), 'ves_gui.ui'))

class Ui_Ves(QDialog, FORM_CLASS):
    def __init__(self, parent=None):
        """Constructor."""
        super(Ui_Ves, self).__init__(parent)
        self.setupUi(self)
        
class GUI_Ves(object):
    """QGIS Plugin Implementation."""

    def __init__(self):
        self.dlg = Ui_Ves()
        
        self.dlg.lineEdit_ev_file.setEnabled(False)
        self.dlg.pushButton_ev_file.setEnabled(False)
    
        self.dlg.pushButton_ev_file.clicked.connect(self.select_ev_file)
        self.dlg.checkBox_EV_prob.stateChanged.connect(self.evs_file_enable)
        #Checks excluyentes
        self.dlg.checkBox_EV_prob.stateChanged.connect(self.evs_change_check_prob)
        self.dlg.checkBox_EV_consum.stateChanged.connect(self.evs_change_check_consum)
        self.dlg.checkBox_EV_aleat.stateChanged.connect(self.evs_change_check_aleat)
        
        #Verificar estado de checks (que esté seleccionado al menos uno)
        self.dlg.checkBox_EV_prob.stateChanged.connect(self.check_study_mode_ev)
        self.dlg.checkBox_EV_consum.stateChanged.connect(self.check_study_mode_ev)
        self.dlg.checkBox_EV_aleat.stateChanged.connect(self.check_study_mode_ev)
        
    #=======================================================================================================
    #=======================================================================================================
    #Cambia el estado de los checks de EVs, ya que cada modo es excluyente
    #Signal check EVs prob
    def evs_change_check_prob(self):
        if self.dlg.checkBox_EV_prob.isChecked():
            self.dlg.checkBox_EV_consum.setChecked(False)
            self.dlg.checkBox_EV_aleat.setChecked(False)
    
    #Signal check EVs consumo
    def evs_change_check_consum(self):
        if self.dlg.checkBox_EV_consum.isChecked():
            self.dlg.checkBox_EV_prob.setChecked(False)
            self.dlg.checkBox_EV_aleat.setChecked(False)
    
    #Signal check EVs aleatorio
    def evs_change_check_aleat(self):
        if self.dlg.checkBox_EV_aleat.isChecked():
            self.dlg.checkBox_EV_prob.setChecked(False)
            self.dlg.checkBox_EV_consum.setChecked(False)
        
    #Verifica que se seleccione una opción válida para el tipo de estudio de EVs. Si no hay ninguno seleccionado elige random
    def check_study_mode_ev(self):
        random = self.dlg.checkBox_EV_aleat.isChecked()
        consum = self.dlg.checkBox_EV_consum.isChecked()
        prob = self.dlg.checkBox_EV_prob.isChecked()
        
        if random == 0 and consum == 0 and prob == 0:
            self.dlg.checkBox_EV_aleat.setChecked(True) #si todos están deshabilitados realiza un estudio random
            
    
    #Enable/disable button EV information
    def evs_file_enable(self):
        if self.dlg.checkBox_EV_prob.isChecked():
            self.dlg.lineEdit_ev_file.setEnabled(True)
            self.dlg.pushButton_ev_file.setEnabled(True)
        else:
            self.dlg.lineEdit_ev_file.setEnabled(False)
            self.dlg.pushButton_ev_file.setEnabled(False)
            
    def select_ev_file(self):
        load_ev_file = QFileDialog.getOpenFileName(None, "Seleccione el archivo .CSV para asignar la información de los vehículos eléctricos", "", "*.csv")
        load_ev_file = load_ev_file[0]
        self.dlg.lineEdit_ev_file.setText(load_ev_file)
    
      
    
if __name__ == "__main__":
	app = QtWidgets.QApplication(sys.argv)
	gui = GUI_Ves()    
	gui.dlg.show()
	result = gui.dlg.exec_()
	
