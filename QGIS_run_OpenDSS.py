# -*- coding: utf-8 -*-
"""
/***************************************************************************
 QGISrunOpenDSS
                                 A QGIS plugin
 This plugin runs network studies using OpenDSS
                              -------------------
        begin                : 2015-12-03
        git sha              : $Format:%H$
        copyright            : (C) 2015 by EPERLab
        email                : eperlab.eie@ucr.ac.cr
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""
""" Import the basic libraries used through the plugin"""
#from __future__ import print_function
#from __future__ import absolute_import
from builtins import next
from builtins import zip
from builtins import str
from builtins import range
from builtins import object
import os
from qgis.PyQt.QtCore import QSettings, QTranslator, qVersion, QCoreApplication, QObject, Qt
from qgis.PyQt.QtWidgets import QAction, QFileDialog, QMessageBox, QComboBox
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt import uic
from . import resources  # Initialize Qt resources from file resources.py
from .QGIS_run_OpenDSS_dialog import QGISrunOpenDSSDialog  # Import the code for the dialog
import numpy as np  # import random
from . import auxiliary_functions as auxfcns  # File to define some usefull class/functions in python to do specific tasks
from . import Harmonics  # File that contains harmonic simulation
from matplotlib import pyplot as plt
import matplotlib.patches as mpatches
import xlrd
import random
from collections import OrderedDict
import time
import csv
from qgis.core import QgsProject
import glob
import re
from PyQt5.QtGui import *  # Para desplegar mensajes, útil para debugin
from .QGIS_run_OpenDSS_transformer import Ui_Transformer
from .QGIS_run_OpenDSS_progress import Ui_Progress
from .ves_gui import GUI_Ves
from .gd_gui import GUI_Gd
import datetime as dt
import matplotlib.dates as mdates

from PyQt5.QtCore import *
from PyQt5 import QtCore 

from PyQt5 import QtGui #Paquetes requeridos para crear ventanas de diálogo e interfaz gráfica.
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMessageBox, QDialog, QStyleFactory
import traceback
import sys
import copy
import fileinput

from .evs_code.create_study_EVS import CreateEVDss #Para estudio random, por consumo o por probabilidad de EV
from .evs_code.EVsFunctions import CreateList_SOC_t, update_storage #Para crear lista de dssnames y t's
from .correct_date_time import correct_date, correct_hour #para corregir la fecha y tiempo

QApplication.setStyle(QStyleFactory.create('WindowsVista'))  # <- Choose the style

# noinspection PyCallByClass
class QGISrunOpenDSS(object):
    def __init__(self, iface):
        # Save reference to the QGIS interface
        self.iface = iface
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        # initialize locale
        locale = QSettings().value('locale/userLocale')[0:2]  # This line is used to select the OS language
        if locale != (u'en'):
            locale = (u'es')

        locale_path = os.path.join(self.plugin_dir, 'i18n',
                                   'QGISrunOpenDSS_{}.qm'.format(locale))  # This line links to the translation file

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)

            if qVersion() > '4.3.3':
                QCoreApplication.installTranslator(self.translator)

        # Create the dialog (after translation) and keep reference
        self.dlg = QGISrunOpenDSSDialog()
        self.transformer = Ui_Transformer()
        self.gui_ev = GUI_Ves()
        self.gui_gd = GUI_Gd()
        
        self.dlg.buttonBox.helpRequested.connect(self.show_help)
        self.dlg.lineEdit_dirOutput.clear()
        self.dlg.lineEdit_dirOutput.setText('C:\Results_QGISPython')
        # self.harm = Ui_harmonics()
        self.progress = Ui_Progress()
        # iface.mainWindow().statusBar().addWidget(self.progress)

        # Declare instance attributes
        self.actions = []
        self.menu = self.tr(QCoreApplication.translate('dialog', u'&Correr estudio de red usando OpenDSS'))

        self.toolbar = self.iface.addToolBar(
            QCoreApplication.translate('dialog', u'Correr estudio de red usando OpenDSS'))
        self.toolbar.setObjectName(QCoreApplication.translate('dialog', u'Correr estudio de red usando OpenDSS'))

        self.dlg.pushButton_output_folder.clicked.connect(self.select_output_folder)
        self.dlg.pushButton_load_curve.clicked.connect(self.select_demand_curve)
        self.dlg.pushButton_transformer.clicked.connect(self.transformer_button)
        # self.dlg.pushButton_harmonics.clicked.connect(self.harmonic_button)
        self.dlg.pushButton_res_spectrum.clicked.connect(self.select_res_spectrum)
        self.dlg.pushButton_gen_spectrum.clicked.connect(self.select_gen_spectrum)
        self.dlg.pushButton_TMT_spectrum.clicked.connect(self.select_TMT_spectrum)
        

        # SC checkboxex
        self.dlg.checkBox_SC_phaseA.toggled.connect(self.short_circuit_enabler)
        self.dlg.checkBox_SC_phaseB.toggled.connect(self.short_circuit_enabler)
        self.dlg.checkBox_SC_phaseC.toggled.connect(self.short_circuit_enabler)
        self.short_circuit_enabler()
        self.dlg.checkBox_SC_phaseLL.setEnabled(True)

        

        # harmonic enabler
        self.dlg.harmonics.toggled.connect(self.harmonics_enabler)
        self.dlg.checkBox_harm_res.toggled.connect(self.harmonics_enabler)
        self.dlg.checkBox_harm_gen.toggled.connect(self.harmonics_enabler)
        self.dlg.checkBox_harm_TMT.toggled.connect(self.harmonics_enabler)

        self.dlg.powerflow_snapshot.setChecked(True)
        self.dlg.powerflow_snapshot.toggled.connect(self.SLUenabler)
        self.dlg.powerflow_daily.toggled.connect(self.SLUenabler)
        self.dlg.powerflow_yearly.toggled.connect(self.SLUenabler)
        self.dlg.short_circuit.toggled.connect(self.SLUenabler)

        self.dlg.SCMVA.toggled.connect(self.scPowerAmps)
        self.dlg.SCkA.toggled.connect(self.scPowerAmps)

        self.SLUenabler()
        

        self.dlg.lineEdit_snapshot_date.clear()
        self.dlg.lineEdit_snapshot_time.clear()
        self.dlg.lineEdit_sc_date.clear()
        self.dlg.lineEdit_sc_time.clear()
        self.dlg.lineEdit_daily_date.clear()
        self.dlg.lineEdit_yearly.clear()
        self.dlg.lineEdit_load_curve.clear()
        self.dlg.lineEdit_dirOutput.clear()

        self.dlg.lineEdit_harmonics.clear()
        self.dlg.lineEdit_harmonic_date.clear()
        self.dlg.lineEdit_harmonic_time.clear()
        self.dlg.lineEdit_res_spectrum.clear()
        
        
        #Flujo de potencia
        self.dlg.powerflow_yearly.toggled.connect( self.change_check_yearly_powerflow )
        self.dlg.powerflow_daily.toggled.connect( self.change_check_daily_powerflow )
        self.dlg.powerflow_snapshot.toggled.connect( self.change_check_snapshot_powerflow )
        
        #Armónicos
        self.dlg.harmonics.toggled.connect( self.change_check_harmonics )
        self.dlg.daily.toggled.connect( self.change_check_daily_harmonics )
        self.dlg.snapshot.toggled.connect( self.change_check_snapshot_harmonics )
        
        #Cortocircuito
        self.dlg.short_circuit.toggled.connect( self.change_check_shortcircuit )
        
        #Botones de EV y GD
        self.dlg.pushButton_EV.clicked.connect(self.EV_button)
        self.dlg.pushButton_dg.clicked.connect(self.GD_button)
        self.dlg.EV.toggled.connect( self.evs_change_check )
        self.dlg.PV.toggled.connect( self.pvs_change_check )
        self.gd_button_push = False
        self.ev_button_push = False
        
        
        #####################################################################
        ComboBoxChildren = self.dlg.findChildren(QComboBox)
        ComboBox_Sub = ComboBoxChildren[0]
        ComboBox_Sub.activated.connect(lambda: self.cktNameDefaults())
        self.substation = ''

        # Combo_Fill_Short_circuit
        self.dlg.lineEdit_short_circuit.clear()
        try:
            buses = auxfcns.busSC()
            self.dlg.lineEdit_short_circuit.addItems(buses)
        except:
            self.dlg.lineEdit_short_circuit.clear()
            buses = [QCoreApplication.translate('SC', "Todas las barras MT"),
                     QCoreApplication.translate('SC', 'Todas las barras BT'),
                     QCoreApplication.translate('SC', 'Todas las barras MT y BT')]
            self.dlg.lineEdit_short_circuit.addItems(buses)
            pass

    # Enable/disable short circuit checkboxes
    def short_circuit_enabler(self):
        if self.dlg.checkBox_SC_phaseA.isChecked() and self.dlg.checkBox_SC_phaseB.isChecked() and not self.dlg.checkBox_SC_phaseC.isChecked():
            self.dlg.checkBox_SC_phaseLL.setEnabled(True)
        if not self.dlg.checkBox_SC_phaseA.isChecked() and self.dlg.checkBox_SC_phaseB.isChecked() and self.dlg.checkBox_SC_phaseC.isChecked():
            self.dlg.checkBox_SC_phaseLL.setEnabled(True)
        if self.dlg.checkBox_SC_phaseA.isChecked() and not self.dlg.checkBox_SC_phaseB.isChecked() and self.dlg.checkBox_SC_phaseC.isChecked():
            self.dlg.checkBox_SC_phaseLL.setEnabled(True)
        if self.dlg.checkBox_SC_phaseA.isChecked() and self.dlg.checkBox_SC_phaseB.isChecked() and self.dlg.checkBox_SC_phaseC.isChecked():
            self.dlg.checkBox_SC_phaseLL.setEnabled(True)
        if self.dlg.checkBox_SC_phaseA.isChecked() and not self.dlg.checkBox_SC_phaseB.isChecked() and not self.dlg.checkBox_SC_phaseC.isChecked():
            self.dlg.checkBox_SC_phaseLL.setEnabled(False)
        if not self.dlg.checkBox_SC_phaseA.isChecked() and self.dlg.checkBox_SC_phaseB.isChecked() and not self.dlg.checkBox_SC_phaseC.isChecked():
            self.dlg.checkBox_SC_phaseLL.setEnabled(False)
        if not self.dlg.checkBox_SC_phaseA.isChecked() and not self.dlg.checkBox_SC_phaseB.isChecked() and self.dlg.checkBox_SC_phaseC.isChecked():
            self.dlg.checkBox_SC_phaseLL.setEnabled(False)

    # Switch between kA and MVA for short circuit
    def scPowerAmps(self):
        if self.dlg.SCkA.isChecked():
            self.dlg.Box_circuit_definition_5.setTitle(QCoreApplication.translate('dialog',
                                                                                  u'Ingrese kA 3p de corto circuito'))
            self.dlg.Box_circuit_definition_6.setTitle(QCoreApplication.translate('dialog',
                                                                                  u'Ingrese kA 1p de corto circuito'))
        if self.dlg.SCMVA.isChecked():
            self.dlg.Box_circuit_definition_5.setTitle(QCoreApplication.translate('dialog',
                                                                                  u'Ingrese MVA 3p de corto circuito'))
            self.dlg.Box_circuit_definition_6.setTitle(QCoreApplication.translate('dialog',
                                                                                  u'Ingrese MVA 1p de corto circuito'))

    # Losses, unbalance and save to shapes checkboxes enabler
    def SLUenabler(self):
        if self.dlg.powerflow_snapshot.isChecked() or self.dlg.powerflow_daily.isChecked() or self.dlg.short_circuit.isChecked():
            self.dlg.checkBoxLosses.setEnabled(True)
            self.dlg.checkBoxUnbalance.setEnabled(True)
            self.dlg.checkBox_capas.setEnabled(True)
        else:
            self.dlg.checkBoxLosses.setEnabled(False)
            self.dlg.checkBoxUnbalance.setEnabled(False)
            self.dlg.checkBox_capas.setEnabled(False)


    #=======================================================================================================
    #=======================================================================================================
    
    #Abre la ventana de EVs cuando se habilita el check la primera vez
    def evs_change_check(self):
        if self.ev_button_push == False and self.dlg.EV.isChecked():
            self.gui_ev.dlg.show()
            self.result_ev = self.gui_ev.dlg.exec_()
            self.ev_button_push = True
            return
            
    #Abre la ventana de PVs cuando se habilita el check la primera vez
    def pvs_change_check(self):
        if self.gd_button_push == False and self.dlg.PV.isChecked():
            self.gui_gd.dlg.show()
            self.result_gd = self.gui_gd.dlg.exec_()
            self.gd_button_push = True
            return

    
    #Funcion para abrir ventada de EVs
    def EV_button(self):
        self.gui_ev.dlg.show()
        self.result_ev = self.gui_ev.dlg.exec_()
        self.ev_button_push = True
        return
    
    #Funcion para abrir ventada de generacion distribuida
    def GD_button(self):
        self.gui_gd.dlg.show()
        self.result_gd = self.gui_gd.dlg.exec_()
        self.gd_button_push = True
        return
		
    #Cheks excluyentes
    #Signal check snapshot powerflow
    def change_check_snapshot_powerflow(self):
        if self.dlg.powerflow_snapshot.isChecked():
            self.dlg.powerflow_daily.setChecked(False)
            self.dlg.powerflow_yearly.setChecked(False)
            self.dlg.short_circuit.setChecked(False)
            self.dlg.harmonics.setChecked(False)
            self.dlg.daily.setChecked(False)
            self.dlg.snapshot.setChecked(False)
        return
    
    #Signal check daily powerflow
    def change_check_daily_powerflow(self):
        if self.dlg.powerflow_daily.isChecked():
            self.dlg.powerflow_snapshot.setChecked(False)
            self.dlg.powerflow_yearly.setChecked(False)
            self.dlg.short_circuit.setChecked(False)
            self.dlg.harmonics.setChecked(False)
            self.dlg.daily.setChecked(False)
            self.dlg.snapshot.setChecked(False)
        return
            
    #Signal check yearly powerflow
    def change_check_yearly_powerflow(self):
        if self.dlg.powerflow_yearly.isChecked():
            self.dlg.powerflow_daily.setChecked(False)
            self.dlg.powerflow_snapshot.setChecked(False)
            self.dlg.short_circuit.setChecked(False)
            self.dlg.harmonics.setChecked(False)
            self.dlg.daily.setChecked(False)
            self.dlg.snapshot.setChecked(False)
        return
            
    #Signal short circuit
    def change_check_shortcircuit(self):
        if self.dlg.short_circuit.isChecked():
            self.dlg.powerflow_yearly.setChecked(False)
            self.dlg.powerflow_daily.setChecked(False)
            self.dlg.powerflow_snapshot.setChecked(False)
            self.dlg.harmonics.setChecked(False)
            self.dlg.daily.setChecked(False)
            self.dlg.snapshot.setChecked(False)
        return
            
    #Signal harmonics
    def change_check_harmonics(self):
        if self.dlg.harmonics.isChecked():
            self.dlg.powerflow_yearly.setChecked(False)
            self.dlg.powerflow_daily.setChecked(False)
            self.dlg.powerflow_snapshot.setChecked(False)
            self.dlg.daily.setChecked(False)
            self.dlg.snapshot.setChecked(True) #por defecto estará habilitado el análisis de armónicos instantáneo
            self.dlg.short_circuit.setChecked(False)
        return
        
    #Signal snapshot harmonics    
    def change_check_snapshot_harmonics(self):
        if self.dlg.snapshot.isChecked():
            self.dlg.powerflow_yearly.setChecked(False)
            self.dlg.powerflow_daily.setChecked(False)
            self.dlg.powerflow_snapshot.setChecked(False)
            self.dlg.daily.setChecked(False)
            self.dlg.short_circuit.setChecked(False)
        return
            
    #Signal daily harmonics
    def change_check_daily_harmonics(self):
        if self.dlg.daily.isChecked():
            self.dlg.powerflow_yearly.setChecked(False)
            self.dlg.powerflow_daily.setChecked(False)
            self.dlg.powerflow_snapshot.setChecked(False)
            self.dlg.snapshot.setChecked(False)
            self.dlg.short_circuit.setChecked(False)
        return
    
    
    def harmonics_enabler(self):
        if self.dlg.harmonics.isChecked():
            self.dlg.lineEdit_res_spectrum.setEnabled(False)
            self.dlg.lineEdit_gen_spectrum.setEnabled(False)
            self.dlg.lineEdit_TMT_spectrum.setEnabled(False)
            self.dlg.pushButton_res_spectrum.setEnabled(False)
            self.dlg.pushButton_gen_spectrum.setEnabled(False)
            self.dlg.pushButton_TMT_spectrum.setEnabled(False)

            if self.dlg.checkBox_harm_res.isChecked():
                self.dlg.lineEdit_res_spectrum.setEnabled(True)
                self.dlg.pushButton_res_spectrum.setEnabled(True)

            if self.dlg.checkBox_harm_gen.isChecked():
                self.dlg.lineEdit_gen_spectrum.setEnabled(True)
                self.dlg.pushButton_gen_spectrum.setEnabled(True)

            if self.dlg.checkBox_harm_TMT.isChecked():
                self.dlg.lineEdit_TMT_spectrum.setEnabled(True)
                self.dlg.pushButton_TMT_spectrum.setEnabled(True)

    # fill the circuit comboBox
    def circuit_calling(self):
        self.dlg.lineEdit_circuit_name.clear()
        prjpath = QgsProject.instance().fileName()
        if not prjpath:
            return
        print( "prjpath = ", prjpath )
        dir_general, dir_general2 = prjpath.split('/GIS', 1)
        dir_network = dir_general + '/DSS'  # Folder where all network models are stored
        os.chdir(dir_network)
        networks = dir_general.split('/')[-1]  # [d for d in os.listdir('.') if os.path.isdir(d)]
        self.dlg.lineEdit_circuit_name.setText(networks)
        # self.dlg.comboBox_circuit_name.addItems(networks)

    # Method to select the output folder
    def select_output_folder(self):
        output_folder = QFileDialog.getExistingDirectory(self.dlg, QCoreApplication.translate('dialog',
                                                                                              "Seleccione carpeta de salida"),
                                                         "", )
        self.dlg.lineEdit_dirOutput.setText(output_folder)

    # Method to select the load curve of the circuit
    def select_demand_curve(self):
        #load_curve_circuit, __, __ = QFileDialog.getOpenFileName(self.dlg, QCoreApplication.translate('dialog', 'Seleccione el archivo .CSV para asignar curva de demanda de circuito'), "", "*.csv")
        load_curve_circuit = QFileDialog.getOpenFileName(self.dlg, QCoreApplication.translate('dialog', 'Seleccione el archivo .CSV para asignar curva de demanda de circuito'), "", "*.csv")
        load_curve_circuit = load_curve_circuit[0]
        print( "load_curve_circuit = ", str( load_curve_circuit ) )
        self.dlg.lineEdit_load_curve.setText(load_curve_circuit)

    def transformer_button(self):  # Method to call the main transformer configuration window
        self.transformer.show()

    def harmonic_button(self):  # Method to call the harmonic configuration window
        self.harm.show()
    
    def select_res_spectrum(self):  # Method to select the residential harmonic spectrum
        load_res_spectrum = QFileDialog.getOpenFileName(self.dlg, QCoreApplication.translate('dialog',
                                                                                             u'Seleccione el archivo .XLS para asignar el espectro de armónicos residencial'),
                                                        "", "*.xls")
        load_res_spectrum = load_res_spectrum[0]
        self.dlg.lineEdit_res_spectrum.setText(load_res_spectrum)
        

    def select_gen_spectrum(self):  # Method to select the residential harmonic spectrum
        load_spectrum = QFileDialog.getOpenFileName(self.dlg, QCoreApplication.translate('dialog',
                                                                                             u'Seleccione el archivo .XLS para asignar el espectro de armónicos comercial e industrial'),
                                                        "", "*.xls")
        load_spectrum = load_spectrum[0]
        self.dlg.lineEdit_gen_spectrum.setText(load_spectrum)

    def select_TMT_spectrum(self):  # Method to select the residential harmonic spectrum
        load_spectrum = QFileDialog.getOpenFileName(self.dlg, QCoreApplication.translate('dialog',
                                                                                             u'Seleccione el archivo .XLS para asignar el espectro de armónicos de media tensión'),
                                                        "", "*.xls")
        load_spectrum = load_spectrum[0]
        self.dlg.lineEdit_TMT_spectrum.setText(load_spectrum)

    # change the parameters according to the selected circuit
    def cktNameDefaults(self):
        prjpath = QgsProject.instance().fileName()
        if not prjpath:
            return
        dir_general, dir_general2 = prjpath.split('/GIS', 1)
        dir_network = dir_general + '/DSS'  # Folder where network model is stored
        # mydir = dir_network + '\\' + self.dlg.comboBox_circuit_name.currentText()
        os.chdir(dir_network)
        for file in glob.glob("*.dss"):
            file_check = re.sub('\.dss$', '', file)
            file_check = file_check[-10:]
            file_name = re.sub('\.dss$', '', file).replace('_' + file_check, '')
            if file_check == 'Substation':
                file_name_path = str(os.getcwd()) + '\\' + '\\' + file
                fp = open(file_name_path)                
                lines = fp.readlines()
                line = lines[1]
                line = line.split(' ')
                self.dlg.lineEdit_frequency.clear()
                self.dlg.lineEdit_frequency.setText('60')

                self.dlg.lineEdit_phase.clear()
                self.dlg.lineEdit_phase.setText('0')

                #self.dlg.lineEdit_3pShortCircuitPower.clear()
                if self.dlg.lineEdit_3pShortCircuitPower.text() == '':
                    self.dlg.lineEdit_3pShortCircuitPower.setText('10')

                #self.dlg.lineEdit_1pShortCircuitPower.clear()
                if self.dlg.lineEdit_1pShortCircuitPower.text() == '':
                    self.dlg.lineEdit_1pShortCircuitPower.setText('10.5')
                if "UNIT" in lines[0]: # substation is single unit
                    self.transformer.checkBox_tertiary.setCheckState(0)
                    for ij in range(len(line)):
                        line_check = line[ij]
                        # Windings fillup
                        if line_check == 'windings=3':
                            self.transformer.checkBox_tertiary.setCheckState(1)
                            # Voltages fillup
                        if line_check[:5] == 'kVs=[':
                            def_V_prim = line_check.replace('kVs=[', '')
                            self.dlg.lineEdit_voltage.clear()
                            self.dlg.lineEdit_voltage.setText(def_V_prim)
                            self.transformer.lineEdit_voltage_high.clear()
                            self.transformer.lineEdit_voltage_high.setText(def_V_prim)
                            V_ter = line[ij + 2]
                            if V_ter[-1:] == ']':
                                V_sec = line[ij + 1]
                                self.transformer.lineEdit_voltage_low.clear()
                                self.transformer.lineEdit_voltage_low.setText(V_sec)
                                V_ter = V_ter.replace(']', '')
                                self.transformer.lineEdit_voltage_tertiary.clear()
                                self.transformer.lineEdit_voltage_tertiary.setText(V_ter)
                            else:
                                V_sec = line[ij + 1]
                                V_sec = V_sec.replace(']', '')
                                self.transformer.lineEdit_voltage_low.clear()
                                self.transformer.lineEdit_voltage_low.setText(V_sec)
                        # Powers fillup
                        if line_check[:6] == 'kVAs=[':
                            P_prim = line_check.replace('kVAs=[', '')
                            self.transformer.lineEdit_power_high.clear()
                            self.transformer.lineEdit_power_high.setText(P_prim)
                            P_ter = line[ij + 2]
                            if P_ter[-1:] == ']':
                                P_sec = line[ij + 1]
                                self.transformer.lineEdit_power_low.clear()
                                self.transformer.lineEdit_power_low.setText(P_sec)
                                P_ter = P_ter.replace(']', '')
                                self.transformer.lineEdit_power_tertiary.clear()
                                self.transformer.lineEdit_power_tertiary.setText(P_ter)
                            else:
                                P_sec = line[ij + 1]
                                P_sec = P_sec.replace(']', '')
                                self.transformer.lineEdit_power_low.clear()
                                self.transformer.lineEdit_power_low.setText(P_sec)
                        # Reactances fillup
                        if line_check[:4] == 'xhl=':
                            X_hl = line_check.replace('xhl=', '')
                            self.transformer.lineEdit_x_hl.clear()
                            self.transformer.lineEdit_x_hl.setText(X_hl)
                        if line_check[:4] == 'xht=':
                            X_ht = line_check.replace('xht=', '')
                            self.transformer.lineEdit_x_ht.clear()
                            self.transformer.lineEdit_x_ht.setText(X_ht)
                        if line_check[:4] == 'xlt=':
                            X_lt = line_check.replace('xlt=', '')
                            self.transformer.lineEdit_x_lt.clear()
                            self.transformer.lineEdit_x_lt.setText(X_lt)
                            # Losses fillup
                        if line_check[:12] == '%noloadloss=':
                            noloadloss = line_check.replace('%noloadloss=', '')
                            self.transformer.lineEdit_losses_no_load.clear()
                            self.transformer.lineEdit_losses_no_load.setText(noloadloss)
                        if line_check[:10] == '%loadloss=':
                            inloadloss = line_check.replace('%loadloss=', '')
                            self.transformer.lineEdit_losses_in_load.clear()
                            self.transformer.lineEdit_losses_in_load.setText(inloadloss)
                        # Main bus fillup
                        if line_check[:7] == 'buses=[':
                            bus = line[ij + 1]
                            bus = bus.split('.')
                            bus = bus[0]
                            self.dlg.lineEdit_name_busbar.clear()
                            self.dlg.lineEdit_name_busbar.setText(bus)
                        # Taps fillup
                        if line_check[:8] == 'numtaps=':
                            taps = line_check.replace('numtaps=', '')
                            self.transformer.lineEdit_no_taps.setText(taps)
                        if line_check[:4] == 'tap=':
                            tap = line_check.replace('tap=', '')
                            self.transformer.lineEdit_tap_position.setText(tap)
                        if line_check[:7] == 'maxtap=':
                            maxtap = line_check.replace('maxtap=', '')
                            self.transformer.lineEdit_max_tap.setText(maxtap)
                        if line_check[:7] == 'mintap=':
                            mintap = line_check.replace('mintap=', '')
                            self.transformer.lineEdit_min_tap.setText(mintap)
                        # Connections fillup
                        if line_check[:7] == 'conns=[':
                            conP = line_check.replace('conns=[', '')
                            if conP == 'wye':
                                self.transformer.comboBox_conn_high.setCurrentIndex(0)
                            elif conP == 'delta':
                                self.transformer.comboBox_conn_high.setCurrentIndex(1)

                        if line_check[:7] == 'conns=[':
                            conP = line_check.replace('conns=[', '')
                            if conP == 'wye':
                                self.transformer.comboBox_conn_high.setCurrentIndex(0)
                            elif conP == 'delta':
                                self.transformer.comboBox_conn_high.setCurrentIndex(1)
                            conT = line[ij + 2]
                            conS = line[ij + 1]
                            if conT[-1:] == ']':
                                if conT == 'wye]':
                                    self.transformer.comboBox_conn_tertiary.setCurrentIndex(0)
                                elif conT == 'delta]':
                                    self.transformer.comboBox_conn_tertiary.setCurrentIndex(1)
                                if conS == 'wye':
                                    self.transformer.comboBox_conn_low.setCurrentIndex(0)
                                else:
                                    self.transformer.comboBox_conn_low.setCurrentIndex(1)
                            else:
                                if conS == 'wye]':
                                    self.transformer.comboBox_conn_low.setCurrentIndex(0)
                                elif conS == 'delta]':
                                    self.transformer.comboBox_conn_low.setCurrentIndex(1)
                        if line_check[:4] == 'wdg=':
                            if line_check == 'wdg=1':
                                self.transformer.comboBox_tap_location.setCurrentIndex(0)
                            else:
                                self.transformer.comboBox_tap_location.setCurrentIndex(1)
                    self.dlg.transformer_modelling.setTitle(QCoreApplication.translate('dialog',
                                                                                       u'Transformador principal'))
                    self.dlg.transformer_modelling.setEnabled(True)
                elif "NOMODEL" in lines[0]: # substation empty model
                    self.dlg.transformer_modelling.setChecked(False)
                    self.dlg.transformer_modelling.setEnabled(True)
                    self.dlg.lineEdit_voltage.clear()
                    self.dlg.lineEdit_voltage.setText(line[2])
                    self.dlg.lineEdit_name_busbar.clear()
                    self.dlg.lineEdit_name_busbar.setText(line[0])
                    self.substation = 'None'
                elif "AUTO" in lines[0]: #substation as autotransformer
                    self.substation = 'Auto'
                    self.dlg.transformer_modelling.setChecked(False)
                    self.dlg.transformer_modelling.setEnabled(False)
                    self.dlg.lineEdit_voltage.clear()
                    line = lines[3]
                    line_check = line.split(" ")
                    for ij in range(len(line_check)):
                        if 'kVs=[' in line_check[ij]:
                            def_V_prim = float(line_check[ij].replace('kVs=[', ''))*np.sqrt(3) # from kVs line to neutral to kVs line to line
                            self.dlg.lineEdit_voltage.clear()
                            self.dlg.lineEdit_voltage.setText(str(def_V_prim))
                    self.dlg.lineEdit_name_busbar.clear()
                    self.dlg.lineEdit_name_busbar.setText('BUSMV' + file_name + '1') # first bus
                    self.dlg.transformer_modelling.setTitle(QCoreApplication.translate('dialog',
                                                                                       u'Autotransformador en subestación'))
                fp.close()
            else:
                pass  # no substation   
    
    # method to extract first MV line name
    def firstLine(self):
        # line = self.dlg.lineEdit_name_busbar.text()
        mainBus = self.dlg.lineEdit_name_busbar.text() + "."
        prjpath = QgsProject.instance().fileName()
        if not prjpath:
            return
        dir_general, dir_general2 = prjpath.split('/GIS', 1)
        dir_network = dir_general + '/DSS'  # Folder where all network models are stored
        # mydir = dir_network + '\\' + self.dlg.comboBox_circuit_name.currentText()
        os.chdir(dir_network)
        line = ''
        
        for file in glob.glob("*.dss"):
            
            file_check = re.sub('\.dss$', '', file)
            file_check = file_check[-7:]
            if file_check == 'LinesMV':
                print( "file_check = 'LinesMV'")
                print( "mainBus = ", str(mainBus) )

                file_name_path = str(os.getcwd()) + '\\' + '\\' + file
                
                fp = open(file_name_path)
                lines = fp.readlines()
                for i in lines:
                    #print i
                    if mainBus in i:
                        
                        
                        line = i.split(" ")[1].replace('line.', '')
                        
                if line == '':
                    QMessageBox.information(None, QCoreApplication.translate('dialog', u"Error al encontrar primera línea"),
                                            QCoreApplication.translate('dialog',
                                                                     u"Verificar conexión de subestación"))        
        return line

    def tr(self, message):
        # noinspection PyTypeChecker,PyArgumentList,PyCallByClass
        return QCoreApplication.translate('Correr estudio de red usando OpenDSS', message)

    def add_action(self, icon_path, text, callback, enabled_flag=True, add_to_menu=True,
                   add_to_toolbar=True, status_tip=None, whats_this=None, parent=None):
        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)
        if status_tip is not None:
            action.setStatusTip(status_tip)
        if whats_this is not None:
            action.setWhatsThis(whats_this)
        if add_to_toolbar:
            self.toolbar.addAction(action)
        if add_to_menu:
            self.iface.addPluginToMenu(self.menu, action)
        self.actions.append(action)
        return action
        
    #Función que se encarga de imprimir los errores que han ocurrido
    def print_error(self):
        exc_info = sys.exc_info()
        print("\nError: ", exc_info )
        print("*************************  Información detallada del error ********************")
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""
        icon_path = ':/plugins/QGISrunOpenDSS/icon.png'
        self.add_action(icon_path,
                        text=self.tr(u'Correr estudio de red usando OpenDSS'),
                        callback=self.run, 
                        parent=self.iface.mainWindow())
    
    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginMenu(
                self.tr(u'Correr estudio de red usando OpenDSS'), action)
            self.iface.removeToolBarIcon(action)
        del self.toolbar  # remove the toolbar
        
    def show_help(self):
        """Display application help to the user."""

        help_file = 'file:///%s/help/Manual_QGISrunOpenDSS.pdf' % self.plugin_dir
        # For testing path:
        # QMessageBox.information(None, 'Help File', help_file)
        # noinspection PyCallByClass,PyTypeChecker
        QDesktopServices.openUrl(QUrl(help_file))
    
    def copy(self, src, dst):
        import shutil
        if os.path.isdir(dst):
            dst = os.path.join(dst, os.path.basename(src))
        shutil.copyfile(src, dst)
    
    """
    Función que se encarga de instalar una librería en la versión de python de QGIS
    -Parámetros de entrada:
    *librarie_name (string): nombre de la librería a instalar (tal como se le debe pasar a pip)
    
    -Valores retornados:
    *1 en caso de finalizar exitosamente
    *0 en caso de ocurrir algún error
    """
    
    def install_libraries(self, librarie_name):
        try:
            import subprocess
            from pathlib import Path
            #Se obtiene el path de QGIS
            directorio = str( os.path )
            fin_dir = directorio.find("\\apps")
            inic_dir = directorio.find("C:\\")
            path = directorio[ inic_dir : fin_dir - 1]
            #Se obtiene version de Python en QGIS
            info = sys.version_info
            verspy1 = str( info[0] )
            verspy2 = str( info[1] )
            carp_python = str( verspy1 + verspy2 )
            carp_python = str( "Python" + carp_python )
            
            #Se copia los archivos
            dir_origen = path + "\\bin\\"
            name_file_or = "python" + verspy1 + ".dll"
            archivo_origen = str( dir_origen + name_file_or )
            dir_destino = str( path + "\\apps\\" + carp_python)        
            name_dest = dir_destino +  name_file_or
            
            my_file = Path( name_dest )
            print("name_dest = ", my_file.exists() )
            if my_file.exists() == False:            
                #Copia python3.dll
                self.copy( archivo_origen, dir_destino )        
            
            #Copia python37.dll
            name_file_or = "python" + verspy1 + verspy2 + ".dll"
            archivo_origen = dir_origen + name_file_or
            name_dest = dir_destino +  name_file_or
            
            my_file = Path( name_dest )
            print("name_dest = ", my_file.exists() )
            
            if my_file.exists() == False:            
                #Copia python37.dll
                self.copy( archivo_origen, dir_destino )
            
            #Instalación de librerías
            
            #Actualización de pip
            subprocess.call('python.exe -m pip install –upgrade pip', cwd = dir_destino, shell = True )
            
            #Instalación libreria
            sentencia = str( "python.exe -m pip install " + librarie_name )
            subprocess.call(sentencia, cwd = dir_destino, shell = True )
            
            print("Instalación de librería ", librarie_name, " finalizada.")
            return 1
        
        except:
            self.print_error()
            return 0
        
        
        
    """
    Función encargada de graficar las tensiones en una simulación instantánea. Se utiliza la librería de matplotlib.
    Grafica la tensión de las barras en función de la distancia entre estas y la subestación, y las separa en baja tensión,
    media tensión y fase A, B y C.
    -Parámetros de entrada:
    *dataframe_voltages (dataframe): dataframe a ser graficado
    
    -Valores retornados:
    *1 en caso de finalizar exitosamente
    *0 en caso de ocurrir algún error
    
    """
    def GraphSnapshotVoltages(self, dataframe_voltages):
        try:
            import matplotlib.pyplot as plt
            import pandas
            
            nodes = list(dataframe_voltages.index)
            
            mv_nodes_A = []
            mv_nodes_B = []
            mv_nodes_C = []
            
            lv_nodes_A = []
            lv_nodes_B = []
            lv_nodes_C = []
            
            for node in nodes:
                if "mv" in str(node).lower() or "source" in str(node).lower():
                    if str(node)[-1] == '1':
                        mv_nodes_A.append(node)
                    elif str(node)[-1] == '2':
                        mv_nodes_B.append(node)
                    elif str(node)[-1] == '3':
                        mv_nodes_C.append(node)
                
                if "lv" in str(node).lower():
                    if str(node)[-1] == '1':
                        lv_nodes_A.append(node)
                    elif str(node)[-1] == '2':
                        lv_nodes_B.append(node)
                    elif str(node)[-1] == '3':
                        lv_nodes_C.append(node)
                
            V_mv_A = pandas.DataFrame(index = mv_nodes_A, columns = ['voltage', 'distance'])
            V_mv_B = pandas.DataFrame(index = mv_nodes_B, columns = ['voltage', 'distance'])
            V_mv_C = pandas.DataFrame(index = mv_nodes_C, columns = ['voltage', 'distance'])
            
            V_lv_A = pandas.DataFrame(index = lv_nodes_A, columns = ['voltage', 'distance'])
            V_lv_B = pandas.DataFrame(index = lv_nodes_B, columns = ['voltage', 'distance'])
            V_lv_C = pandas.DataFrame(index = lv_nodes_C, columns = ['voltage', 'distance'])
            
            #Asignar valores
            for node in list(dataframe_voltages.index):
                #FASE A MEDIA TENSIÓN
                if str(node)[-1] == '1' and node in mv_nodes_A:
                    V_mv_A.loc[node] = dataframe_voltages.loc[node].values
                #FASE A BAJA TENSIÓN
                elif str(node)[-1] == '1' and node in lv_nodes_A:
                    V_lv_A.loc[node] = dataframe_voltages.loc[node].values
                #FASE B MEDIA TENSIÓN
                elif str(node)[-1] == '2' and node in mv_nodes_B:
                    V_mv_B.loc[node] = dataframe_voltages.loc[node].values
                #FASE B BAJA TENSIÓN
                elif str(node)[-1] == '2' and node in lv_nodes_B:
                    V_lv_B.loc[node] = dataframe_voltages.loc[node].values
                #FASE C MEDIA TENSIÓN
                elif str(node)[-1] == '3' and node in mv_nodes_C:
                    V_mv_C.loc[node] = dataframe_voltages.loc[node].values
                #FASE C BAJA TENSIÓN
                elif str(node)[-1] == '3' and node in lv_nodes_C:
                    V_lv_C.loc[node] = dataframe_voltages.loc[node].values
            
            #PLOT 
            
            leyenda = []
            fign = plt.figure("Tensiones pu con respecto a la distancia de la subestación")
            if V_mv_A.empty == False:	
                plt.plot(V_mv_A['distance'], V_mv_A['voltage'], 'ro', markersize = 3.5)
                leyenda.append("MV fase A")
            if V_mv_B.empty == False:
                plt.plot(V_mv_B['distance'], V_mv_B['voltage'], 'ko', markersize = 3.5)
                leyenda.append("MV fase B")
            if V_mv_C.empty == False:
                plt.plot(V_mv_C['distance'], V_mv_C['voltage'], 'bo', markersize = 3.5)
                leyenda.append("MV fase C")
            if V_lv_A.empty == False:
                plt.plot(V_lv_A['distance'], V_lv_A['voltage'], 'r.', markersize = 2.5)
                leyenda.append("LV fase A (vivo 1)")
            if V_lv_B.empty == False:
                plt.plot(V_lv_B['distance'], V_lv_B['voltage'], 'k.', markersize = 2.5)
                leyenda.append("LV fase B (vivo 2)")
            if V_lv_C.empty == False:
                plt.plot( V_lv_C['distance'], V_lv_C['voltage'], 'b.', markersize = 2.5)
                leyenda.append("LV fase C")
            
            #dia = self.dlg.lineEdit_snapshot_date.text() # read sim hour
            dia = self.snapshotdate
            hora = self.dlg.lineEdit_snapshot_time.text() # read sim hour
            titulo_graph = str( "Tensión pu por bus. Día: " + dia + ". Hora: "+ hora + ".")
            
            tiempo_sim = str( dia + "_" + hora )
            tiempo_sim = tiempo_sim.replace("/", "_")
            tiempo_sim = tiempo_sim.replace(" ", "_")
            tiempo_sim = tiempo_sim.replace(":", "_")
            
            plt.legend(leyenda)
            plt.ylabel("Tensión (pu)")
            plt.xlabel("Distancia (km)")
            
            plt.title( titulo_graph )
            mng = plt.get_current_fig_manager()
            mng.window.showMaximized()
            
            
            output_file = str( self.output_folder + "\\Snapshot\\Snapshot_Voltage_" + tiempo_sim + ".csv" )
            dataframe_voltages.to_csv( output_file)
            output_file = str( self.output_folder + "\\Snapshot\\Snapshot_Voltages_"+ tiempo_sim + '.pdf' )
            plt.savefig(output_file, format='pdf', dpi=6000)
            
            print("Graficación de tensiones exitosa")
            return 1
        except:
            exc_info = sys.exc_info()
            print("\nError: ", exc_info )
            print("*************************  Información detallada del error ********************")
            
            for tb in traceback.format_tb(sys.exc_info()[2]):
                print(tb)
            return 0
        
    """
    Función encargada de graficar las tensiones en una simulación diaria. Se utiliza la librería de matplotlib.
    Grafica la tensión de las barras en función del tiempo, y las separa en baja tensión, media tensión y fase A, B y C.
    -Parámetros de entrada:
    *dataframe_voltages (dataframe): dataframe a ser graficado
    
    -Valores retornados:
    *1 en caso de finalizar exitosamente
    *0 en caso de ocurrir algún error
    
    """
    
    def GraphDailyVoltages(self, DSScircuit, dataframe_voltages):
        try:
            import matplotlib.pyplot as plt
            import pandas
            import datetime
            from .Dataframe_Operations import plot_dataframe
            
            #Primero, modificar los dataframes para que devuelvan los valores promedio y por fase
            nodes = list(DSScircuit.AllNodeNames) #PARA RETORNAR VALORES DE FASE
            buses = list(DSScircuit.AllBusNames) # PARA RETORNAR VALORES PROMEDIO
            
            mv_nodes_A = []
            mv_nodes_B = []
            mv_nodes_C = []
            
            lv_nodes_A = []
            lv_nodes_B = []
            lv_nodes_C = []
            
            mv_buses = []
            lv_buses = []
            
            for node in nodes:
                if "mv" in str(node).lower() or "source" in str(node).lower():
                    if str(node)[-1] == '1':
                        mv_nodes_A.append(node)
                    elif str(node)[-1] == '2':
                        mv_nodes_B.append(node)
                    elif str(node)[-1] == '3':
                        mv_nodes_C.append(node)
                if "lv" in str(node).lower():
                    if str(node)[-1] == '1':
                        lv_nodes_A.append(node)
                    elif str(node)[-1] == '2':
                        lv_nodes_B.append(node)
                    elif str(node)[-1] == '3':
                        lv_nodes_C.append(node)
            
            for bus in buses:
                if "mv" in str(bus).lower() or "source" in str(bus).lower():
                    mv_buses.append(bus)
                elif "lv" in str(bus).lower():
                    lv_buses.append(bus)
            # Ya se tienen las listas 
            
            #Creación de dataframes
            columnas = [ x for x in range(0,96)]
            #Media tensión
            V_mv_prom = pandas.DataFrame(index = mv_buses, columns = columnas)
            V_mv_A = pandas.DataFrame(index = mv_nodes_A, columns = columnas)
            V_mv_B = pandas.DataFrame(index = mv_nodes_B, columns = columnas)
            V_mv_C = pandas.DataFrame(index = mv_nodes_C, columns = columnas)
            
            #Baja tensión
            V_lv_prom = pandas.DataFrame(index = lv_buses, columns = columnas)
            V_lv_A = pandas.DataFrame(index = lv_nodes_A, columns = columnas)
            V_lv_B = pandas.DataFrame(index = lv_nodes_B, columns = columnas)
            V_lv_C = pandas.DataFrame(index = lv_nodes_C, columns = columnas)
            
            #Asignación de valores con respecto a V_buses
            for node in list(dataframe_voltages.index):
                
                if str(node)[-1] == '1':
                    if node in mv_nodes_A: #FASE MEDIA TENSIÓN
                        V_mv_A.loc[node] = dataframe_voltages.loc[node].values
                    else: #FASE BAJA TENSIÓN
                        V_lv_A.loc[node] = dataframe_voltages.loc[node].values
                        
                elif str(node)[-1] == '2':
                    if node in mv_nodes_B: #FASE MEDIA TENSIÓN
                        V_mv_B.loc[node] = dataframe_voltages.loc[node].values
                    else: #FASE BAJA TENSIÓN
                        V_lv_B.loc[node] = dataframe_voltages.loc[node].values
                        
                elif str(node)[-1] == '3': 
                    if node in mv_nodes_C: #FASE MEDIA TENSIÓN
                        V_mv_C.loc[node] = dataframe_voltages.loc[node].values
                    else: #FASE BAJA TENSIÓN
                        V_lv_C.loc[node] = dataframe_voltages.loc[node].values
            
            #PROMEDIOS
            #MT
            if  self.dlg.checkBox_prom_mv.isChecked():
            
                set_mv = set(mv_nodes_A + mv_nodes_B + mv_nodes_C)
            
                for bus in range(len(mv_buses)):
                    temp_list = []
                    
                    for num in range(1,4):
                        if mv_buses[bus]+'.'+str(num) in set_mv:
                            temp_list.append(mv_buses[bus]+'.'+str(num))
                            set_mv.remove(mv_buses[bus]+'.'+str(num))
                 
                    #Asignar valor
                    V_mv_prom.loc[mv_buses[bus]] = pandas.DataFrame(dataframe_voltages.loc[temp_list].mean(axis=0)).T.values
                
            
            #BT
            if  self.dlg.checkBox_prom_lv.isChecked():
            
                set_lv = set(lv_nodes_A + lv_nodes_B + lv_nodes_C)
            
                for bus in range(len(lv_buses)):
                    temp_list = []
                    
                    for num in range(1,4):
                        if lv_buses[bus]+'.'+str(num) in set_lv:
                            temp_list.append(lv_buses[bus]+'.'+str(num))
                            set_lv.remove(lv_buses[bus]+'.'+str(num))
                    
                    #Asignar valor
                    V_lv_prom.loc[lv_buses[bus]] = pandas.DataFrame(dataframe_voltages.loc[temp_list].mean(axis=0)).T.values
            
            #Vector de tiempo
            n_iteraciones = len( dataframe_voltages.columns )
            time = []
            
            min_dia = 60*24
            delta_min = min_dia/n_iteraciones
            
            tiemp = 0
            
            for i in range( n_iteraciones ):                
                tiemp_d = datetime.timedelta( minutes = tiemp )
                sec_temp = str(tiemp_d.seconds)
                hours = int( int(sec_temp) /3600 )
                minutes = int( int(sec_temp)/60 - hours*60 )
                
                if hours <= 10:
                    hours = "0" + str(hours)
                if minutes <= 10:
                    minutes = "0" + str(minutes)
                hours = str( hours )
                minutes = str( minutes )
                tiemp_add = str( hours + ":" + minutes )
                time.append( tiemp_add )
                tiemp = tiemp + delta_min
            
            
            
            #Se reajusta el vector para graficar
            dataframe_voltages.columns = time
            V_mv_A.columns = time
            V_mv_B.columns = time
            V_mv_C.columns = time
            V_mv_prom.columns = time
            
            V_lv_A.columns = time
            V_lv_B.columns = time
            V_lv_C.columns = time
            V_lv_prom.columns = time 
            
            #Se grafica según lo que solicite el usuario
            #fecha = self.dlg.lineEdit_daily_date.text()
            fecha = self.dailydate
            fecha_sim = fecha.replace("/", "_")
            fecha_sim = fecha_sim.replace(" ", "_")
            fecha_sim = fecha_sim.replace(":", "_")
            
            ylabel = "Tensión (pu)"
            xlabel = "Hora"
            
            
            #Media tensión
            if  self.dlg.checkBox_faseA_mv.isChecked() and V_mv_A.empty == False:
                titulo_ventana = "Tensión en barras MV fase A"
                titulo_grafica = str("Tensión pu MV fase A por bus. Fecha: " + fecha)                   
                output_file = str( self.output_folder + "\\Daily\\Daily_Voltages_MVphaseA_"+ fecha_sim + '.pdf' )
                plot_dataframe( V_mv_A, titulo_grafica, titulo_ventana = titulo_ventana, x_label = xlabel, y_label = ylabel, output_file = output_file, save_csv = True )
                
            if  self.dlg.checkBox_faseB_mv.isChecked() and V_mv_B.empty == False:
                titulo_ventana = ("Tensión en barras MV fase B")
                titulo_grafica = str( "Tensión pu MV fase B por bus. Fecha: " + fecha)
                output_file = str( self.output_folder + "\\Daily\\Daily_Voltages_MVphaseB_"+ fecha_sim + '.pdf' )
                #Se grafica el dataframe
                plot_dataframe( V_mv_B, titulo_grafica, titulo_ventana = titulo_ventana, x_label = xlabel, y_label = ylabel, output_file = output_file, save_csv = True )
                
                
            if  self.dlg.checkBox_faseC_mv.isChecked() and V_mv_C.empty == False:
                titulo_ventana = "Tensión en barras MV fase C"
                titulo_grafica = str("Tensión pu MV fase C por bus. Fecha: " + fecha)
                output_file = str( self.output_folder + "\\Daily\\Daily_Voltages_MVphaseC_"+ fecha_sim + '.pdf' )
                #Se grafica el dataframe
                plot_dataframe( V_mv_C, titulo_grafica, titulo_ventana = titulo_ventana, x_label = xlabel, y_label = ylabel, output_file = output_file, save_csv = True )
                
                
            if  self.dlg.checkBox_prom_mv.isChecked() and V_mv_prom.empty == False:
                titulo_ventana = "Tensiones en barras MV"
                titulo_grafica = str("Tensión pu promedio MV por bus. Fecha: " + fecha)      
                output_file = str( self.output_folder + "\\Daily\\Daily_Voltages_MVprom_"+ fecha_sim + '.pdf' )
                #Se grafica el dataframe
                plot_dataframe( V_mv_prom, titulo_grafica, titulo_ventana = titulo_ventana, x_label = xlabel, y_label = ylabel, output_file = output_file, save_csv = True )
            
            #Baja tensión
            if  self.dlg.checkBox_faseA_lv.isChecked() and V_mv_A.empty == False:
                titulo_ventana = "Tensión en barras LV fase A"
                titulo_grafica = str("Tensión pu LV fase A (vivo 1) por bus. Fecha: " + fecha) 
                output_file = str( self.output_folder + "\\Daily\\Daily_Voltages_LVphaseA_"+ fecha_sim + '.pdf' )
                #Se grafica el dataframe
                plot_dataframe( V_lv_A, titulo_grafica, titulo_ventana = titulo_ventana, x_label = xlabel, y_label = ylabel, output_file = output_file, save_csv = True )
                
                
            if  self.dlg.checkBox_faseB_lv.isChecked() and V_mv_B.empty == False:
                titulo_ventana = "Tensión en barras LV fase B (vivo 2)"
                titulo_grafica = str( "Tensión pu LV fase B (vivo 2) por bus. Fecha: " + fecha)
                output_file = str( self.output_folder + "\\Daily\\Daily_Voltages_LVphaseB_"+ fecha_sim + '.pdf' )
                #Se grafica el dataframe
                plot_dataframe( V_lv_B, titulo_grafica, titulo_ventana = titulo_ventana, x_label = xlabel, y_label = ylabel, output_file = output_file, save_csv = True )
                
            if  self.dlg.checkBox_faseC_lv.isChecked() and V_lv_C.empty == False:
                titulo_ventana = "Tensión en barras LV fase 3"
                titulo_grafica =  str("Tensión pu LV fase 3 por bus. Fecha: " + fecha)         
                output_file = str( self.output_folder + "\\Daily\\Daily_Voltages_LVphase3_"+ fecha_sim + '.pdf' )
                #Se grafica el dataframe
                plot_dataframe( V_lv_C, titulo_grafica, titulo_ventana = titulo_ventana, x_label = xlabel, y_label = ylabel, output_file = output_file, save_csv = True )
            
            if  self.dlg.checkBox_prom_lv.isChecked() and V_lv_prom.empty == False: 
                titulo_ventana = "Tensiones promedio en barras LV"
                titulo_grafica =  str("Tensión pu promedio LV por bus. Fecha: " + fecha)          
                output_file = str( self.output_folder + "\\Daily\\Daily_Voltages_LVprom_"+ fecha_sim + '.pdf' )
                #Se grafica el dataframe
                plot_dataframe( V_lv_prom, titulo_grafica, titulo_ventana = titulo_ventana, x_label = xlabel, y_label = ylabel, output_file = output_file, save_csv = True )
                
            
            output_file = str( self.output_folder + "\\Daily\\Daily_VoltagesTot_" + fecha_sim + ".csv" )
            dataframe_voltages.to_csv( output_file )
            print("Graficación de tensiones exitosa")
            return 1
        except PermissionError:
            self.print_error()
            aviso = "Por favor cierre los csv anteriormente generados por este plugin para poder generarlos nuevamente"
            QMessageBox.warning(None, QCoreApplication.translate('dialog', 'Error escritura csv'), aviso)
        
        except:
            self.print_error()

    """
    Función que cuenta el número de trafos sobrecargados en una simulación diaria
    -Parámetros de entrada:
    *dict_trafos (dict): diccionario con las tensiones de los trafos en cada momento del día
    *self.output_folder: directorio de salida de los resultados
    
    -Valores retornados:
    *n_trafos_sobrecargados (int): cantidad de trafos sobrecargados en todo el día
    *list_trafos (list): lista con el DSSName del trafo sobrecargado, la cantidad de veces en el día que estuvo sobrecargado
    y los momentos del día en que estuvo sobrecargado.
    *Si hay algún error retorna 0, 0
    """
    
    def overload_trafos(self, dict_trafos, study = "daily"):
        try:
            
            #Carpeta donde se guardará, según tipo de estudio
            if study == "daily":
                folder = "\Daily"
            elif study == "snapshot":
                folder = "\Snapshot"
            else:
                folder = ""
            
            if not os.path.exists( self.output_folder + folder ):
                os.makedirs( self.output_folder + folder )
            
            #Inicialización de variables
            list_trafos = []
            n_trafos_sobrecargados = 0
            
            #Recorre todas las filas de la capa
            for dssname_trafo, vector_datos in dict_trafos.items():
                sobrecargas_trafo_actual = 0
                time_trafo = []
                
                #Ciclo para recorrer los valores en el vector
                n_elementos_vect = len( vector_datos )
                for i in range(  n_elementos_vect ):
                    value = vector_datos[ i ]
                    value = float( value )
                    if value > 1.0:
                        sobrecargas_trafo_actual += 1
                        time_trafo.append( i )
                
                #Cuando termina de recorrer las columnas asigna (o no) el DSSName, según sea el caso
                if time_trafo != []:
                    n_trafos_sobrecargados += 1
                    dato_vector = [ dssname_trafo,  sobrecargas_trafo_actual, time_trafo ]
                    list_trafos.append( dato_vector )
                
            #Se escribe un txt con los resultados
            line = "Número de trafos con sobrecargas = " + str( n_trafos_sobrecargados ) + "\n"
            print( line )
            line += "Lista de trafos sobrecargados:\n" + str( list_trafos )
            
            
            path_file = self.output_folder + folder + "\OverloadTrafos.txt"
            f = open( path_file,'w' )            
            f.write( line )
            f.close()
            return n_trafos_sobrecargados, list_trafos	
        except:
            self.print_error()
            return 0, 0

    """
    Función que cuenta el número de buses con caídas de tensión en una simulación diaria. Para ello observa si
    en una simulación diaria el bus presenta una caída de tensión 5 o más veces.
    La caída de tensión se establece cuando su tensión está fuera del rango 0.95 - 1.05
    -Parámetros de entrada:
    *dictBuses (dict): diccionario con las tensiones en los buses en cada momento del día
    *self.output_folder: directorio de salida de los resultados
    
    -Valores retornados:
    *n_buses_con_problemas (int): cantidad de buses sobrecargados en todo el día
    *list_buses (list): lista con el DSSName de los buses con caídas de tensión, la cantidad de veces en el día que tuvo esas caídas,
    y los momentos del día en que tuvo dicho problema.
    *Si hay algún error retorna 0, 0
    """
    
    def buses_withVoltageProblems(self, dictBuses, folder = "\Daily" ):
        try:
            
            if not os.path.exists( self.output_folder + folder ):
                os.makedirs( self.output_folder + folder )
            
            #Inicialización de variables
            list_BusesWithProblems = []
            list_DSSNameBusesWithProblems = []
            n_buses_con_problemas = 0
            
            #Recorre todas las filas del diccionario
            for dssname, vector_datos in dictBuses.items():
                cant_problemas_bus_act = 0
                time_bus = []
                dssname_bus = str( dssname )
                
                #Ciclo para recorrer los valores en el vector
                n_elementos_vect = len( vector_datos )
                for i in range(  n_elementos_vect ):
                    value = vector_datos[ i ]
                    try:
                        value = float( value )
                    except: #caso en que no hayan datos para ese bus
                        break
                    if value >= 1.05 or value <= 0.95:
                        cant_problemas_bus_act += 1
                        time_bus.append( i )
                
                #Cuando termina de recorrer las columnas asigna (o no) el DSSName, según sea el caso
                if cant_problemas_bus_act >= 5:
                    n_buses_con_problemas += 1
                    dato_vector = [ dssname_bus,  cant_problemas_bus_act, time_bus ]
                    list_BusesWithProblems.append( dato_vector )
                    list_DSSNameBusesWithProblems.append( dssname_bus )
                    
            #Se escribe un txt con los resultados
            line = "Número de buses con problemas de tensión = " + str( n_buses_con_problemas ) + "\n"
            line += "Lista de buses con problemas de tensión:\n" + str( list_DSSNameBusesWithProblems )
            line += "\nInformación adicional de buses con problemas de tensión:\n" + str( list_BusesWithProblems )
            path_file = self.output_folder + folder + "\BusesWithProblems.txt"
            f = open( path_file,'w' )            
            f.write( line )
            f.close()
            
            return n_buses_con_problemas, list_DSSNameBusesWithProblems, list_BusesWithProblems	
        except:
            self.print_error()
            return 0, 0, 0
    
    """
    Función encargada de obtener la lista de buses con caídas de tensión (llamando a la función buses_withVoltageProblems) y luego 
    leyendo el dss de cargas obtiene el bus asociado a las cargas.
    Si la carga tiene un bus con caída de tensión se verifica que sea una carga residencial y se agrega a la lista de casas
    con problemas.
    -Parámetros de entrada:
    *dictBuses (string): nombre de la capa de buses
    *self.name_lvloads (string): path del dss de cargas
    
    -Datos retornados:
    *n_casas_con_problemas (int): número de casas con problemas de tensión
    *list_DSSNameBuseswithProblems (list): lista de casas (dssnames) con problemas de tensión.
    
    
    """
    def GetVoltageProblems( self, dictBuses, study = "daily"):
        try:
            
            #Carpeta donde se guardará, según tipo de estudio
            if study == "daily":
                folder = "\Daily"
            elif study == "snapshot":
                folder = "\Snapshot"
            else:
                folder = ""
            
            if not os.path.exists( self.output_folder + folder ):
                os.makedirs( self.output_folder + folder )
            
            n_buses_withProblems, list_DSSNameBuseswithProblems, list_withProblemsBuses = self.buses_withVoltageProblems( dictBuses, folder )
            
            #Inicialización de variables
            n_casas_con_problemas = 0
            list_DSSNamewithProblemsHouses = []
            
            with open( self.name_lvloads ) as f:
                lineList = f.readlines()		
            for line in lineList:
                #Búsqueda de nombre del bus asociado a cargas en archivo dss
                in_bus = line.find("bus1=")
                in_bus += len("bus1=")
                fin_bus = line.find(".", in_bus)
                name_bus = line[in_bus:fin_bus]
            
                #Búsqueda del dssname de la carga
                in_dss = line.find("load.")
                in_dss += len("load.")
                fin_dss = line.find(" ", in_dss)
                dssname_load = line[in_dss:fin_dss]
                dssname_load = str( dssname_load )
                
                #Búsqueda de la "clase" de carga
                in_class = line.find("class=")
                in_class += len("class=")
                fin_class = line.find(" ", in_class)
                
                if fin_class == -1: #Caso en que no haya encontrado un espacio después del class buscará un cambio de línea
                    fin_class = line.find("\n",in_class)
                    if fin_class == -1: #Caso en que no haya encontrado un cambio de línea ni un espacio significa que es el elemento final
                        fin_class = in_class + 1
                class_ = line[in_class:fin_class]
                
                #Se determina si la carga es del tipo residencial y que además el bus asociado esté en la lista de buses con problemas
                if class_.lower() == "r" and name_bus in list_DSSNameBuseswithProblems:
                    n_casas_con_problemas += 1
                    list_DSSNamewithProblemsHouses.append( dssname_load )
                    
                    
            #Se escribe un txt con los resultados
            line = "Número de casas con problemas de tensión = " + str( n_casas_con_problemas ) + "\n"
            print( line )
            line += "Lista de casas con problemas de tensión:\n" + str( list_DSSNamewithProblemsHouses )            
            path_file = self.output_folder + folder + "\HousesWithVoltageProblems.txt"
            f = open( path_file,'w' )            
            f.write( line )
            f.close()
                
            return n_casas_con_problemas, list_DSSNamewithProblemsHouses
        except:
            self.print_error()
            return 0, 0
        

    
    ##################################################
    ##################################################
    ############## Función principal #################
    ##################################################
    ##################################################
    
    def run(self):
        plt.close('all')
        #Install the required libraries
        #Instalación comtypes
        try:
            import comtypes.client as cc
        except:
            self.install_libraries("comtypes")
            import comtypes.client as cc
            
        #Instalación pandas
        try:
            import pandas
        except:
            self.install_libraries("pandas")
            import pandas
        try:
            from sklearn.mixture import GaussianMixture
        except:
            self.install_libraries("scikit-learn") #Instalación de librerías extra requeridas
            from sklearn.mixture import GaussianMixture
        
        print("Instalación de librerías finalizada")
     
        #Se debe poner aquí porque utiliza pandas
        from .Dataframe_Operations import plot_dataframe
        
        #Funciones que cargan datos iniciales
        try:
            self.cktNameDefaults()
        except:
            pass
        
        #Try principal para función run
        try:
            def_time = '18:00' # default time
    
            try:
                self.circuit_calling()  # call the networks
            except ValueError:
                exc_info = sys.exc_info()
                print("\nError: ", exc_info )
                print("*************************  Información detallada del error ********************")
                
                for tb in traceback.format_tb(sys.exc_info()[2]):
                    print(tb)
                
                aviso = "El projecto en QGIS debe poseer el formato establecido en el manual\ny ya deben existir los archivos *.dss"
                #QMessageBox.warning(None, QCoreApplication.translate('dialog', 'Alerta Transformadores'), aviso)
                QMessageBox.information(None, QCoreApplication.translate('dialog', "Error al abrir el archivo"), aviso)
            str_time = time.strftime("%Y%m%d%H%M%S",
                                     time.gmtime())  # This string is used to create an unique time string to be adopted when saving results
            prjpath = QgsProject.instance().fileName()
            if not prjpath:
                QMessageBox.information(None, QCoreApplication.translate('dialog', "Error al abrir el archivo"),
                                        QCoreApplication.translate('dialog',
                                                                   u"El projecto en QGIS debe estar abierto") + "\n" + \
                                        QCoreApplication.translate('dialog', u"para poder analizar la red de distribución"))
                return
            dir_general, dir_general2 = prjpath.split('/GIS', 1)
            dir_network = dir_general + '/DSS'  # Folder where all network models are stored
            self.dlg.lineEdit_dirOutput.clear()
            self.dlg.lineEdit_dirOutput.setText(dir_general)
            os.chdir(dir_network)
            networks = [d for d in os.listdir('.') if os.path.isdir(d)]
            
            files_names = list(glob.glob(os.path.join(dir_network, '*.dss')))
            for file in files_names:
                if len(file.split('\\')[1].split('.')[0].split('_')) > 1:
                    if file.split('\\')[1].split('.')[0].split('_')[1] == 'OutputQGIS2OpenDSS':
                        name_file_created = file.split('\\')[1].split('.')[0]
    
            # Combo_Fill_Short_circuit
            try:
                try:
                    ib = self.dlg.lineEdit_short_circuit.itemText(3)
                except IndexError:
                    ib = ''
                if not name_file_created in ib:
                    buses = auxfcns.busSC()
                    self.dlg.lineEdit_short_circuit.clear()
                    self.dlg.lineEdit_short_circuit.addItems(buses)
            except:
                self.dlg.lineEdit_short_circuit.clear()
                buses = [QCoreApplication.translate('SC', "Todas las barras MT"),
                         QCoreApplication.translate('SC', 'Todas las barras BT'),
                         QCoreApplication.translate('SC', 'Todas las barras MT y BT')]
                self.dlg.lineEdit_short_circuit.addItems(buses)
                pass        
    
    
            ########NUEVOS PARAMETROS DE ARMONICOS
            type_parameters = ['GMM', 'GMM']
            # self.dlg.comboBox_parameters.addItems(type_parameters)
    
            type_tx_conns = [QCoreApplication.translate('dialog', 'Estrella'),
                             QCoreApplication.translate('dialog', 'Delta')]
            self.transformer.comboBox_conn_high.clear()
            self.transformer.comboBox_conn_high.addItems(type_tx_conns)
            self.transformer.comboBox_conn_tertiary.clear()
            self.transformer.comboBox_conn_tertiary.addItems(type_tx_conns)
            self.transformer.comboBox_conn_low.clear()
            self.transformer.comboBox_conn_low.addItems(type_tx_conns)
            type_tx_conns = ['wye', 'delta']
            tap_location = [QCoreApplication.translate('dialog', 'Alta'), QCoreApplication.translate('dialog', 'Baja')]
            self.transformer.comboBox_tap_location.clear()
            self.transformer.comboBox_tap_location.addItems(tap_location)
            
            
            #//////////////////////////////////////////////////////////////////////////
            #//////////////////////////////////////////////////////////////////////////
            # show the dialog
            self.dlg.show()
            # Run the dialog event loop
            result = self.dlg.exec_()        
            # See if OK was pressed
            
            if result:
                if self.dlg.transformer_modelling.isChecked():
                    #Esto se usa más adelante para hacer las simulaciones de daily
                    self.substation_sel = True
                else:
                    self.substation_sel = False
                firstLine = self.firstLine()  # first line name        
                if firstLine == '':
                    QMessageBox.information(None, QCoreApplication.translate('dialog', "Error al iniciar simulación"),
                                            QCoreApplication.translate('dialog',
                                                                       u"Debe seleccionarse una barra del alimentador válida"))                       
                    return
                ################ All variables defined in the GUI are read here ################
                # Basic circuit data
                
                tinitial = time.time()
                network = self.dlg.lineEdit_circuit_name.text()  # self.dlg.comboBox_circuit_name.currentIndex()
                self.circuitName = network
                # network = selected_network
                files_names = list(glob.glob(os.path.join(dir_network, '*.dss')))
                for file in files_names:                
                    if len(file.split('\\')[1].split('.')[0].split('_')) > 1:
                        if file.split('\\')[1].split('.')[0].split('_')[1] == 'OutputQGIS2OpenDSS':
                            name_file_created = file.split('\\')[1].split('.')[0]                        
    
                frequency = self.dlg.lineEdit_frequency.text()
                if not frequency:
                    frequency = 60  # def_f_nom
    
                volt_nom = self.dlg.lineEdit_voltage.text()
                if not volt_nom:
                    volt_nom = 230  # def_V_prim
    
                phase_angle = self.dlg.lineEdit_phase.text()
                if not phase_angle:
                    phase_angle = 0  # def_angle_nom
    
                three_phase_SC = self.dlg.lineEdit_3pShortCircuitPower.text()
                if not three_phase_SC:
                    three_phase_SC = 10  # def_3pSCP_nom
    
                single_phase_SC = self.dlg.lineEdit_1pShortCircuitPower.text()
                if not single_phase_SC:
                    single_phase_SC = 10.5  # def_1pSCP_nom
    
                number_simulations = 1
    
                load_curve_circuit = self.dlg.lineEdit_load_curve.text()  # Name of the load curve of the circuit
                
                
                if load_curve_circuit:
                    import io
                    with io.open(load_curve_circuit, 'rt', encoding = "ascii") as workbook:
                        try:
                            reader = csv.reader(workbook)                    
                            next(reader)                    
                        
                        
                            circuit_demand = [[row[3], row[2], row[0], row[1]] for row in
                                              reader]  #day, hour, P (kW), Q (kVAr)
                            workbook.closed
                        except:
                            self.print_error()
                            QMessageBox.information(None, "Informacion errónea", "Curva de demanda del alimentador errónea")
                            
                            return                
    
                tx_main_bus_circuit = self.dlg.lineEdit_name_busbar.text().upper()  # Name of the main busbar
    
                if not load_curve_circuit and not tx_main_bus_circuit:
                    QMessageBox.information(None, "Informacion Requerida", "Seleccione la curva de demanda del alimentador y \n defina el nombre de la barra principal")
                    return
    
                if not load_curve_circuit:
                    QMessageBox.information(None, "Informacion Requerida", "Seleccione la curva de demanda del alimentador")
                    return
    
                if not tx_main_bus_circuit:
                    QMessageBox.information(None, "Informacion Requerida", "Defina el nombre de la barra principal")
                    return
                    
               
    
                ########################################################
                ########################################################
                ############### Creación de capas de EVs ###############
                ########################################################
                ########################################################
                #Elimina archivos previamente creados anteriormente por simulaciones de EVs aleatorias
                study_types = ["random", "consum", "prob"]
                name_dss_evshape = ""
                for name in files_names:
                    nombre = name.lower()                    
                    for study in study_types:
                        name_search = "_ev" + str( study )
                        if name_search in nombre:
                            os.remove( name ) #Elimina el archivo creado antes por estudios aleatorios
                    
                    if "loadslv" in nombre: #averigua el nombre del archivo de cargas
                        name_lvloads = name
                        self.name_lvloads = name_lvloads #parámetro utilizado en búsqueda de caídas de tensión
                    #Determina el nombre del archivo de salida del azul
                    if "outputqgis2opendss" in nombre:
                        name_output_azul = name
                        self.name_output_azul = name_output_azul
                    if "_ev.dss" in nombre:
                        name_ev_tmp = name
                    
                
                #Plantel de buses
                plantel = False
                #Borra líneas creadas anteriormente en el archivo de salida del azul por simulaciones de EVs aleatorias
                with open(name_output_azul, "r+") as f:
                    lines = f.readlines()
                    f.seek(0)
                    for line in lines:
                        linea_ = line.lower()
                        bandera_study_types = False
                        #Se fija si el plantel de buses está agregado a la salida del azul
                        if "plantelesbuses.dss" in linea_:
                            plantel = True
                        if "_ev.dss" in linea_ and linea_[0] != "!": #Busca el nombre del archivo de EVs creado por el azul (si es que existe)
                            name_dss_evshape = name_ev_tmp
                        for study in study_types:
                            name_search = "_ev" + str( study ) 
                            if name_search in linea_:
                                bandera_study_types = True
                        if bandera_study_types == False:
                            f.write(line)
                    f.truncate()
                
                #Inicialización de vector de SOC_t
                vector_soc_t = {}
                for x in range(0,96):
                    vector_soc_t[x] = {}   
                if name_dss_evshape != "":
                    vector_soc_t = CreateList_SOC_t( name_dss_evshape, vector_soc_t )
                
                #Creación de EVs aleatorios
                if self.dlg.EV.isChecked():
                    list_prob = ""
                    #Tipo de estudio aleatorio
                    if self.gui_ev.dlg.checkBox_EV_aleat.isChecked():
                        study_type = "random"
                    #Por consumo
                    elif self.gui_ev.dlg.checkBox_EV_consum.isChecked():
                        study_type = "consumo"
                    
                    #Por probabilidad
                    elif self.gui_ev.dlg.checkBox_EV_prob.isChecked():
                        study_type = "prob"
                        path_csv = self.gui_ev.dlg.lineEdit_ev_file.text()
                        if  path_csv == "" or path_csv == None:
                            aviso = "Debe seleccionar un csv de probabilidades"
                            QMessageBox.warning(None, "Error csv de probabilidades", aviso)
                            return
                        else:
                            try:
                                with open(path_csv, 'r') as f:
                                    reader = csv.reader(f)
                                    list_prob = list(reader)
                            except:
                                self.print_error()
                                aviso = "Debe seleccionar una dirección válida para el csv de probabilidades"
                                QMessageBox.warning(None, "Error csv de probabilidades", aviso)
                                return
                                    
                    
                    percent_evs = self.gui_ev.dlg.lineEdit_ev_prob.text()
                    
                    if not percent_evs:
                        percent_evs = 20 # default to 20% of penetration
                    
                    percent_evs = int( percent_evs )
                    if percent_evs > 100 or percent_evs <= 0:
                        aviso = "Debe indicar un valor entre 0 y 100 para el porcentaje de penetración de VEs"
                        QMessageBox.warning(None, "Error porcentaje de penetración VEs", aviso)
                        return
                        
                    vector_soc_t = CreateEVDss( name_lvloads, study_type, percent_evs, self.circuitName, name_output_azul, vector_soc_t, name_dss_evshape, list_prob = list_prob  )
                    if vector_soc_t == 0:
                        return
                # Transformer data
                line_tx_definition = ''
                if self.dlg.transformer_modelling.isChecked():
                    tx_active = True
                    if self.transformer.checkBox_tertiary.isChecked():
                        tx_tertiary_bus_circuit = 'tertiary_no_use'
                        tx_voltage_high = self.transformer.lineEdit_voltage_high.text().upper()
                        tx_voltage_low = self.transformer.lineEdit_voltage_low.text().upper()
                        tx_voltage_tertiary = self.transformer.lineEdit_voltage_tertiary.text().upper()
                        selected_conn = self.transformer.comboBox_conn_high.currentIndex()
                        tx_conn_high = type_tx_conns[selected_conn]
                        selected_conn = self.transformer.comboBox_conn_low.currentIndex()
                        tx_conn_low = type_tx_conns[selected_conn]
                        selected_conn = self.transformer.comboBox_conn_tertiary.currentIndex()
                        tx_conn_tertiary = type_tx_conns[selected_conn]
                        tx_power_high = self.transformer.lineEdit_power_high.text().upper()
                        tx_power_low = self.transformer.lineEdit_power_low.text().upper()
                        tx_power_tertiary = self.transformer.lineEdit_power_tertiary.text().upper()
                        tx_reactance_hl = self.transformer.lineEdit_x_hl.text().upper()
                        tx_reactance_ht = self.transformer.lineEdit_x_ht.text().upper()
                        tx_reactance_lt = self.transformer.lineEdit_x_lt.text().upper()
                        tx_lossess_noload = self.transformer.lineEdit_losses_no_load.text().upper()
                        tx_lossess_inload = self.transformer.lineEdit_losses_in_load.text().upper()
                        ltc_location = self.transformer.comboBox_tap_location.currentIndex()
                        tx_ltc_location = tap_location[ltc_location]
                        if tx_ltc_location == 'Alta':
                            tx_ltc_location_definition = '1'
                        else:
                            tx_ltc_location_definition = '2'
                        tx_number_taps = self.transformer.lineEdit_no_taps.text().upper()
                        tx_current_tap_position = self.transformer.lineEdit_tap_position.text().upper()
                        tx_max_tap_position = self.transformer.lineEdit_max_tap.text().upper()
                        tx_min_tap_position = self.transformer.lineEdit_min_tap.text().upper()
                        # ------------------------------------
                        setNew = 'Edit'
                        if self.substation == 'None':
                            setNew = 'New'
                        line_tx_definition = setNew + ' transformer.HVMV phases=3 windings=3 buses=[sourcebus,' + tx_main_bus_circuit + ',' + tx_tertiary_bus_circuit + '] conns=[' \
                                             + tx_conn_high + ',' + tx_conn_low + ',' + tx_conn_tertiary + '] kvs=[' + tx_voltage_high + ',' + tx_voltage_low + ',' + tx_voltage_tertiary \
                                             + '] kvas=[' + tx_power_high + ',' + tx_power_low + ',' + tx_power_tertiary + '] xhl=' + tx_reactance_hl + ' xht=' + tx_reactance_ht + ' xlt=' + tx_reactance_lt \
                                             + ' %loadloss=' + tx_lossess_inload + ' %noloadloss=' + tx_lossess_noload + ' wdg=' + tx_ltc_location_definition + ' numtaps=' + tx_number_taps \
                                             + ' tap=' + tx_current_tap_position + ' maxtap=' + tx_max_tap_position + ' mintap=' + tx_min_tap_position
                        # ------------------------------------    
                    else:
                        tx_voltage_high = self.transformer.lineEdit_voltage_high.text().upper()
                        tx_voltage_low = self.transformer.lineEdit_voltage_low.text().upper()
                        selected_conn = self.transformer.comboBox_conn_high.currentIndex()
                        tx_conn_high = type_tx_conns[selected_conn]
                        selected_conn = self.transformer.comboBox_conn_low.currentIndex()
                        tx_conn_low = type_tx_conns[selected_conn]
                        tx_power_high = self.transformer.lineEdit_power_high.text().upper()
                        tx_power_low = self.transformer.lineEdit_power_low.text().upper()
                        tx_reactance_hl = self.transformer.lineEdit_x_hl.text().upper()
                        tx_lossess_noload = self.transformer.lineEdit_losses_no_load.text().upper()
                        tx_lossess_inload = self.transformer.lineEdit_losses_in_load.text().upper()
                        selected_loc = self.transformer.comboBox_tap_location.currentIndex()
                        tx_tap_location = tap_location[selected_loc]
                        if tx_tap_location == 'Alta':
                            tx_ltc_location_definition = '1'
                        else:
                            tx_ltc_location_definition = '2'
                        tx_number_taps = self.transformer.lineEdit_no_taps.text().upper()
                        tx_current_tap_position = self.transformer.lineEdit_tap_position.text().upper()
                        tx_max_tap_position = self.transformer.lineEdit_max_tap.text().upper()
                        tx_min_tap_position = self.transformer.lineEdit_min_tap.text().upper()
                        # ------------------------------------
                        setNew = 'Edit'
                        if self.substation == 'None': # if no substation was modelled on QGIS2OpenDSS but is desired to add one
                            setNew = 'New'
                        line_tx_definition = setNew + ' transformer.HVMV phases=3 windings=2 buses=[sourcebus,' + tx_main_bus_circuit + '] conns=[' \
                                             + tx_conn_high + ',' + tx_conn_low + '] kvs=[' + tx_voltage_high + ',' + tx_voltage_low + '] kvas=[' + tx_power_high \
                                             + ',' + tx_power_low + '] xhl=' + tx_reactance_hl + ' %loadloss=' + tx_lossess_inload + ' %noloadloss=' + tx_lossess_noload \
                                             + ' wdg=' + tx_ltc_location_definition + ' numtaps=' + tx_number_taps + ' tap=' + tx_current_tap_position + ' maxtap=' + tx_max_tap_position \
                                             + ' mintap=' + tx_min_tap_position
                        # ------------------------------------    
                    
                    #
                    
                    
                elif self.substation == 'Auto': # If transformer substation is an autotransformer
                    # QMessageBox.information(None, "Debug", self.substation)
                    line_tx_definition = ""
                    tx_active = True
                    tx_conn_low = 'wye'
                    pass
                else: # self.substation == 'None':  # If transformer is not modelled, then the sourcebus is connected straight to a MV line
                    tx_active = False
                    line_tx_definition = ''
                    new_lines_file = open(dir_network + '/' + name_file_created.split('_')[0] + '_LinesMV_mod.dss', 'w')
                    with open(dir_network + '/' + name_file_created.split('_')[0] + '_LinesMV.dss', 'r') as temp:
                        for aux in temp:
                            linetemp = re.sub(r"\b%s\b" % tx_main_bus_circuit + '.1.2.3', 'Sourcebus', aux)
                            new_lines_file.write(linetemp)
                    new_lines_file.close()
    
                # add PV systems            
                if self.dlg.PV.isChecked():
                    PVS = list()
                    # study type
                    pv_study_type = self.gui_gd.dlg.comboBoxPVTipoEstudio.currentText()
                    if QCoreApplication.translate('dialog', "Mayor consumo") in pv_study_type:
                        pv_study_type = "Maximum"
                    if QCoreApplication.translate('dialog', "Por probabilidad") in pv_study_type:
                        pv_study_type = "Bernoulli"
                    if QCoreApplication.translate('dialog', "Aleatorio") in pv_study_type:
                        pv_study_type = "Random"
    
                    # Residential
                    if self.gui_gd.dlg.checkBox_PV_res.isChecked() and self.dlg.PV.isChecked():                    
                        total_capacity = self.gui_gd.dlg.lineEdit_pv_capacity.text()
                        pv_information = self.gui_gd.dlg.lineEdit_pv_info.text()
    
                        if not total_capacity:
                            total_capacity = str(200) # default to 200 kW
                        if not pv_information:
                            QMessageBox.information(None, "Informacion Requerida", "Seleccione el archivo de información de generación distribuida para clientes residenciales")
                            return
    
                    # General
                    if self.gui_gd.dlg.checkBox_PV_gen.isChecked() and self.dlg.PV.isChecked():
                        total_capacity_gen = self.gui_gd.dlg.lineEdit_pv_capacity_gen.text()
                        pv_information_gen = self.gui_gd.dlg.lineEdit_pv_info_gen.text()
    
                        if not total_capacity_gen:
                            total_capacity_gen = str(200)
                        if not pv_information_gen:
                            QMessageBox.information(None, "Informacion Requerida", "Seleccione el archivo de información de generación distribuida para clientes comercial e industrial")
                            return
    
                    # TMT
                    if self.gui_gd.dlg.checkBox_PV_TMT.isChecked() and self.dlg.PV.isChecked():
                        total_capacity_TMT = self.gui_gd.dlg.lineEdit_pv_capacity_TMT.text()
                        pv_information_TMT = self.gui_gd.dlg.lineEdit_pv_info_TMT.text()
    
                        if not total_capacity_TMT:
                            total_capacity_TMT = str(200)
                        if not pv_information_TMT:
                            QMessageBox.information(None, "Informacion Requerida", "Seleccione el archivo de información de generación distribuida para clientes TMT")
                            return
    
                output_folder = self.dlg.lineEdit_dirOutput.text()  # Output folder
                if not output_folder:  # Default path is otherwise used
                    output_folder = "C:/QGISrunOpenDSS_Results"  # def_out_folder
                    if not os.path.exists(output_folder):
                        os.mkdir(output_folder)
                self.output_folder = output_folder
    
                # loadshape files path
                dir_profiles = os.path.join(dir_network,
                                            'profiles')
                # update loadshapes files path on *.dss
                
                load_profiles = auxfcns.ReadLoadProfiles(self, dir_profiles,
                                                         dir_network,
                                                         name_file_created)
                # exit if an error occurred
                if load_profiles[0] == "ERROR":
                    print("Error load_profiles")
                    return
                
                [DSSobj, DSSstart, DSStext, DSScircuit,
                 DSSprogress] = auxfcns.SetUpCOMInterface()  # Inicializacion de la interfaz COM de OpenDSS
    
                # Master.dss file creation
                created_files = open(dir_network + '/' + name_file_created + '.dss', 'r')
                created_files = created_files.read()
                if tx_active == False:
                    created_files = created_files.replace(name_file_created.split('_')[0] + '_LinesMV',
                                                          name_file_created.split('_')[0] + '_LinesMV_mod')
                    created_files = created_files.replace('redirect ' + name_file_created.split('_')[0] + '_Substation.dss', '')
                
                elif (tx_active == True) and (self.dlg.powerflow_snapshot.isChecked() or self.dlg.powerflow_daily.isChecked()):
                    
                    with open(dir_network + '/' + name_file_created.split('_')[0] + '_LinesMV.dss', 'r+') as f:   #save file's original settings
                        original_lines = f.readlines()
                        f.close
                    
                    if self.dlg.powerflow_snapshot.isChecked():
                    
                        with open(dir_network + '/' + name_file_created.split('_')[0] + '_LinesMV_snap.dss', 'w') as the_file:
                            lines = copy.deepcopy(original_lines)
                            lines.insert(0,'new line.MV3P'+name_file_created.split('_')[0]+'00 bus1=BUSMV'+name_file_created.split('_')[0]+'1.1.2.3 bus2=AFTERMETER r1=0.00001 x1=0.00001 length=0.0001 units=m \n')
                            the_file.seek(0)
                                
                            for line_num in range(1,len(lines)):
                                whole_line = lines[line_num].split(' ')
                                bus1 = lines[line_num].split(' ')[2]
                                bus2 = lines[line_num].split(' ')[3]
                                
                                if bus1.split('=')[1].split('.')[0].split(name_file_created.split('_')[0])[1] == '1':
                                    lines[line_num] = ''
                                    for elem in range(len(whole_line)):
                                        if elem == 2: #campo del bus 1
                                            lines[line_num] = lines[line_num] + 'bus1=AFTERMETER '
                                        elif elem == len(whole_line)-1:
                                            lines[line_num] = lines[line_num] +  whole_line[elem]
                                        else:
                                            lines[line_num] = lines[line_num] + whole_line[elem] + ' '
                                
                                elif bus2.split('=')[1].split('.')[0].split(name_file_created.split('_')[0])[1] == '1':
                                    lines[line_num] = ''
                                    for elem in range(len(whole_line)):
                                        if elem == 3: #campo del bus 1
                                            lines[line_num] = lines[line_num] + 'bus2=AFTERMETER '
                                        elif elem == len(whole_line)-1:
                                            lines[line_num] = lines[line_num] +  whole_line[elem]
                                        else:
                                            lines[line_num] = lines[line_num] + whole_line[elem] + ' '
                           
                            the_file.writelines(lines) 
                            
                            the_file.close
                            
                        created_files = created_files.replace(name_file_created.split('_')[0] + '_LinesMV',
                                                          name_file_created.split('_')[0] + '_LinesMV_snap')
                    
                    elif self.dlg.powerflow_daily.isChecked():
                    
                        with open(dir_network + '/' + name_file_created.split('_')[0] + '_LinesMV_daily.dss', 'w') as the_file:
                            lines = copy.deepcopy(original_lines)
                            lines.insert(0,'new line.MV3P'+name_file_created.split('_')[0]+'00 bus1=BUSMV'+name_file_created.split('_')[0]+'1.1.2.3 bus2=AFTERMETER r1=0.00001 x1=0.00001 length=0.0001 units=m \n')
                            the_file.seek(0)
                                
                            for line_num in range(1,len(lines)):
                                whole_line = lines[line_num].split(' ')
                                bus1 = lines[line_num].split(' ')[2]
                                bus2 = lines[line_num].split(' ')[3]
                                
                                if bus1.split('=')[1].split('.')[0].split(name_file_created.split('_')[0])[1] == '1':
                                    lines[line_num] = ''
                                    for elem in range(len(whole_line)):
                                        if elem == 2: #campo del bus 1
                                            lines[line_num] = lines[line_num] + 'bus1=AFTERMETER '
                                        elif elem == len(whole_line)-1:
                                            lines[line_num] = lines[line_num] +  whole_line[elem]
                                        else:
                                            lines[line_num] = lines[line_num] + whole_line[elem] + ' '
                                
                                elif bus2.split('=')[1].split('.')[0].split(name_file_created.split('_')[0])[1] == '1':
                                    lines[line_num] = ''
                                    for elem in range(len(whole_line)):
                                        if elem == 3: #campo del bus 1
                                            lines[line_num] = lines[line_num] + 'bus2=AFTERMETER '
                                        elif elem == len(whole_line)-1:
                                            lines[line_num] = lines[line_num] +  whole_line[elem]
                                        else:
                                            lines[line_num] = lines[line_num] + whole_line[elem] + ' '
                           
                            the_file.writelines(lines) 
                            
                            the_file.close
                            
                        created_files = created_files.replace(name_file_created.split('_')[0] + '_LinesMV',
                                                          name_file_created.split('_')[0] + '_LinesMV_daily')
                
                
                
                file = open(dir_network + '/Master.dss', 'w')
                file.write('set defaultbasefrequency=' + frequency + '\n')
                if self.dlg.SCMVA.isChecked():  # SC MVAs
                    file.write('Edit Vsource.Source BasekV=' + volt_nom + ' pu=1.00 angle=' + phase_angle +
                               ' frequency=' + frequency + ' phases=3 MVAsc3='
                               + three_phase_SC + ' MVAsc1=' + single_phase_SC + '\n')
                if self.dlg.SCkA.isChecked():  # SC kAs
                    file.write('Edit Vsource.Source BasekV=' + volt_nom + ' pu=1.00 angle=' + phase_angle +
                               ' frequency=' + frequency + ' phases=3 Isc3='
                               + str(float(three_phase_SC) * 1000) + ' Isc1=' + str(float(single_phase_SC) * 1000) + '\n')
                tx_modelling = False
                if self.dlg.transformer_modelling.isChecked():  # see if transformer_modelling checkbox is active
                    tx_modelling = True
    
                file.write(created_files + '\n')
                file.close()          
                
    
                time_common_for_all = time.time() - tinitial
                #Inicialización de variables
                P_to_be_matched = 0
                Q_to_be_matched = 0
    
                ######################################################################
                ###########################    SNAPSHOT     ##########################
                ######################################################################
                # Routine to perform snapshot power flows
                if self.dlg.powerflow_snapshot.isChecked():
    
                    DSSprogress.Show()
                    DSSprogress.Caption = QCoreApplication.translate('progress', u'Snapshot')
                    DSSprogress.PctProgress = 0              
    
                    tinitial_snapshot = time.time() # time counter
                    snapshotdate = self.dlg.lineEdit_snapshot_date.text().upper()
                    snapshotdate = correct_date( load_curve_circuit, snapshotdate )
                    if not snapshotdate:
                        snapshotdate = auxfcns.selection_representative_day(load_curve_circuit, 'weekday')
                    self.snapshotdate = snapshotdate
    
                    snapshottime = self.dlg.lineEdit_snapshot_time.text() # read sim hour
                    snapshottime = correct_hour( load_curve_circuit, snapshottime )
    
                    if not snapshottime:
                        snapshottime = def_time  # Default is 6pm
    
                    h, m = snapshottime.split(':')
                    if m is not '00' or '15' or '30' or '45':  # round sim minutes
                        if int(m) <= 7:
                            m = '00'
                        elif int(m) <= 22:
                            m = '15'
                        elif int(m) <= 37:
                            m = '30'
                        elif int(m) <= 52:
                            m = '45'
                        else:
                            m = '00'
                            h = str(int(h) + 1)
                            if int(h) == 24:  # last round on 23:45
                                h = '23'
                                m = '45'
                    snapshottime = h + ':' + m
                    
                    day_ = snapshotdate.replace('/', '')
                    day_ = day_.replace('-', '')               
    
                    daily_strtime = str(day_ + snapshottime.replace(':', ''))                
                    hora_sec = snapshottime.split(':')
                    for ij in range(len(circuit_demand)):
                        temp_a = circuit_demand[ij][0]  # day
                        temp_b = circuit_demand[ij][1]  # hour                    
                        if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:                        
                            P_to_be_matched = circuit_demand[ij][2]  # Active power
                            Q_to_be_matched = circuit_demand[ij][3]  # Reactive power
                    
                    print( "Progreso = 17 " )
                    DSSprogress.PctProgress = 17
                    
                    
                    #changes on mvlines
                    if self.substation_sel == True:
                        
                        circuit_name = name_file_created.split('_')[0]
                        firstLine = 'MV3P'+circuit_name+'00'
                        
                    ######## FIRST SIMULATION: SNAPSHOT TO GET LN VOLTAGE BASES

                    DSStext.Command = 'clear'  # clean previous circuits
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
                    DSStext.Command = 'Set mode=snapshot'  # Type of Simulation
                    # DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    # DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                    DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time
                    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
                    DSStext.Command = 'batchedit load..* enabled = no' # No load simulation
                    DSStext.Command = 'batchedit storage..* enabled = no' # No load simulation
                    DSStext.Command = 'batchedit PVSystem..* enabled = no' # No load simulation
                    DSStext.Command = 'batchedit Generator..* enabled = no' # No load simulation
                    
                    
                    DSScircuit.Solution.Solve()  # solve the circuit
                    
                    all_buses = list(DSScircuit.AllBusVmag)
                    index = list(DSScircuit.AllNodeNames)
                    tam_buses = len(all_buses)
                    #borrar datos innecesarios
                    del_bus = []
                        
                    for i in range(len(index)):
                        if 'AFTERMETER' in str(index[i]).upper():
                            del_bus.append(i)
                    
                    num_del_buses = len(del_bus)
                    
                    for j in range(num_del_buses-1,-1,-1):
                        index.remove(index[del_bus[j]])
                        all_buses.remove(all_buses[del_bus[j]])
                    
                    Base_V = pandas.DataFrame(0, index = index, columns=['base'])
                    base_vals = [120, 138, 208, 240, 254, 277, 416, 440, 480, 7620, 7967, 13200, 13800, 14380, 19920, 24900, 34500, 79670, 132790]
                    for i in range(len(all_buses)):
                        busV = all_buses[i]
                        Base_V['base'].iloc[i] = base_vals[[abs(val-busV) for val in base_vals].index(min([abs(val-busV) for val in base_vals]))] #selecciona el index que de la mínima resta
                    
                    Base_V.to_csv('bases.csv')
                    #End simulation to get bases
                    
                    DSStext.Command = 'clear'
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
                    DSStext.Command = 'Set mode=daily'  # Type of Simulation
                    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time                
                    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the transformer secondary side to monitor P and Q
    
                    if tx_modelling and self.substation != 'Auto': # check tx_modelling and autotransformer
                        DSStext.Command = line_tx_definition
                    
                    for MC_iteration in range(int(number_simulations)):  # Run Monte Carlo Simulations, defaults to 1
                        # Run the daily power flow for a particular moment
                        gen_p = 0
                        gen_q = 0
                        DSScircuit.Solution.Solve()  # Initialization solution                        
                        DSScircuit.setActiveElement('line.' + firstLine)
                        temp_powers = DSScircuit.ActiveElement.Powers                        

                        errorP = 0.003  # Maximum desired correction error for active power
                        errorQ = 0.01  # Maximum desired correction error for reactive power
                        max_it_correction = 10  # Maximum number of allowed iterations
                        study = 'snapshot'  # Study type for PQ_corrector
                        # distributed generators powers
                        gen_powers = np.zeros(1)
                        gen_rpowers = np.zeros(1)
                        GenNames = DSScircuit.Generators.AllNames
                        PVNames = DSScircuit.PVSystems.AllNames
                        if GenNames[0] != 'NONE':
                            for i in GenNames: # extract power from generators
                                DSScircuit.setActiveElement('generator.' + i)
                                p = DSScircuit.ActiveElement.Powers
                                for w in range(0, len(p), 2):
                                    gen_p += -p[w] # P
                                    gen_q += -p[w + 1] # Q
                            gen_powers[0] += gen_p
                            gen_rpowers[0] += gen_q
                        if PVNames[0] != 'NONE':
                            for i in PVNames: # extract power from PVSystems
                                DSScircuit.setActiveElement('PVSystem.' + i)
                                p = DSScircuit.ActiveElement.Powers
                                for w in range(0, len(p), 2):
                                    gen_p += -p[w] # P
                                    gen_q += -p[w + 1] # Q
                            gen_powers[0] += gen_p
                            gen_rpowers[0] += gen_q

                        #DSStext.Command = 'batchedit storage..* enabled = no' # No storage simulation
                        # load allocation algorithm
                        [DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_sim,
                         kVAr_sim] = auxfcns.PQ_corrector(DSSprogress, DSScircuit, DSStext, errorP, errorQ,
                                                          max_it_correction,
                                                          P_to_be_matched, Q_to_be_matched, hora_sec, study,
                                                          dir_network, tx_modelling, 1, firstLine, self.substation,
                                                          line_tx_definition, gen_powers, gen_rpowers)
                        if DSScircuit == None:
                            self.iface.messageBar().pushCritical("QGIS2RunOpenDSS", "Sucedió un error grave y no fue posible completar la operación")
                            return
                            
                        print( "Progreso = 50" )           
                        
                        DSSprogress.PctProgress = 50
                        #### Post load allocation simmulation
                        DSStext.Command = 'clear'
                        DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                        
                        DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
                        if plantel == True: #se agregan los buses si hay un archivo del plantel de buses
                            DSStext.Command = 'redirect ' + str(name_file_created.split('_')[0]) +  '_StorageBuses.dss'
                        DSStext.Command = 'New Energymeter.Sub line.' + firstLine #Energymeter in the substation. It allows to perform the voltage profile plot
                        DSStext.Command = 'Set mode=daily'  # Type of Simulation
                        DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                        DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                        DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                        
                        if tx_modelling and self.substation != 'Auto':
                            DSStext.Command = line_tx_definition
                        if self.dlg.PV.isChecked():
                            if self.gui_gd.dlg.checkBox_PV_res.isChecked():                            
                                PVS += auxfcns.PV_allocation(total_capacity, pv_information, name_file_created, "R",
                                                             pv_study_type)
    
                            if self.gui_gd.dlg.checkBox_PV_gen.isChecked():
                                PVS += auxfcns.PV_allocation(total_capacity_gen, pv_information_gen, name_file_created, "C",
                                                             pv_study_type)
    
                            if self.gui_gd.dlg.checkBox_PV_TMT.isChecked():
                                PVS += auxfcns.PV_allocation(total_capacity_TMT, pv_information_TMT, name_file_created,
                                                             "TMT",
                                                             pv_study_type)
    
                            for pv in range(len(PVS)):
                                if PVS[pv] == "Error":
                                    print("Error")
                                    return
                                DSStext.Command = PVS[pv]
                                #print("PVS[pv] = ", PVS[pv])
                                #DSStext.Command = 'solve'
    
                        DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the transformer secondary side to monitor P and Q
                        DSStext.Command = 'calcvoltagebases'
    
                        nodeVoltages_ph1 = []
                        nodeVoltages_ph2 = []
                        nodeVoltages_ph3 = []
                        currentList = []
    
                        try:
                            DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
                        except TypeError: #error debido a que no se encontró la fecha
                            QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), QCoreApplication.translate('dialog',
                                                                           u'Debe introducir una fecha y hora que coincida con alguna fecha en el .csv de la curva del alimentador.\nIgualmente puede dejar en blanco y se seleccionará el día más representativo'))
                            return
                            
                        DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
                        DSScircuit.Solution.Solve()
    
                        DSSprogress.PctProgress = 51
                        ###############################################################
                        all_buses = list(DSScircuit.AllBusVmag)
                        index = list(DSScircuit.AllNodeNames)
                        distance = list(DSScircuit.AllNodeDistances)
                        #eliminar los buses creados por el energy meter
                        del_bus = []
                        
                        for i in range(len(index)):
                            if 'AFTERMETER' in str(index[i]).upper():
                                del_bus.append(i)
                        
                        num_del_buses = len(del_bus)
                        
                        for j in range(num_del_buses-1,-1,-1):
                            index.remove(index[del_bus[j]])
                            all_buses.remove(all_buses[del_bus[j]])
                            distance.remove(distance[del_bus[j]])
                            
                        buses_pu = []
                        
                        if len(all_buses) != len(Base_V['base']):
                            QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), QCoreApplication.translate('dialog',
                                                                                                   u'Verifique que todas las cargas estén conectadas correctamente a las líneas.\nNota: confirme que las cargas estén conectadas a las fases adecuadas. '))
                            DSSprogress.Close()
                            return
                        
                        for i in range(len(all_buses)):
                            valor_pu = all_buses[i]/Base_V['base'].iloc[i]
                            buses_pu.append(valor_pu)

                        V_buses = pandas.DataFrame(buses_pu, index = index, columns = ['voltage'] )
                        V_buses['distance'] = distance
                        V_buses.to_csv('voltages.csv')
                        ###############################################################
                        # kV base for all buses, must go after solution
                        DSStext.Command = 'redirect ' + name_file_created.split('_')[0] + '_LV_KVBaseLN.dss'
                        DSStext.Command = 'redirect ' + name_file_created.split('_')[0] + '_MV_BaseKV_LN.dss'
                        # Extract Voltage Values
                        
                        # Extract V from all nodes
                        nodeNames_ph1 = DSScircuit.AllNodeNamesByPhase(1)
                        nodeNames_ph2 = DSScircuit.AllNodeNamesByPhase(2)
                        nodeNames_ph3 = DSScircuit.AllNodeNamesByPhase(3)
                        # all line names
                        lineNames = DSScircuit.Lines.allnames
                        normalAmpsDic = auxfcns.normalAmps(DSScircuit, lineNames)  # ampacities dictionary
                        nodeVoltages_ph1.append([DSScircuit.AllNodeVmagPUByPhase(1)])  # Buses pu voltages ph1
                        nodeVoltages_ph2.append([DSScircuit.AllNodeVmagPUByPhase(2)])  # Buses pu voltages ph2
                        nodeVoltages_ph3.append([DSScircuit.AllNodeVmagPUByPhase(3)])  # Buses pu voltages ph3
                        
                        trafosDict, trafosNames = auxfcns.ReadTrafosLoad(self, DSScircuit, DSStext, name_file_created)
                        busesDict = auxfcns.ReadBusVolts(self, nodeVoltages_ph1, nodeVoltages_ph2, nodeVoltages_ph3, nodeNames_ph1, nodeNames_ph2, nodeNames_ph3)
                        
                        #Se determina el número de sobrecargas de los trafos y problemas de tensión en casas
                        n_sobrecargas, lista_sobrecargas = self.overload_trafos( trafosDict, study = "snapshot" )
                        
                        # save results on shapes
                        if self.dlg.checkBox_capas.isChecked():
                            currentList.append(
                                auxfcns.lineCurrents(self, DSScircuit, lineNames, normalAmpsDic, study))  # Normalized currents
                            # voltages shapes update
                            auxfcns.WriteBusVolts(self, busesDict, name_file_created, study)
    
                            DSSprogress.PctProgress = 68
                            # currents shape update
                            list_layersTrafos, list_layersLinesMV, list_layersLinesLV = auxfcns.get_layersnames( self.name_output_azul )
                            auxfcns.lineCurrentsResults(self, lineNames, currentList, study, list_layersLinesMV, list_layersLinesLV)
                            # transformer overload shape update
                            auxfcns.WriteTrafosLoad(self, trafosDict, trafosNames, study, list_layersTrafos)
                        DSSprogress.PctProgress = 75
    
                        final_output_folder = output_folder + '/Snapshot/'# + 'MC_' + str(MC_iteration + 1) + '/'
                        if not os.path.exists(final_output_folder):
                            os.makedirs(final_output_folder)
    
                        # openDSS results path
                        DSStext.Command = 'set datapath=' + final_output_folder
    
                        DSStext.Command = 'monitor.HVMV_PQ_vs_Time.action=take'
                        DSStext.Command = 'export monitors HVMV_PQ_vs_Time'
    
                        # circuit losses
                        if self.dlg.checkBoxLosses.isChecked():
                            DSStext.Command = 'export losses'
                            DSStext.Command = 'show losses'
    
                        # voltage unbalance
                        if self.dlg.checkBoxUnbalance.isChecked():
                            buses = DSScircuit.ALLBusNames
                            unbalance = {}
                            orderedUnbalance = OrderedDict()
                            orderedUnbalance = auxfcns.unbalance(self, DSScircuit, buses, unbalance, orderedUnbalance)
    
                            x, y = list(zip(*list(orderedUnbalance.items())))
                            x_labels = []  # create an empty list to store the labels
                            if len(x) <= 7:
                                steps = 1
                            else:
                                steps = len(x) / 7
                            ticksIndex = np.arange(start = 0, stop = len(x), step = steps, dtype = int)
                            for key in ticksIndex:
                                x_labels.append(x[key])
                            fig, ax = plt.subplots( num = "Desbalance de tensión")
                            plt.plot(y, label=QCoreApplication.translate('graphs', u'Desbalance'))
                            plt.title(QCoreApplication.translate('graphs', u'Desbalance de tensión'))
                            DSScircuit.Monitors.ResetAll()
                            # plt.xlim([0, 96])
                            ax.set_xticks(ticksIndex)
                            ax.set_xticklabels(x_labels, rotation=90)
                            plt.xlabel(QCoreApplication.translate('graphs', u'Barras'))
                            plt.ylabel(
                                QCoreApplication.translate('graphs', u'Desbalance (%)'))
                            plt.legend(fontsize=8, loc=2)
                            plt.tight_layout()
                            mng = plt.get_current_fig_manager()
                            mng.window.showMaximized() 
                            #fig.show()
                            fig.savefig(final_output_folder + '/' + network + '_Unbalance_' + str_time + '.pdf',
                                        format='pdf',
                                        dpi=6000)
                            csvFile = final_output_folder + '/' + network + '_Unbalance_' + str_time + '.csv'
                            with open(csvFile, 'w') as f:
                                w = csv.writer(f, lineterminator='\n')
                                w.writerow(["Bus", "Unbalance (%)"])
                                for row in range(len(x)):
                                    # line = str(x[row]) + "," + str(y[row])
                                    w.writerow([str(x[row]), str(y[row])])
    
                        # print DSScircuit.Losses
                        DSSprogress.PctProgress = 90
                        if not self.dlg.checkBox_capas.isChecked():
                            DSStext.Command = 'visualize What={voltages} element=line.' + firstLine
                            DSStext.Command = 'visualize What={currents} element=line.' + firstLine                    
                        DSScircuit.Monitors.ResetAll()
                    
                    # voltage profile plot
                    self.GraphSnapshotVoltages( V_buses )
                    DSSprogress.PctProgress = 100
                    DSSprogress.Close()
                    
                    
    
                    tfinal_snapshot = time.time() - tinitial_snapshot # snapshot total time
                    if not self.dlg.powerflow_daily.isChecked() and not self.dlg.short_circuit.isChecked() \
                            and not self.dlg.harmonics.isChecked() and not self.dlg.powerflow_yearly.isChecked():
                        QMessageBox.information(None,
                                                QCoreApplication.translate('dialog', "Simulacion Instantanea Terminada"),
                                                QCoreApplication.translate('dialog', "Tiempo de simulacion: ") + str(
                                                    tfinal_snapshot + time_common_for_all) + " s" + "\n" + \
                                                QCoreApplication.translate('dialog',
                                                                           "Los archivos han sido guardados en: ") + output_folder)
                    
                    with open(dir_network+'\\'+circuit_name+'_LinesMV.dss', 'r+') as the_file:   #save file's original settings
                        the_file.truncate(0)
                        the_file.writelines(original_lines)
                        the_file.close
                    
                    os.system('clear')
    
                ###################################################################################
                #################################      DAILY      #################################
                ###################################################################################
                # Routine to perform daily power flows
                if self.dlg.powerflow_daily.isChecked():
                    DSSprogress.Show()
                    DSSprogress.Caption = QCoreApplication.translate('progress', u'Daily')
                    DSSprogress.PctProgress = 0
                    
                    tinitial_daily = time.time()  # daily time start
                    dailydate = self.dlg.lineEdit_daily_date.text().upper()  # simulation date
                    dailydate = correct_date( load_curve_circuit, dailydate )
                    self.dailydate = dailydate
                    if not dailydate:  # representative day selection routine
                        dailydate = auxfcns.selection_representative_day(load_curve_circuit, 'weekday')
                    self.dailydate = dailydate
    
                    SabcMC = []
                    DP_to_be_matched = []  # Real P
                    DQ_to_be_matched = []  # Real Q
                    for ij in range(len(circuit_demand)):
                        temp_a = circuit_demand[ij][0]
                        if str(temp_a.replace('/', '')) == str(dailydate.replace('/', '')):
                            DP_to_be_matched.append(circuit_demand[ij][2])
                            DQ_to_be_matched.append(circuit_demand[ij][3])
                            
                    if self.substation_sel == True:
                        
                        circuit_name = name_file_created.split('_')[0]
                        firstLine = 'MV3P'+circuit_name+'00'
                    
                    ######## FIRST SIMULATION: SNAPSHOT TO GET LN VOLTAGE BASES

                    DSStext.Command = 'clear'  # clean previous circuits
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Daily'
                    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
                    DSStext.Command = 'Set mode=snapshot'  # Type of Simulation
                    # DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    # DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                    DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time                
                    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
                    DSStext.Command = 'batchedit load..* enabled = no' # No load simulation
                    DSStext.Command = 'batchedit storage..* enabled = no' # No load simulation
                    DSStext.Command = 'batchedit PVSystem..* enabled = no' # No load simulation
                    DSStext.Command = 'batchedit Generator..* enabled = no' # No load simulation
                    
                    DSScircuit.Solution.Solve()  # solve the circuit
                    
                    
                    all_buses = list(DSScircuit.AllBusVmag)
                    index = list(DSScircuit.AllNodeNames)
                    tam_buses = len(all_buses)
                    #borrar datos innecesarios
                    del_bus = []
                        
                    for i in range(len(index)):
                        if 'AFTERMETER' in str(index[i]).upper():
                            del_bus.append(i)
                    
                    num_del_buses = len(del_bus)
                    
                    for j in range(num_del_buses-1,-1,-1):
                        index.remove(index[del_bus[j]])
                        all_buses.remove(all_buses[del_bus[j]])
                        
                    Base_V = pandas.DataFrame(0, index = index, columns=['base'])
                    base_vals = [120, 138, 208, 240, 254, 277, 416, 440, 480, 7620, 7967, 13200, 13800, 14380, 19920, 24900, 34500, 79670, 132790]
                    
                    for i in range(len(all_buses)):
                        busV = all_buses[i]
                        Base_V['base'].iloc[i] = base_vals[[abs(val-busV) for val in base_vals].index(min([abs(val-busV) for val in base_vals]))] #selecciona el index que de la mínima resta
                    #End get voltage bases
    
                    DSStext.Command = 'clear'  # clean previous circuits
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Daily'  # create a new circuit
                    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # master file compilation
                    DSStext.Command = 'Set mode = daily'  # daily simulation mode
                    DSStext.Command = 'Set number= 1'  # steps by solve
                    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s
                    DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time
                    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
                    
                    if tx_modelling and self.substation != 'Auto':
                        DSStext.Command = line_tx_definition
    
                    # generators powers
                    gen_powers = np.zeros(96)
                    gen_rpowers = np.zeros(96)
                    GenNames = DSScircuit.Generators.AllNames
                    PVNames = DSScircuit.PVSystems.AllNames
                    
                    monitor_data = []
    
                    # solver for-loop
                    for t in range(96):
                        gen_p = 0
                        gen_q = 0
                        DSScircuit.Solution.Solve()
                        if GenNames[0] != 'NONE':
                            for i in GenNames:  # extract power from existing generators
                                DSScircuit.setActiveElement('generator.' + i)
                                p = DSScircuit.ActiveElement.Powers
                                for w in range(0, len(p), 2):
                                    gen_p += -p[w]
                                    gen_q += -p[w+1]
                            gen_powers[t] += gen_p
                            gen_rpowers[t] += gen_q
                        if PVNames[0] != 'NONE':
                            for i in PVNames:  # extract power from existing PVSystems
                                DSScircuit.setActiveElement('PVSystem.' + i)
                                p = DSScircuit.ActiveElement.Powers
                                for w in range(0, len(p), 2):
                                    gen_p += -p[w]
                                    gen_q += -p[w + 1]
                            gen_powers[t] += gen_p
                            gen_rpowers[t] += gen_q
    
                    DSSprogress.PctProgress = 20
                    errorP = 0.003  # Maximum desired correction error for active power
                    errorQ = 0.01  # Maximum desired correction error for reactive power
                    max_it_correction = 15  # Maximum number of allowed iterations
                    study = 'daily'  # Study type for PQ_corrector
                    #DSStext.Command = 'batchedit storage..* enabled = no' # No storage simulation
    
                    # load allocation algorithm            
                    
                    [DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_sim, kVAr_sim] = auxfcns.PQ_corrector(
                        DSSprogress, DSScircuit, DSStext, errorP, errorQ, max_it_correction,
                        DP_to_be_matched, DQ_to_be_matched, dailydate, study, dir_network, tx_modelling, 1, firstLine,
                        self.substation, line_tx_definition, gen_powers, gen_rpowers)
                    
                    if DSScircuit == -1:
                        QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), QCoreApplication.translate('dialog',
                                                                                                   u'Debe introducir una fecha que coincida con alguna fecha en el .csv de la curva del alimentador.\nIgualmente puede dejar en blanco y se seleccionará el día más representativo'))
                        return
                    if DSScircuit == None or DSScircuit == 0:
                        self.iface.messageBar().pushCritical("QGISRunOpenDSS", QCoreApplication.translate('dialog',
                                                                                                   u'Sucedió un error crítico y no fue posible completar la operación'))  
                        return
                    
                    DSSprogress.PctProgress = 50
                    ################################################################
                    
                    DSScircuit.Monitors.ResetAll()
                    # post load allocation simulation
                    DSStext.Command = 'clear'  # clean previous circuits
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Daily'  # create a new circuit
                    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # master file compilation
                    if plantel == True: #se agregan los buses si hay un archivo del plantel de buses
                        DSStext.Command = 'redirect ' + str(name_file_created.split('_')[0]) +  '_StorageBuses.dss'
                    DSStext.Command = 'Set mode = daily'  # daily simulation mode
                    DSStext.Command = 'Set number= 1'  ## steps by solve
                    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s
                    DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time
                    if tx_modelling and self.substation != 'Auto':
                        DSStext.Command = line_tx_definition
    
                    # add pvsystems if checked
                    if self.dlg.PV.isChecked():
                        if self.gui_gd.dlg.checkBox_PV_res.isChecked():
                            PVS += auxfcns.PV_allocation(total_capacity, pv_information, name_file_created, "R",
                                                         pv_study_type)
    
                        if self.gui_gd.dlg.checkBox_PV_gen.isChecked():
                            PVS += auxfcns.PV_allocation(total_capacity_gen, pv_information_gen, name_file_created, "C",
                                                         pv_study_type)
    
                        if self.gui_gd.dlg.checkBox_PV_TMT.isChecked():
                            PVS += auxfcns.PV_allocation(total_capacity_TMT, pv_information_TMT, name_file_created, "TMT",
                                                         pv_study_type)
    
                        for pv in range(len(PVS)):
                            if PVS[pv] == "Error":  # exit if an error occurred
                                return
                            DSStext.Command = PVS[pv]
                            # print(PVS[pv])
                            # DSStext.Command = 'solve'
                        DSSprogress.PctProgress = 60
    
                    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first MV line to monitor P and Q
                    # calc pu voltages
                    DSStext.Command = 'calcvoltagebases'
                    # node voltages lists
                    nodeVoltages_ph1 = []
                    nodeVoltages_ph2 = []
                    nodeVoltages_ph3 = []
                    # other lists
                    currentList = []
                    lossesWList = list()
                    # current simulation powers
                    temp_powersP = []
                    temp_powersQ = []
                    buses = DSScircuit.ALLBusNames  # all buses names
                    unbalance = {}
                    orderedUnbalance = OrderedDict()
                    
                    ############################################
                    for t in range(96):                    
                        DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[t])  # set kW corrector
                        DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[t])  # set kVAr corrector
                        
                        update_storage(DSScircuit, vector_soc_t, t)
                        DSScircuit.Solution.Solve()  # solve the circuit
                        
                        ########################################################
                        all_buses = list(DSScircuit.AllBusVmag)
                        index = list(DSScircuit.AllNodeNames)
                        #eliminar los buses creados por el energy meter
                        del_bus = []
                        
                        for i in range(len(index)):
                            if 'AFTERMETER' in str(index[i]).upper():
                                del_bus.append(i)
                        
                        num_del_buses = len(del_bus)
                        
                        for j in range(num_del_buses-1,-1,-1):
                            index.remove(index[del_bus[j]])
                            all_buses.remove(all_buses[del_bus[j]])

                        ###########################################################
                        
                        if t == 0:  # set pu voltages
                            DSStext.Command = 'redirect ' + name_file_created.split('_')[0] + '_LV_KVBaseLN.dss'
                            DSStext.Command = 'redirect ' + name_file_created.split('_')[0] + '_MV_BaseKV_LN.dss'
                            V_buses = pandas.DataFrame(np.nan, index = index, columns=[x for x in range(96)])  
                        
                        v_allbuses = np.array(all_buses)
                        v_base = np.array(Base_V['base'])
                        
                        if len(v_allbuses) != len(v_base):
                            QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), QCoreApplication.translate('dialog',
                                                                                                   u'Verifique que todas las cargas estén conectadas correctamente a las líneas.\nNota: confirme que las cargas estén conectadas a las fases adecuadas. '))
                            DSSprogress.Close()
                            return
                           
                        V_buses[t] = np.divide(v_allbuses, v_base) #to create a chart of the voltages of all the circuit 
                        ############################################
                        
                        DSScircuit.setActiveElement('line.' + firstLine)
                        temp_powers = DSScircuit.ActiveElement.Powers  # extract power from first line
                        temp_powersP.append((temp_powers[2] + temp_powers[4] + temp_powers[0]))
                        temp_powersQ.append((temp_powers[3] + temp_powers[5] + temp_powers[1]))
    
                        if self.dlg.checkBoxLosses.isChecked():  # append losses for every instant
                            lossesWList.append(DSScircuit.Losses[0] / 1000)
    
                        if self.dlg.checkBoxUnbalance.isChecked():  # calc voltage unbalances
                            orderedUnbalance = auxfcns.unbalance(self, DSScircuit, buses, unbalance, orderedUnbalance)
                        # save on results shape
                        if t == 0:                            
                            # node voltages for every instant
                            nodeNames_ph1 = DSScircuit.AllNodeNamesByPhase(1)
                            nodeNames_ph2 = DSScircuit.AllNodeNamesByPhase(2)
                            nodeNames_ph3 = DSScircuit.AllNodeNamesByPhase(3)
                            lineNames = DSScircuit.Lines.allnames  # all lines names
                            normalAmpsDic = auxfcns.normalAmps(DSScircuit, lineNames)  # all lines ampacities
                        nodeVoltages_ph1.append([DSScircuit.AllNodeVmagPUByPhase(1)])  # Buses pu voltages ph1
                        nodeVoltages_ph2.append([DSScircuit.AllNodeVmagPUByPhase(2)])  # Buses pu voltages ph2
                        nodeVoltages_ph3.append([DSScircuit.AllNodeVmagPUByPhase(3)])  # Buses pu voltages ph3
                        
                        currentList.append( auxfcns.lineCurrents(self, DSScircuit, lineNames, normalAmpsDic, study))  ##Normalized currents
                    
                    
                    #####################################
                    #Dato potencia en secundario del trafo
                    
                    """
                    monitor_name = "MonSRA3P_plant_ruta51"
                    monitor_data = auxfcns.ExtractMonitorData(DSScircuit, monitor_name, [1, 3, 5], 1)
                    
                    import pandas as pd
                    name_file = "potencia_" + monitor_name + ".csv"
                    output_file = self.output_folder + "\\" + name_file
                    pd.DataFrame( monitor_data ).to_csv( output_file )
                    
                    for i in range(34):
                        monitor_name = "mon_plant_ruta51_" + str(i)
                        name_file = "potencia_" + monitor_name + ".csv"
                        monitor_data = auxfcns.ExtractMonitorData(DSScircuit, monitor_name, [1, 3, 5], 1)
                        output_file = self.output_folder + "\\" + name_file
                        pd.DataFrame( monitor_data ).to_csv( output_file )
                    """
                        
                    
                    DSSprogress.PctProgress = 70
                    trafosDict, trafosNames = auxfcns.ReadTrafosLoad(self, DSScircuit, DSStext, name_file_created)
                    busesDict = auxfcns.ReadBusVolts(self, nodeVoltages_ph1, nodeVoltages_ph2, nodeVoltages_ph3, nodeNames_ph1, nodeNames_ph2, nodeNames_ph3)
                    
                    #Se determina el número de sobrecargas de los trafos y problemas de tensión en casas
                    n_sobrecargas, lista_sobrecargas = self.overload_trafos( trafosDict )
                    n_casas_con_problemas, list_DSSNameBuseswithProblems = self.GetVoltageProblems( busesDict )
    
                    # Results shapes update
                    if self.dlg.checkBox_capas.isChecked():
                        list_layersTrafos, list_layersLinesMV, list_layersLinesLV = auxfcns.get_layersnames( self.name_output_azul )
                        auxfcns.lineCurrentsResults(self, lineNames, currentList, study, list_layersLinesMV, list_layersLinesLV)  # currents shape
                        DSSprogress.PctProgress = 80
                        auxfcns.WriteTrafosLoad(self, trafosDict, trafosNames, study, list_layersTrafos)
                        auxfcns.WriteBusVolts(self, busesDict, name_file_created, study)
                        DSSprogress.PctProgress = 90
    
                    # Aparent power calculation
                    S_sim = np.sqrt(np.power(temp_powersP, 2) + np.power(temp_powersQ, 2))
                    SabcMC.append(S_sim)
                    # P, Q and S from substation curve
                    P_real = [float(x) for x in DP_to_be_matched]
                    Q_real = [float(x) for x in DQ_to_be_matched]
                    S_real = np.sqrt(np.power(P_real, 2) + np.power(Q_real, 2))
                    # results output path
                    final_output_folder = output_folder + '/Daily/'
                    if not os.path.exists(final_output_folder):
                        os.makedirs(final_output_folder)
    
                    # Run Monte Carlo Simulations
                    for MC_iteration in range(int(number_simulations)):
    
                        # Generate a series of dates (these are in matplotlib's internal date format)
                        hours = mdates.drange(dt.datetime(2015, 1, 1), dt.datetime(2015, 1, 2), dt.timedelta(minutes=15))
    
                        # Tell matplotlib to interpret the x-axis values as dates
                        fmtr = mdates.DateFormatter("%H:%M")
                        # Aparent power graph creation
                        fig1 = plt.figure(1)
                        ax = plt.gca()
                        ax.plot(hours, S_real, label=QCoreApplication.translate('graphs', u'Real'))
                        ax.plot(hours, S_sim, label=QCoreApplication.translate('graphs', u'Simulación'))
                        ax.xaxis.set_major_formatter(fmtr)
                        ax.set_xlabel(QCoreApplication.translate('graphs', u'Hora'))
                        ax.set_title(QCoreApplication.translate('graphs', u'Potencia Aparente'))
                        DSScircuit.Monitors.ResetAll()
                        # Real power graph creation
                        fig2 = plt.figure(2)
                        ax = plt.gca()
                        ax.plot(hours, P_real, label=QCoreApplication.translate('graphs', u'Real'))
                        ax.plot(hours, temp_powersP, label=QCoreApplication.translate('graphs', u'Simulación'))
                        ax.xaxis.set_major_formatter(fmtr)
                        ax.set_xlabel(QCoreApplication.translate('graphs', u'Hora'))
                        ax.set_title(QCoreApplication.translate('graphs', u'Potencia Real'))
                        DSScircuit.Monitors.ResetAll()
                        # Reactive power graph creation
                        fig3 = plt.figure(3)
                        ax = plt.gca()
                        ax.plot(hours, Q_real, label=QCoreApplication.translate('graphs', u'Real'))
                        ax.plot(hours, temp_powersQ, label=QCoreApplication.translate('graphs', u'Simulación'))
                        ax.xaxis.set_major_formatter(fmtr)
                        ax.set_xlabel(QCoreApplication.translate('graphs', u'Hora'))
                        ax.set_title(QCoreApplication.translate('graphs', u'Potencia Reactiva'))
                        DSScircuit.Monitors.ResetAll()
    
                        # losses graph creation
                        if self.dlg.checkBoxLosses.isChecked():
                            fig4 = plt.figure(4)
                            ax = plt.gca()
                            ELP = np.round(np.sum(lossesWList) * .25, 2)
                            ax.plot(hours, lossesWList, label=str(ELP) + ' kWh')
                            ax.xaxis.set_major_formatter(fmtr)
                            ax.set_xlabel(QCoreApplication.translate('graphs', u'Hora'))
                            ax.set_title(QCoreApplication.translate('graphs', u'Pérdidas Reales'))
                            # fig4 = plt.figure(4)
                            # ELQ = np.round(np.sum(lossesQList) * .25,2)
                            # plt.plot(lossesQList, label=QCoreApplication.translate('graphs', u'Reactiva [kVAr], ') + str(ELQ) + ' kVArh')
                            DSScircuit.Monitors.ResetAll()
    
                    DSSprogress.PctProgress = 100
                    # aparent power graph
                    plt.figure(1)
                    plt.xticks(np.arange(min(hours), max(hours) + 0.04166666663 * 2, 0.04167 * 3))
                    plt.xlim([hours[0], hours[len(hours) - 1]])
                    # plt.xlabel(QCoreApplication.translate('graphs', u'Tiempo'))
                    plt.ylabel(QCoreApplication.translate('graphs', u'Potencia Aparente [kVA]'))  # 'Aparent Power [kVA]'
                    plt.tight_layout()
                    plt.legend(fontsize=8, loc=2)
                    mng = plt.get_current_fig_manager()
                    mng.window.showMaximized() 
                    #fig1.show()
                    # real power graph
                    plt.figure(2)
                    plt.xticks(np.arange(min(hours), max(hours) + 0.04166666663 * 2, 0.04167 * 3))
                    plt.xlim([hours[0], hours[len(hours) - 1]])
                    # plt.xlabel(QCoreApplication.translate('graphs', u'Tiempo'))
                    plt.ylabel(QCoreApplication.translate('graphs', u'Potencia Activa [kW]'))  # 'Active Power [kW]'
                    plt.tight_layout()
                    plt.legend(fontsize=8, loc=2)
                    mng = plt.get_current_fig_manager()
                    mng.window.showMaximized() 
                    #fig2.show()
                    # reactive power graph
                    plt.figure(3)
                    plt.xticks(np.arange(min(hours), max(hours) + 0.04166666663 * 2, 0.04167 * 3))
                    plt.xlim([hours[0], hours[len(hours) - 1]])
                    # plt.xlabel(QCoreApplication.translate('graphs', u'Tiempo'))
                    plt.ylabel(QCoreApplication.translate('graphs', u'Potencia Reactiva [kVAr]'))  # 'Reactive Power [kVAr]'
                    plt.tight_layout()
                    plt.legend(fontsize=8, loc=2)
                    #fig3.show()
                    mng = plt.get_current_fig_manager()
                    mng.window.showMaximized() 
    
                    # losses graph
                    if self.dlg.checkBoxLosses.isChecked():
                        plt.figure(4)
                        plt.xticks(np.arange(min(hours), max(hours) + 0.04166666663 * 2, 0.04167 * 3))
                        plt.xlim([hours[0], hours[len(hours) - 1]])
                        # plt.xlabel(QCoreApplication.translate('graphs', u'Tiempo'))
                        plt.ylabel(
                            QCoreApplication.translate('graphs', u'Pérdidas Activas [kW]'))  # 'Reactive Power [kVAr]'
                        plt.tight_layout()
                        plt.legend(fontsize=8, loc=2)
                        mng = plt.get_current_fig_manager()
                        mng.window.showMaximized() 
                        #fig4.show()
                        fig4.savefig(final_output_folder + '/' + network + '_Losses_' + str_time + '.pdf',
                                     format='pdf',
                                     dpi=6000)
                        csvFile = final_output_folder + '/' + network + '_Losses_' + str_time + '.csv'
                        with open(csvFile, 'w') as f:
                            w = csv.writer(f, lineterminator='\n')
                            #w.writerow(["Losses (kW)"])
                            for row in lossesWList:
                                line = str(row)
                                w.writerow([line])
    
                    # voltage unbalance graph
                    if self.dlg.checkBoxUnbalance.isChecked():
                        x, y = list(zip(*list(orderedUnbalance.items())))
                        x_labels = []  # create an empty list to store the labels
                        if len(x) <= 7:
                            steps = 1
                        else:
                            steps = len(x) / 7
                        ticksIndex = np.arange(start = 0, stop = len(x), step = steps, dtype = int)
                        for key in ticksIndex:  # buses names for graph's ticks
                            x_labels.append(x[key])
    
                        fig5 = plt.figure(5)
                        fig5, ax = plt.subplots(num = "Desbalance de tensión")
                        plt.plot(y, label=QCoreApplication.translate('graphs', u'Desbalance'))
                        DSScircuit.Monitors.ResetAll()
                        # plt.xlim([0, 96])
                        ax.set_xticks(ticksIndex)
                        ax.set_xticklabels(x_labels, rotation=90)
                        plt.xlabel(QCoreApplication.translate('graphs', u'Barras'))
                        plt.ylabel(
                            QCoreApplication.translate('graphs', u'Desbalance (%)'))
                        plt.legend(fontsize=8, loc=2)
                        plt.title(QCoreApplication.translate('graphs', u'Desbalance de tensión'))
                        plt.tight_layout()
                        mng = plt.get_current_fig_manager()
                        mng.window.showMaximized() 
                        #fig5.show()
                        fig5.savefig(final_output_folder + '/' + network + '_Unbalance_' + str_time + '.pdf',
                                     format='pdf',
                                     dpi=6000)
                        csvFile = final_output_folder + '/' + network + '_Unbalance_' + str_time + '.csv'
                        with open(csvFile, 'w') as f:
                            w = csv.writer(f, lineterminator='\n')
                            w.writerow(["Bus", "Unbalance (%)"])
                            for row in range(len(x)):
                                #line = str(x[row]) + "," + str(y[row])
                                w.writerow([str(x[row]), str(y[row])])
                    # real power figure creation
                    fig2.savefig(final_output_folder + '/' + network + '_MC_Result_plot_' + str_time + '.pdf', format='pdf',
                                 dpi=6000)
    
                    # powers: csv file creation
                    csvFile = final_output_folder + '/' + network + '_Powers_' + str_time + '.csv'
                    with open(csvFile, 'w') as f:
                        w = csv.writer(f, lineterminator='\n')
                        w.writerow(["P (kW)", "Q (kVAr)"])
                        for row in range(96):
                            lineP = str(temp_powersP[row])
                            lineQ = str(temp_powersQ[row])
                            w.writerow([lineP, lineQ])
                
                    #Se llama a la función que grafica las tensiones de los buses
                    self.GraphDailyVoltages(DSScircuit, V_buses)
                    tfinal_daily = time.time() - tinitial_daily
                    DSSprogress.Close()
                    
                    
                    if not self.dlg.powerflow_snapshot.isChecked() and not self.dlg.short_circuit.isChecked() \
                            and not self.dlg.harmonics.isChecked() and not self.dlg.powerflow_yearly.isChecked():
                        QMessageBox.information(None, QCoreApplication.translate('dialog', "Simulacion Diaria Terminada"), \
                                                QCoreApplication.translate('dialog', "Tiempo de simulacion: ") + str(
                                                    tfinal_daily + time_common_for_all) + " s" + "\n" + \
                                                QCoreApplication.translate('dialog',
                                                                           "Los archivos han sido guardados en: ") + output_folder)
                                                                           
                    with open(dir_network+'\\'+circuit_name+'_LinesMV.dss', 'r+') as the_file:   #save file's original settings
                        the_file.truncate(0)
                        the_file.writelines(original_lines)
                        the_file.close
                        
                    os.system('clear')
    
                ###################################################################################
                #################################      YEARLY      ################################
                ###################################################################################
                # yearly power flow simulation routine
                if self.dlg.powerflow_yearly.isChecked():
                    self.progress.show()
                    self.progress.progressBar.setValue(0)
    
                    tinitial_yearly = time.time() # yearly simulation init time
                    yearly_resolution = self.dlg.lineEdit_yearly.text().upper() # yearly simulation resolution
                    if not yearly_resolution:  # default: 1h
                        yearly_resolution = '1'
    
                    final_output_folder = output_folder + '/Yearly/'
                    if not os.path.exists(final_output_folder):
                        os.makedirs(final_output_folder)
    
                    #
                    lossesWList = list()
                    # lossesQList = list()
                    SabcMC = []
                    DP_to_be_matched = []
                    DQ_to_be_matched = []
                    steps = 4 * int(yearly_resolution)  # *4 because substation info have 15min steps
                    for i in range(1, len(circuit_demand), steps):
                        temp_a = circuit_demand[i][0]
                        DP_to_be_matched.append(circuit_demand[i][2])
                        DQ_to_be_matched.append(circuit_demand[i][3])
    
                    DSSprogress.PctProgress = 5
                    self.progress.progressBar.setRange(0, (35040 / steps) + 100)
                    DSStext.Command = 'clear' # clean previous circuits
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot' # create a new circuit
                    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
                    DSStext.Command = 'Set mode=yearly'  # Type of Simulation
                    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    DSStext.Command = 'Set h=' + yearly_resolution + 'h'
                    # DSStext.Command = 'Solve'
                    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the transformer secondary side to monitor P and Q
                    if tx_modelling and self.substation != 'Auto':
                        DSStext.Command = line_tx_definition
    
                    # generators powers arrays
                    gen_powers = np.zeros( int(35040 / steps) )
                    gen_rpowers = np.zeros( int(35040 / steps) )
                    GenNames = DSScircuit.Generators.AllNames
                    PVNames = DSScircuit.PVSystems.AllNames
                                   
    
                    for t in range( int(35040 / steps) ):
                        gen_p = 0
                        gen_q = 0
                        DSScircuit.Solution.Solve()
                        if GenNames[0] != 'NONE': # extract power from existing generators
                            for i in GenNames:
                                DSScircuit.setActiveElement('generator.' + i)
                                p = DSScircuit.ActiveElement.Powers
                                for w in range(0, len(p), 2):
                                    gen_p += -p[w]
                                    gen_q += -p[w + 1]
                            gen_powers[t] += gen_p
                            gen_rpowers[t] += gen_q
                        if PVNames[0] != 'NONE': # extract power from existing PVSystems
                            for i in PVNames:
                                DSScircuit.setActiveElement('PVSystem.' + i)
                                p = DSScircuit.ActiveElement.Powers
                                for w in range(0, len(p), 2):
                                    gen_p += -p[w]
                                    gen_q += -p[w + 1]
                            gen_powers[t] += gen_p
                            gen_rpowers[t] += gen_q
    
                    DSSprogress.PctProgress = 20
                    self.progress.progressBar.setValue(100)
    
                    errorP = 0.003  # Maximum desired correction error for active power
                    errorQ = 0.001  # Maximum desired correction error for reactive power
                    max_it_correction = 100  # Maximum number of allowed iterations
                    study = 'yearly'  # Study type for PQ_corrector
                    #DSStext.Command = 'batchedit storage..* enabled = no' # No storage simulation
                    
    
                    # load allocation algorithm
                    [DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_sim,
                     kVAr_sim] = auxfcns.PQ_corrector(DSSprogress,
                                                      DSScircuit, DSStext, errorP, errorQ, max_it_correction,
                                                      DP_to_be_matched, DQ_to_be_matched, '000000', study, dir_network,
                                                      tx_active, steps, firstLine, self.substation, line_tx_definition, gen_powers, gen_rpowers)
                    
                    if DSScircuit == -1:
                        QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), QCoreApplication.translate('dialog',
                                                                                                   u'Debe introducir una fecha que coincida con alguna fecha en el .csv de la curva del alimentador.\nIgualmente puede dejar en blanco y se seleccionará el día más representativo'))
                        return
                    if DSScircuit == None:
                        self.iface.messageBar().pushCritical("QGIS2RunOpenDSS", QCoreApplication.translate('dialog',
                                                                                                   u'Sucedió un error grave y no fue posible completar la operación'))  
                        DSSprogress.Close()
                        return
            
    
                    DSSprogress.PctProgress = 50
    
                    ################################################################
                    DSScircuit.Monitors.ResetAll()
                    DSStext.Command = 'clear'  # clean previous circuits
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'  # create a new circuit
                    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
                    if plantel == True: #se agregan los buses si existe el plantel de buses
                        DSStext.Command = 'redirect ' + str(name_file_created.split('_')[0]) +  '_StorageBuses.dss'
                    print( str(name_file_created.split('_')[0] +  '_StorageBuses.dss' ), " yearly" )
                    DSStext.Command = 'Set mode=yearly'  # Type of Simulation
                    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    DSStext.Command = 'Set h=' + yearly_resolution + 'h'  # simulation resolution, in hours
                    # DSStext.Command = 'Solve'
                    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
                    DSStext.Command = 'calcvoltagebases'  # calc pu voltages
                    if tx_modelling and self.substation != 'Auto':
                        DSStext.Command = line_tx_definition  # edit or add a substation transformer
    
                    # Add PV systems with pv allocation algorithm
                    if self.dlg.PV.isChecked():
                        if self.gui_gd.dlg.checkBox_PV_res.isChecked():
                            PVS += auxfcns.PV_allocation(total_capacity, pv_information, name_file_created, "R",
                                                         pv_study_type)
    
                        if self.gui_gd.dlg.checkBox_PV_gen.isChecked():
                            PVS += auxfcns.PV_allocation(total_capacity_gen, pv_information_gen, name_file_created, "C",
                                                         pv_study_type)
    
                        if self.gui_gd.dlg.checkBox_PV_TMT.isChecked():
                            PVS += auxfcns.PV_allocation(total_capacity_TMT, pv_information_TMT, name_file_created, "TMT",
                                                         pv_study_type)
    
                        for pv in range(len(PVS)):
                            if PVS[pv] == "Error":  # exit if an error occurred
                                return
                            DSStext.Command = PVS[pv]
                            # print(PVS[pv])
                            # DSStext.Command = 'solve'
                    DSSprogress.PctProgress = 60
                    # Iteración de control
                    temp_powersP = []
                    temp_powersQ = []
    
                    # voltage unbalance init
                    buses = DSScircuit.ALLBusNames
                    unbalance = {}
                    orderedUnbalance = OrderedDict()
    
                    for t in range(35040 / steps):
                        self.progress.progressBar.setValue(100 + t)
                        DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[t])  # kW corrector
                        DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[t])  # kVAr corrector
                        DSScircuit.Solution.Solve()
                        DSScircuit.setActiveElement('line.' + firstLine)
                        temp_powers = DSScircuit.ActiveElement.Powers  # circuit powers at instant t
                        temp_powersP.append((temp_powers[2] + temp_powers[4] + temp_powers[0]))
                        temp_powersQ.append((temp_powers[3] + temp_powers[5] + temp_powers[1]))
                        # circuit losses at instant t
                        if self.dlg.checkBoxLosses.isChecked():
                            lossesWList.append(DSScircuit.Losses[0] / 1000)
                            # lossesQList.append(DSScircuit.Losses[1] / 1000)
                        # voltage unbalance at instant t
                        if self.dlg.checkBoxUnbalance.isChecked():
                            orderedUnbalance = auxfcns.unbalance(self, DSScircuit, buses, unbalance, orderedUnbalance)
                    DSSprogress.PctProgress = 80
                    # aparent power calc
                    S_sim = np.sqrt(np.power(temp_powersP, 2) + np.power(temp_powersQ, 2))
                    SabcMC.append(S_sim)
    
                    P_real = [float(x) for x in DP_to_be_matched]
                    Q_real = [float(x) for x in DQ_to_be_matched]
                    S_real = np.sqrt(np.power(P_real, 2) + np.power(Q_real, 2))
    
                    # Run Monte Carlo Simulations
                    for MC_iteration in range(int(number_simulations)):
    
                        # aparent power graph init
                        fig1 = plt.figure(1)
                        plt.plot(S_real, label=QCoreApplication.translate('graphs', u'Real'))
                        fig1 = plt.figure(1)
                        plt.plot(S_sim, label=QCoreApplication.translate('graphs', u'Simulación'))
                        plt.title(QCoreApplication.translate('graphs', u'Potencia Aparente'))
                        DSScircuit.Monitors.ResetAll()
                        # real power graph init
                        fig2 = plt.figure(2)
                        plt.plot(P_real, label=QCoreApplication.translate('graphs', u'Real'))
                        fig2 = plt.figure(2)
                        plt.plot(temp_powersP, label=QCoreApplication.translate('graphs', u'Simulación'))
                        plt.title(QCoreApplication.translate('graphs', u'Potencia Real'))
                        DSScircuit.Monitors.ResetAll()
                        # reactive power graph init
                        fig3 = plt.figure(3)
                        plt.plot(Q_real, label=QCoreApplication.translate('graphs', u'Real'))
                        fig3 = plt.figure(3)
                        plt.plot(temp_powersQ, label=QCoreApplication.translate('graphs', u'Simulación'))
                        plt.title(QCoreApplication.translate('graphs', u'Potencia Reactiva'))
                        DSScircuit.Monitors.ResetAll()
                        # losses graph init
                        if self.dlg.checkBoxLosses.isChecked():
                            fig4 = plt.figure(4)
                            ELP = np.round(np.sum(lossesWList) * int(yearly_resolution), 2)
                            plt.plot(lossesWList,
                                     label=str(ELP) + ' kWh')
                            plt.title(QCoreApplication.translate('graphs', u'Pérdidas Reales'))
                            # fig4 = plt.figure(4)
                            # ELQ = np.round(np.sum(lossesQList) * int(yearly_resolution),2)
                            # plt.plot(lossesQList, label=QCoreApplication.translate('graphs', u'Reactiva [kVAr], ') + str(ELQ) + ' kVArh')
                            DSScircuit.Monitors.ResetAll()
    
                    DSSprogress.PctProgress = 90
    
                    # aparent power graph
                    plt.figure(1)
                    plt.xlim([0, 35040 / steps])
                    plt.xlabel(QCoreApplication.translate('graphs', u'Tiempo'))
                    plt.ylabel(QCoreApplication.translate('graphs', u'Potencia Aparente [kVA]'))  # 'Aparent Power [kVA]'
                    plt.tight_layout()
                    plt.legend(fontsize=8, loc=2)
                    #fig1.show()
                    mng = plt.get_current_fig_manager()
                    mng.window.showMaximized() 
                    # real power graph
                    plt.figure(2)
                    plt.xlim([0, 35040 / steps])
                    plt.xlabel(QCoreApplication.translate('graphs', u'Tiempo'))
                    plt.ylabel(QCoreApplication.translate('graphs', u'Potencia Activa [kW]'))  # 'Active Power [kW]'
                    plt.tight_layout()
                    plt.legend(fontsize=8, loc=2)
                    #fig2.show()
                    mng = plt.get_current_fig_manager()
                    mng.window.showMaximized() 
                    # reactive power graph
                    plt.figure(3)
                    plt.xlim([0, 35040 / steps])
                    plt.xlabel(QCoreApplication.translate('graphs', u'Tiempo'))
                    plt.ylabel(QCoreApplication.translate('graphs', u'Potencia Reactiva [kVAr]'))  # 'Reactive Power [kVAr]'
                    plt.tight_layout()
                    plt.legend(fontsize=8, loc=2)
                    mng = plt.get_current_fig_manager()
                    mng.window.showMaximized() 
                    #fig3.show()
                    # results final path
                    final_output_folder = output_folder + '/Yearly/'
                    if not os.path.exists(final_output_folder):
                        os.makedirs(final_output_folder)
                    # losses graph
                    if self.dlg.checkBoxLosses.isChecked():
                        plt.figure(4)
                        plt.xlim([0, 96])
                        plt.xlabel(QCoreApplication.translate('graphs', u'Tiempo'))
                        plt.ylabel(
                            QCoreApplication.translate('graphs', u'Pérdidas Activas [kW]'))  # 'Reactive Power [kVAr]'
                        plt.tight_layout()
                        plt.legend(fontsize=8, loc=2)
                        mng = plt.get_current_fig_manager()
                        mng.window.showMaximized() 
                        #fig4.show()
                        fig4.savefig(final_output_folder + '/' + network + '_Losses_' + str_time + '.pdf',
                                     format='pdf',
                                     dpi=6000)
    
                    # voltage unbalance graph and csv
                    if self.dlg.checkBoxUnbalance.isChecked():
                        x, y = list(zip(*list(orderedUnbalance.items())))
                        x_labels = []  # create an empty list to store the labels
                        ticksIndex = np.arange(0, len(x), len(x) / 7, int)
                        for key in ticksIndex:  # buses names for graph ticks
                            x_labels.append(x[key])
                        fig5 = plt.figure(5)
                        fig5, ax = plt.subplots()
                        plt.plot(y, label=QCoreApplication.translate('graphs', u'Desbalance'))
                        DSScircuit.Monitors.ResetAll()
                        ax.set_xticks(ticksIndex)
                        ax.set_xticklabels(x_labels, rotation=90)
                        plt.xlabel(QCoreApplication.translate('graphs', u'Barras'))
                        plt.ylabel(
                            QCoreApplication.translate('graphs', u'Desbalance (%)'))
                        plt.legend(fontsize=8, loc=2)
                        plt.tight_layout()
                        mng = plt.get_current_fig_manager()
                        mng.window.showMaximized() 
                        #fig5.show()
                        fig5.savefig(final_output_folder + '/' + network + '_Unbalance_' + str_time + '.pdf',
                                     format='pdf',
                                     dpi=6000)
                        csvFile = final_output_folder + '/' + network + '_Unbalance_' + str_time + '.csv'
                        with open(csvFile, 'w') as f:  # write csv file
                            w = csv.writer(f, lineterminator='\n')
                            w.writerow(["Bus,Unbalance (%)"])
                            for row in range(len(x)):
                                line = str(x[row]) + "," + str(y[row])
                                w.writerow([line])
                    DSSprogress.PctProgress = 100
                    # real power figure creation
                    fig2.savefig(final_output_folder + '/' + network + '_MC_Result_plot_' + str_time + '.pdf', format='pdf',
                                 dpi=6000)
    
                    # powers: csv file creation
                    csvFile = final_output_folder + '/' + network + '_Powers_' + str_time + '.csv'
                    with open(csvFile, 'w') as f:
                        w = csv.writer(f, lineterminator='\n')
                        w.writerow(["P (kW)", "Q (kVAr)"])
                        for row in range(35040 / steps):
                            lineP = str(temp_powersP[row])
                            lineQ = str(temp_powersQ[row])
                            w.writerow([lineP, lineQ])
                    DSSprogress.Close()
                    self.progress.show()
                    tfinal_yearly = time.time() - tinitial_yearly  # yearly simulation end time
                    if not self.dlg.powerflow_snapshot.isChecked() and not self.dlg.powerflow_daily.isChecked() \
                            and not self.dlg.short_circuit.isChecked() and not self.dlg.harmonics.isChecked():
                        QMessageBox.information(None, QCoreApplication.translate('dialog', "Simulacion Anual Terminada"), \
                                                QCoreApplication.translate('dialog', "Tiempo de simulacion: ") + str(
                                                    tfinal_yearly + time_common_for_all) + " s" + "\n" + \
                                                QCoreApplication.translate('dialog',
                                                                           "Los archivos han sido guardados en: ") + output_folder)
                    os.system('clear')
    
                ###################################################################################
                #############################      SHORT CIRCUIT      #############################
                ###################################################################################
                # short circuit simulation routine
                if self.dlg.short_circuit.isChecked():
                    tinitial_short_circuit = time.time()
    
                    DSSprogress.Show()
                    DSSprogress.Caption = QCoreApplication.translate('progress', u'Short Circuit')
                    DSSprogress.PctProgress = 0
    
                    # simulation date
                    scdate = self.dlg.lineEdit_sc_date.text().upper()
                    scdate = correct_date( load_curve_circuit, scdate )
                    if not scdate: # if empty: representative day selection
                        scdate = auxfcns.selection_representative_day(load_curve_circuit, 'weekday')
                    # simulation time
                    sctime = self.dlg.lineEdit_sc_time.text()  # .upper()
                    sctime = correct_hour( load_curve_circuit, sctime )
                    if not sctime:
                        sctime = def_time  # Default is 6pm
    
                    h, m = sctime.split(':')
                    if m is not '00' or '15' or '30' or '45':  # round time to 15min multiples
                        if int(m) <= 7:
                            m = '00'
                        elif int(m) <= 22:
                            m = '15'
                        elif int(m) <= 37:
                            m = '30'
                        elif int(m) <= 52:
                            m = '45'
                        else:
                            m = '00'
                            h = str(int(h) + 1)
                            if int(h) == 24:  # last rount stays on 23:45
                                h = '23'
                                m = '45'
                    sctime = h + ':' + m
    
                    daily_strtime = str(scdate.replace('/', '') + sctime.replace(':', ''))
                    hora_sec = sctime.split(':')
                    for i in range(len(circuit_demand)):
                        temp_a = circuit_demand[i][0]  # day
                        temp_b = circuit_demand[i][1]  # hour
                        if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:
                            P_to_be_matched = circuit_demand[i][2]  # Active power
                            Q_to_be_matched = circuit_demand[i][3]  # Reactive power
    
                    # SC bus
                    short_circuit_bus = self.dlg.lineEdit_short_circuit.currentText().upper()
    
                    # SC 3ph
                    terminales = ''
                    terminales2 = ''
                    terminales3 = ''
                    short_circuit_phases = '3'
                    if self.dlg.checkBox_SC_phaseA.isChecked() == True and self.dlg.checkBox_SC_phaseB.isChecked() == True and self.dlg.checkBox_SC_phaseC.isChecked() == True:
                        short_circuit_phases = '3'
                        terminales = '.1.2.3'
                        terminales2 = '.0.0.0'
                        if self.dlg.checkBox_SC_phaseLL.isChecked() == True:  # line to line
                            short_circuit_phases = '1'
                            terminales = '.1'
                            terminales2 = '.2'
                            terminales3 = '.3'
                    # SC 2ph
                    elif self.dlg.checkBox_SC_phaseA.isChecked() == True and self.dlg.checkBox_SC_phaseB.isChecked() == True and self.dlg.checkBox_SC_phaseC.isChecked() == False:
                        short_circuit_phases = '1'
                        if self.dlg.checkBox_SC_phaseLL.isChecked() == True:  # line to line
                            terminales = '.1'
                            terminales2 = '.2'
                        else:  # Line to Ground (LG)
                            short_circuit_phases = '2'
                            terminales = '.1.1'
                            terminales2 = '.2.0'
                    elif self.dlg.checkBox_SC_phaseA.isChecked() == True and self.dlg.checkBox_SC_phaseB.isChecked() == False and self.dlg.checkBox_SC_phaseC.isChecked() == True:
                        short_circuit_phases = '1'
                        if self.dlg.checkBox_SC_phaseLL.isChecked() == True:  # line to line
                            terminales = '.1'
                            terminales2 = '.3'
                        else:  # LG
                            short_circuit_phases = '2'
                            terminales = '.1.1'
                            terminales2 = '.3.0'
                    elif self.dlg.checkBox_SC_phaseA.isChecked() == False and self.dlg.checkBox_SC_phaseB.isChecked() == True and self.dlg.checkBox_SC_phaseC.isChecked() == True:
                        short_circuit_phases = '1'
                        if self.dlg.checkBox_SC_phaseLL.isChecked() == True:  # line to line
                            terminales = '.2'
                            terminales2 = '.3'
                        else:  # LG
                            short_circuit_phases = '2'
                            terminales = '.2.2'
                            terminales2 = '.3.0'
                    # SC 1ph
                    elif self.dlg.checkBox_SC_phaseA.isChecked() == True and self.dlg.checkBox_SC_phaseB.isChecked() == False and self.dlg.checkBox_SC_phaseC.isChecked() == False:
                        short_circuit_phases = '1'
                        terminales = '.1'
                        terminales2 = '.0'
                    elif self.dlg.checkBox_SC_phaseA.isChecked() == False and self.dlg.checkBox_SC_phaseB.isChecked() == True and self.dlg.checkBox_SC_phaseC.isChecked() == False:
                        short_circuit_phases = '1'
                        terminales = '.2'
                        terminales2 = '.0'
                    elif self.dlg.checkBox_SC_phaseA.isChecked() == False and self.dlg.checkBox_SC_phaseB.isChecked() == False and self.dlg.checkBox_SC_phaseC.isChecked() == True:
                        short_circuit_phases = '1'
                        terminales = '.3'
                        terminales2 = '.0'
    
                    # bus selection simplifier
                    if short_circuit_bus == QCoreApplication.translate('SC', "Todas las barras MT").upper():
                        short_circuit_bus = QCoreApplication.translate('SC', 'MT')
                    elif short_circuit_bus == QCoreApplication.translate('SC', 'Todas las barras BT').upper():
                        short_circuit_bus = QCoreApplication.translate('SC', 'BT')
                    elif short_circuit_bus == QCoreApplication.translate('SC', 'Todas las barras MT y BT').upper():
                        short_circuit_bus = QCoreApplication.translate('SC', 'MT_BT')
    
                    # results output folder
                    final_output_folder = output_folder + '/Short_Circuit/' + 'BUS_' + str(short_circuit_bus) + '/'
                    if not os.path.exists(final_output_folder):
                        os.makedirs(final_output_folder)
                    DSSprogress.PctProgress = 20
    
                    DSStext.Command = 'clear'  # clean previous circuits
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'  # new circuit creation
                    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
                    DSStext.Command = 'Set mode=daily'  # Type of Simulation
                    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation
                    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the transformer secondary side to monitor P and Q
                    if tx_modelling and self.substation != 'Auto':
                        DSStext.Command = line_tx_definition
                    for MC_iteration in range( int(number_simulations) ):  # Run Monte Carlo Simulations
                        # Run the daily power flow for a particular moment
                        DSScircuit.Solution.Solve()  # Initialization solution
                        busNames = DSScircuit.ALLBusNames  # all buses names
                        gen_p = 0
                        gen_q = 0
                        gen_powers = np.zeros(1)
                        gen_rpowers = np.zeros(1)
                        GenNames = DSScircuit.Generators.AllNames
                        PVNames = DSScircuit.PVSystems.AllNames
                        if GenNames[0] != 'NONE':  # existing distributed generators powers
                            for i in GenNames:
                                DSScircuit.setActiveElement('generator.' + i)
                                p = DSScircuit.ActiveElement.Powers
                                for w in range(0, len(p), 2):
                                    gen_p += -p[w]
                                    gen_q += -p[w + 1]
                            gen_powers[0] += gen_p
                            gen_rpowers[0] += gen_q
                        if PVNames[0] != 'NONE': # existing PV systems powers
                            for i in PVNames:
                                DSScircuit.setActiveElement('PVSystem.' + i)
                                p = DSScircuit.ActiveElement.Powers
                                for w in range(0, len(p), 2):
                                    gen_p += -p[w]
                                    gen_q += -p[w + 1]
                            gen_powers[0] += gen_p
                            gen_rpowers[0] += gen_q

                        errorP = 0.03  # Maximum desired correction error for active power
                        errorQ = 0.1  # Maximum desired correction error for reactive power
                        max_it_correction = 10  # Maximum number of allowed iterations
                        study = 'snapshot'  # Study type for PQ_corrector
                        #DSStext.Command = 'batchedit storage..* enabled = no' # No storage simulation

                        # load allocation algorithm
                        [DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_sim,
                         kVAr_sim] = auxfcns.PQ_corrector(DSSprogress, DSScircuit, DSStext, errorP, errorQ,
                                                          max_it_correction,
                                                          P_to_be_matched, Q_to_be_matched, hora_sec, study,
                                                          dir_network, tx_modelling, 1, firstLine, self.substation,
                                                          line_tx_definition, gen_powers, gen_rpowers)

                        if DSScircuit == -1:
                            QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), QCoreApplication.translate('dialog',
                                                                                                       u'Debe introducir una fecha y hora que coincida con alguna fecha en el .csv de la curva del alimentador.\nIgualmente puede dejar en blanco y se seleccionará el día más representativo'))
                            return
                        DSSprogress.PctProgress = 50
                        # post load allocation simulation
                        DSStext.Command = 'clear'  # clean previous circuits
                        DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'  # new circuit creation
                        DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
                        DSStext.Command = 'Set mode=daily'  # Type of Simulation
                        DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                        DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                        DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the simulation time
                        # PV allocation algorithm
                        if self.dlg.PV.isChecked():
                            if self.gui_gd.dlg.checkBox_PV_res.isChecked():
                                PVS += auxfcns.PV_allocation(total_capacity, pv_information, name_file_created, "R",
                                                             pv_study_type)
    
                            if self.gui_gd.dlg.checkBox_PV_gen.isChecked():
                                PVS += auxfcns.PV_allocation(total_capacity_gen, pv_information_gen, name_file_created, "C",
                                                             pv_study_type)
    
                            if self.gui_gd.dlg.checkBox_PV_TMT.isChecked():
                                PVS += auxfcns.PV_allocation(total_capacity_TMT, pv_information_TMT, name_file_created,
                                                             "TMT",
                                                             pv_study_type)
    
                            for pv in range(len(PVS)):
                                if PVS[pv] == "Error":  # exit if error occurred
                                    DSSprogress.Close()
                                    return
                                DSStext.Command = PVS[pv]
                            #QMessageBox.information(None, "Debug", PVS[0])
                                # DSStext.Command = 'solve'
                        DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
                        DSStext.Command = 'calcvoltagebases'  # pu voltages calc
                        if tx_modelling and self.substation != 'Auto':
                            DSStext.Command = line_tx_definition  # edit/create substation transformer
                        # create node voltages lists
                        nodeVoltages_ph1 = []
                        nodeVoltages_ph2 = []
                        nodeVoltages_ph3 = []
                        currentList = []  # current list
                        DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0])  # set kW corrector
                        DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0])  # set kVAr corrector
                        DSStext.Command = 'batchedit PVSystem..* %R=75'  # set %R to PV systems for 1.1 pu SC current
                        DSStext.Command = 'batchedit PVSystem..* %X=75'  # set %X to PV systems for 1.1 pu SC current
                        if not 5 < int(hora_sec[0]) < 18:
                            DSStext.Command = 'batchedit PVSystem..* enabled=no'  # set PVSystem enable
                        # set results output folder
                        DSStext.Command = 'redirect ' + name_file_created.split('_')[0] + '_LV_KVBaseLN.dss'
                        DSStext.Command = 'redirect ' + name_file_created.split('_')[0] + '_MV_BaseKV_LN.dss'
                        DSStext.Command = 'set datapath=' + final_output_folder
                        DSScircuit.Solution.Solve()
    
                        # create a new fault
                        DSStext.Command = 'New Fault.F1'
                        DSStext.Command = 'New Fault.F2'
                        if short_circuit_bus == "":  # exit on empty information
                            QMessageBox.information(None, QCoreApplication.translate('dialog',
                                                                                     "Analisis de Corto Circuito Terminado"), \
                                                    QCoreApplication.translate('dialog', "No se ha seleccionado barra de estudio"))
                            DSSprogress.Close()
                            return
                        elif short_circuit_bus == QCoreApplication.translate('SC', 'MT'):  # all MV buses simulation
                            for i in busNames:
    
                                if 'MV' in i.upper():
                                    DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
                                    sc_bus1 = i + terminales
                                    sc_bus2 = i + terminales2
                                    sc_bus3 = i + terminales3
                                    if self.dlg.checkBox_SC_phaseLL.isChecked() and short_circuit_phases == '2':
                                        sc_bus2 = i + '.' + terminales.split('.')[2]
    
                                    # edit fault object
                                    DSStext.Command = 'Edit Fault.F1 phases=' + short_circuit_phases + ' Bus1=' + sc_bus1 + ' Bus2=' + sc_bus2 + ' R=0.002'
                                    if terminales3 != '':
                                        DSStext.Command = 'Edit Fault.F2 phases=' + short_circuit_phases + ' Bus1=' + sc_bus2 + ' Bus2=' + sc_bus3 + ' R=0.002'
                                    DSStext.Command = 'Solve'
                                    DSStext.Command = 'Export Current ' + i + '.csv'
                            DSSprogress.PctProgress = 70
                        elif short_circuit_bus == QCoreApplication.translate('SC', 'BT'): # all LV buses simulation
                            busNames = DSScircuit.ALLBusNames
                            for i in busNames:
    
                                if 'LV' in i.upper():
                                    DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
                                    sc_bus1 = i + terminales
                                    sc_bus2 = i + terminales2
                                    sc_bus3 = i + terminales3
                                    if self.dlg.checkBox_SC_phaseLL.isChecked() and short_circuit_phases == '2':
                                        sc_bus2 = i + '.' + terminales.split('.')[2]
                                    # edit fault object
                                    DSStext.Command = 'Edit Fault.F1 phases=' + short_circuit_phases + ' Bus1=' + sc_bus1 + ' Bus2=' + sc_bus2 + ' R=0.002'
                                    if terminales3 != '':
                                        DSStext.Command = 'Edit Fault.F2 phases=' + short_circuit_phases + ' Bus1=' + sc_bus2 + ' Bus2=' + sc_bus3 + ' R=0.002'
                                    DSStext.Command = 'Solve'
                                    DSStext.Command = 'Export Current ' + i + '.csv'
                            DSSprogress.PctProgress = 80
                        elif short_circuit_bus == QCoreApplication.translate('SC', 'MT_BT'): # all circuit buses simulation
                            busNames = DSScircuit.ALLBusNames
                            for i in busNames:
    
                                DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
                                sc_bus1 = i + terminales
                                sc_bus2 = i + terminales2
                                sc_bus3 = i + terminales3
                                if self.dlg.checkBox_SC_phaseLL.isChecked() and short_circuit_phases == '2':
                                    sc_bus2 = i + '.' + terminales.split('.')[2]
                                DSStext.Command = 'Edit Fault.F1 phases=' + short_circuit_phases + ' Bus1=' + sc_bus1 + ' Bus2=' + sc_bus2 + ' R=0.002'
                                if terminales3 != '':
                                    DSStext.Command = 'Edit Fault.F2 phases=' + short_circuit_phases + ' Bus1=' + sc_bus2 + ' Bus2=' + sc_bus3 + ' R=0.002'
                                DSStext.Command = 'Solve'
                                DSStext.Command = 'Export Current ' + i + '.csv'
                            DSSprogress.PctProgress = 90
                        else:  # specific bus simulation
    
                            DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
                            i = short_circuit_bus
                            sc_bus1 = i + terminales
                            sc_bus2 = i + terminales2
                            sc_bus3 = i + terminales3
                            if self.dlg.checkBox_SC_phaseLL.isChecked() and short_circuit_phases == '2':
                                sc_bus2 = i + '.' + terminales.split('.')[2]
                            DSStext.Command = 'Edit Fault.F1 phases=' + short_circuit_phases + ' Bus1=' + sc_bus1 + ' Bus2=' + sc_bus2 + ' R=0.002'
                            if terminales3 != '':
                                DSStext.Command = 'Edit Fault.F2 phases=' + short_circuit_phases + ' Bus1=' + sc_bus2 + ' Bus2=' + sc_bus3 + ' R=0.002'
                            DSStext.Command = 'Solve'
                            DSStext.Command = 'Export Current ' + i + '.csv'
    
                            if self.dlg.checkBox_capas.isChecked():
    
                                nodeNames_ph1 = DSScircuit.AllNodeNamesByPhase(1)
                                nodeNames_ph2 = DSScircuit.AllNodeNamesByPhase(2)
                                nodeNames_ph3 = DSScircuit.AllNodeNamesByPhase(3)
                                nodeVoltages_ph1.append([DSScircuit.AllNodeVmagPUByPhase(1)])  # Buses pu voltages ph1
                                nodeVoltages_ph2.append([DSScircuit.AllNodeVmagPUByPhase(2)])  # Buses pu voltages ph2
                                nodeVoltages_ph3.append([DSScircuit.AllNodeVmagPUByPhase(3)])  # Buses pu voltages ph3
                                # voltages shapes update
                                busesDict = auxfcns.ReadBusVolts(self, nodeVoltages_ph1, nodeVoltages_ph2, nodeVoltages_ph3, nodeNames_ph1, nodeNames_ph2, nodeNames_ph3)
                                auxfcns.WriteBusVolts(self, busesDict, name_file_created, "shortcircuit")
                                #auxfcns.busVolts(self, nodeVoltages_ph1, nodeVoltages_ph2, nodeVoltages_ph3, nodeNames_ph1, nodeNames_ph2, nodeNames_ph3, name_file_created, 'shortcircuit')
    
                            DSStext.Command = 'Visualize Currents Fault.F1'
                            if terminales3 != '':
                                DSStext.Command = 'Visualize Currents Fault.F2'
                            DSSprogress.PctProgress = 100
                        tfinal_short_circuit = time.time() - tinitial_short_circuit
                        DSSprogress.Close()
                        if not self.dlg.powerflow_snapshot.isChecked() and not self.dlg.powerflow_daily.isChecked() \
                                and not self.dlg.harmonics.isChecked() and not self.dlg.powerflow_yearly.isChecked():
                            QMessageBox.information(None, QCoreApplication.translate('dialog',
                                                                                     "Analisis de Corto Circuito Terminado"), \
                                                    QCoreApplication.translate('dialog', "Tiempo de simulacion: ") + str(
                                                        tfinal_short_circuit + time_common_for_all) + " s" + "\n" + \
                                                    QCoreApplication.translate('dialog',
                                                                               "Los archivos han sido guardados en: ") + output_folder)
                        os.system('clear')
    
                ###################################################################################
                ###############################      HARMONICS      ###############################
                ###################################################################################
                # harmonics simulation routine
                try:
                    if self.dlg.harmonics.isChecked():
                        tinitial_harmonics = time.time()
                        counter = 0
                        harmonic_num = self.dlg.lineEdit_harmonics.text().upper()
                        DSSprogress.Show()
                        DSSprogress.Caption = QCoreApplication.translate('progress', u'Load Allocation')
                        DSSprogress.PctProgress = 0
                        # input harmonic numbers reading                
                                                
                        if harmonic_num == "":
                            harmonic_num = "1-25" #si el usuario no introdujo un valor se establece el valor por defecto
                        harmonics, NumHarm = Harmonics.harmonic_choice(harmonic_num)                
        
                        # spectrum creation
                        if self.dlg.checkBox_harm_res.isChecked():
                            load_res_spectrum = self.dlg.lineEdit_res_spectrum.text()
                            if not load_res_spectrum:
                                DSSprogress.Close()
                                QMessageBox.information(None, QCoreApplication.translate('dialog', u"Información Requerida"),
                                                        QCoreApplication.translate('dialog',
                                                                                   u"Seleccione el espectro de armónicos residencial") + '\n')
                                return # exit on except
        
                        if self.dlg.checkBox_harm_gen.isChecked():
                            load_gen_spectrum = self.dlg.lineEdit_gen_spectrum.text()
                            if not load_gen_spectrum:
                                DSSprogress.Close()
                                QMessageBox.information(None, QCoreApplication.translate('dialog', u"Información Requerida"),
                                                        QCoreApplication.translate('dialog',
                                                                                   u"Seleccione el espectro de armónicos comercial e industrial") + '\n')
                                return # exit on except
        
                        if self.dlg.checkBox_harm_TMT.isChecked():
                            load_TMT_spectrum = self.dlg.lineEdit_TMT_spectrum.text()
                            if not load_TMT_spectrum:
                                DSSprogress.Close()
                                QMessageBox.information(None,
                                                        QCoreApplication.translate('dialog', u"Información Requerida"),
                                                        QCoreApplication.translate('dialog',
                                                                                   u"Seleccione el espectro de armónicos TMT") + '\n')
                                return  # exit on except   
        
                        selected_parameter = 0;
                        harm_parameter = type_parameters[selected_parameter]  # always use GMM
        
                        
                        ##################################
                        ##### HARMONIC SNAPSHOT ##########
                        ##################################
        
                        if self.dlg.snapshot.isChecked():
							# Date
                            harmonicdate = self.dlg.lineEdit_harmonic_date.text().upper()
                            harmonicdate = correct_date( load_curve_circuit, harmonicdate )
                            if not harmonicdate:
                                harmonicdate = auxfcns.selection_representative_day(load_curve_circuit, 'weekday')
                            harm_daily_date = str(harmonicdate.replace('/', ''))
							
                            harmonictime = self.dlg.lineEdit_harmonic_time.text().upper()
                            harmonictime = correct_hour( load_curve_circuit, harmonictime )
                            if not harmonictime:
                                harmonictime = def_time  # default: 18:00
        
                            h, m = harmonictime.split(':')  # round hour to 10min multiples
                            if list(m)[1] != '0':
                                if int(list(m)[1]) <= 5:
                                    m = list(m)[0] + str(0)
                                elif int(list(m)[1]) > 5:
                                    if int(list(m)[0]) < 5:
                                        m = list(m)[0] + str(int(list(m)[1]) + 1)
                                    else:
                                        h = list(h)[0] + str(int(list(h)[1]) + 1)
                                        if int(h) >= 24:
                                            h = '00'
                                        m = '00'
                                harmonictime = h + ':' + m
                            H_Time = [harmonictime]
        
                        ##################################
                        ######## HARMONIC DAILY ##########
                        ##################################            
        
                        if self.dlg.daily.isChecked():
							# Date
                            harmonicdate = self.dlg.lineEdit_harmonic_date_daily.text().upper()
                            harmonicdate = correct_date( load_curve_circuit, harmonicdate )
                            if not harmonicdate:
                                harmonicdate = auxfcns.selection_representative_day(load_curve_circuit, 'weekday')
            
                            harm_daily_date = str(harmonicdate.replace('/', ''))
                            
                            H_Time = []  ## time array from 00:00:00 to 23:50:00, 10 minutes step
                            for h in range(24):
                                for m in range(6):
                                    if m == 0:
                                        if h < 10:
                                            H_Time.append('0' + str(h) + ':00')
                                        else:
                                            H_Time.append(str(h) + ':00')
                                    else:
                                        if h < 10:
                                            H_Time.append('0' + str(h) + ':' + str(m * 10))
                                        else:
                                            H_Time.append(str(h) + ':' + str(m * 10))
        
                        ##################################
                        ######### HARMONIC SIM ###########
                        ##################################
                        # real P, Q from substation
                        DP_to_be_matched = []
                        DQ_to_be_matched = []
                        for ij in range(len(circuit_demand)):
                            temp_a = circuit_demand[ij][0]
                            if str(temp_a.replace('/', '')) == str(harm_daily_date.replace('/', '')):
                                DP_to_be_matched.append(circuit_demand[ij][2])
                                DQ_to_be_matched.append(circuit_demand[ij][3])
        
                        DSStext.Command = 'clear'  # Clean previous circuits
                        DSStext.Command = 'New Circuit.Circuito_Distribucion_Harmonics'  # Create a new circuit
                        DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Master file compilation
                        DSStext.Command = 'set mode=daily'  # Daily sim for load allocation
                        DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the transformer secondary side to monitor P and Q
                        DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation
                        DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time
                        if tx_modelling and self.substation != 'Auto':
                            DSStext.Command = line_tx_definition  # Edit/create substation transformer
        
                        # dist gen powers init
                        gen_powers = np.zeros(96)
                        gen_rpowers = np.zeros(96)
                        GenNames = DSScircuit.Generators.AllNames
                        PVNames = DSScircuit.PVSystems.AllNames
        
                        for t in range(96):
                            gen_p = 0
                            gen_q = 0
                            DSScircuit.Solution.Solve()
                            if GenNames[0] != 'NONE':  # existing generators powers
                                for i in GenNames:
                                    DSScircuit.setActiveElement('generator.' + i)
                                    p = DSScircuit.ActiveElement.Powers
                                    for w in range(0, len(p), 2):
                                        gen_p += -p[w]
                                        gen_q += -p[w + 1]
                                gen_powers[t] += gen_p
                                gen_rpowers[t] += gen_q
                            if PVNames[0] != 'NONE':  # existing PVSystems powers
                                for i in PVNames:
                                    DSScircuit.setActiveElement('PVSystem.' + i)
                                    p = DSScircuit.ActiveElement.Powers
                                    for w in range(0, len(p), 2):
                                        gen_p += -p[w]
                                        gen_q += -p[w + 1]
                                gen_powers[t] += gen_p
                                gen_rpowers[t] += gen_q
        
                        errorP = 0.03  # Maximum desired correction error for active power
                        errorQ = 0.1  # Maximum desired correction error for reactive power
                        max_it_correction = 10  # Maximum number of allowed iterations
                        study = 'daily'  # Study type for PQ_corrector
                        #DSStext.Command = 'batchedit storage..* enabled = no' # Preguntar Orlo
                        # load allocations algorithm
                        [DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_sim, kVAr_sim] = auxfcns.PQ_corrector(
                            DSSprogress, DSScircuit, DSStext, errorP, errorQ, max_it_correction,
                            DP_to_be_matched, DQ_to_be_matched, harm_daily_date, study, dir_network, tx_modelling, 1, firstLine,
                            self.substation, line_tx_definition, gen_powers, gen_rpowers)
                        DSSprogress.Close()
                        
                        if DSScircuit == None:
                            self.iface.messageBar().pushCritical("QGISRunOpenDSS", QCoreApplication.translate('dialog',
                                                                                                       u'Sucedió un error grave y no fue posible completar la operación'))  
                            return
                        
                        # interpolate from 96 pts to 144 pts
                        x_96 = np.arange(0, 96, 1)
                        x_144 = np.arange(0, 96, 0.67)              
                       
                        try:    
                            kW_sim_144 = np.interp(x_144, x_96, kW_sim)
                            kVAr_sim_144 = np.interp(x_144, x_96, kVAr_sim)          
        
                            PQ_Time = []  ## time array from 00:00:00 to 23:50:00, 10 minutes step
                            i = 0
                            for h in range(24):
                                for m in range(6):
                                    if m == 0:
                                        if h < 10:
                                            PQ_Time.append(['0' + str(h) + ':00', kW_sim_144[i], kVAr_sim_144[i]])
                                        else:
                                            PQ_Time.append([str(h) + ':00', kW_sim_144[i], kVAr_sim_144[i]])
                                    else:
                                        if h < 10:
                                            PQ_Time.append(['0' + str(h) + ':' + str(m * 10), kW_sim_144[i], kVAr_sim_144[i]])
                                        else:
                                            PQ_Time.append([str(h) + ':' + str(m * 10), kW_sim_144[i], kVAr_sim_144[i]])
                                    i += 1
                        except:
                            QMessageBox.information(None,
                                                    QCoreApplication.translate('dialog', u"Análisis de armónicos terminado"),
                                                    QCoreApplication.translate('dialog', "Potencia simulada es cero. \nFavor verifique lo siguiente: \n-Que la fecha de simulación sea correcta\n -Que todas las cargas se encuentren conectadas") )        
        
                        PVS = []
                        readerPV = []
        
                        # PV allocation algorithm
                        if self.dlg.PV.isChecked():
                            if self.gui_gd.dlg.checkBox_PV_res.isChecked():
                                PVS += auxfcns.PV_allocation(total_capacity, pv_information, name_file_created, "R",
                                                             pv_study_type)
        
                            if self.gui_gd.dlg.checkBox_PV_gen.isChecked():
                                PVS += auxfcns.PV_allocation(total_capacity_gen, pv_information_gen, name_file_created, "C",
                                                             pv_study_type)
        
                            if self.gui_gd.dlg.checkBox_PV_TMT.isChecked():
                                PVS += auxfcns.PV_allocation(total_capacity_TMT, pv_information_TMT, name_file_created, "I",
                                                             pv_study_type)
                        for pv in range(len(PVS)):
                            if PVS[pv] == "Error":  # exit on error
                                return
        
                        # results ouput folder                
                        final_output_folder = output_folder + '\\Harmonics\\'
                        if not os.path.exists(final_output_folder):
                            os.makedirs(final_output_folder)
        
                        if not self.dlg.transformer_modelling.isChecked():
                            tx_conn_low = 'delta'
                            line_tx_definition = ''
                        try:  # harmonic simulation
                            Harmonics.harmonic_sim(self, NumHarm, harmonics, harmonicdate, H_Time, name_file_created,
                                                   type_parameters,
                                                   harm_parameter, dir_network, created_files, harm_daily_date,
                                                   frequency, volt_nom, phase_angle, line_tx_definition, circuit_demand,
                                                   DSStext, DSSprogress,
                                                   DSScircuit, final_output_folder, counter, tx_conn_low, PVS, readerPV,
                                                   str_time, tx_active, firstLine, PQ_Time, self.substation)
                            tfinal_harmonics = time.time() - tinitial_harmonics
                            QMessageBox.information(None,
                                                    QCoreApplication.translate('dialog', u"Análisis de armónicos terminado"),
                                                    QCoreApplication.translate('dialog', "Tiempo de simulacion: ") + str(
                                                        tfinal_harmonics + time_common_for_all) + " s" + "\n" + \
                                                    QCoreApplication.translate('dialog',
                                                                               "Los archivos han sido guardados en: ") + "\n" + output_folder)
                        except UnboundLocalError as e:
                            # fix_print_with_import
                            print(e)
                            
                            exc_info = sys.exc_info()
                            print("\nError: ", exc_info )
                            print("*************************  Información detallada del error ********************")
                            
                            for tb in traceback.format_tb(sys.exc_info()[2]):
                                print(tb)
                                
                            QMessageBox.information(None, QCoreApplication.translate('dialog', u"Análisis de armónicos"),
                                                    QCoreApplication.translate('dialog',
                                                                               "No se ha seleccionado el tipo de estudio"))
                            os.system('clear')
                            return
                    
                except:            
                    DSSprogress.Close()
                    self.progress.close()
                    exc_info = sys.exc_info()
                    print("\nError: ", exc_info )
                    print("*************************  Información detallada del error ********************")
                    
                    for tb in traceback.format_tb(sys.exc_info()[2]):
                        print(tb)
                    
                    QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"),
                                        QCoreApplication.translate('dialog',
                                                                   u"No se finalizar la simulación de armónicos. Favor revisar el código de error en la consola"))                                                                                             
                    return # exit on except
                print( "Simulations end" )
        except: 
            exc_info = sys.exc_info()
            print("\nError: ", exc_info )
            print("*************************  Información detallada del error ********************")
            
            for tb in traceback.format_tb(sys.exc_info()[2]):
                print(tb)
            
            QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"),
                                QCoreApplication.translate('dialog',
                                                           u"Error crítico. Favor revisar el código de error en la consola"))                                                                                             
            return # exit on except

