# -*- coding: utf-8 -*-
from __future__ import absolute_import
from builtins import str
from builtins import range
import os
from qgis.PyQt.QtCore import QSettings, QTranslator, qVersion, QCoreApplication, QObject
from qgis.PyQt.QtWidgets import QAction, QFileDialog, QMessageBox, QDialog
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt import uic
from . import resources  # Initialize Qt resources from file resources.py
from .QGIS_run_OpenDSS_dialog import QGISrunOpenDSSDialog  # Import the code for the dialog
import numpy as np  # import random
from . import auxiliary_functions as auxfcns  # File to define some usefull class/functions in python to do specific tasks
from matplotlib import pyplot as plt
import matplotlib
import matplotlib.patches as mpatches
import datetime as dt
import matplotlib.dates as mdates
import xlrd
import random
import collections
import time
import csv
from qgis.core import QgsProject
import glob
import re
from PyQt5.QtGui import *
from .QGIS_run_OpenDSS_transformer import Ui_Transformer
from .QGIS_run_OpenDSS_progress import Ui_Progress

matplotlib.rcParams.update({'font.size': 12})

def harmonic_sim(self, NumHarm, harmonics, harmonicdate, H_Time, name_file_created, type_parameters,
                 harm_parameter, dir_network, created_files, harm_daily_date,
                 frequency, volt_nom, phase_angle, line_tx_definition, circuit_demand, DSStext, DSSprogress,
                 DSScircuit, final_output_folder, counter, tx_conn_low, PVS, readerPV, str_time, tx_active, firstLine,
                 PQ_Time, substation):
    print( "firstLine = ", firstLine )
					 
    max_it = 10
    reader = []
    self.progress.show()
    self.progress.progressBar.setValue(0)

    # .xls file reader
    if self.dlg.checkBox_harm_res.isChecked():
        load_res_spectrum = self.dlg.lineEdit_res_spectrum.text()
        xlsfile_res = xlrd.open_workbook(load_res_spectrum, 'r')
    if self.dlg.checkBox_harm_gen.isChecked():
        load_gen_spectrum = self.dlg.lineEdit_gen_spectrum.text()
        xlsfile_gen = xlrd.open_workbook(load_gen_spectrum, 'r')
    if self.dlg.checkBox_harm_TMT.isChecked():
        load_TMT_spectrum = self.dlg.lineEdit_TMT_spectrum.text()
        xlsfile_TMT = xlrd.open_workbook(load_TMT_spectrum, 'r')

    for T in range(len(H_Time)):

        self.progress.progressBar.setRange(0, len(H_Time))

        # residential harmonic spectra creation
        file_object = open(dir_network + '/' + name_file_created.split('_')[0] + '_Spectra.dss', 'w')
        file_object.write('Spectrum.default.numharm=1 \n')  # clean all spectrums

        daily_strtime = str(harmonicdate.replace('/', '') + H_Time[T].replace(':', ''))
        hora_sec = H_Time[T].split(':')

        # residential harmonic spectra creation
        if self.dlg.checkBox_harm_res.isChecked():
            xls_mean = xlsfile_res.sheet_by_name('pdf_mean')
            xls_covar = xlsfile_res.sheet_by_name('pdf_covar')
            xls_weigth = xlsfile_res.sheet_by_name('pdf_weigth')
            for i in range(10000):
                mag, angle = gmm(harmonics, H_Time[T], xls_mean, xls_covar, xls_weigth)  #gmm routine
                file_object.write('new Spectrum.spectrum_res_loads_' + str(
                    i + 1) + ' NumHarm=' + NumHarm + ' harmonic=(' + harmonics + ') %mag=(' + mag + ') angle=(' + angle + ') \n')
            # file_object.close()
        # general harmonic spectra creation
        if self.dlg.checkBox_harm_gen.isChecked():
            xls_mean = xlsfile_gen.sheet_by_name('pdf_mean')
            xls_covar = xlsfile_gen.sheet_by_name('pdf_covar')
            xls_weigth = xlsfile_gen.sheet_by_name('pdf_weigth')
            for i in range(10000):
                mag, angle = gmm(harmonics, H_Time[T], xls_mean, xls_covar, xls_weigth)  #gmm routine
                file_object.write('new Spectrum.spectrum_gen_loads_' + str(
                    i + 1) + ' NumHarm=' + NumHarm + ' harmonic=(' + harmonics + ') %mag=(' + mag + ') angle=(' + angle + ') \n')
            # file_object.close()
        # MV harmonic spectra creation
        if self.dlg.checkBox_harm_TMT.isChecked():
            xls_mean = xlsfile_TMT.sheet_by_name('pdf_mean')
            xls_covar = xlsfile_TMT.sheet_by_name('pdf_covar')
            xls_weigth = xlsfile_TMT.sheet_by_name('pdf_weigth')
            for i in range(10000):
                mag, angle = gmm(harmonics, H_Time[T], xls_mean, xls_covar, xls_weigth)  #gmm routine
                file_object.write('new Spectrum.spectrum_TMT_loads_' + str(
                    i + 1) + ' NumHarm=' + NumHarm + ' harmonic=(' + harmonics + ') %mag=(' + mag + ') angle=(' + angle + ') \n')
        file_object.close()


        # LoadsLV_mod.dss writing
        new_loads_file = open(dir_network + '/' + name_file_created.split('_')[0] + '_LoadsLV_mod.dss', 'w')
        with open(dir_network + '/' + name_file_created.split('_')[0] + '_LoadsLV.dss', 'r') as temp:
            num_lines = temp.readlines()
            for line in num_lines:
                #i += 1
                if 'class=R' in line:
                    if self.dlg.checkBox_harm_res.isChecked():
                        try:
                            l, g = line.split('!kWh=')
                        except ValueError:
                            l = line
                            g = '\n'
                        linetemp = l + ' Spectrum = spectrum_res_loads_' + str(random.randint(1, 10000)) + ' !kWh=' + g
                        new_loads_file.write(linetemp)
                    else:
                        new_loads_file.write(line)
                elif 'class=C' or 'class=I' in line:
                    if self.dlg.checkBox_harm_gen.isChecked():
                        try:
                            l, g = line.split('!kWh=')
                        except ValueError:
                            l = line
                            g = '\n'
                        linetemp = l + ' Spectrum = spectrum_gen_loads_' + str(random.randint(1, 10000)) + ' !kWh=' + g
                        new_loads_file.write(linetemp)
                    else:
                        new_loads_file.write(line)
                elif 'class=TMT' in line:
                    if self.dlg.checkBox_harm_TMT.isChecked():
                        try:
                            l, g = line.split('!kWh=')
                        except ValueError:
                            l = line
                            g = '\n'
                        linetemp = l + ' Spectrum = spectrum_TMT_loads_' + str(random.randint(1, 10000)) + ' !kWh=' + g
                        new_loads_file.write(linetemp)
                    else:
                        new_loads_file.write(line)
        new_loads_file.close()
        load_files = created_files.replace(name_file_created.split('_')[0] + '_LoadsLV',
                                           name_file_created.split('_')[0] + '_LoadsLV_mod')

        # master file modification
        mlines = []
        with open(dir_network + '/Master.dss', 'r') as mt:
            num_lines = mt.readlines()
            for line in num_lines:
                if '_LoadsLV' in line:
                    mlines.append('redirect ' + name_file_created.split('_')[0] + '_Spectra.dss' + '\n')
                    line = line.replace('_LoadsLV.dss\n', '_LoadsLV_mod.dss\n')
                mlines.append(line)
        with open(dir_network + '/Master.dss', 'w') as mt:
            for line in mlines:
                mt.write(line)

        DSStext.Command = 'clear'  # clean old circuits
        DSStext.Command = 'New Circuit.Circuito_Distribucion_Harmonics'  # Create new circuit
        DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # compile master file
        DSStext.Command = 'set mode=daily'  # first simulation as daily
        DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
        DSStext.Command = 'Set number= 1'  # number of steps for each solve
        DSStext.Command = 'Set stepsize=10m'  # Stepsize of the simulation
        DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
        if self.dlg.transformer_modelling.isChecked() and substation != 'Auto':
            DSStext.Command = line_tx_definition

        for i in range(144):
            if PQ_Time[i][0] == H_Time[T]:
                DSStext.Command = 'batchedit load..* kW=' + str(PQ_Time[i][1])  # kW corrector
                DSStext.Command = 'batchedit load..* kVAr=' + str(PQ_Time[i][2])  # kVAr corrector

        # assign spectra to pv systems
        if self.dlg.PV.isChecked():
            hpv = harmonics.split(',')
            for sv in range(1, 101):
                pv_spectrum = []

                for i in range(len(hpv) - 1):
                    s = int(hpv[i])
                    if s > 1:
                        if s < 11:
                            if s % 2 == 0:
                                pv_spectrum.append(np.round(random.uniform(0, 4 * 0.25) / s, 2))
                            else:
                                pv_spectrum.append(np.round(random.uniform(1, 4) / s, 2))
                        elif s == 11 or s < 17:
                            if s % 2 == 0:
                                pv_spectrum.append(np.round(random.uniform(0, 2 * 0.25) / s, 2))
                            else:
                                pv_spectrum.append(np.round(random.uniform(0, 2) / s, 2))
                        elif s == 17 or s < 23:
                            if s % 2 == 0:
                                pv_spectrum.append(np.round(random.uniform(0, 1.5 * 0.25) / s, 2))
                            else:
                                pv_spectrum.append(np.round(random.uniform(0, 1.5) / s, 2))
                        elif s == 23 or s < 35:
                            if s % 2 == 0:
                                pv_spectrum.append(np.round(random.uniform(0, 0.6 * 0.25) / s, 2))
                            else:
                                pv_spectrum.append(np.round(random.uniform(0, 0.6) / s, 2))
                        elif s >= 35:
                            if s % 2 == 0:
                                pv_spectrum.append(np.round(random.uniform(0, 0.3 * 0.25) / s, 2))
                            else:
                                pv_spectrum.append(np.round(random.uniform(0, 0.3) / s, 2))
                    magPV = str(pv_spectrum).split('[')[1].split(']')[0]
                    spectrum = 'new Spectrum.spectrum_PV' + str(
                        sv) + ' NumHarm=' + NumHarm + ' harmonic=(' + harmonics + ') %mag=(100,' + magPV + ') angle=(' + angle + ')'
                    DSStext.Command = spectrum
            for pv in PVS:

                DSStext.Command = pv.split('!')[0] + ' spectrum=spectrum_PV'+ str(random.randint(1,100)) + '!' + pv.split('!')[1]
        if int(hora_sec[0]) < 5 or int(hora_sec[0]) > 18:
            DSStext.Command = 'batchedit PVSystem..* enabled=no'

        DSStext.Command = 'Solve'
        DSStext.Command = 'set mode=Harmonic'
        # Monitor in the first line to monitor P and Q,  residual=yes show neutral voltages/currents
        DSStext.Command = 'Edit Monitor.HVMV_PQ_vs_Time Mode=0 residual=yes'
        DSStext.Command = 'Solve'
        DSStext.Command = 'set datapath=' + final_output_folder

        DSStext.Command = 'export mon HVMV_PQ_vs_Time'
        
        
        #Busca archivo creado por DSS con nombre igual o similar a Circuito_Distribucion_Harmonics_Mon_hvmv_pq_vs_time.csv     
        import glob
        extension = 'csv'
        os.chdir( final_output_folder )
        result = glob.glob('*.{}'.format(extension))
        
        #Si no hay ningún csv en el directorio retorna
        if len(result) == 0:
            return 0
        
        name_file = "Circuito_Distribucion_Harmonics_Mon_hvmv_pq_vs_time"
        mon_file = ""
        
        #Busca un archivo con nombre similar a name_file dentro de la lista de archivos en el directorio
        for nom_file in result:
            nom_file = str( nom_file )
            if name_file in nom_file:
                mon_file = nom_file
                break                        
        
        results_file = open(os.path.join(final_output_folder, mon_file), 'r')
        reader.append(list(csv.reader(results_file)))
        results_file.close()
        DSScircuit.Monitors.ResetAll()
        counter += 1
        self.progress.progressBar.setValue(counter)

    self.progress.close()

    #########Graphics###########
    if self.dlg.snapshot.isChecked():
        ymax = harm_snapshot(reader, harmonics, tx_conn_low, daily_strtime, final_output_folder, str_time, 0)
    if self.dlg.daily.isChecked():
        imax = harm_daily(reader, harmonics, tx_conn_low, harm_daily_date, final_output_folder, str_time, 0)

    if self.dlg.PV.isChecked():
        if self.dlg.snapshot.isChecked():
            harm_snapshot(readerPV, harmonics, tx_conn_low, 'wPV_' + daily_strtime, final_output_folder, str_time, ymax)
        if self.dlg.daily.isChecked():
            harm_daily(readerPV, harmonics, tx_conn_low, 'wPV_' + harm_daily_date, final_output_folder, str_time, imax)


def weighted_choice(choices):  #
    total = sum(w for c, w in choices)
    r = random.uniform(0, total)
    upto = 0
    for c, w in choices:
        if upto + w >= r:
            return c
        upto += w
    assert False


def harmonic_choice(harmonic_num):
    harmonics = '1,'
    if ',' in harmonic_num or '-' in harmonic_num or ';' in harmonic_num:
        commas = harmonic_num.split(',') if ',' in harmonic_num else harmonic_num.split(';')
        for i in range(len(commas)):
            if commas[i] != '1':
                if '-' in commas[i]:
                    a, b = commas[i].split('-')
                    for j in range(int(a), int(b) + 1):
                        harmonics += str(j) + ','
                else:
                    harmonics += str(commas[i]) + ','
    elif ':' in harmonic_num:
        steps = harmonic_num.split(':')
        if str(steps[0]) is not '1':
            harmonics += str(steps[0]) + ','
        ini = int(steps[0])
        for i in range(int(steps[0]), 50):
            ini = ini + int(steps[1])
            if ini > int(steps[2]):
                break
            harmonics += str(ini) + ','
    elif ',' not in harmonic_num or '-' not in harmonic_num or ':' not in harmonic_num:
        for j in range(2, int(harmonic_num) + 1):
            harmonics += str(j) + ','
    NumHarm = str(len(harmonics.split(',')) - 1)

    return harmonics, NumHarm


def gmm(harmonics, H_Time, xls_mean, xls_covar, xls_weigth):
    mag = '100,'
    angle = '0,'
    angleMean = [0, 12.86, 0, 6.43, 0, 0, 0, -10.71, 0, -4.29, 0, -38.57, 0, -111.43, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    angleStd = [0.1, 6.43, 0.1, 5.71, 0.1, 7.14, 0.1, 11.43, 0.1, 7.14, 0.1, 18.57, 0.1, 13.57, 0.1, 0.1, 0.1, 0.1, 0.1,
                0.1, 0.1, 0.1, 0.1, 0.1]
    for p in range(1, 145):
        if H_Time in xls_mean.cell(p, 0).value:
            for harm in harmonics.split(','):
                for h in range(26):
                    try:
                        if xls_mean.cell(0, h).value == 'I_h'+harm:
                            # print(xls_mean.cell(0, h).value)
                            means_ = [float(x) for x in
                                      xls_mean.cell(p, h).value.replace("[", "").replace("]", "").replace("\n", "").replace(
                                          " -", "  -").replace(",", "").split(' ') if x is not u'']
                            covars_ = [float(x) for x in
                                       xls_covar.cell(p, h).value.replace("[", "").replace("]", "").replace("\n", "").replace(
                                           " -", "  -").replace(",", "").split(' ') if x is not u'']
                            weights_ = [float(x) for x in
                                        xls_weigth.cell(p, h).value.replace("[", "").replace("]", "").replace("\n","").replace(
                                            " -", "  -").replace(",", "").split(' ') if x is not u'']
                            choices = []
                            for wc in range(len(weights_)):
                                choices.append([wc, weights_[wc]])
                            w = weighted_choice(choices)
                            sp = -1
                            while sp < 0:
                                if means_[w] == 0 or covars_[w] == 0:
                                    sp = float(0)
                                else:
                                    sp = np.round(np.random.normal(means_[w], np.sqrt(covars_[w])), 2)

                            mag = mag + str(sp) + ','
                            angle = angle + str(np.round(np.random.normal(angleMean[h - 2], angleStd[h - 2]), 2)) + ','
                            break
                    except IndexError:
                        pass
            # break
    return mag, angle


def mstd(harmonics, H_Time, mean, std):
    mag = '100,'
    angle = '0,'
    for p in range(1, 145):
        if mean.cell(p, 0).value == H_Time:
            for h in range(2, 26):
                if str(h) in harmonics.split(','):
                    # busqueda de datos para creacion de espectro
                    Ih = float(mean.cell(p, h).value)
                    Ih_std = float(std.cell(p, h).value)
                    sp = -1
                    while sp < 0:
                        sp = random.normalvariate(Ih, Ih_std)
                    mag = mag + str(sp) + ','
                    angle = angle + str(0) + ','
            break
    return mag, angle


def harm_snapshot(reader, harmonics, tx_conn_low, daily_strtime, final_output_folder, str_time, ymax):
    i = 0
    Freq = []
    Harmonic = []
    V1 = []
    VAngle1 = []
    V2 = []
    VAngle2 = []
    V3 = []
    VAngle3 = []
    V4 = []
    VAngle4 = []
    I1 = []
    IAngle1 = []
    I2 = []
    IAngle2 = []
    I3 = []
    IAngle3 = []
    I4 = []
    IAngle4 = []
    results_dict = collections.OrderedDict()
    # print reader[i]
    idxFreq = None
    idxHarmonic = None
    idxV1 = None
    idxVAngle1 = None
    idxV2 = None
    idxVAngle2 = None
    idxV3 = None
    idxVAngle3 = None
    idxV4 = None
    idxVAngle4 = None
    idxI1 = None
    idxIAngle1 = None
    idxI2 = None
    idxIAngle2 = None
    idxI3 = None
    idxIAngle3 = None
    idxI4 = None
    idxIAngle4 = None

    for n in range(len(reader[0][0])):
        if reader[0][0][n] == "Freq":
            idxFreq = n
        elif reader[0][0][n] == " Harmonic":
            idxHarmonic = n
        elif reader[0][0][n] == " V1":
            idxV1 = n
        elif reader[0][0][n] == " VAngle1":
            idxVAngle1 = n
        elif reader[0][0][n] == " V2":
            idxV2 = n
        elif reader[0][0][n] == " VAngle2":
            idxVAngle2 = n
        elif reader[0][0][n] == " V3":
            idxV3 = n
        elif reader[0][0][n] == " VAngle3":
            idxVAngle3 = n
        elif reader[0][0][n] == " VN":
            idxV4 = n
        elif reader[0][0][n] == " VNAngle":
            idxVAngle4 = n
        elif reader[0][0][n] == " I1":
            idxI1 = n
        elif reader[0][0][n] == " IAngle1":
            idxIAngle1 = n
        elif reader[0][0][n] == " I2":
            idxI2 = n
        elif reader[0][0][n] == " IAngle2":
            idxIAngle2 = n
        elif reader[0][0][n] == " I3":
            idxI3 = n
        elif reader[0][0][n] == " IAngle3":
            idxIAngle3 = n
        elif reader[0][0][n] == " IN":
            idxI4 = n
        elif reader[0][0][n] == " INAngle":
            idxIAngle4 = n

    for n in range(1, len(reader[i])):
        try:
            Freq.append(float(reader[i][n][idxFreq]))
        except TypeError:
            pass
        try:
            Harmonic.append(float(reader[i][n][idxHarmonic]))
        except TypeError:
            pass
        try:
            V1.append(float(reader[i][n][idxV1]))
        except TypeError:
            pass
        try:
            VAngle1.append(float(reader[i][n][idxVAngle1]))
        except TypeError:
            pass
        try:
            V2.append(float(reader[i][n][idxV2]))
        except TypeError:
            pass
        try:
            VAngle2.append(float(reader[i][n][idxVAngle2]))
        except TypeError:
            pass
        try:
            V3.append(float(reader[i][n][idxV3]))
        except TypeError:
            pass
        try:
            VAngle3.append(float(reader[i][n][idxVAngle3]))
        except TypeError:
            pass
        try:
            V4.append(float(reader[i][n][idxV4]))
        except TypeError:
            pass
        try:
            VAngle4.append(float(reader[i][n][idxVAngle4]))
        except TypeError:
            pass
        try:
            I1.append(float(reader[i][n][idxI1]))
        except TypeError:
            pass
        try:
            IAngle1.append(float(reader[i][n][idxIAngle1]))
        except TypeError:
            pass
        try:
            I2.append(float(reader[i][n][idxI2]))
        except TypeError:
            pass
        try:
            IAngle2.append(float(reader[i][n][idxIAngle2]))
        except TypeError:
            pass
        try:
            I3.append(float(reader[i][n][idxI3]))
        except TypeError:
            pass
        try:
            IAngle3.append(float(reader[i][n][idxIAngle3]))
        except TypeError:
            pass
        try:
            I4.append(float(reader[i][n][idxI4]))
        except TypeError:
            pass
        try:
            IAngle4.append(float(reader[i][n][idxIAngle4]))
        except TypeError:
            pass

    results_dict['Freq'] = Freq
    results_dict['Harmonic'] = Harmonic
    results_dict['V1'] = V1
    results_dict['VAngle1'] = VAngle1
    results_dict['V2'] = V2
    results_dict['VAngle2'] = VAngle2
    results_dict['V3'] = V3
    results_dict['VAngle3'] = VAngle3
    results_dict['V4'] = V4
    results_dict['VAngle4'] = VAngle4
    results_dict['I1'] = I1
    results_dict['IAngle1'] = IAngle1
    results_dict['I2'] = I2
    results_dict['IAngle2'] = IAngle2
    results_dict['I3'] = I3
    results_dict['IAngle3'] = IAngle3
    results_dict['I4'] = I4
    results_dict['IAngle4'] = IAngle4

    I1p = [x * 100 / I1[0] for x in I1]
    I2p = [x * 100 / I2[0] for x in I2]
    I3p = [x * 100 / I3[0] for x in I3]
    I1p[0] = 100
    I2p[0] = 100
    I3p[0] = 100
    results_dict['I1p'] = I1p
    results_dict['I2p'] = I2p
    results_dict['I3p'] = I3p

    if tx_conn_low == 'wye':
        I4p = [x * 100 / I4[0] for x in I4]
        I4p[0] = 100
        results_dict['I4p'] = I4p

        ##calculo de THDv y THDi #hay que grabarlo en csv y calcular el TDDi
    thdi1 = ((np.sqrt(np.sum(np.square(I1[1:len(I1)])))) / I1[0]) * 100
    thdi2 = ((np.sqrt(np.sum(np.square(I2[1:len(I2)])))) / I2[0]) * 100
    thdi3 = ((np.sqrt(np.sum(np.square(I3[1:len(I3)])))) / I3[0]) * 100

    thdv1 = ((np.sqrt(np.sum(np.square(V1[1:len(V1)])))) / V1[0]) * 100
    thdv2 = ((np.sqrt(np.sum(np.square(V2[1:len(V2)])))) / V2[0]) * 100
    thdv3 = ((np.sqrt(np.sum(np.square(V3[1:len(V3)])))) / V3[0]) * 100
    if tx_conn_low == 'wye':
        thdi4 = ((np.sqrt(np.sum(np.square(I4[1:len(I4)])))) / I4[0]) * 100
        thdv4 = ((np.sqrt(np.sum(np.square(V4[1:len(V4)])))) / V4[0]) * 100

    #######################################
    ############ currents RMS #############
    #######################################

    fig, ax = plt.subplots(figsize=[10, 6])
    ## the data
    N = len(results_dict['I1'])
    ph1 = results_dict['I1']
    ph2 = results_dict['I2']
    ph3 = results_dict['I3']
    if tx_conn_low == 'wye':
        ph4 = results_dict['I4']

    ## necessary variables
    ind = np.arange(N)  # the x locations for the groups
    width = 0.2  # the width of the bars

    ## the bars
    rects1 = ax.bar(ind, ph1, width, color='#e70505')
    rects2 = ax.bar(ind + width, ph2, width, color='yellow')
    rects3 = ax.bar(ind + 2 * width, ph3, width, color='#305bf8')
    if tx_conn_low == 'wye':
        rects4 = ax.bar(ind + 3 * width, ph4, width, color='gray')

    if ymax == 0:
        ymax = max(I1[0], I2[0], I3[0]) + 5
    # axes and labels
    ax.set_xlim(-width, len(ind) + width)
    ax.set_ylim(0, ymax)
    ax.set_ylabel(QCoreApplication.translate('graphs', u'Corriente (A)'))
    ax.set_title(QCoreApplication.translate('graphs', u'Corrientes armónicas (RMS)'))
    xTickMarks = ['h' + str(harmonics.split(',')[i]) for i in range(0, N)]
    ax.set_xticks(ind + width)
    xtickNames = ax.set_xticklabels(xTickMarks)
    plt.setp(xtickNames, rotation=45, fontsize=10)

    ## add a legend
    empty = mpatches.Patch(color='white')
    if tx_conn_low == 'wye':
        plt.legend((rects1[0], rects2[0], rects3[0], rects4[0], empty, empty, empty, empty, empty, empty), \
                   ('I1', 'I2', 'I3', 'N'))
    else:  # TODO verificar esto
        plt.legend((rects1[0], rects2[0], rects3[0], empty, empty, empty, empty, empty), \
                   ('I1', 'I2', 'I3'))

    plt.grid()
    plt.savefig(final_output_folder + 'Harmonic_currents_RMS_' + daily_strtime + '_' + str_time + '.pdf', format='pdf',
                dpi=6000)
    plt.show()

    #######################################
    ######### currents percentage##########
    #######################################
    fig, ax = plt.subplots(figsize=[10, 6])
    ## the data
    N = len(results_dict['I1p'])
    ph1 = results_dict['I1p']
    ph2 = results_dict['I2p']
    ph3 = results_dict['I3p']
    if tx_conn_low == 'wye':
        ph4 = results_dict['I4p']

    ## necessary variables
    ind = np.arange(N)  # the x locations for the groups
    width = 0.2  # the width of the bars

    ## the bars
    rects1 = ax.bar(ind, ph1, width, color='#e70505')

    rects2 = ax.bar(ind + width, ph2, width, color='yellow')

    rects3 = ax.bar(ind + 2 * width, ph3, width, color='#305bf8')
    if tx_conn_low == 'wye':
        rects4 = ax.bar(ind + 3 * width, ph4, width, color='gray')


    # axes and labels
    ax.set_xlim(-width, len(ind) + width)
    ax.set_ylim(0, 110)
    ax.set_ylabel(QCoreApplication.translate('graphs', u'Corriente (%)'))
    ax.set_title(QCoreApplication.translate('graphs', u'Corrientes armónicas %'))
    xTickMarks = ['h' + str(harmonics.split(',')[i]) for i in range(0, N)]
    ax.set_xticks(ind + width)
    xtickNames = ax.set_xticklabels(xTickMarks)
    plt.setp(xtickNames, rotation=45, fontsize=10)

    ## add a legend
    empty = mpatches.Patch(color='white')
    if tx_conn_low == 'wye':
        plt.legend((rects1[0], rects2[0], rects3[0], rects4[0], empty, empty, empty, empty, empty, empty), \
                   ('I1', 'I2', 'I3', 'N'))
    else:  # TODO verificar esto
        plt.legend((rects1[0], rects2[0], rects3[0], empty, empty, empty, empty, empty), \
                   ('I1', 'I2', 'I3'))

    plt.grid()
    plt.savefig(final_output_folder + 'Harmonic_currents_percentage_' + daily_strtime + '_' + str_time + '.pdf',
                format='pdf', dpi=6000)
    plt.show()

    ##### CSV #######
    with open(final_output_folder + '\snapshot_' + daily_strtime + '_' + str_time + '.csv',
              'w') as f:  # Just use 'w' mode in 3.x
        w = csv.writer(f)
        w.writerow(['RMS'])
        w.writerow(['Phase'] + xTickMarks)
        w.writerow(['V1'] + results_dict['V1'])
        w.writerow(['V2'] + results_dict['V2'])
        w.writerow(['V3'] + results_dict['V3'])
        if tx_conn_low == 'wye':
            w.writerow(['VN'] + results_dict['V4'])
        w.writerow(['I1'] + results_dict['I1'])
        w.writerow(['I2'] + results_dict['I2'])
        w.writerow(['I3'] + results_dict['I3'])
        if tx_conn_low == 'wye':
            w.writerow(['IN'] + results_dict['I4'])
        w.writerow(['Percentage of fundamental'])
        w.writerow(['Phase'] + xTickMarks)
        w.writerow(['I1'] + results_dict['I1p'])
        w.writerow(['I2'] + results_dict['I2p'])
        w.writerow(['I3'] + results_dict['I3p'])
        if tx_conn_low == 'wye':
            w.writerow(['IN'] + results_dict['I4p'])

        w.writerow(['Total harmonic distortion'])
        w.writerow(['THDv1'] + [thdv1])
        w.writerow(['THDv2'] + [thdv2])
        w.writerow(['THDv3'] + [thdv3])
        if tx_conn_low == 'wye':
            w.writerow(['THDv4'] + [thdv4])
        w.writerow(['THDi1'] + [thdi1])
        w.writerow(['THDi2'] + [thdi2])
        w.writerow(['THDi3'] + [thdi3])
        if tx_conn_low == 'wye':
            w.writerow(['THDi4'] + [thdi4])

    return ymax


def harm_daily(reader, harmonics, tx_conn_low, harm_daily_date, final_output_folder, str_time, imax):
    h1 = []
    h2 = []
    h3 = []
    h4 = []
    thdi1 = []
    thdi2 = []
    thdi3 = []
    thdi4 = []
    thdv1 = []
    thdv2 = []
    thdv3 = []
    thdv4 = []
    idxFreq = None
    idxHarmonic = None
    idxV1 = None
    idxVAngle1 = None
    idxV2 = None
    idxVAngle2 = None
    idxV3 = None
    idxVAngle3 = None
    idxV4 = None
    idxVAngle4 = None
    idxI1 = None
    idxIAngle1 = None
    idxI2 = None
    idxIAngle2 = None
    idxI3 = None
    idxIAngle3 = None
    idxI4 = None
    idxIAngle4 = None

    for n in range(len(reader[0][0])):
        if reader[0][0][n] == "Freq":
            idxFreq = n
        elif reader[0][0][n] == " Harmonic":
            idxHarmonic = n
        elif reader[0][0][n] == " V1":
            idxV1 = n
        elif reader[0][0][n] == " VAngle1":
            idxVAngle1 = n
        elif reader[0][0][n] == " V2":
            idxV2 = n
        elif reader[0][0][n] == " VAngle2":
            idxVAngle2 = n
        elif reader[0][0][n] == " V3":
            idxV3 = n
        elif reader[0][0][n] == " VAngle3":
            idxVAngle3 = n
        elif reader[0][0][n] == " VN":
            idxV4 = n
        elif reader[0][0][n] == " VNAngle":
            idxVAngle4 = n
        elif reader[0][0][n] == " I1":
            idxI1 = n
        elif reader[0][0][n] == " IAngle1":
            idxIAngle1 = n
        elif reader[0][0][n] == " I2":
            idxI2 = n
        elif reader[0][0][n] == " IAngle2":
            idxIAngle2 = n
        elif reader[0][0][n] == " I3":
            idxI3 = n
        elif reader[0][0][n] == " IAngle3":
            idxIAngle3 = n
        elif reader[0][0][n] == " IN":
            idxI4 = n
        elif reader[0][0][n] == " INAngle":
            idxIAngle4 = n

    for i in range(len(reader)):  ##caso de prueba
        Freq = []
        Harmonic = []
        V1 = []
        VAngle1 = []
        V2 = []
        VAngle2 = []
        V3 = []
        VAngle3 = []
        V4 = []
        VAngle4 = []
        I1 = []
        IAngle1 = []
        I2 = []
        IAngle2 = []
        I3 = []
        IAngle3 = []
        I4 = []
        IAngle4 = []

        for n in range(1, len(reader[i])):
            try:
                Freq.append(float(reader[i][n][idxFreq]))
            except TypeError:
                pass
            try:
                Harmonic.append(float(reader[i][n][idxHarmonic]))
            except TypeError:
                pass
            try:
                V1.append(float(reader[i][n][idxV1]))
            except TypeError:
                pass
            try:
                VAngle1.append(float(reader[i][n][idxVAngle1]))
            except TypeError:
                pass
            try:
                V2.append(float(reader[i][n][idxV2]))
            except TypeError:
                pass
            try:
                VAngle2.append(float(reader[i][n][idxVAngle2]))
            except TypeError:
                pass
            try:
                V3.append(float(reader[i][n][idxV3]))
            except TypeError:
                pass
            try:
                VAngle3.append(float(reader[i][n][idxVAngle3]))
            except TypeError:
                pass
            try:
                V4.append(float(reader[i][n][idxV4]))
            except TypeError:
                pass
            try:
                VAngle4.append(float(reader[i][n][idxVAngle4]))
            except TypeError:
                pass
            try:
                I1.append(float(reader[i][n][idxI1]))
            except TypeError:
                pass
            try:
                IAngle1.append(float(reader[i][n][idxIAngle1]))
            except TypeError:
                pass
            try:
                I2.append(float(reader[i][n][idxI2]))
            except TypeError:
                pass
            try:
                IAngle2.append(float(reader[i][n][idxIAngle2]))
            except TypeError:
                pass
            try:
                I3.append(float(reader[i][n][idxI3]))
            except TypeError:
                pass
            try:
                IAngle3.append(float(reader[i][n][idxIAngle3]))
            except TypeError:
                pass
            try:
                I4.append(float(reader[i][n][idxI4]))
            except TypeError:
                pass
            try:
                IAngle4.append(float(reader[i][n][idxIAngle4]))
            except TypeError:
                pass

        # fundamental currents
        h1.append(I1[0])
        h2.append(I2[0])
        h3.append(I3[0])

        ##calculo de THDv y THDi #hay que grabarlo en csv y calcular el TDDi
        thdi1.append((np.sqrt(np.sum(np.square(I1[1:len(I1)])))) * 100)
        thdi2.append((np.sqrt(np.sum(np.square(I2[1:len(I2)])))) * 100)
        thdi3.append((np.sqrt(np.sum(np.square(I3[1:len(I3)])))) * 100)

        thdv1.append(((np.sqrt(np.sum(np.square(V1[1:len(V1)])))) / V1[0]) * 100)
        thdv2.append(((np.sqrt(np.sum(np.square(V2[1:len(V2)])))) / V2[0]) * 100)
        thdv3.append(((np.sqrt(np.sum(np.square(V3[1:len(V3)])))) / V3[0]) * 100)

        if tx_conn_low == 'wye':
            h4.append(I4[0])
            thdi4.append((np.sqrt(np.sum(np.square(I4[1:len(I4)])))) * 100)
            thdv4.append(((np.sqrt(np.sum(np.square(V4[1:len(V4)])))) / V4[0]) * 100)

    if imax == 0:
        imax = np.max([np.max(h1), np.max(h2), np.max(h3)])
    tddi1 = thdi1 / imax
    tddi2 = thdi2 / imax
    tddi3 = thdi3 / imax

    if tx_conn_low == 'wye':
        tddi4 = thdi4 / imax

    # Generate a series of dates (these are in matplotlib's internal date format)
    hours = mdates.drange(dt.datetime(2015, 1, 1), dt.datetime(2015, 1, 2), dt.timedelta(minutes=10))
    fig = plt.subplots(figsize=[16, 9])
    # Tell matplotlib to interpret the x-axis values as dates
    fmtr = mdates.DateFormatter("%H:%M")
    # need a handle to the current axes to manipulate it
    ax = plt.gca()
    ## plot
    ax.plot(hours, tddi1, lw=3, color='r', label='I1')
    ax.plot(hours, tddi2, lw=3, color='y', label='I2')
    ax.plot(hours, tddi3, lw=3, color='b', label='I3')
    if tx_conn_low == 'wye':
        ax.plot(hours, tddi4, lw=3, color='gray', label='In')
    # set this formatter to the axis
    ax.xaxis.set_major_formatter(fmtr)
    ax.set_xlabel(QCoreApplication.translate('graphs', u'Hora'))
    plt.ylabel('TDD (%)')  # plt.title('Perfil de carga, cliente '+str(i+1) +' - ' + string_dias[d])
    # Make space for and rotate the x-axis tick labels
    # fig.autofmt_xdate()
    plt.xticks(np.arange(min(hours), max(hours) + 0.04166666663 * 2, 0.04167 * 3))
    plt.xlim([hours[0], hours[len(hours) - 1]])
    plt.grid()
    plt.legend(bbox_to_anchor=(1, 1), loc='upper right', ncol=1)
    plt.savefig(final_output_folder + 'TDD_' + harm_daily_date + '_' + str_time + '.pdf', format='pdf', dpi=1000)
    plt.show()

    fig = plt.subplots(figsize=[16, 9])
    # Tell matplotlib to interpret the x-axis values as dates
    fmtr = mdates.DateFormatter("%H:%M")
    # need a handle to the current axes to manipulate it
    ax = plt.gca()
    ## plot
    ax.plot(hours, thdv1, lw=3, color='r', label='I1')
    ax.plot(hours, thdv2, lw=3, color='y', label='I2')
    ax.plot(hours, thdv3, lw=3, color='b', label='I3')
    # if tx_conn_low == 'wye':
    #     ax.plot(hours, thdv4, lw=3, color='gray', label='In')
    # set this formatter to the axis
    ax.xaxis.set_major_formatter(fmtr)
    ax.set_xlabel(QCoreApplication.translate('graphs', u'Hora'))
    plt.ylabel('THDv (%)')  # plt.title('Perfil de carga, cliente '+str(i+1) +' - ' + string_dias[d])
    # Make space for and rotate the x-axis tick labels
    # fig.autofmt_xdate()
    plt.xticks(np.arange(min(hours), max(hours) + 0.04166666663 * 2, 0.04167 * 3))
    plt.xlim([hours[0], hours[len(hours) - 1]])
    plt.grid()
    plt.legend(bbox_to_anchor=(1, 1), loc='upper right', ncol=1)
    plt.savefig(final_output_folder + 'THDv_' + harm_daily_date + '_' + str_time + '.pdf', format='pdf', dpi=1000)
    plt.show()

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

    with open(final_output_folder + '\TDD_' + harm_daily_date + '_' + str_time + '.csv',
              'w') as f:  # Just use 'w' mode in 3.x
        w = csv.writer(f)
        w.writerow(['Time,TDD1,TDD2,TDD3,TDDn'])
        for value in range(144):
            if tx_conn_low == 'wye':
                row = str(H_Time[value]) + ',' + str(tddi1[value]) + ',' + str(tddi2[value]) + ',' + str(
                    tddi3[value]) + ',' + str(tddi4[value])

            else:
                row = str(H_Time[value]) + ',' + str(tddi1[value]) + ',' + str(tddi2[value]) + ',' + str(
                    tddi3[value]) + ',0.0'
            w.writerow([row])

    with open(final_output_folder + '\THDv_' + harm_daily_date + '_' + str_time + '.csv',
              'w') as f:  # Just use 'w' mode in 3.x
        w = csv.writer(f)
        w.writerow(['Time,THD1,THD2,THD3,THDn'])
        for value in range(144):
            if tx_conn_low == 'wye':
                row = str(H_Time[value]) + ',' + str(thdv1[value]) + ',' + str(thdv2[value]) + ',' + str(
                    thdv3[value]) + ',' + str(thdv4[value])

            else:
                row = str(H_Time[value]) + ',' + str(thdv1[value]) + ',' + str(thdv2[value]) + ',' + str(
                    thdv3[value]) + ',0.0'
            w.writerow([row])

    return imax
