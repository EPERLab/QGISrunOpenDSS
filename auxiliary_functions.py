# -*- coding: utf-8 -*-
from __future__ import absolute_import
from builtins import next
from builtins import str
from builtins import range
import os
import numpy as np
import glob
import csv
import collections
from collections import OrderedDict
import datetime
import random
import subprocess
import datetime as dt
from PyQt5.QtGui import *  # Para desplegar mensajes, util para debugin
from PyQt5.QtCore import *
#from qgis.gui import QgsMessageBar  # Paquete requerido para desplegar mensajes en la ventana principal de QGIS.
from qgis.core import *
from qgis.utils import iface
from .Harmonics import weighted_choice
from PyQt5.QtWidgets import QMessageBox
# Para acelerar el COMM
# import win32com.client
# from win32com.client import makepy
import sys
import re
import traceback


def natural_sort(l):  # sort list in alphanumeric order
    try:
        convert = lambda text: int(text) if text.isdigit() else text.lower()
        alphanum_key = lambda key: [convert(c) for c in re.split('([0-9]+)', key)]
        return sorted(l, key=alphanum_key)
    except:
        exc_info = sys.exc_info()
        print("\nError: ", exc_info )
        print("*************************  Información detallada del error ********************")
        
        
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)      
        
        return 0


def resolve(name, basepath=None):  # choose external *.py file
    if not basepath:
        basepath = os.path.dirname(os.path.realpath(__file__))
    return os.path.join(basepath, name)


def installCOM():  # comtypes library installation
    try:
        import comtypes  # check if comtypes is already installed
        QMessageBox.information(None, QCoreApplication.translate('dialog', u"Instalación correcta"), \
                                QCoreApplication.translate('dialog', "COM ya se encuentra instalado"))
    except ImportError:
        pyFile = resolve("get-pip.py") # pip installation file
        try:
            p = subprocess.check_output(["python", pyFile])  # execute pip installation file
            # stdout, stderr = p.communicate()
        except subprocess.CalledProcessError as e:  # admin mode error
            # if p.returncode != 0:
            QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), \
                                 QCoreApplication.translate('dialog', "Para instalar el servidor COM, iniciar") + "\n" + \
                                 QCoreApplication.translate('dialog', "QGIS en modo administrador"))
            #return
        try:  # if pip installed, try to install comtypes via console
            i = subprocess.check_output(["python", "-m", "pip", "install", "comtypes"])
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), \
                                 QCoreApplication.translate('dialog', "Para instalar el servidor COM, iniciar") + "\n" + \
                                 QCoreApplication.translate('dialog', "QGIS en modo administrador"))


def SetUpCOMInterface():
    try:
        import comtypes.client as cc  # import comtypes library
    except ImportError:
        QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), \
                             QCoreApplication.translate('dialog', "No se encuentra instalado COM") + "\n" + \
                             QCoreApplication.translate('dialog',
                                                        u"Ejecutar QGIS como administrador para realizar su correcta instalación"))

    DSSobj = cc.CreateObject("OpenDSSEngine.DSS")  # create OpenDSS object
    DSSstart = DSSobj.Start(0)  # start the object
    DSStext = DSSobj.Text  # DSS command introduction
    DSScircuit = DSSobj.ActiveCircuit  # DSS active circuit calling
    DSSprogress = DSSobj.DSSProgress  # DSS progress object
    return DSSobj, DSSstart, DSStext, DSScircuit, DSSprogress  # , DSSobj_sim


def ReadLoadProfiles(self, perfilespath, dir_network, name_file_created):  # read loadshapes file and update path
    try:
        curves = [""]
        filenames = list()
        # open loadshapes file and read old path in first line
        with open(dir_network + '/' + name_file_created.split('_')[0] + '_LoadShapes.dss', 'r') as file_object:
            filenames = file_object.readlines()
            old_path = filenames[0].replace('!', '').replace('\\', '/').replace('\n', '')
            old_dir_profiles_gd= os.path.join(old_path, 'DG').replace('\\', '/')
            if 'curvas' in old_path:
                perfilespath = perfilespath.replace('profiles', 'curvas')
            dir_profiles_gd = os.path.join(perfilespath, 'DG')
            perfilespath = perfilespath.replace('\\', '/')
        file_object.closed
    
        # replace old path with new for all lines list
        for i in range(len(filenames)):
            filenames[i] = str(filenames[i]).replace('\\', '/').replace(old_path, perfilespath).replace('/', '\\')
    
        # open loadshapes file and replace old path with new for all lines in file
        with open(dir_network + '/' + name_file_created.split('_')[0] + '_LoadShapes.dss', 'w') as file_object:
            for loadshape in filenames:
                file_object.write(loadshape.replace('\\', '/').replace(old_path, perfilespath).replace('/', '\\'))
    
            # create P vs T curve if not in file
            if 'New XYCurve.MyPvsT npts=4 xarray=[.001 25 75 100] yarray=[1.2 1.0 0.8 0.6]\n' not in filenames:
                file_object.writelines(
                    'New XYCurve.MyPvsT npts=4 xarray=[.001 25 75 100] yarray=[1.2 1.0 0.8 0.6]\n')  # pmpp for 25C
                file_object.writelines('New XYCurve.MyEff npts=4 xarray=[.1 .2 .4 1.0] yarray=[.86 .9 .93 .97]\n')
    
            # add PV profile and PV temp shapes if not in file
            if self.dlg.PV.isChecked():
                try :
                    prof = open(old_dir_profiles_gd.replace('/', '\\') + '\\' + 'PVprofile.txt')
                    prof.close()
                except IOError:
                    curves = ["ERROR"]
                    QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"), \
                                         QCoreApplication.translate('dialog', u"No se encuentra perfil de sistemas fotovoltaicos") + "\n" + \
                                         QCoreApplication.translate('dialog',
                                                                    u"Verificar ubicación o existencia de archivos PVprofile.txt y PVtemp.txt"))
                    return curves
    
                st = 'New Loadshape.MyIrrad npts=96 minterval=15 csvfile=' + old_dir_profiles_gd.replace('/', '\\')\
                     + '\\' + 'PVprofile.txt\n'
                if st not in filenames:
                    file_object.write('New Loadshape.MyIrrad npts=96 minterval=15 csvfile=' + dir_profiles_gd.replace('/', '\\')
                                      + '\\' + 'PVprofile.txt\n')  # normalized with 1.045180145 kW/m2
                    file_object.write('New Tshape.MyTemp npts=96 minterval=15 csvfile=' + dir_profiles_gd.replace('/', '\\')
                                      + '\\' + 'PVtemp.txt\n')
        file_object.closed
    
        return curves
    
    except:
        exc_info = sys.exc_info()
        print("\nError: ", exc_info )
        print("*************************  Información detallada del error ********************")
        
        
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)      
        
        return [0]
    

def ExtractMonitorData(DSScircuit, name_monitor, Channel, Base):  # extract DSS monitor data
    DSScircuit.Monitors.Name = name_monitor
    n = DSScircuit.Monitors.SampleCount
    ba = DSScircuit.Monitors.ByteStream
    idata = np.array(ba[0:16], dtype=np.uint8).view(np.int32)
    nrec = idata[2]
    mode = idata[3]
    sdata = np.array(ba[272:], dtype=np.uint8).view(np.single)
    y = np.reshape(sdata, (n, nrec + 2))
    if Channel:
        sa = y[:, np.array(Channel) + 1] / Base
    else:
        sa = (3600 * y[:, 0] + y[1]) / Base
    return sa


def movingaverage(values, window):  # moving average
    weights = np.repeat(1.0, window) / window
    sma = np.convolve(values, weights, 'valid')

    return sma


def ensure_dir(f):  # create directory
    d = os.path.dirname(f)
    if not os.path.exists(d):
        os.makedirs(d)
        

def ExtractVoltageData(DSScircuit, V_buses, Base_V, t): #FUNCIÓN MODIFICADA
    a = np.array(DSScircuit.AllBusVmag) 
    b = np.array(Base_V['base'])   
    V_buses[t] = np.divide(a,b) 
    return V_buses
    

def PQ_corrector(DSSprogress, DSScircuit, DSStext, errorP, errorQ, max_it, P_to_be_matched, Q_to_be_matched, hora_sec,
                 study, dir_network, tx_active, yearly_steps, firstLine, substation, line_tx_definition, gen_powers, gen_rpowers):
    """
    Load allocation algorithm
    :param errorP: maximum P error desired
    :param errorQ: maximum Q error desired
    :param max_it: number of iteration permited
    :param P_to_be_matched: substation metered P
    :param Q_to_be_matched: substation metered Q
    :param hora_sec: date
    :param study: study type
    :param dir_network: network directory
    :param tx_active: substation transformer modelling boolean
    :param yearly_steps: yearly simulation step size on substation file
    :param firstLine: name of first line 
    :param substation: substation type
    :param line_tx_definition: transformer substation definition
    :param gen_powers: existing generators P
    :param gen_rpowers: existing generator Q
    :return: kW_corrector, kVAr_corrector arrays
    """
    try:

        # counter assignation according to study type
        if (study == 'snapshot') or (study == 'shortCircuit'):
            counter = 1
            DP_to_be_matched = [float(P_to_be_matched)]
            DQ_to_be_matched = [float(Q_to_be_matched)]
        if study == 'daily':
            counter = 96
            DP_to_be_matched = [float(x) for x in P_to_be_matched]
            DQ_to_be_matched = [float(x) for x in Q_to_be_matched]
        if study == 'yearly':
            counter = 35040 / yearly_steps
            DP_to_be_matched = [float(x) for x in P_to_be_matched]
            DQ_to_be_matched = [float(x) for x in Q_to_be_matched]
        
        if DP_to_be_matched == [] or DP_to_be_matched == [0] or DP_to_be_matched[0] == 0:
            return -1, 0, 0, 0, 0, 0, 0
        
    
        current_it = max_it # maximum iteration number        
        it_sec = 1
        prog = 20
    
        # first simulation powers
        DSScircuit.setActiveElement('line.' + firstLine)
        temp_powers = DSScircuit.ActiveElement.Powers
        temp_powersP = ExtractMonitorData(DSScircuit, 'HVMV_PQ_vs_Time', [1, 3, 5], 1)
        temp_powersQ = ExtractMonitorData(DSScircuit, 'HVMV_PQ_vs_Time', [2, 4, 6], 1)
        
        """
        print( "temp_powers = ", temp_powers )
        print( "temp_powersP = ", temp_powersP )
        print( "temp_powersQ = ", temp_powersQ )
        """
        
        # print temp_powersP
        temp_powersPtot = []
        temp_powersQtot = []
        
        for i in range( int(counter) ):
            temp_powersPtot.append(np.sum(temp_powersP[i][:]))
            temp_powersQtot.append(np.sum(temp_powersQ[i][:]))
        temp_powersP = temp_powersPtot
        temp_powersQ = temp_powersQtot
        
        kW_corrector = np.ones( int(counter) )  # kW corrector array init
        # pf_corrector = DPF_to_be_matched
        kVAr_corrector = np.ones( int(counter) )  # kVAr corrector array init
        errorP_i = np.ones( int(counter) ) * 100  # P error array init
        errorP_av = 100
        errorQ_i = np.ones( int(counter) ) * 100  # Q error array init
        errorQ_av = 100
    
        while errorQ_av > errorQ:  # and errorP_av > errorP:  # corrector iteration
            prog += 2
            DSSprogress.PctProgress = prog
            #( "counter = ", int(counter),  ", len(DQ_to_be_matched) = ", len(DQ_to_be_matched) )
            
            """
            if int(counter) != len(DQ_to_be_matched):            
               QMessageBox.critical(None, QCoreApplication.translate('dialog', "Error"),
                                            QCoreApplication.translate('dialog', u"El archivo de curvas de demanda no tiene la cantidad de datos necesaria para hacer el estudio solicitado "))
                
               return None, None, None, None, None, None, None
            """
            
    
            for t in range( int(counter) ): # corrector loop
                if DP_to_be_matched[t] != 0:                    
                    # P correction calc
                    kW_corrector[t] = kW_corrector[t] * ((DP_to_be_matched[t] + gen_powers[t]) / (temp_powersP[t] + gen_powers[t]))                    
                    
                    # Q correction calc
                    kVAr_corrector[t] = kVAr_corrector[t] * ( (DQ_to_be_matched[t] + gen_rpowers[t]) / (temp_powersQ[t] + gen_rpowers[t]) )
                    
                    
                else:
                    kW_corrector[t] = 0 
                    kVAr_corrector[t] = 0
    
            # if np.isnan(np.array(kW_corrector)).any():
            #     return
            # New simulation setting
            DSStext.Command = 'clear'  # erase old circuits
            DSStext.Command = 'New Circuit.Circuito_Distribucion_Daily'  # create new circuit
            DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compila el archivo master de OpenDSS
            if (study == 'daily') or (study == 'snapshot') or (study == 'shortCircuit'):
                DSStext.Command = 'Set mode = daily'  # Define el tipo de simulacion a realizar (diaria en este caso)
            if study == 'yearly':
                DSStext.Command = 'Set mode = yearly'  # Define el tipo de simulacion a realizar (diaria en este caso)
            DSStext.Command = 'Set number= 1'  # Define el numero de pasos a realizar
            if (study == 'snapshot') or (study == 'shortCircuit'):
                DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s
                DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
            if study == 'daily':
                DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s
                DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time
            if study == 'yearly':
                DSStext.Command = 'Set h=' + str(yearly_steps / 4) + 'h'
            DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
            if tx_active and substation != 'Auto':
                DSStext.Command = line_tx_definition            
    
            temp_powersP = []
            temp_powersQ = []
    
            for t in range( int(counter) ):  # corrector loop
    
                # DSStext.Command = 'batchedit load..* pf='+str(pf_corrector[t])
                
                DSStext.Command = 'batchedit load..* kW=' + str(kW_corrector[t])  # apply kW corrector previously calculated
                DSStext.Command = 'batchedit load..* kvar=' + str(kVAr_corrector[t])  # apply kVAr corrector previously calculated
                DSScircuit.Solution.Solve()
                # Result Query
                DSScircuit.setActiveElement('line.' + firstLine)
                temp_powers = DSScircuit.ActiveElement.Powers  # extract circuit powers
                temp_powersP.append((temp_powers[2] + temp_powers[4] + temp_powers[0]))
                temp_powersQ.append((temp_powers[3] + temp_powers[5] + temp_powers[1]))
    
                # calc P error
                errorP_i[t] = abs((DP_to_be_matched[t] - temp_powersP[t]) / DP_to_be_matched[t])
                # calc Q error
                errorQ_i[t] = abs((DQ_to_be_matched[t] - temp_powersQ[t]) / DQ_to_be_matched[t])
    
            errorP_av = np.max(errorP_i)
            errorQ_av = np.max(errorQ_i)
    
            current_it += - 1
            it_sec += 1
            if current_it == 0:  # exit condition
                break
    
        errorP_i = errorP_av
        errorQ_i = errorQ_av
    
        return DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_corrector, kVAr_corrector
    
    except:
        exc_info = sys.exc_info()
        print("\nError: ", exc_info )
        print("*************************  Información detallada del error ********************")
        
        
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)      
        
        return 0, 0, 0, 0, 0, 0, 0
    
    
def busSC():  # extract all buses names for SC analysis, add them to UI
    buses = []
    try:
        a = 0
        # look for buses layers
        layerMVNames = ['Bus_MV_Layer', 'Bus_MT_Layer']
        layerLVNames = ['Bus_LV_Layer', 'Bus_BT_Layer']
        i = 0
        layerMV = ''
        layerLV = ''
        while a == 0:
            try:  # activate layer by name
                layerMV = QgsProject.instance().mapLayersByName(layerMVNames[i])[0]
                layerLV = QgsProject.instance().mapLayersByName(layerLVNames[i])[0]
                a = 1
            except IndexError:
                i += 1
                if i > len(layerMVNames):
                    a = 1
        MV = layerMV.getFeatures()
        LV = layerLV.getFeatures()
        for bus in MV:
            buses.append(bus['BUS'])
        for bus in LV:
            buses.append(bus['BUS'])
        #buses = natural_sort(buses)  # sort buses in alphanumerical order
        #buses = list(reversed( buses ))
        buses_sort = ["",QCoreApplication.translate('SC', "Todas las barras MT"),
        QCoreApplication.translate('SC', 'Todas las barras BT'),
        QCoreApplication.translate('SC', 'Todas las barras MT y BT')] + buses  # add other options
    except IndexError:  # no bus shape
        QgsMessageLog.logMessage('No exite capa de buses', QCoreApplication.translate('dialog', 'Alerta Buses'),
                                 QgsMessageLog.WARNING)
    return buses_sort


def getAttributeIndex(self, aLayer, attrName):  # create attribute and get respective ID
    """Find the attribute index, adding a new Int column, if necessary"""
    if len(attrName) > 10 and aLayer.storageType() == 'ESRI Shapefile':
        self.iface.messageBar().pushMessage("Error",
                                            "For ESRI Shapefiles, the maximum length of any attribute name is 10. Please choose a shorter attribute name.",
                                            level=QgsMessageBar.CRITICAL)
        return -3
    AttrIdx = aLayer.dataProvider().fieldNameIndex(attrName)

    if AttrIdx == -1:  # attribute doesn't exist, so create it
        caps = aLayer.dataProvider().capabilities()
        if caps & QgsVectorDataProvider.AddAttributes:
            res = aLayer.dataProvider().addAttributes([QgsField(attrName, QVariant.String, len=8)])
            AttrIdx = aLayer.dataProvider().fieldNameIndex(attrName)
            aLayer.updateFields()
            if AttrIdx == -1:
                self.iface.messageBar().pushMessage("Error", "Failed to create attribute!",
                                                    level=QgsMessageBar.CRITICAL)
                return -1
        else:
            self.iface.messageBar().pushMessage("Error", "Failed to add attribute!", level=QgsMessageBar.CRITICAL)
            return -1
    else:
        pass
    return AttrIdx


#Función que obtiene el vector de los buses con sus tensiones asociadas para cada momento de la simulación
def ReadBusVolts(self, nodeVoltages_ph1, nodeVoltages_ph2, nodeVoltages_ph3, nodeNames_ph1, nodeNames_ph2, nodeNames_ph3):
    """ add pu voltages to buses shapes """
    try:
        busesDicTot = {}
        for i in range(len(nodeVoltages_ph1)):
            busesDic = {}
            #PH1
            for j in range(len(nodeNames_ph1)):
                key = nodeNames_ph1[j].split('.')[0].upper()
                dato = float(nodeVoltages_ph1[i][0][j])
                if key in busesDic:
                    busesDic[key].append( dato )
                else:
                    busesDic[key] = [ dato ]
                
            #PH2
            vector_dato = []
            for j in range(len(nodeNames_ph2)):
                key = nodeNames_ph2[j].split('.')[0].upper()
                dato = float(nodeVoltages_ph2[i][0][j])
                if key in busesDic:
                    busesDic[key].append( dato )
                else:
                    busesDic[key] = [ dato ]
                
            #PH3
            vector_dato = []
            for j in range(len(nodeNames_ph3)):
                key = nodeNames_ph3[j].split('.')[0].upper()
                dato = float(nodeVoltages_ph3[i][0][j])
                if key in busesDic:
                    busesDic[key].append( dato )
                else:
                    busesDic[key] = [ dato ]
            #Saca promedios a casos en que sea un vector de varios datos para esta iteración
            for dss_name, datos in  busesDic.items():
                prom = closeOne( datos )
                if dss_name in busesDicTot:
                    busesDicTot[ dss_name ].append( prom )
                else:
                    busesDicTot[ dss_name ] = [ prom ]
        return busesDicTot
        
    except:
        exc_info = sys.exc_info()
        print("\nError: ", exc_info )
        print("*************************  Información detallada del error ********************")
        
        
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)      
        
        return 0

def WriteBusVolts(self, busesDic, name_file_created, study):
    """ add pu voltages to buses shapes """
    try:
        a = 0
        # locate buses layers
        layerMVNames = [name_file_created.split('_')[0] + '_BusListMV', 'Bus_MV_Layer', 'Bus_MT_Layer']
        layerLVNames = [name_file_created.split('_')[0] + '_BusListLV', 'Bus_LV_Layer', 'Bus_BT_Layer']
        i = 0
        layerMV = ''
        layerLV = ''
        while a == 0:
            try:				
                layerMV = QgsProject.instance().mapLayersByName(layerMVNames[i])[0]
                layerLV = QgsProject.instance().mapLayersByName(layerLVNames[i])[0]
                a = 1
            except IndexError:
                i += 1
                if i > len(layerMVNames):
                    a = 1
        try:
            capsMV = layerMV.dataProvider().capabilities()
            capsLV = layerMV.dataProvider().capabilities()
        except AttributeError:
            return
        layerMV.startEditing()
        layerLV.startEditing()
        
        if study == 'snapshot':
            limit = 1
        elif study == 'daily':
            limit = 96
        elif study == 'shortcircuit':
            limit = 1
        
        for i in range(limit):
            MV = layerMV.getFeatures()
            LV = layerLV.getFeatures()
            # attribute names
            if study == 'snapshot':
                name = 'Vsnap'
            elif study == 'daily':
                name = "V" + str(i)
            elif study == 'shortcircuit':
                name = "Vsc"
            Idx_MV = getAttributeIndex(self, layerMV, name)
            Idx_LV = getAttributeIndex(self, layerLV, name)
            # write attribute values on MV shape
            
            try:
                if capsMV & QgsVectorDataProvider.ChangeAttributeValues:
                    for bus in MV:
                        try:
                            bus_name = str( bus['BUS'] )
                            dato = busesDic[ bus_name ][ i ]
                            dato = str( '{0:.6f}'.format( dato ) )
                            layerMV.changeAttributeValue( bus.id(), Idx_MV, dato )
                        except:
                            pass
                          
                # write attribute values on LV shape
                if capsLV & QgsVectorDataProvider.ChangeAttributeValues:
                    for bus in LV:
                        try:
                            bus_name = str( bus['BUS'] )
                            dato = busesDic[ bus_name ][ i ]
                            dato = str( '{0:.6f}'.format( dato ) )
                            layerLV.changeAttributeValue( bus.id(), Idx_LV, dato )
                        except:
                            pass
            except:
                exc_info = sys.exc_info()
                print("\nError: ", exc_info )
                print("*************************  Información detallada del error ********************")
                
                
                for tb in traceback.format_tb(sys.exc_info()[2]):
                    print(tb)        
        
        layerMV.commitChanges()
        layerLV.commitChanges()
        return
    except IndexError:  # no layer encountered
        QgsMessageLog.logMessage('No existe capa de buses', QCoreApplication.translate('dialog', 'Alerta Buses'),
                                 QgsMessageLog.WARNING)
    except:
        exc_info = sys.exc_info()
        print("\nError: ", exc_info )
        print("*************************  Información detallada del error ********************")
        
        
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)      
        
        return
        


def closeOne(vector):  # identifies which value write on shape
    value = 0
    for i in range(len(vector)):
        if 0.95 < vector[i] < 1.05:
            value = vector[i]
    if value == 0:
        value = np.mean(vector)
    return value

"""
This function read the output file of Qgis2Opendss and get the names of the layers of trafos and lines
Parámetros de entrada:
*name_output_azul (string): directorio completo del archivo de salida del azul

Valores retornados
*list_layersTrafos (list): lista con los nombres de las capas de los trafos
*, list_layersLinesMV (list): lista con los nombres de las capas de las líneas de baja tensión
*list_layersLinesLV (list): lista con los nombres de las capas de las líneas de media tensión
"""

def get_layersnames( name_output_azul ):
    list_layersTrafos = []
    list_layersLinesMV = []
    list_layersLinesLV = []
    
    with open( name_output_azul, "r" ) as f:
        lines = f.readlines()
        
    for line in lines:
        #Búsqueda de nombres de capas de líneas media tensión
        in_linesmv = line.find("Layers LinesMV: ")
        if in_linesmv != -1: #si encontró la frase anterior entra al if
            in_linesmv += len("Layers LinesMV: ")
            fin_linesmv = line.find("\n", in_linesmv)
            layers_lines_mv = line[in_linesmv:fin_linesmv]
            if layers_lines_mv != "": #si existía alguna capa de líneas de mt
                list_layersLinesMV = layers_lines_mv.split(",")
                list_layersLinesMV = list( filter(None, list_layersLinesMV) ) #se eliminan los espacios en blanco en la lista
        
        #Búsqueda de nombres de capas de líneas baja tensión
        in_lineslv = line.find("Layers LinesLV: ")
        if in_lineslv != -1: #si encontró la frase anterior entra al if
            in_lineslv += len("Layers LinesLV: ")
            fin_lineslv = line.find("\n", in_lineslv)
            layers_lines_lv = line[in_lineslv:fin_lineslv]
            if layers_lines_lv != "": #si existía alguna capa de líneas de bt
                list_layersLinesLV = layers_lines_lv.split(",")
                list_layersLinesLV = list( filter(None, list_layersLinesLV) ) #se eliminan los espacios en blanco en la lista
        
        #Búsqueda de nombres de capas de transformadores
        in_transformers = line.find("Layers Transformers: ")
        if in_transformers != -1: #si encontró la frase anterior entra al if
            in_transformers += len("Layers Transformers: ")
            fin_transformers = line.find("\n", in_transformers)
            layers_transformers = line[in_transformers:fin_transformers]
            if layers_transformers != "": #si existía alguna capa de transformadores
                list_layersTrafos = layers_transformers.split(",")
                list_layersTrafos = list( filter(None, list_layersTrafos) ) #se eliminan los espacios en blanco en la lista
    
    return list_layersTrafos, list_layersLinesMV, list_layersLinesLV

def newShape(self, DSSNames, shapeName):  # new shape file creation
    try:  # read layer
        layer = QgsProject.instance().mapLayersByName(shapeName)[0]
    except IndexError:
        layer = QgsVectorLayer('Point?crs=epsg:5367&field=DSSName:string(15)', shapeName, 'memory')
        layer.updateFields()
        self.iface.mapCanvas().refresh()
        QgsProject.instance().addMapLayers([layer])  # add memory layer to project
        pr = layer.dataProvider()
        fet = QgsFeature()
        for name in DSSNames:  # write DSS objects names on Attribute
            fet.setAttributes([str(name).upper()])
            pr.addFeatures([fet])
        layer.updateExtents()
        layer.commitChanges()
        shp = QgsProject.instance().fileName().split('GIS')[0] + 'GIS/' + shapeName + '.shp'  # shape path
        QgsVectorFileWriter.writeAsVectorFormat(layer, shp, "epsg:5367", QgsCoordinateReferenceSystem(),
                                                "ESRI Shapefile")
        QgsProject.instance().removeMapLayer(layer.id())  # delete memory layer
        wb = QgsVectorLayer(shp, shapeName, 'ogr')  # *.shp file
        QgsProject.instance().addMapLayer(wb)  # readd layer by path


def normalAmps(DSScircuit, lineNames):  # lines nominal ampacity
    normalAmpsDic = {}
    for name in lineNames:
        DSScircuit.SetActiveElement('line.' + name)
        lineNormalAmps = DSScircuit.ActiveCktElement.NormalAmps  # read normal Amps for each line
        normalAmpsDic[str(name).upper()] = lineNormalAmps  # add normal amps in dictionary
    return normalAmpsDic


def lineCurrents(self, DSScircuit, lineNames, normalAmpsDic, study):  # lineCurrents results calculation
    currentList = []
    for name in lineNames:        
        DSScircuit.SetActiveElement('line.' + name)
        lineAmps = DSScircuit.ActiveCktElement.CurrentsMagAng  # read line currents
        meanCurrent = '{0:.6f}'.format(  # calculate mean current
            np.max([lineAmps[x] for x in range(0, len(lineAmps), 2)]) / normalAmpsDic[str(name.upper())])
        currentList.append(meanCurrent)
       
    return currentList


def lineCurrentsResults(self, lineNames, currentList, study, list_layersLinesMV, list_layersLinesLV):  # add results to lineCurrents.shp
    tot_listlines = list_layersLinesMV + list_layersLinesLV #Lista de capas de líneas de baja y media tensión
    for name_layer in tot_listlines:
        layer = QgsProject.instance().mapLayersByName(name_layer)[0]
        caps = layer.dataProvider().capabilities()
        layer.startEditing()    
        for i in range(len(currentList)):        
            # attribute name
            if study.lower() == 'snapshot':
                name = 'PFsnap'
            elif study.lower() == 'daily':
                name = "PF" + str(i)
            Idx = getAttributeIndex(self, layer, str(name))
            if caps & QgsVectorDataProvider.ChangeAttributeValues:
                j = 0
                lineFeat = layer.getFeatures()            
                # DSS object current
                for line in lineFeat:                
                    layer.changeAttributeValue(line.id(), Idx, str(currentList[i][j]))  # id, Index, Value
                    if j == len( currentList[i] ) - 1:                    
                        break #si currentList y lineFeat difieren de tamaño debe terminar el ciclo                
                    j += 1
        layer.commitChanges()
        layer.updateFields()
    
    return 1

#Función que crea un vector con todos los trafos y los resultados de las simulaciones
def ReadTrafosLoad(self, DSScircuit, DSStext, name_file_created):  # transformers overload writing
    
    trafosDict = {}
    trafosNames = list(DSScircuit.Transformers.AllNames)

    # identifies 3-unit 3ph transformers on dss code
    for idx, name in enumerate(trafosNames):
        if 'u' in str(name.replace(name_file_created.split('_')[0].lower(), '')) and 'auto' not in str(name.replace(name_file_created.split('_')[0].lower(), '')):
            # QMessageBox.information(None, "Debug", name)
            trafosNames[idx] = str(name).split('_')[0] + '_' + str(name).split('_')[2]
            try:
                trafosNames.remove(str(name).split('_')[0] + '_2_' + str(name).split('_')[2])
            except ValueError:
                pass
            try:
                trafosNames.remove(str(name).split('_')[0] + '_3_' + str(name).split('_')[2])
            except ValueError:
                pass
    
    DSScircuit.Transformers.First
    nextTx = 1
    while nextTx != 0:  # results writing by transformer
        temp = 0
        trafo = DSScircuit.Transformers.Name  # get transformer name
        trafo = str(trafo).upper()
        DSScircuit.SetActiveElement('transformer.' + trafo)  # set active transformer
        # read transformer monitor
        if 'HVMV' in trafo:
            mon = 'HVMV_PQ_vs_Time'
        else:
            mon = 'mon' + trafo
        DSStext.Command = 'transformer.' + trafo + '.wdg=1'
        capacity = int(DSScircuit.Transformers.kva)  # transformer capacity
        if 'auto' in trafo.lower():
            capacity = capacity * 3
        try: # 3ph trafo
            valor = ExtractMonitorData(DSScircuit, mon, [1, 3, 5], 1)
            temp = np.abs(np.sum(valor, axis=1)) / capacity
        except IndexError:
            try:  # 2ph trafo
                valor = ExtractMonitorData(DSScircuit, mon, [1, 3], 1)
                temp = np.abs(np.sum(valor, axis=1)) / capacity
            except IndexError:
                try:  # 1ph trafo
                    valor = ExtractMonitorData(DSScircuit, mon, [1], 1)
                    temp = np.abs(np.sum(valor, axis=1)) / capacity
                except IndexError:
                    pass
        # for 3-units 3ph transformers there is only one point representing it
        # this code writes the worst case for the unit
        if 'U' in trafo.replace(name_file_created.split('_')[0].upper(), '') and 'auto' not in str(trafo.lower()):
            trafo = trafo.split('_')[0] + '_' + trafo.split('_')[2]
            try:
                if np.mean(trafosDict[trafo]) > np.mean(temp):
                    temp = trafosDict[trafo]
            except KeyError:
                pass
        trafosDict[trafo] = temp
        nextTx = DSScircuit.Transformers.Next  # set active next transformer
    return trafosDict, trafosNames

def WriteTrafosLoad(self, trafosDict, trafosNames, study, list_layersTrafos):  # transformers overload writing
    for name_layer in list_layersTrafos:
        layerTx = QgsProject.instance().mapLayersByName(name_layer)[0]  # select layer
        layerTx.startEditing()
        caps = layerTx.dataProvider().capabilities()
        
        #Determina el nombre de las columnas y el límite sobre el que va a iterar dependiendo del estudio
        if study == 'snapshot':
            lim = 1
        elif study == 'daily':
            lim = 96
        
        for i in range( lim ):  # results writing        
            if study == 'snapshot':
                name = 'PFsnap'
            elif study == 'daily':
                name = "PF" + str(i)
            Idx = getAttributeIndex(self, layerTx, str(name))
            
            if caps & QgsVectorDataProvider.ChangeAttributeValues:
                Tx = layerTx.getFeatures()
                for txs in Tx:
                    nameTx = str(txs['DSSName'])
                    dato = str('{0:.6f}'.format(trafosDict[nameTx][i]))               
                    layerTx.changeAttributeValue(txs.id(), Idx, dato)  # writing attribute value on shape
        layerTx.commitChanges()
        layerTx.updateFields()
        
    return 1

def selection_representative_day(file_path, type_day):
    # routine to select the most representative year on the year
    ############################################################################################    
    try:
        print( "file_path, type_day = ", file_path, ", ", type_day)
        with open(file_path, 'rt') as workbook:
            
            reader = csv.reader(workbook)  # read csv file
            next(reader)
            circuit_demand = [[row[3], row[2], row[0], row[1]] for row in reader]
            year = int((circuit_demand[0][0]).split('/')[2]);
        
            d1 = datetime.datetime(year, 1, 1)
            d2 = datetime.datetime(year, 12, 31)
            no_days = (d2 - d1).days + 1            
            
            num_zeros = ( int(len(circuit_demand) / no_days), no_days)           
            DP_to_be_matched = np.zeros( num_zeros )
            DQ_to_be_matched = np.zeros( num_zeros )
            type_of_day = np.zeros((1, no_days))
            date_monitored = np.zeros((3, no_days))

            weekend_days = []
            for i in range( int(no_days) ):
                type_of_day[0][i] = datetime.date(
                    int((circuit_demand[int(i * int(len(circuit_demand) / no_days))][0]).split('/')[2]),
                    int((circuit_demand[int(i * int(len(circuit_demand) / no_days))][0]).split('/')[1]),
                    int((circuit_demand[int(i * int(len(circuit_demand) / no_days))][0]).split('/')[0])).weekday()
                date_monitored[0][i] = int((circuit_demand[int(i * int(len(circuit_demand) / no_days))][0]).split('/')[2])
                date_monitored[1][i] = int((circuit_demand[int(i * int(len(circuit_demand) / no_days))][0]).split('/')[1])
                date_monitored[2][i] = int((circuit_demand[int(i * int(len(circuit_demand) / no_days))][0]).split('/')[0])
                if type_of_day[0][i] > 4:
                    weekend_days.append(i)
                for j in range(int(len(circuit_demand) / no_days)):
                    DP_to_be_matched[j][i] = circuit_demand[int(i * len(circuit_demand) / no_days) + j][2]
                    DQ_to_be_matched[j][i] = circuit_demand[int(i * len(circuit_demand) / no_days) + j][3]
            if type_day == 'weekday':
                DP_to_be_matched = np.delete(DP_to_be_matched, weekend_days, axis=1)
                DQ_to_be_matched = np.delete(DP_to_be_matched, weekend_days, axis=1)
                date_monitored = np.delete(date_monitored, weekend_days, axis=1)
            else:
                DP_to_be_matched = DP_to_be_matched[:, weekend_days]
                DQ_to_be_matched = DQ_to_be_matched[:, weekend_days]
                date_monitored = date_monitored[:, weekend_days]

            average_daily_demand = np.average(DP_to_be_matched, axis=1)
            differences_days = np.zeros((len(DP_to_be_matched), len(DP_to_be_matched[0][:])))

            for i in range(len(DP_to_be_matched[0][:])):
                differences_days[:, i] = DP_to_be_matched[:, i] - average_daily_demand

            sum_errors = np.sum(differences_days, axis=0)
            min_error = np.argmin(np.absolute(sum_errors), axis=0)
            DP_to_be_matched = DP_to_be_matched[:, min_error]
            DQ_to_be_matched = DQ_to_be_matched[:, min_error]

            s = "/";
            representative_day = "/".join((str(int(date_monitored[2, min_error])) if len(
                str(int(date_monitored[2, min_error]))) > 1 else '0' + str(int(date_monitored[2, min_error])), \
                                           str(int(date_monitored[1, min_error])) if len(
                                               str(int(date_monitored[1, min_error]))) > 1 else '0' + str(
                                               int(date_monitored[1, min_error])), str(int(date_monitored[0, min_error]))))
        workbook.closed
        return representative_day
    except ZeroDivisionError:  # if error, select the first day on file
        workbook.closed
        return circuit_demand[0][0]
        
    except:
        workbook.closed        
        exc_info = sys.exc_info()
        print("\nError: ", exc_info )
        print("*************************  Información detallada del error ********************")
        
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)
        return 0



def PV_allocation(total_capacity, pv_information, name_file_created, loadClass, studyType):
    # PV allocation algorithm
    PVS = [] # PVSystems list
    total_capacity = float(total_capacity)  #capacity to install

    myfile = open(pv_information, 'r')
    reader = list(csv.reader(myfile))

    lclass1 = ""
    lclass2 = ""
    # class separation by type
    if loadClass == "R":
        lclass1 = "class=R"
        lclass2 = "class=R"
    elif loadClass == "TMT":
        lclass1 = "class=TMT"
        lclass2 = "class=TMT"
    elif loadClass == "C":
        lclass1 = "class=C"
        lclass2 = "class=I"

    forbiddenBuses = []  # forbidden buses --> buses with PVsystems already installed
    try:
        gd_file = open(name_file_created.split('_')[0] + '_DG.dss')  # DG file reading
        gds = gd_file.read().split('\n')
        for i in gds:
            try:
                forbiddenBuses.append(i.split(" ")[2].replace("bus1=", "").split(".")[0])  # forbidden bus
            except IndexError:
                pass
    except IOError:
        pass

    probDict = {}  # probabilities dictionaty, order [capacity, probability]
    for n in range(1, len(reader)):
        try:  # try to read input file
            probDict[float(reader[n][0])] = [float(reader[n][1]), float(reader[n][2])]
        except ValueError:
            exc_info = sys.exc_info()
            print("\nError: ", exc_info )
            print("*************************  Información detallada del error ********************")
            
            for tb in traceback.format_tb(sys.exc_info()[2]):
                print(tb)
                        
            QMessageBox.information(None, QCoreApplication.translate('dialog', u"Información PV"),
                                    QCoreApplication.translate('dialog',
                                                               u"Archivo de información erróneo:") + "\n" + pv_information)
            PVS = ["Error"]  # exit on error            
            return PVS

    # loads file reader
    try:
        loads_file = open(name_file_created.split('_')[0] + '_loadslv.dss', 'r')
    except IOError:
        loads_file = open(name_file_created.split('_')[0] + '_LoadsLV.dss', 'r')
    lines = loads_file.read().split('\n')

    # dictionary with loads and kWh to order from max to min
    loadsDict = {}  # dictionary without order
    for i in range(len(lines) - 1):
        kWh_load = float(lines[i].split('kWh=')[1].split(' ')[0])
        loadsDict[lines[i]] = kWh_load
    # ordering dictionary from max to min
    orderedLoads = OrderedDict(sorted(list(loadsDict.items()), key=lambda t: t[1], reverse=True))
    oLKlist = list(orderedLoads)  #ordered dict keys/loads list
    capacidadTotal = np.zeros(len(oLKlist))
    sumaAnt = 0
    stop = 0  # stop condition

    ##### random
    if "Random" in studyType or "Maximum" in studyType:
        # shuffle the loads
        if "Random" in studyType:
            random.shuffle(oLKlist)

        while np.sum(capacidadTotal) < total_capacity:
            for idx, i in enumerate(oLKlist):
                if capacidadTotal[idx] == 0 and (lclass1 or lclass2) in i:
                    kWh = orderedLoads[i]  # load monthly energy
                    try:
                        capacidad, probabilidad = probDict[np.round(kWh, -1)]
                        bus = i.split(' ')[2]
                        if bus not in forbiddenBuses:
                            capacidadTotal[idx] = capacidad
                            if capacidadTotal[idx] != 0: # bus without PV system
                                # PV System sentence
                                load_class = i.split(' ')[12].split('=')[1]
                                name = i.split(' ')[1].split('.')[1]

                                voltage = i.split(' ')[3]
                                conn = i.split(' ')[5]
                                phases = i.split(' ')[9]
                                PVS.append(
                                    'New PVSystem.PV' + loadClass + name + ' ' + bus + ' ' + voltage + ' ' + phases + ' kVA='
                                    + str(capacidad) + ' PF=1 ' + conn + ' irrad=0.90' + ' Pmpp=' + str(capacidad)
                                    + ' temperature=25 effcurve=Myeff P-TCurve=MyPvsT Daily=MyIrrad '
                                    + 'TDaily=MyTemp %cutin=0.01 %cutout=0.01 enabled=yes !class=' + load_class)
                    except KeyError:
                        pass
                    if np.sum(capacidadTotal) >= total_capacity:
                        break
                else:
                    pass
            if sumaAnt == np.sum(capacidadTotal):
                stop += 1
                if stop == 5:
                    QMessageBox.information(None, QCoreApplication.translate('dialog', u"Información PV"),
                                            QCoreApplication.translate('dialog',
                                                                       u"Mayor potencia instalada lograda según archivo de información \nde instalaciones óptimas: ") + str(
                                                sumaAnt) + " kW" + '\n')
                    break
            sumaAnt = np.sum(capacidadTotal)

    # Bernoulli
    if studyType == "Bernoulli":
        while np.sum(capacidadTotal) < total_capacity:
            for idx, i in enumerate(oLKlist):
                if capacidadTotal[idx] == 0 and (lclass1 or lclass2) in i:
                    kWh = orderedLoads[i]
                    try:
                        capacidad, probabilidad = probDict[np.round(kWh, -1)]
                        choices = [[capacidad, probabilidad], [0, 1 - probabilidad]]
                        capacidadTotal[idx] = weighted_choice(choices)
                        bus = i.split(' ')[2]
                        if bus not in forbiddenBuses:
                            if capacidadTotal[idx] != 0:  # bus without PV system
                                # PV System sentence
                                load_class = i.split(' ')[12].split('=')[1]
                                name = i.split(' ')[1].split('.')[1]
                                voltage = i.split(' ')[3]
                                conn = i.split(' ')[5]
                                phases = i.split(' ')[9]
                                PVS.append(
                                    'New PVSystem.PV' + loadClass + name + ' ' + bus + ' ' + voltage + ' ' + phases + ' kVA=' + str(
                                        capacidad) + ' PF=1 ' + conn + ' irrad=0.90' + ' Pmpp=' + str(
                                        capacidad) + ' temperature=25 effcurve=Myeff P-TCurve=MyPvsT '
                                                     'Daily=MyIrrad TDaily=MyTemp %cutin=0.01 %cutout=0.01 enabled=yes '
                                                     '!class=' + load_class)
                    except KeyError:
                        pass
                    if np.sum(capacidadTotal) >= total_capacity:
                        break
                else:
                    pass
            if sumaAnt == np.sum(capacidadTotal):
                stop += 1
                if stop == 10:
                    QMessageBox.information(None, QCoreApplication.translate('dialog', u"Información PV"),
                                            QCoreApplication.translate('dialog',
                                                                       u"Mayor potencia instalada lograda según "
                                                                       u"archivo de información \nde instalaciones "
                                                                       u"óptimas: ") + str(
                                                sumaAnt) + " kW" + '\n')
                    break
            sumaAnt = np.sum(capacidadTotal)

    print("Capacidad: " + str(sumaAnt))
    loads_file.close()
    return PVS


def unbalance(self, DSScircuit, buses, unbalance, orderedUnbalance):
    if orderedUnbalance:  # uses dictionary that already exist
        for i in buses:
            DSScircuit.setActiveBus(i)
            if len(DSScircuit.ActiveBus.Nodes) == 3:  # use only 3ph buses
                curVal = orderedUnbalance[i]
                unb = (DSScircuit.ActiveBus.SeqVoltages[2] / DSScircuit.ActiveBus.SeqVoltages[1]) * 100  # (V-/V+)*100
                if unb > curVal:
                    orderedUnbalance[i] = unb  # stores maximum registered unbalances
    else:  # create unbalance dictionary
        for i in buses:
            DSScircuit.setActiveBus(i)
            if len(DSScircuit.ActiveBus.Nodes) == 3:  # use only 3ph buses
                unb = (DSScircuit.ActiveBus.SeqVoltages[2] / DSScircuit.ActiveBus.SeqVoltages[1]) * 100  # (V-/V+)*100
                unbalance[i] = unb  # write voltage unbalance
        # se ordena el diccionario unbalance en orden alfabetico
        orderedUnbalance = OrderedDict(sorted(list(unbalance.items()), key=lambda t: t[0], reverse=True))

    return orderedUnbalance
