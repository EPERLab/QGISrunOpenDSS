# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'QGIS_run_OpenDSS_transformer.ui'
#
# Created: Mon Mar 21 17:23:48 2016
#      by: PyQt4 UI code generator 4.10.2
#
# WARNING! All changes made in this file will be lost!

import os

from qgis.PyQt import QtGui, uic
from PyQt5.QtWidgets import QDialog

FORM_CLASS, _ = uic.loadUiType(os.path.join(
os.path.dirname(__file__), 'QGIS_run_OpenDSS_transformer.ui'))

class Ui_Transformer(QDialog, FORM_CLASS):

   def __init__(self, parent=None):
      """Constructor."""
      super(Ui_Transformer, self).__init__(parent)
      self.setupUi(self)
