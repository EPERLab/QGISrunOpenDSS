# -*- coding: utf-8 -*-
"""
/***************************************************************************
 QGISrunOpenDSSDialog
                                 A QGIS plugin
 This plugin runs power flows using OpenDSS
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

import os

from qgis.PyQt import QtGui, uic
from PyQt5.QtWidgets import QDialog

FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'QGIS_run_OpenDSS_dialog_base.ui'))


class QGISrunOpenDSSDialog(QDialog, FORM_CLASS):
    def __init__(self, parent=None):
        """Constructor."""
        super(QGISrunOpenDSSDialog, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
