# -*- coding: utf-8 -*-
"""
/***************************************************************************
 QGISrunOpenDSS
                                 A QGIS plugin
 This plugin runs power flows using OpenDSS
                             -------------------
        begin                : 2015-12-03
        copyright            : (C) 2015 by Jairo Quiros
        email                : jquiros@eie.ucr.ac.cr
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load QGISrunOpenDSS class from file QGISrunOpenDSS.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .QGIS_run_OpenDSS import QGISrunOpenDSS
    return QGISrunOpenDSS(iface)
