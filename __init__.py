# -*- coding: utf-8 -*-
"""
/***************************************************************************
 ODKGeo_QgisWkt
                                 A QGIS plugin
 This plugin converts ODK geo coordinate values to QGIS (flipped) WKT values.
 Generated by Plugin Builder: http://g-sherman.github.io/Qgis-Plugin-Builder/
                             -------------------
        begin                : 2025-01-24
        copyright            : (C) 2025 by Junaid Abdul Jabbar
        email                : junaid.abdul.jabbar@gmail.com
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
    """Load ODKGeo_QgisWkt class from file ODKGeo_QgisWkt.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .odk_geo_qgis_wkt import ODKGeo_QgisWkt
    return ODKGeo_QgisWkt(iface)
