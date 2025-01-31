# -*- coding: utf-8 -*-
"""
/***************************************************************************
 ODKGeo_QgisWktDialog
                                 A QGIS plugin
 This plugin converts ODK geo coordinate values to QGIS (flipped) WKT values.
 It processes points, traces (lines), and polygons.
 ***************************************************************************/
"""

import os
import gc
from openpyxl import load_workbook
from qgis.PyQt import uic
from qgis.PyQt import QtWidgets
from qgis.PyQt.QtWidgets import QMessageBox
from shapely.geometry import Point, LineString, Polygon

# Load UI file
FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'odk_geo_qgis_wkt_dialog_base.ui'))

class ODKGeo_QgisWktDialog(QtWidgets.QDialog, FORM_CLASS):
    def __init__(self, parent=None):
        """Constructor - Initializes the UI and connects signals to slots."""
        super(ODKGeo_QgisWktDialog, self).__init__(parent)
        self.setupUi(self)

        # Ensure convertButton exists in the UI
        if hasattr(self, "convertButton"):
            self.convertButton.clicked.connect(self.convert_coordinates)
        else:
            raise AttributeError("convertButton is missing from the UI file. Check your .ui design.")

        # Connect file and dropdown widgets to corresponding functions
        self.xlsFileWidget.fileChanged.connect(self.load_sheets)
        self.sheetDropdown.currentIndexChanged.connect(self.load_columns)

    def load_sheets(self):
        """Load sheets from the selected .xlsx file into the sheet dropdown and auto-load first sheet's columns."""
        file_path = self.xlsFileWidget.filePath()
        if not file_path.endswith(".xlsx"):
            QMessageBox.warning(self, "Invalid File", "Please select a valid .xlsx file.")
            return

        try:
            workbook = load_workbook(file_path, read_only=False)  # Open in normal mode for writing
            self.sheetDropdown.clear()
            self.sheetDropdown.addItems(workbook.sheetnames)
            self.workbook = workbook

            # Auto-select the first sheet and load columns immediately
            if workbook.sheetnames:
                self.sheetDropdown.setCurrentIndex(0)  # Select first sheet
                self.load_columns()  # Populate columns automatically

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read .xlsx file: {e}")

    def load_columns(self):
        """Loads column headers from the selected sheet into dropdowns."""
        selected_sheet = self.sheetDropdown.currentText()
        if not selected_sheet or not hasattr(self, 'workbook'):
            return

        try:
            # Get the headers from the first row
            sheet = self.workbook[selected_sheet]
            headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1)) if cell.value]

            # Populate dropdowns with column names
            self.pointColumnDropdown.clear()
            self.traceColumnDropdown.clear()
            self.polygonColumnDropdown.clear()

            self.pointColumnDropdown.addItems(headers)
            self.traceColumnDropdown.addItems(headers)
            self.polygonColumnDropdown.addItems(headers)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load columns: {e}")

    def convert_coordinates(self):
        """Converts selected coordinate columns into WKT format (Point, LineString, Polygon)."""
        selected_sheet = self.sheetDropdown.currentText()
        point_column = self.pointColumnDropdown.currentText()
        trace_column = self.traceColumnDropdown.currentText()
        polygon_column = self.polygonColumnDropdown.currentText()

        if not selected_sheet or not hasattr(self, 'workbook'):
            QMessageBox.warning(self, "Error", "No sheet selected or workbook not loaded.")
            return

        file_path = self.xlsFileWidget.filePath()
        try:
            # Open workbook for editing
            workbook = load_workbook(file_path)
            sheet = workbook[selected_sheet]

            # Get existing headers and determine new column positions
            headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            existing_columns = {header: idx + 1 for idx, header in enumerate(headers) if header}

            # Determine column positions for new WKT values
            point_col_idx = existing_columns.get("QGIS Point WKT", len(headers) + 1)
            trace_col_idx = existing_columns.get("QGIS Trace WKT", len(headers) + (2 if point_column else 1))
            polygon_col_idx = existing_columns.get("QGIS Polygon WKT", len(headers) + 
                                                   (3 if point_column and trace_column else 
                                                    2 if point_column or trace_column else 1))

            # Add WKT column headers if they don’t exist
            if "QGIS Point WKT" not in existing_columns and point_column:
                sheet.cell(row=1, column=point_col_idx, value="QGIS Point WKT")
            if "QGIS Trace WKT" not in existing_columns and trace_column:
                sheet.cell(row=1, column=trace_col_idx, value="QGIS Trace WKT")
            if "QGIS Polygon WKT" not in existing_columns and polygon_column:
                sheet.cell(row=1, column=polygon_col_idx, value="QGIS Polygon WKT")

            # Iterate over rows and convert coordinates
            for row_index, row in enumerate(sheet.iter_rows(min_row=2, max_row=sheet.max_row), start=2):
                if point_column and row[headers.index(point_column)].value:
                    point = self.flip_coordinates(row[headers.index(point_column)].value)
                    sheet.cell(row=row_index, column=point_col_idx, value=Point(point).wkt)

                if trace_column and row[headers.index(trace_column)].value:
                    trace = [self.flip_coordinates(coord) for coord in row[headers.index(trace_column)].value.split(';')]
                    sheet.cell(row=row_index, column=trace_col_idx, value=LineString(trace).wkt)

                if polygon_column and row[headers.index(polygon_column)].value:
                    polygon = [self.flip_coordinates(coord) for coord in row[headers.index(polygon_column)].value.split(';')]
                    sheet.cell(row=row_index, column=polygon_col_idx, value=Polygon(polygon).wkt)

            # Save changes and clean up
            workbook.save(file_path)
            workbook.close()
            del workbook
            gc.collect()  # Free memory

            QMessageBox.information(self, "Success", "Coordinates converted and saved successfully!")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to convert coordinates: {e}")

    def flip_coordinates(self, coordinate):
        """
        Flips ODK coordinates from (latitude, longitude) to (longitude, latitude).
        Expects coordinates in "lat lon" format.
        """
        coords = coordinate.split()
        if len(coords) >= 2:
            return float(coords[1]), float(coords[0])  # Swap lat and lon
        raise ValueError(f"Invalid coordinate format: {coordinate}")
