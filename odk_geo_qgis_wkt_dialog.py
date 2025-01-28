# -*- coding: utf-8 -*-
"""
/***************************************************************************
 ODKGeo_QgisWktDialog
                                 A QGIS plugin
 This plugin converts ODK geo coordinate values to QGIS (flipped) WKT values.
 ***************************************************************************
"""

import os
from openpyxl import load_workbook  # Ensure this import is here for reading/writing Excel files
from qgis.PyQt import uic
from qgis.PyQt import QtWidgets
from qgis.PyQt.QtWidgets import QMessageBox
from shapely.geometry import Point, LineString, Polygon

# Load UI file
FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'odk_geo_qgis_wkt_dialog_base.ui'))

class ODKGeo_QgisWktDialog(QtWidgets.QDialog, FORM_CLASS):
    def __init__(self, parent=None):
        """Constructor."""
        super(ODKGeo_QgisWktDialog, self).__init__(parent)
        self.setupUi(self)

        # Ensure convertButton exists in the UI
        if hasattr(self, "convertButton"):
            self.convertButton.clicked.connect(self.convert_coordinates)
        else:
            raise AttributeError("convertButton is missing from the UI file. Check your .ui design.")

        # Connect signals to slots for other UI components
        self.xlsFileWidget.fileChanged.connect(self.load_sheets)
        self.sheetDropdown.currentIndexChanged.connect(self.load_columns)

    def load_sheets(self):
        """Load sheets from the selected .xlsx file into the sheet dropdown."""
        file_path = self.xlsFileWidget.filePath()
        if not file_path.endswith(".xlsx"):
            QMessageBox.warning(self, "Invalid File", "Please select a valid .xlsx file.")
            return

        try:
            workbook = load_workbook(file_path)  # Open in normal mode, not read-only
            self.sheetDropdown.clear()
            self.sheetDropdown.addItems(workbook.sheetnames)
            self.workbook = workbook
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read .xlsx file: {e}")

    def load_columns(self):
        """Load column headers from the selected sheet into dropdowns."""
        selected_sheet = self.sheetDropdown.currentText()
        if not selected_sheet or not hasattr(self, 'workbook'):
            return

        try:
            # Get the selected sheet and read the first row for headers
            sheet = self.workbook[selected_sheet]
            headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1)) if cell.value]

            # Populate dropdowns
            self.pointColumnDropdown.clear()
            self.traceColumnDropdown.clear()
            self.polygonColumnDropdown.clear()

            self.pointColumnDropdown.addItems(headers)
            self.traceColumnDropdown.addItems(headers)
            self.polygonColumnDropdown.addItems(headers)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load columns: {e}")

    def convert_coordinates(self):
        """Process the coordinates, flip and create valid WKT records for polygons only."""
        selected_sheet = self.sheetDropdown.currentText()
        polygon_column = self.polygonColumnDropdown.currentText()

        if not selected_sheet or not hasattr(self, 'workbook'):
            QMessageBox.warning(self, "Error", "No sheet selected or workbook not loaded.")
            return

        try:
            # Open the workbook in regular (writeable) mode instead of read-only mode
            file_path = self.xlsFileWidget.filePath()
            workbook = load_workbook(file_path)  # Now open in normal mode (not read-only)
            sheet = workbook[selected_sheet]  # Access the sheet directly by name

            # Extract headers and determine the new column index
            headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            if None in headers:
                raise ValueError("Empty column headers detected. Please ensure all columns have names.")
            
            new_column_index = len(headers) + 1

            # Add a new header for the WKT column
            sheet.cell(row=1, column=new_column_index).value = "QGIS Geoshape WKT"

            # Process each row to create WKT values for polygons only
            for row_index, row in enumerate(sheet.iter_rows(min_row=2, max_row=sheet.max_row), start=2):
                # Get the cell value for the polygon column
                polygon = row[headers.index(polygon_column)].value if polygon_column in headers else None

                wkt_value = None
                if polygon:
                    # Split the coordinates by semicolon to get individual points
                    coords = polygon.split(';')
                    
                    # Flip each coordinate and store them in a list
                    flipped_polygon = [self.flip_coordinates(coord.strip()) for coord in coords if coord.strip()]
                    
                    # Create the Polygon WKT from the flipped coordinates
                    wkt_value = Polygon(flipped_polygon).wkt

                # Explicitly write the WKT value to the new column
                target_cell = sheet.cell(row=row_index, column=new_column_index)
                target_cell.value = wkt_value

            # Save the modified workbook
            workbook.save(file_path)  # Save the workbook with the new WKT values
            QMessageBox.information(self, "Success", "Polygons converted and saved successfully!")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to convert coordinates: {e}")

    def flip_coordinates(self, coordinate):
        """Flip the coordinates (longitude, latitude)."""
        coords = coordinate.split(' ')  # ODK coordinates use space to separate values
        if len(coords) >= 2:
            # Only use the first two values (longitude, latitude)
            return float(coords[1].strip()), float(coords[0].strip())
        raise ValueError(f"Invalid coordinate format: {coordinate}")
