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
from openpyxl import load_workbook
from qgis.PyQt import uic, QtWidgets
from qgis.PyQt.QtWidgets import QMessageBox
from shapely.geometry import LineString, Polygon

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

        # Connect file selection and dropdowns
        self.xlsFileWidget.fileChanged.connect(self.load_sheets)
        self.sheetDropdown.currentIndexChanged.connect(self.load_columns)

        # Auto-selection checkbox logic
        if hasattr(self, 'autoSelectCheckbox'):
            self.autoSelectCheckbox.stateChanged.connect(self.load_columns)

        # Geometry type selection logic
        self.geometryTypeDropdown.currentIndexChanged.connect(self.toggle_geometry_fields)
        self.toggle_geometry_fields()

    def toggle_geometry_fields(self):
        """Enable/disable fields based on selected geometry type instead of hiding them."""
        selected_option = self.geometryTypeDropdown.currentText()
        
        is_line_selected = selected_option in ["Line Only", "Both"]
        is_polygon_selected = selected_option in ["Polygon Only", "Both"]

        # Enable/Disable line-related fields
        self.labelTrace.setEnabled(is_line_selected)
        self.traceColumnDropdown.setEnabled(is_line_selected)
        self.labelTraceResult.setEnabled(is_line_selected)
        self.traceResultColumnName.setEnabled(is_line_selected)

        # Enable/Disable polygon-related fields
        self.labelPolygon.setEnabled(is_polygon_selected)
        self.polygonColumnDropdown.setEnabled(is_polygon_selected)
        self.labelPolyResult.setEnabled(is_polygon_selected)
        self.polyResultColumnName.setEnabled(is_polygon_selected)


    def load_sheets(self):
        """Load sheets from the selected .xlsx file into the sheet dropdown."""
        file_path = self.xlsFileWidget.filePath()
        if not file_path.endswith(".xlsx"):
            QMessageBox.warning(self, "Invalid File", "Please select a valid .xlsx file.")
            return

        try:
            workbook = load_workbook(file_path, read_only=False)
            self.sheetDropdown.clear()
            self.sheetDropdown.addItems(workbook.sheetnames)
            self.workbook = workbook

            if workbook.sheetnames:
                self.sheetDropdown.setCurrentIndex(0)
                self.load_columns()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read .xlsx file: {e}")

    def load_columns(self):
        """Loads column headers from the selected sheet into dropdowns."""
        selected_sheet = self.sheetDropdown.currentText()
        if not selected_sheet or not hasattr(self, 'workbook'):
            return

        try:
            sheet = self.workbook[selected_sheet]
            headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1)) if cell.value]

            # Populate dropdowns
            self.traceColumnDropdown.clear()
            self.polygonColumnDropdown.clear()

            # Add 'No Column Selected' option
            self.traceColumnDropdown.addItem("No Column Selected", None)
            self.polygonColumnDropdown.addItem("No Column Selected", None)

            # Add actual column names
            self.traceColumnDropdown.addItems(headers)
            self.polygonColumnDropdown.addItems(headers)

            # Auto-select predefined columns if checkbox is checked
            if hasattr(self, 'autoSelectCheckbox') and self.autoSelectCheckbox.isChecked():
                odk_geotrace_col = "Record surface feature as line."
                odk_geoshape_col = "Record surface feature as polygon."

                if odk_geotrace_col in headers:
                    self.traceColumnDropdown.setCurrentText(odk_geotrace_col)
                if odk_geoshape_col in headers:
                    self.polygonColumnDropdown.setCurrentText(odk_geoshape_col)
                
                self.traceResultColumnName.setText("qgis_line_wkt")
                self.polyResultColumnName.setText("qgis_polygon_wkt")
            else:
                self.traceResultColumnName.clear()
                self.polyResultColumnName.clear()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load columns: {e}")

    def convert_coordinates(self):
        """Converts selected coordinate columns into WKT format."""
        selected_option = self.geometryTypeDropdown.currentText()
        process_lines = selected_option in ["Line Only", "Both"]
        process_polygons = selected_option in ["Polygon Only", "Both"]

        if not process_lines and not process_polygons:
            QMessageBox.warning(self, "Error", "Please select a geometry type to convert.")
            return

        if process_lines and self.traceColumnDropdown.currentText() == "No Column Selected":
            QMessageBox.warning(self, "Error", "Please select a column for line conversion.")
            return
        if process_polygons and self.polygonColumnDropdown.currentText() == "No Column Selected":
            QMessageBox.warning(self, "Error", "Please select a column for polygon conversion.")
            return

        self._convert_coordinates_logic(process_lines, process_polygons)

    def _convert_coordinates_logic(self, process_lines, process_polygons):
        """Internal logic for coordinate conversion."""
        selected_sheet = self.sheetDropdown.currentText()
        trace_column = self.traceColumnDropdown.currentText() if process_lines else None
        polygon_column = self.polygonColumnDropdown.currentText() if process_polygons else None

        if not selected_sheet or not hasattr(self, 'workbook'):
            QMessageBox.warning(self, "Error", "No sheet selected or workbook not loaded.")
            return

        sheet = self.workbook[selected_sheet]

        # Iterate over rows and convert coordinates
        for row in sheet.iter_rows(min_row=2):
            if process_lines and trace_column:
                trace_data = self.get_cell_value(sheet, row, trace_column)
                if trace_data:
                    try:
                        wkt_line = self.convert_to_wkt(trace_data, "LINESTRING")
                        print(f"Converted Line WKT: {wkt_line}")
                    except ValueError as e:
                        QMessageBox.warning(self, "Conversion Error", str(e))

            if process_polygons and polygon_column:
                polygon_data = self.get_cell_value(sheet, row, polygon_column)
                if polygon_data:
                    try:
                        wkt_polygon = self.convert_to_wkt(polygon_data, "POLYGON")
                        print(f"Converted Polygon WKT: {wkt_polygon}")
                    except ValueError as e:
                        QMessageBox.warning(self, "Conversion Error", str(e))

    def get_cell_value(self, sheet, row, column_name):
        """Get the value of a specific column from a row."""
        col_index = [cell.column for cell in sheet[1] if cell.value == column_name]
        if col_index:
            return row[col_index[0] - 1].value
        return None

    def convert_to_wkt(self, coordinates, geom_type):
        """Converts ODK coordinates to WKT format."""
        flipped_coords = [self.flip_coordinates(coord) for coord in coordinates.split(";") if coord]
        if geom_type == "LINESTRING":
            return LineString(flipped_coords).wkt
        elif geom_type == "POLYGON":
            return Polygon(flipped_coords).wkt
        raise ValueError("Invalid geometry type")

    def flip_coordinates(self, coordinate):
        """Flips ODK coordinates from (latitude, longitude) to (longitude, latitude)."""
        coords = coordinate.split()
        if len(coords) >= 2:
            return float(coords[1]), float(coords[0])
        raise ValueError(f"Invalid coordinate format: {coordinate}")
