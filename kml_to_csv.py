import sys
import csv
import simplekml
from shapely.geometry import Point
from shapely import wkt
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QFileDialog, QLineEdit, QComboBox, QColorDialog,
                             QCheckBox, QSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
                             QMessageBox, QRadioButton, QButtonGroup, QGroupBox, QScrollArea)
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtCore import Qt, QPoint, QRect 

import numpy as np
import pandas as pd
import re


MISSING_VALS = {"", "null", "none", "nan", "na", "n/a"}


def jenks_breaks(data, num_classes):
    """Calculate Jenks natural breaks for the given data."""
    if not data or num_classes <= 0:
        return []

    data = sorted(data)
    num_data = len(data)
    if num_classes > num_data:
        num_classes = num_data

    mat1 = [[0] * (num_classes + 1) for _ in range(num_data + 1)]
    mat2 = [[0] * (num_classes + 1) for _ in range(num_data + 1)]

    for i in range(1, num_classes + 1):
        mat1[0][i] = 1
        mat2[0][i] = 0
        for j in range(1, num_data + 1):
            mat2[j][i] = float('inf')

    for l in range(1, num_data + 1):
        s1 = s2 = w = 0.0
        for m in range(l, 0, -1):
            val = data[m - 1]
            s1 += val
            s2 += val * val
            w += 1
            variance = s2 - (s1 * s1) / w
            if m > 1:
                for j in range(2, num_classes + 1):
                    if mat2[l][j] >= variance + mat2[m - 1][j - 1]:
                        mat1[l][j] = m
                        mat2[l][j] = variance + mat2[m - 1][j - 1]
        mat1[l][1] = 1
        mat2[l][1] = variance

    breaks = [0] * (num_classes + 1)
    breaks[num_classes] = data[-1]
    k = num_data
    for j in range(num_classes, 1, -1):
        idx = int(mat1[k][j] - 2)
        breaks[j - 1] = data[idx]
        k = int(mat1[k][j] - 1)
    breaks[0] = data[0]

    return breaks


def parse_filter_expression(expr: str) -> str:
    """Convert a user-friendly filter expression to a pandas query string."""
    if not expr:
        return ""

    # Replace single '=' with '==' for equality checks
    expr = re.sub(r'(?<![<>=!])=(?!=)', '==', expr)

    # Quote bare words on the right side of comparisons
    def repl(match):

        col = match.group(1).strip()
        op = match.group(2)
        val = match.group(3).strip()

        col_token = f'`{col}`' if not re.fullmatch(r'\w+', col) else col

        if re.fullmatch(r'-?\d+(\.\d+)?', val):
            return f"{col_token}{op}{val}"
        if not (val.startswith('"') or val.startswith("'")):
            val = f'"{val}"'
        return f"{col_token}{op}{val}"

    expr = re.sub(r'([\w ]+)\s*(==|!=|>=|<=|>|<)\s*([^&|]+)', repl, expr)
    return expr

    col, op, val = match.group(1), match.group(2), match.group(3)
    if re.fullmatch(r'-?\d+(\.\d+)?', val):
        return f"{col}{op}{val}"
    # Wrap strings that are not already quoted
    if not (val.startswith('"') or val.startswith("'")):
        val = f'"{val}"'
    return f"{col}{op}{val}"

    expr = re.sub(r'([\w ]+)\s*(==|!=|>=|<=|>|<)\s*([^&|]+)', repl, expr)
    return expr

def jenks_breaks(data, num_classes):
    """Calculate Jenks natural breaks for the given data."""
    if not data or num_classes <= 0:
        return []

    data = sorted(data)
    num_data = len(data)
    if num_classes > num_data:
        num_classes = num_data

    mat1 = [[0] * (num_classes + 1) for _ in range(num_data + 1)]
    mat2 = [[0] * (num_classes + 1) for _ in range(num_data + 1)]

    for i in range(1, num_classes + 1):
        mat1[0][i] = 1
        mat2[0][i] = 0
        for j in range(1, num_data + 1):
            mat2[j][i] = float('inf')

    for l in range(1, num_data + 1):
        s1 = s2 = w = 0.0
        for m in range(l, 0, -1):
            val = data[m - 1]
            s1 += val
            s2 += val * val
            w += 1
            variance = s2 - (s1 * s1) / w
            if m > 1:
                for j in range(2, num_classes + 1):
                    if mat2[l][j] >= variance + mat2[m - 1][j - 1]:
                        mat1[l][j] = m
                        mat2[l][j] = variance + mat2[m - 1][j - 1]
        mat1[l][1] = 1
        mat2[l][1] = variance

    breaks = [0] * (num_classes + 1)
    breaks[num_classes] = data[-1]
    k = num_data
    for j in range(num_classes, 1, -1):
        idx = int(mat1[k][j] - 2)
        breaks[j - 1] = data[idx]
        k = int(mat1[k][j] - 1)
    breaks[0] = data[0]

    return breaks

class KmlGeneratorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.data = []
        self.filtered_data = []  # Store data after applying filter expressions
        self.headers = []
        self.group_colors = {}
        self.groups = []
        self.end_color = QColor(255, 0, 0)
        self.field_types = {}
        self.encoding = 'utf-8'
        self.manual_group_bounds = {}
        self.current_header_combo = None # To keep track of the currently open QComboBox for header editing
        self.current_header_combo_column = -1 # Track which column the combo belongs to

        self.initUI()

    def initUI(self):
        """Инициализация пользовательского интерфейса с обновленным дизайном."""
        self.setWindowTitle('KML Generator')
        self.setGeometry(100, 100, 950, 850)

        self.setStyleSheet("""
            QWidget { background-color: #F8F8F8;
font-family: Arial; }
            QGroupBox {
                border: 1px solid #D0D0D0;
border-radius: 7px; margin-top: 10px;
                padding-top: 15px; background-color: #FFFFFF;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
left: 10px; padding: 0 5px;
                color: #333333; font-weight: bold; font-size: 11px;
            }
            QScrollArea {
                border: none;
            }
        """)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        main_container = QWidget()
        layout = QVBoxLayout(main_container)
        main_app_layout = QVBoxLayout(self)
        main_app_layout.addWidget(scroll_area)
        self.setLayout(main_app_layout)

        bold_large_font = QFont()
        bold_large_font.setPointSize(10)
        bold_large_font.setBold(True)

        label_style = "QLabel { color: #333333;font-weight: bold;}"
        button_style = "QPushButton { background-color: #555555; color: white; border-radius: 5px;padding: 5px 10px;font-weight: bold; }"
        button_hover_style = "QPushButton:hover { background-color: #777777}"
        lineedit_style = "QLineEdit { background-color: #EEEEEE; border: 1px solid #CCCCCC; padding: 3px; }"
        combobox_style = "QComboBox { background-color: #EEEEEE;border: 1px solid #CCCCCC; padding: 1px; }"
        spinbox_style = "QSpinBox { background-color: #EEEEEE; border: 1px solid #CCCCCC; padding: 3px; }"
        checkbox_style = "QCheckBox { color: #333333;font-weight: bold;}"
        radio_button_style = "QRadioButton { color: #333333;font-weight: bold;}"
        table_style = """
            QTableWidget { background-color: #FFFFFF; border: 1px solid #CCCCCC; gridline-color: #E0E0E0;
}
            QHeaderView::section { background-color: #E0E0E0; color: #333333; padding: 4px;
border: 1px solid #CCCCCC; font-weight: bold; }
        """


        file_group_box = QGroupBox("File Input/Output Settings")
        file_group_box.setFont(bold_large_font)
        file_group_layout = QVBoxLayout()
        file_group_layout.setContentsMargins(10, 15, 10, 15)
        file_group_box.setLayout(file_group_layout)
        
        file_layout = QHBoxLayout()
        self.file_label = QLabel('Input Data File:')
        self.file_label.setStyleSheet(label_style)
        file_layout.addWidget(self.file_label)
        self.file_path_input = QLineEdit()
        self.file_path_input.setStyleSheet(lineedit_style)
        file_layout.addWidget(self.file_path_input)
        self.browse_button = QPushButton('Browse')
        self.browse_button.clicked.connect(self.browse_file)
        self.browse_button.setStyleSheet(button_style + button_hover_style)
        file_layout.addWidget(self.browse_button)
        file_group_layout.addLayout(file_layout)

        output_file_layout = QHBoxLayout()
        self.output_file_label = QLabel('Output KML File:')
        self.output_file_label.setStyleSheet(label_style)
        output_file_layout.addWidget(self.output_file_label)
        self.output_file_path_input = QLineEdit()
        self.output_file_path_input.setStyleSheet(lineedit_style)
        output_file_layout.addWidget(self.output_file_path_input)
        self.browse_output_button = QPushButton('Browse')
        self.browse_output_button.clicked.connect(self.browse_output_file)
        self.browse_output_button.setStyleSheet(button_style + button_hover_style)
        output_file_layout.addWidget(self.browse_output_button)
        file_group_layout.addLayout(output_file_layout)

        options_layout = QHBoxLayout()
        self.delimiter_label = QLabel('Delimiter:')
        self.delimiter_label.setStyleSheet(label_style)
        options_layout.addWidget(self.delimiter_label)
        self.delimiter_input = QLineEdit(';')
        self.delimiter_input.textChanged.connect(self.on_file_settings_changed)
        self.delimiter_input.setFixedWidth(30)
        self.delimiter_input.setStyleSheet(lineedit_style)
        options_layout.addWidget(self.delimiter_input)

        self.has_header_checkbox = QCheckBox('File has header')
        self.has_header_checkbox.setChecked(True)
        self.has_header_checkbox.stateChanged.connect(self.on_file_settings_changed)
        self.has_header_checkbox.setStyleSheet(checkbox_style)
        options_layout.addWidget(self.has_header_checkbox)

        self.start_row_label = QLabel('Data starts from row:')
        self.start_row_label.setStyleSheet(label_style)
        options_layout.addWidget(self.start_row_label)
        self.start_row_spinbox = QSpinBox()
        self.start_row_spinbox.setMinimum(1)
        self.start_row_spinbox.setValue(1)
        self.start_row_spinbox.valueChanged.connect(self.on_file_settings_changed)
        self.start_row_spinbox.setStyleSheet(spinbox_style)
        options_layout.addWidget(self.start_row_spinbox)
        file_group_layout.addLayout(options_layout)

        encoding_layout = QHBoxLayout()
        self.encoding_label = QLabel('File encoding:')
        self.encoding_label.setStyleSheet(label_style)
        encoding_layout.addWidget(self.encoding_label)
        self.encoding_group = QButtonGroup(self)
        self.utf8_radio = QRadioButton('UTF-8')
        self.utf8_radio.setChecked(True)
        self.utf8_radio.toggled.connect(self.on_file_settings_changed)
        self.cp1251_radio = QRadioButton('CP1251')
        self.cp1251_radio.toggled.connect(self.on_file_settings_changed)
        self.encoding_group.addButton(self.utf8_radio)
        self.encoding_group.addButton(self.cp1251_radio)
        
        self.utf8_radio.setStyleSheet(radio_button_style)
        self.cp1251_radio.setStyleSheet(radio_button_style)
        encoding_layout.addWidget(self.utf8_radio)
        encoding_layout.addWidget(self.cp1251_radio)
        file_group_layout.addLayout(encoding_layout)
        layout.addWidget(file_group_box)
        
        coord_group_box = QGroupBox("Coordinate and KML Settings")
        coord_group_box.setFont(bold_large_font)
        coord_layout = QVBoxLayout()
        coord_layout.setContentsMargins(10, 15, 10, 15)
        coord_group_box.setLayout(coord_layout)
        self.coord_system_label = QLabel('Coordinate System:')
        self.coord_system_label.setStyleSheet(label_style)
        coord_layout.addWidget(self.coord_system_label)
        self.coord_system_group = QButtonGroup(self)
        coord_system_radio_layout = QHBoxLayout()
        self.wkt_radio = QRadioButton('WKT')
        self.wkt_radio.setChecked(True)
        self.wkt_radio.toggled.connect(self.on_coord_system_changed)
        self.wkt_radio.setStyleSheet(radio_button_style)
        coord_system_radio_layout.addWidget(self.wkt_radio)
        self.coord_system_group.addButton(self.wkt_radio)
        self.lonlat_radio = QRadioButton('Longitude/Latitude')
        self.lonlat_radio.toggled.connect(self.on_coord_system_changed)
        self.lonlat_radio.setStyleSheet(radio_button_style)
        coord_system_radio_layout.addWidget(self.lonlat_radio)
        self.coord_system_group.addButton(self.lonlat_radio)
        coord_system_radio_layout.addStretch(1)
        coord_layout.addLayout(coord_system_radio_layout)
        
        self.wkt_field_layout = QHBoxLayout()
        self.wkt_field_label = QLabel('WKT Field:')
        self.wkt_field_label.setStyleSheet(label_style)
        self.wkt_field_combo = QComboBox()
        self.wkt_field_combo.setStyleSheet(combobox_style)
        self.wkt_field_layout.addWidget(self.wkt_field_label)
        self.wkt_field_layout.addWidget(self.wkt_field_combo)
        coord_layout.addLayout(self.wkt_field_layout)

        self.lon_lat_field_layout = QHBoxLayout()
        self.lon_field_label = QLabel('Longitude Field:')
        self.lon_field_label.setStyleSheet(label_style)
        self.lon_field_combo = QComboBox()
        self.lon_field_combo.setStyleSheet(combobox_style)
        self.lat_field_label = QLabel('Latitude Field:')
        self.lat_field_label.setStyleSheet(label_style)
        self.lat_field_combo = QComboBox()
        self.lat_field_combo.setStyleSheet(combobox_style)
        
        self.lon_lat_field_layout.addWidget(self.lon_field_label)
        self.lon_lat_field_layout.addWidget(self.lon_field_combo)
        self.lon_lat_field_layout.addWidget(self.lat_field_label)
        self.lon_lat_field_layout.addWidget(self.lat_field_combo)
        coord_layout.addLayout(self.lon_lat_field_layout)

        self.lon_field_label.setVisible(False)
        self.lon_field_combo.setVisible(False)
        self.lat_field_label.setVisible(False)
        self.lat_field_combo.setVisible(False)
        
        self.add_label_button = QPushButton('Select KML Label Field')
        self.add_label_button.clicked.connect(self.toggle_kml_label_field)
        self.add_label_button.setStyleSheet(button_style + button_hover_style)
        coord_layout.addWidget(self.add_label_button)

        self.kml_label_field_layout = QHBoxLayout()
        self.kml_label_field_label = QLabel('KML Label Field:')
        self.kml_label_field_label.setStyleSheet(label_style)
        self.kml_label_field_combo = QComboBox()
        self.kml_label_field_combo.setStyleSheet(combobox_style)
        self.kml_label_field_layout.addWidget(self.kml_label_field_label)
        self.kml_label_field_layout.addWidget(self.kml_label_field_combo)
        coord_layout.addLayout(self.kml_label_field_layout)
        
        self.kml_label_field_label.setVisible(False)
        self.kml_label_field_combo.setVisible(False)

        self.use_custom_icon_checkbox = QCheckBox('Use Custom Icon')
        self.use_custom_icon_checkbox.setChecked(False)
        self.use_custom_icon_checkbox.stateChanged.connect(self.toggle_custom_icon_input)
        self.use_custom_icon_checkbox.setStyleSheet(checkbox_style)
        coord_layout.addWidget(self.use_custom_icon_checkbox)
        self.icon_url_layout = QHBoxLayout()
        self.icon_url_label = QLabel('Icon URL:')
        self.icon_url_label.setStyleSheet(label_style)
        self.icon_url_input = QLineEdit('http://maps.google.com/mapfiles/kml/pal2/icon18.png') # Changed default icon URL
        self.icon_url_input.setStyleSheet(lineedit_style)
        self.icon_url_layout.addWidget(self.icon_url_label)
        self.icon_url_layout.addWidget(self.icon_url_input)
        coord_layout.addLayout(self.icon_url_layout)
        self.toggle_custom_icon_input()
        coord_group_box.setLayout(coord_layout)
        layout.addWidget(coord_group_box)
        
        self.on_coord_system_changed() 

        grouping_group_box = QGroupBox("Numerical Grouping Settings")
        grouping_group_box.setFont(bold_large_font)
        grouping_options_layout = QVBoxLayout()
        grouping_options_layout.setContentsMargins(10, 15, 10, 15)
        grouping_group_box.setLayout(grouping_options_layout)
        numerical_field_selection_layout = QHBoxLayout()
        self.numerical_group_label = QLabel('Numerical Grouping Field:')
        self.numerical_group_label.setStyleSheet(label_style)
        numerical_field_selection_layout.addWidget(self.numerical_group_label)
        self.numerical_group_field_combo = QComboBox()
        self.numerical_group_field_combo.currentIndexChanged.connect(self.on_numerical_grouping_field_changed)
        self.numerical_group_field_combo.setStyleSheet(combobox_style)
        numerical_field_selection_layout.addWidget(self.numerical_group_field_combo)
        grouping_options_layout.addLayout(numerical_field_selection_layout)
        num_groups_layout = QHBoxLayout()
        self.num_groups_label = QLabel('Number of Groups:')
        self.num_groups_label.setStyleSheet(label_style)
        num_groups_layout.addWidget(self.num_groups_label)
        self.num_groups_spinbox = QSpinBox()
        self.num_groups_spinbox.setMinimum(1)
        self.num_groups_spinbox.setMaximum(20)
        self.num_groups_spinbox.setValue(3)
        self.num_groups_spinbox.valueChanged.connect(self.on_numerical_grouping_field_changed)
        self.num_groups_spinbox.setStyleSheet(spinbox_style)
        num_groups_layout.addWidget(self.num_groups_spinbox)
        grouping_options_layout.addLayout(num_groups_layout)
        end_color_layout = QHBoxLayout()
        self.end_color_label = QLabel('End Color for Gradient:')
        self.end_color_label.setStyleSheet(label_style)
        end_color_layout.addWidget(self.end_color_label)
        self.end_color_button = QPushButton()
        self.end_color_button.clicked.connect(self.pick_end_color)
        self.end_color_button.setFixedSize(20, 20)
        end_color_layout.addWidget(self.end_color_button)
        grouping_options_layout.addLayout(end_color_layout)
        self.update_end_color_button()
        self.numerical_color_display_layout = QVBoxLayout()
        self.numerical_color_label = QLabel('Numerical Group Ranges and Colors:')
        self.numerical_color_label.setStyleSheet(label_style)
        self.numerical_color_display_layout.addWidget(self.numerical_color_label)
        grouping_options_layout.addLayout(self.numerical_color_display_layout)
        layout.addWidget(grouping_group_box)

        # Data filtering controls
        filter_group_box = QGroupBox("Data Filtering")
        filter_group_box.setFont(bold_large_font)
        filter_layout = QHBoxLayout()
        self.filter_label = QLabel('Filter formula:')
        self.filter_label.setStyleSheet(label_style)
        filter_layout.addWidget(self.filter_label)
        self.filter_input = QLineEdit()
        self.filter_input.setStyleSheet(lineedit_style)
        self.filter_input.setPlaceholderText("e.g., Column=Value and Other>5")
        filter_layout.addWidget(self.filter_input)
        self.apply_filter_button = QPushButton('Apply Filter')
        self.apply_filter_button.setStyleSheet(button_style + button_hover_style)
        self.apply_filter_button.clicked.connect(self.apply_filter)
        filter_layout.addWidget(self.apply_filter_button)
        filter_group_box.setLayout(filter_layout)
        layout.addWidget(filter_group_box)

        self.data_table = QTableWidget()
        self.data_table.setMinimumHeight(300)
        self.data_table.setStyleSheet(table_style)
        self.data_table.horizontalHeader().sectionDoubleClicked.connect(self.on_header_double_clicked)
        layout.addWidget(self.data_table)

        self.generate_button = QPushButton('Generate KML')
        self.generate_button.clicked.connect(self.generate_kml)
        self.generate_button.setStyleSheet(button_style + button_hover_style)
        layout.addWidget(self.generate_button)

        main_container.setLayout(layout)
        scroll_area.setWidget(main_container)

    def update_file_options_state(self, is_excel):
        """Включает или отключает опции файла в зависимости от его типа."""
        self.delimiter_input.setEnabled(not is_excel)
        self.delimiter_label.setEnabled(not is_excel)
        self.utf8_radio.setEnabled(not is_excel)
        self.cp1251_radio.setEnabled(not is_excel)
        self.encoding_label.setEnabled(not is_excel)

    def browse_file(self):
        """Открывает диалог выбора файла и загружает данные."""
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select Data File', '',
                                                   'All Supported Files (*.txt *.csv *.xlsx *.xls *.xlsm);;'
                                                   'Text Files (*.txt *.csv);;'
                                                   'Excel Files (*.xlsx *.xls *.xlsm);;'
                                                   'All Files (*)')
        if file_name:
            self.file_path_input.setText(file_name)
            is_excel = file_name.lower().endswith(('.xlsx', '.xls', '.xlsm'))
            self.update_file_options_state(is_excel)
            self.load_data(file_name)

    def browse_output_file(self):
        """Открывает диалог выбора KML файла для сохранения."""
        file_name, _ = QFileDialog.getSaveFileName(self, 'Save KML File', '', 'KML Files (*.kml);;All Files (*)')
        if file_name:
            self.output_file_path_input.setText(file_name)

    def on_file_settings_changed(self):
        """Вызывает перезагрузку данных при изменении настроек файла."""
        file_path = self.file_path_input.text()
        if file_path:
            self.load_data(file_path)
            
    def load_data(self, file_path):
        """Загружает данные из файла с учетом выбранных параметров."""
        self.data = []
        self.filtered_data = []
        self.headers = []
        self.field_types = {}
        self.manual_group_bounds = {} # Clear manual bounds on new file load

        is_excel = file_path.lower().endswith(('.xlsx', '.xls', '.xlsm'))
        self.update_file_options_state(is_excel)

        has_header = self.has_header_checkbox.isChecked()
        start_row = self.start_row_spinbox.value() - 1

        try:
            if is_excel:
                header_row = 0 if has_header else None
                engine = 'openpyxl' if file_path.lower().endswith(('.xlsx', '.xlsm')) else 'xlrd'
                df = pd.read_excel(file_path, header=header_row, skiprows=start_row, engine=engine)
                
                if not has_header:
                    df.columns = [f'Column {i}' for i in range(len(df.columns))]

                df = df.fillna('')
                self.headers = df.columns.astype(str).tolist()
                self.data = df.values.tolist()

                self._auto_cast_numeric()
                self.filtered_data = [row[:] for row in self.data]


                self._auto_cast_numeric()
                self.filtered_data = [row[:] for row in self.data]

                self.filtered_data = self.data[:]


                if hasattr(self, 'filter_input'):
                    self.filter_input.setText('')
            else:
                delimiter = self.delimiter_input.text()
                self.encoding = 'utf-8' if self.utf8_radio.isChecked() else 'cp1251'
                with open(file_path, 'r', encoding=self.encoding, errors='ignore') as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    all_lines = list(reader)

                    data_start_index = start_row
                    if has_header:
                        if start_row < len(all_lines):
                            self.headers = all_lines[start_row]
                            data_start_index += 1
                    
                    if data_start_index < len(all_lines):
                        self.data = all_lines[data_start_index:]

                        self._auto_cast_numeric()
                        self.filtered_data = [row[:] for row in self.data]


                        self._auto_cast_numeric()
                        self.filtered_data = [row[:] for row in self.data]

                        self.filtered_data = self.data[:]


                        if hasattr(self, 'filter_input'):
                            self.filter_input.setText('')
                    else:
                        self.data = []
                        self.filtered_data = []
                        if hasattr(self, 'filter_input'):
                            self.filter_input.setText('')

                    if not has_header and self.data:
                        self.headers = [f'Column {i}' for i in range(len(self.data[0]))]
                    elif not has_header and not self.data and all_lines:
                        if start_row < len(all_lines) and len(all_lines[start_row]) > 0:
                            self.headers = [f'Column {i}' for i in range(len(all_lines[start_row]))]
                        else:
                            self.headers = []
                    elif not has_header and not self.data:
                        self.headers = []
            
            inferred = self._infer_field_types(self.data, self.headers)
            for field, f_type in inferred.items():
                self.field_types[field] = f_type
            
            self.preview_data()
            self.update_field_combos()
            self.on_numerical_grouping_field_changed()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ошибка загрузки файла: {e}\n\nДля файлов Excel убедитесь, что установлены 'pandas' и 'openpyxl'.")
            self.data, self.filtered_data, self.headers, self.field_types = [], [], [], {}
            self.preview_data()
            self.update_field_combos()

    def generate_kml(self):
        """Генерирует KML-файл на основе текущих настроек."""
        output_file = self.output_file_path_input.text()
        if not output_file:
            QMessageBox.warning(self, "Warning", "Пожалуйста, укажите выходной KML-файл.")
            return
        if not self.filtered_data:
            QMessageBox.warning(self, "Warning", "Нет данных для генерации KML.")
            return

        kml = simplekml.Kml()

        # Получение индексов полей
        use_wkt = self.wkt_radio.isChecked()
        field_indices = {name: i for i, name in enumerate(self.headers)}
        wkt_idx = field_indices.get(self.wkt_field_combo.currentText(), -1)
        lon_idx = field_indices.get(self.lon_field_combo.currentText(), -1)
        lat_idx = field_indices.get(self.lat_field_combo.currentText(), -1)
        
        # Check if KML label field is visible and get its index
        label_idx = -1
        if self.kml_label_field_combo.isVisible():
            label_idx = field_indices.get(self.kml_label_field_combo.currentText(), -1)


        num_group_idx = field_indices.get(self.numerical_group_field_combo.currentText(), -1)

        if use_wkt and wkt_idx == -1:
            QMessageBox.critical(self, "Error", "Выбранное поле WKT не найдено.")
            return
        if not use_wkt and (lon_idx == -1 or lat_idx == -1):
            QMessageBox.critical(self, "Error", "Выбранные поля долготы/широты не найдены.")
            return

        kml_folders = {}
        grouping_active = num_group_idx != -1 and self.groups
        if grouping_active:
            for group in self.groups:
                kml_folders[group['label']] = kml.newfolder(name=group['label'])

        try:
            for i, row in enumerate(self.filtered_data):
                # Skip rows where the selected numerical grouping field is empty
                if num_group_idx != -1:
                    if num_group_idx >= len(row) or str(row[num_group_idx]).strip() == '':
                        continue

                target_container = kml
                assigned_group = None
                kml_object = None

                if grouping_active and num_group_idx != -1 and num_group_idx < len(row):
                    try:
                        value = float(str(row[num_group_idx]).replace(',', '.'))
                        for group in self.groups:
                            lower, upper = group['range']
                            if (lower <= value < upper) or (group == self.groups[-1] and np.isclose(value, upper)):
                                target_container = kml_folders[group['label']]
                                assigned_group = group
                                break
                    except (ValueError, TypeError, IndexError):
                        pass

                label_text = ''
                if self.kml_label_field_combo.isVisible() and label_idx != -1 and label_idx < len(row):
                    label_text = str(row[label_idx])

                if use_wkt:
                    if wkt_idx != -1 and wkt_idx < len(row):
                        try:
                            geom = wkt.loads(str(row[wkt_idx]))
                            if geom.geom_type == 'Point':
                                kml_object = target_container.newpoint(name=label_text, coords=[(geom.x, geom.y)])
                            elif geom.geom_type == 'LineString':
                                kml_object = target_container.newlinestring(name=label_text, coords=list(geom.coords))
                            elif geom.geom_type == 'Polygon':
                                kml_object = target_container.newpolygon(name=label_text, outerboundaryis=list(geom.exterior.coords))
                        except Exception as e:
                            print(f"Row {i+1} WKT Error: {e}"); continue
                else:
                    if lon_idx != -1 and lat_idx != -1 and lon_idx < len(row) and lat_idx < len(row):
                        try:
                            lon = float(str(row[lon_idx]).replace(',', '.'))
                            lat = float(str(row[lat_idx]).replace(',', '.'))
                            kml_object = target_container.newpoint(name=label_text, coords=[(lon, lat)])
                        except (ValueError, TypeError):
                            print(f"Row {i+1} Lon/Lat Error"); continue

                if not kml_object:
                    continue

                if assigned_group:
                    color = self.group_colors.get(assigned_group['label'])
                    if color:
                        kml_color = simplekml.Color.rgb(color.red(), color.green(), color.blue())
                        if isinstance(kml_object, simplekml.Point):
                            kml_object.style.iconstyle.color = kml_color
                        elif isinstance(kml_object, simplekml.LineString):
                            kml_object.style.linestyle.color = kml_color
                        elif isinstance(kml_object, simplekml.Polygon):
                            kml_object.style.polystyle.color = kml_color
                            kml_object.style.linestyle.color = kml_color

                if isinstance(kml_object, simplekml.Point):
                    use_custom_icon = self.use_custom_icon_checkbox.isChecked()
                    custom_icon_url = self.icon_url_input.text()
                    if use_custom_icon and custom_icon_url:
                        kml_object.style.iconstyle.icon.href = custom_icon_url
                    else:
                        kml_object.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/paddle/wht-blank.png' # Changed default icon URL
            kml.save(output_file)
            QMessageBox.information(self, "KML Generated", f"KML-файл '{output_file}' успешно создан!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ошибка при генерации KML: {e}")

    def _auto_cast_numeric(self):
        """Convert columns with only numeric-looking values to int or float."""
        if not self.data or not self.headers:
            return

        num_cols = len(self.headers)
        int_pattern = re.compile(r"^-?\d+$")
        float_pattern = re.compile(r"^-?\d+(?:[.,]\d+)?$")

        missing_vals = MISSING_VALS


        for col_idx in range(num_cols):
            is_int_col = True
            is_numeric_col = True
            for row in self.data:
                if col_idx >= len(row):
                    continue
                val = str(row[col_idx]).strip()

                if val.lower() in missing_vals:

                    if val == "":

                        continue
                val_dot = val.replace(",", ".")
                if int_pattern.fullmatch(val_dot):
                    continue
                elif float_pattern.fullmatch(val_dot):
                    is_int_col = False
                else:
                    is_numeric_col = False
                    break

            if not is_numeric_col:
                continue

            for row in self.data:
                if col_idx >= len(row):
                    continue
                val = str(row[col_idx]).strip()

                if val.lower() in missing_vals:

                    if val == "":

                        row[col_idx] = ""
                        continue
                val_dot = val.replace(",", ".")
                try:
                    row[col_idx] = int(val_dot) if is_int_col else float(val_dot)
                except ValueError:
                    pass

            self.field_types[self.headers[col_idx]] = "int" if is_int_col else "float"

    def _infer_field_types(self, data, headers):
        """
        Infers the data types for each field based on a sample of the data.
        Determines if a field is 'float', 'geometry', or 'varchar'.
        """
        if not data:
            return {}
        
        num_columns = len(headers) if headers else (len(data[0]) if data else 0)

        inferred_types = {}
        current_fields = headers if headers else [f'Column {i}' for i in range(num_columns)]

        for i in range(num_columns):
            field_name = current_fields[i]
            is_numerical = True
            is_wkt = False
            
            sample_values = []
            for r_idx in range(min(len(data), 100)):
                if i < len(data[r_idx]):
                    value = str(data[r_idx][i]).strip()
                    if value.lower() in MISSING_VALS or value == "":
                        continue
                    sample_values.append(value)

            if not sample_values:
                inferred_types[field_name] = 'varchar'
                continue

            for value in sample_values:
                try:
                    geom = wkt.loads(value)
                    if geom.geom_type in ['Point', 'LineString', 'Polygon', 'MultiPoint', 'MultiLineString', 'MultiPolygon', 'GeometryCollection']:
                        is_wkt = True
                        break
                except Exception:
                    pass
            
            if is_wkt:
                inferred_types[field_name] = 'geometry'
                continue

            for value in sample_values:
                try:
                    float(value.replace(',', '.'))
                except (ValueError, TypeError):
                    is_numerical = False
                    break
            
            if is_numerical:
                # Further refine to 'int' if all sampled numerical values are integers
                is_integer = True
                for value in sample_values:
                    try:
                        if float(value.replace(',', '.')) != int(float(value.replace(',', '.'))):
                            is_integer = False
                            break
                    except (ValueError, TypeError):
                        is_integer = False # Should not happen if it's already numerical, but defensive
                        break
                inferred_types[field_name] = 'int' if is_integer else 'float'
            else:
                inferred_types[field_name] = 'varchar'
        
        return inferred_types


    def on_header_double_clicked(self, column_index):
        """
        Handles double-click on a table header section.
        Opens a QComboBox to select the data type for the column.
        """
        if not self.headers or column_index >= len(self.headers):
            return

        # Close any existing combo box if open
        self.close_header_combo()

        header_pos = self.data_table.horizontalHeader().sectionViewportPosition(column_index)
        header_width = self.data_table.horizontalHeader().sectionSize(column_index)
        header_height = self.data_table.horizontalHeader().height()
        
        header_rect = QRect(header_pos, 0, header_width, header_height) 

        combo = QComboBox(self) 
        data_types = ['auto', 'int', 'float', 'varchar', 'date', 'geometry', 'text'] 
        combo.addItems(data_types)
        combo.setStyleSheet("QComboBox { background-color: #DDDDDD; border: 1px solid #AAAAAA; padding: 1px; }")

        field_name = self.headers[column_index]
        current_type = self.field_types.get(field_name, 'auto')
        combo.setCurrentText(current_type)
        
        global_pos = self.data_table.horizontalHeader().mapToGlobal(header_rect.topLeft())
        local_pos = self.mapFromGlobal(global_pos)

        combo.setGeometry(local_pos.x(), local_pos.y(), header_rect.width(), header_rect.height())
        
        combo.currentIndexChanged.connect(lambda index, col=column_index, cb=combo: self.update_field_type_from_header_combo(col, cb.currentText()))
        
        # Install event filter on the combo box itself
        combo.installEventFilter(self)

        self.current_header_combo = combo
        self.current_header_combo_column = column_index
        self.current_header_combo.show()
        # Corrected: Use Qt.FocusReason.PopupFocusReason instead of Qt.FocusReason.Popup
        self.current_header_combo.setFocus(Qt.FocusReason.PopupFocusReason) # Set focus to the combo box, hinting it's a popup

        # Add a flag to ensure it's deleted when closed
        self.current_header_combo.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

    def eventFilter(self, obj, event):
        """
        Event filter to detect when the QComboBox loses focus.
        """
        if obj == self.current_header_combo:
            if event.type() == event.Type.FocusOut:
                # Check if the combo box's popup view is still visible.
                # If it is, the focus might have moved internally to the popup,
                # and we should not close the combo box yet.
                if self.current_header_combo.view().isVisible():
                    return False # Let the event be processed normally, don't close yet
                else:
                    self.close_header_combo()
                    return True # Event handled, stop propagation
            # Handle other events if needed, e.g., closing on Escape key
            elif event.type() == event.Type.KeyPress and event.key() == Qt.Key.Key_Escape:
                self.close_header_combo()
                return True # Event handled
        return super().eventFilter(obj, event)

    def close_header_combo(self):
        """Closes and deletes the currently active header combo box."""
        if self.current_header_combo:
            # Disconnect signals to prevent errors during deletion
            try:
                self.current_header_combo.currentIndexChanged.disconnect()
            except TypeError:
                pass # Already disconnected or not connected
            
            # Remove the event filter before deleting, although WA_DeleteOnClose handles this
            self.current_header_combo.removeEventFilter(self)
            
            self.current_header_combo.close() # Close will trigger WA_DeleteOnClose
            self.current_header_combo = None
            self.current_header_combo_column = -1
            
            # Re-update preview to ensure header text is correct and consistent
            self.preview_data()


    def update_field_type_from_header_combo(self, column_index, new_type):
        """
        Updates the field type based on the selection from the QComboBox in the header.
        This function is called when the combo box's selected index changes.
        """
        if column_index == -1 or column_index >= len(self.headers):
            return

        field_name = self.headers[column_index]
        self.field_types[field_name] = new_type
        
        if new_type == 'auto':
            inferred_type = self._infer_field_types_for_column(column_index)
            self.field_types[field_name] = inferred_type
            
        self.update_field_combos()
        self.on_numerical_grouping_field_changed()
        
    def _infer_field_types_for_column(self, column_index):
        """
        Infers the data type for a single column.
        Helper function for 'auto' type selection.
        """
        if not self.data or column_index >= len(self.headers):
            return 'varchar' # Default if no data or invalid column

        is_numerical = True
        is_wkt = False
        
        sample_values = []
        for r_idx in range(min(len(self.data), 100)):
            if column_index < len(self.data[r_idx]):
                value = str(self.data[r_idx][column_index]).strip()
                if value.lower() in MISSING_VALS or value == "":
                    continue
                sample_values.append(value)

        if not sample_values:
            return 'varchar'

        for value in sample_values:
            try:
                geom = wkt.loads(value)
                if geom.geom_type in ['Point', 'LineString', 'Polygon', 'MultiPoint', 'MultiLineString', 'MultiPolygon', 'GeometryCollection']:
                    is_wkt = True
                    break
            except Exception:
                pass
        
        if is_wkt:
            return 'geometry'

        for value in sample_values:
            try:
                float(value.replace(',', '.'))
            except (ValueError, TypeError):
                is_numerical = False
                break
        
        if is_numerical:
            is_integer = True
            for value in sample_values:
                try:
                    if float(value.replace(',', '.')) != int(float(value.replace(',', '.'))):
                        is_integer = False
                        break
                except (ValueError, TypeError):
                    is_integer = False
                    break
            return 'int' if is_integer else 'float'
        else:
            return 'varchar'


    def update_field_combos(self):
        """
        Updates the available fields in all QComboBox widgets based on loaded headers
        and inferred field types.
        """
        if not self.data and not self.headers:
            for combo in [self.wkt_field_combo, self.lon_field_combo, self.lat_field_combo, self.numerical_group_field_combo, self.kml_label_field_combo]:
                combo.clear()
            return

        num_columns = len(self.headers) if self.headers else (len(self.data[0]) if self.data else 0)
        all_fields = self.headers[:]
        
        numerical_fields = [field for field in all_fields if self.field_types.get(field) in ['int', 'float']]

        combos = [self.wkt_field_combo, self.lon_field_combo, self.lat_field_combo, self.numerical_group_field_combo, self.kml_label_field_combo]
        current_texts = [c.currentText() for c in combos]

        for c in combos:
            c.clear()

        for c in [self.wkt_field_combo, self.lon_field_combo, self.lat_field_combo, self.kml_label_field_combo]:
            c.addItems(all_fields)
        self.numerical_group_field_combo.addItems(numerical_fields)

        if current_texts[0] in all_fields:
            self.wkt_field_combo.setCurrentText(current_texts[0])
        if current_texts[1] in all_fields:
            self.lon_field_combo.setCurrentText(current_texts[1])
        if current_texts[2] in all_fields:
            self.lat_field_combo.setCurrentText(current_texts[2])
        
        if current_texts[3] in numerical_fields:
            self.numerical_group_field_combo.setCurrentText(current_texts[3])
        elif numerical_fields:
            self.numerical_group_field_combo.setCurrentText(numerical_fields[0])
        else:
            self.numerical_group_field_combo.setCurrentText('')


        if current_texts[4] in all_fields:
            self.kml_label_field_combo.setCurrentText(current_texts[4])
        elif all_fields:
            self.kml_label_field_combo.setCurrentText(all_fields[0])


    def apply_filter(self):
        """Apply the filter expression from the input field to the data."""
        formula = self.filter_input.text().strip()
        if not formula:
            self.filtered_data = self.data[:]
        else:
            try:
                df = pd.DataFrame(self.data, columns=self.headers)

                numeric_cols = set()
                for col, t in self.field_types.items():
                    if t in ["int", "float"]:
                        numeric_cols.add(col)

                pattern = r"([\w ]+)\s*(==|!=|>=|<=|>|<)\s*([^&|]+)"
                for col, op, val in re.findall(pattern, formula):
                    col = col.strip()
                    val_clean = val.strip().strip("'\"")
                    if op in [">", "<", ">=", "<="] or re.fullmatch(r"-?\d+(\.\d+)?", val_clean):
                        numeric_cols.add(col)

                df_numeric = df.copy()
                for col in numeric_cols:
                    if col in df_numeric.columns:

                        df_numeric = df.copy()
                        for col, t in self.field_types.items():
                            if t in ["int", "float"] and col in df_numeric.columns:



                                df_numeric[col] = pd.to_numeric(
                                    df_numeric[col].astype(str).str.replace(",", "."),
                                    errors="coerce",
                                )

                parsed = parse_filter_expression(formula)
                filtered_indices = df_numeric.query(parsed).index
                self.filtered_data = df.loc[filtered_indices].values.tolist()





                parsed = parse_filter_expression(formula)
                filtered_df = df.query(parsed)

                filtered_df = df.query(formula)

                self.filtered_data = filtered_df.values.tolist()




            except Exception as e:
                QMessageBox.critical(self, "Error", f"Invalid filter: {e}")
                return
        self.preview_data()
        self.on_numerical_grouping_field_changed()


    def preview_data(self):
        """
        Displays a preview of the loaded data in a QTableWidget.
        Headers will show field name and its inferred/selected type.
        """
        self.data_table.clear()
        if not self.filtered_data and not self.headers:
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            return

        num_columns = len(self.headers) if self.headers else (len(self.filtered_data[0]) if self.filtered_data else 0)
        
        if num_columns == 0:
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            return

        self.data_table.setColumnCount(num_columns)
        
        header_labels = []
        for col_idx, header_name in enumerate(self.headers):
            inferred_type = self.field_types.get(header_name, 'auto')
            header_labels.append(f"{header_name}\n({inferred_type})")
        self.data_table.setHorizontalHeaderLabels(header_labels)

        preview_rows = self.filtered_data[:20]
        self.data_table.setRowCount(len(preview_rows))

        for i, row in enumerate(preview_rows):
            for j, item in enumerate(row):
                if j < self.data_table.columnCount():
                    self.data_table.setItem(i, j, QTableWidgetItem(str(item)))

        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


    def on_coord_system_changed(self):
        """Переключает видимость полей WKT или Longitude/Latitude."""
        use_wkt = self.wkt_radio.isChecked()

        for i in range(self.wkt_field_layout.count()):
            widget = self.wkt_field_layout.itemAt(i).widget()
            if widget:
                widget.setVisible(use_wkt)

        for i in range(self.lon_lat_field_layout.count()):
            widget = self.lon_lat_field_layout.itemAt(i).widget()
            if widget:
                widget.setVisible(not use_wkt)

    def toggle_kml_label_field(self):
        """Показать или скрыть поле для выбора поля KML метки."""
        is_visible = self.kml_label_field_label.isVisible() 
        
        self.kml_label_field_label.setVisible(not is_visible)
        self.kml_label_field_combo.setVisible(not is_visible)
        
        if not is_visible:
            self.add_label_button.setText('Hide KML Label Field')
        else:
            self.add_label_button.setText('Select KML Label Field')

    def toggle_custom_icon_input(self):
        """Показать или скрыть поле ввода URL иконки."""
        is_checked = self.use_custom_icon_checkbox.isChecked()
        self.icon_url_label.setVisible(is_checked)
        self.icon_url_input.setVisible(is_checked)
    
    def pick_end_color(self):
        """Открывает диалог выбора цвета для конечного цвета градиента."""
        color = QColorDialog.getColor(self.end_color, self, "Select End Color")
        if color.isValid():
            self.end_color = color
            self.update_end_color_button()
            self.on_numerical_grouping_field_changed()

    def update_end_color_button(self):
        """Обновляет цвет кнопки, отображающей конечный цвет."""
        self.end_color_button.setStyleSheet(f"background-color: {self.end_color.name()}; border: 1px solid #888888;")

    def on_numerical_grouping_field_changed(self):
        """
        Handles the change of the numerical grouping field or number of groups.
        Resets manual bounds and recalculates group ranges and colors.
        """
        self.manual_group_bounds = {} 
        self.groups = []

        selected_field = self.numerical_group_field_combo.currentText()
        if not selected_field or not self.filtered_data or selected_field not in self.headers:
            self.update_group_display()
            return

        col_index = self.headers.index(selected_field)
        numerical_values = []
        for row in self.filtered_data:
            if col_index < len(row):
                try:
                    numerical_values.append(float(str(row[col_index]).replace(',', '.')))
                except (ValueError, TypeError):
                    pass

        if not numerical_values:
            self.update_group_display()
            return

        num_groups = self.num_groups_spinbox.value()

        # Determine the full range of values, explicitly including zeros
        min_val = min(numerical_values)
        max_val = max(numerical_values)

        unique_values = sorted(set(numerical_values))

        if len(unique_values) > 2 and num_groups > 1:
            bins = jenks_breaks(numerical_values, num_groups)

            if len(set(bins)) < len(bins) or len(bins) != num_groups + 1:


                if len(set(bins)) < len(bins) or len(bins) != num_groups + 1:


                    if len(set(bins)) < len(bins) or len(bins) != num_groups + 1:


                        if len(set(bins)) < len(bins) or len(bins) != num_groups + 1:


                            if len(set(bins)) < len(bins) or len(bins) != num_groups + 1:

                                if len(set(bins)) < len(bins):

                                    if min_val == max_val:
                                        bins = [min_val, min_val + 1] if num_groups > 1 else [min_val, min_val]
                                    else:
                                        bins = np.linspace(min_val, max_val, num_groups + 1)
        else:
            if min_val == max_val:
                bins = [min_val, min_val + 1] if num_groups > 1 else [min_val, min_val]
            else:
                bins = np.linspace(min_val, max_val, num_groups + 1)

        bins = list(bins)

        if len(bins) != num_groups + 1:
            if min_val == max_val:
                bins = [min_val, min_val + 1] if num_groups > 1 else [min_val, min_val]
            else:
                bins = np.linspace(min_val, max_val, num_groups + 1)
            bins = list(bins)


        bins = list(bins)


        if len(bins) != num_groups + 1:
            if min_val == max_val:
                bins = [min_val, min_val + 1] if num_groups > 1 else [min_val, min_val]
            else:
                bins = np.linspace(min_val, max_val, num_groups + 1)


            bins = list(bins)

            bins = list(bins)


        bins = list(bins)


        if len(bins) != num_groups + 1:
            if min_val == max_val:
                bins = [min_val, min_val + 1] if num_groups > 1 else [min_val, min_val]
            else:
                bins = np.linspace(min_val, max_val, num_groups + 1)


        if len(bins) != num_groups + 1:
            if min_val == max_val:
                bins = [min_val, min_val + 1] if num_groups > 1 else [min_val, min_val]
            else:
                bins = np.linspace(min_val, max_val, num_groups + 1)
                
        self.groups = []
        start_color = QColor(255, 255, 255)
        end_color = self.end_color

        for i in range(num_groups):
            idx = i if i < len(bins) else len(bins) - 1
            lower_bound = bins[idx]
            upper_bound = bins[idx + 1] if idx + 1 < len(bins) else bins[-1]
            
            if i == num_groups - 1 and numerical_values:
                upper_bound = max(numerical_values)

            if i in self.manual_group_bounds:
                lower_bound = self.manual_group_bounds[i].get('lower', lower_bound)
                upper_bound = self.manual_group_bounds[i].get('upper', upper_bound)
            
            if i > 0 and (i-1) in self.manual_group_bounds:
                prev_upper = self.manual_group_bounds[i-1].get('upper')
                if prev_upper is not None:
                    lower_bound = prev_upper

            r = start_color.red() + i * (end_color.red() - start_color.red()) // (num_groups - 1) if num_groups > 1 else start_color.red()
            g = start_color.green() + i * (end_color.green() - start_color.green()) // (num_groups - 1) if num_groups > 1 else start_color.green()
            b = start_color.blue() + i * (end_color.blue() - start_color.blue()) // (num_groups - 1) if num_groups > 1 else start_color.blue()
            group_color = QColor(r, g, b)

            label = f'{lower_bound:.2f} - {upper_bound:.2f}'
            self.groups.append({
                'label': label,
                'range': [lower_bound, upper_bound],
                'color': group_color
            })
            self.group_colors[label] = group_color

        self.update_group_display()

    def update_group_display(self):
        """
        Updates the display of numerical groups, their colors, and item counts.
        Dynamically creates QLineEdit for bounds and QLabel for counts.
        """
        while self.numerical_color_display_layout.count() > 1:
            child = self.numerical_color_display_layout.takeAt(1)
            if child.widget():
                child.widget().deleteLater()
            elif child.layout():
                self.clear_layout(child.layout())

        if not self.groups:
            no_groups_label = QLabel("No numerical groups defined or data unavailable.")
            no_groups_label.setStyleSheet("QLabel { color: #555555;margin-left: 10px; }")
            self.numerical_color_display_layout.addWidget(no_groups_label)
            return

        selected_field = self.numerical_group_field_combo.currentText()
        numerical_values = []
        if selected_field and self.filtered_data and self.headers and selected_field in self.headers:
            col_index = self.headers.index(selected_field)
            for row in self.filtered_data:
                if col_index < len(row):
                    try:
                        numerical_values.append(float(str(row[col_index]).replace(',', '.')))
                    except (ValueError, TypeError):
                        pass

        for i, group in enumerate(self.groups):
            group_layout = QHBoxLayout()
            color_swatch = QLabel()
            color_swatch.setFixedSize(20, 20)
            color_swatch.setStyleSheet(f"background-color: {group['color'].name()}; border: 1px solid #888888;")
            group_layout.addWidget(color_swatch)

            lower_bound_label = QLabel(f"{group['range'][0]:.2f} - ")
            lower_bound_label.setStyleSheet("QLabel { color: #333333;}")
            group_layout.addWidget(lower_bound_label)

            upper_bound_input = QLineEdit(f"{group['range'][1]:.2f}")
            upper_bound_input.setFixedWidth(80)
            upper_bound_input.setStyleSheet("QLineEdit { background-color: #EEEEEE; border: 1px solid #CCCCCC; padding: 3px; }")
            
            if i == len(self.groups) - 1:
                upper_bound_input.setReadOnly(True)
                upper_bound_input.setStyleSheet("QLineEdit { background-color: #E0E0E0; border: 1px solid #CCCCCC; padding: 3px; color: #888888; }")
            else:
                upper_bound_input.editingFinished.connect(lambda idx=i, sender=upper_bound_input: self.on_group_bound_edited(idx, sender))
            
            group_layout.addWidget(upper_bound_input)

            item_count = 0
            lower, upper = group['range']
            for val in numerical_values:
                if i == len(self.groups) - 1:
                    if lower <= val <= upper:
                        item_count += 1
                else:
                    if lower <= val < upper:
                        item_count += 1
            
            count_label = QLabel(f" ({item_count} items)")
            count_label.setStyleSheet("QLabel { color: #555555;font-size: 9px; }")
            group_layout.addWidget(count_label)
            
            group_layout.addStretch(1)
            self.numerical_color_display_layout.addLayout(group_layout)

    def clear_layout(self, layout):
        """
        Recursively clears all widgets and sub-layouts from a given layout.
        """
        if layout is not None:
            while layout.count():
                child = layout.takeAt(0)
                if child.widget() is not None:
                    child.widget().deleteLater()
                elif child.layout() is not None:
                    self.clear_layout(child.layout())

    def on_group_bound_edited(self, group_index, sender):
        """
        Handles manual editing of a group's upper boundary.
        Adjusts the upper boundary of the current group and the lower boundary of the next group.
        Includes validation to prevent invalid boundary settings.
        """
        if group_index == len(self.groups) - 1:
            QMessageBox.warning(self, "Invalid Action", "The upper bound of the last group cannot be manually edited.")
            sender.setText(f"{self.groups[group_index]['range'][1]:.2f}")
            return

        try:
            new_value = float(sender.text().replace(',', '.'))
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid number for the boundary.")
            sender.setText(f"{self.groups[group_index]['range'][1]:.2f}")
            return

        current_lower_bound = self.groups[group_index]['range'][0]
        if new_value <= current_lower_bound:
            QMessageBox.warning(self, "Invalid Boundary", f"Upper bound must be greater than the lower bound ({current_lower_bound:.2f}).")
            sender.setText(f"{self.groups[group_index]['range'][1]:.2f}")
            return
        
        if group_index + 1 < len(self.groups):
            next_upper_bound = self.groups[group_index + 1]['range'][1]
            if new_value >= next_upper_bound:
                QMessageBox.warning(self, "Invalid Boundary", f"Upper bound must be less than the next group's upper bound ({next_upper_bound:.2f}).")
                sender.setText(f"{self.groups[group_index]['range'][1]:.2f}")
                return

        if group_index not in self.manual_group_bounds:
            self.manual_group_bounds[group_index] = {}
        self.manual_group_bounds[group_index]['upper'] = new_value

        self.groups[group_index]['range'][1] = new_value
        
        if group_index + 1 < len(self.groups):
            self.groups[group_index + 1]['range'][0] = new_value
            if group_index + 1 not in self.manual_group_bounds:
                self.manual_group_bounds[group_index + 1] = {}
            self.manual_group_bounds[group_index + 1]['lower'] = new_value

        for i, group in enumerate(self.groups):
            manual_lower = self.manual_group_bounds.get(i, {}).get('lower')
            lower_display = manual_lower if manual_lower is not None else group['range'][0]

            manual_upper = self.manual_group_bounds.get(i, {}).get('upper')
            upper_display = manual_upper if manual_upper is not None else group['range'][1]

            group['label'] = f"{lower_display:.2f} - {upper_display:.2f}"
            
            num_groups = len(self.groups)
            start_color = QColor(255, 255, 255)
            end_color = self.end_color
            r = start_color.red() + i * (end_color.red() - start_color.red()) // (num_groups - 1) if num_groups > 1 else start_color.red()
            g = start_color.green() + i * (end_color.green() - start_color.green()) // (num_groups - 1) if num_groups > 1 else start_color.green()
            b = start_color.blue() + i * (end_color.blue() - start_color.blue()) // (num_groups - 1) if num_groups > 1 else start_color.blue()
            group_color = QColor(r, g, b)
            self.groups[i]['color'] = group_color
            self.group_colors[group['label']] = group_color

        self.update_group_display()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = KmlGeneratorApp()
    ex.show()
    sys.exit(app.exec())
