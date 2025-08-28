import sys
import csv

# Allow very large geometry strings when reading CSV files
csv.field_size_limit(1000000)
import simplekml
from shapely.geometry import Point
from shapely import wkt
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QFileDialog, QLineEdit, QComboBox, QColorDialog,
                             QCheckBox, QSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
                             QMessageBox, QRadioButton, QButtonGroup, QGroupBox, QScrollArea)
from PyQt6.QtGui import QColor, QFont, QStandardItemModel, QStandardItem
from PyQt6.QtCore import Qt, QPoint, QRect, pyqtSignal, QEvent

import numpy as np
import pandas as pd
import re
import colorsys


MISSING_VALS = {"", "null", "none", "nan", "na", "n/a"}

WKT_FIELD_NAMES = {
    "wkt", "geom", "geometry", "thegeom", "shape", "geomwkt",
    "geometrywkt", "geometria", "геометрия", "геом"
}

LAT_FIELD_NAMES = {
    "lat", "latitude", "y", "ycoord", "ycoordinate", "latwgs84",
    "latwgs", "широта", "широты", "latdeg", "latdd", "коордy",
    "yкоорд", "coordy"
}

LON_FIELD_NAMES = {
    "lon", "long", "longitude", "lng", "x", "xcoord",
    "xcoordinate", "lonwgs84", "lonwgs", "долгота", "долготы",
    "долг", "londeg", "коордx", "xкоорд", "coordx"
}


def normalize_field_name(name: str) -> str:
    return re.sub(r"[^a-zA-Zа-яА-Я0-9]", "", name).lower()


class CheckableComboBox(QComboBox):
    """A QComboBox allowing multiple selection via checkable items."""

    selection_changed = pyqtSignal()

    def __init__(self, parent=None, show_count=False):
        super().__init__(parent)
        self.setModel(QStandardItemModel(self))
        # Use the pressed signal and handle state changes manually so that
        # the check state is always updated before we emit signals.
        # This prevents issues where a column could not be re-enabled after
        # being toggled off (especially for the first column in the list).
        self.view().pressed.connect(self.handle_item_pressed)
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self.lineEdit().setPlaceholderText("")
        self.lineEdit().installEventFilter(self)
        self.show_count = show_count
        self.select_all_text = "Выбрать все"

    def addItem(self, text, data=None):
        item = QStandardItem(text)
        item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
        item.setData(Qt.CheckState.Unchecked, Qt.ItemDataRole.CheckStateRole)
        if data is not None:
            item.setData(data, Qt.ItemDataRole.UserRole)
        self.model().appendRow(item)
        self.update_display_text()

    def clear(self):
        self.model().clear()
        self.update_display_text()

    def handle_item_pressed(self, index):
        """Toggle check state for the pressed item and update dependent UI."""
        item = self.model().itemFromIndex(index)
        if item.text() == self.select_all_text:
            # Toggle all items based on the current state of the select-all item
            new_state = (Qt.CheckState.Unchecked
                         if item.checkState() == Qt.CheckState.Checked
                         else Qt.CheckState.Checked)
            item.setCheckState(new_state)
            for i in range(self.model().rowCount()):
                cur_item = self.model().item(i)
                if cur_item.text() == self.select_all_text:
                    continue
                cur_item.setCheckState(new_state)
        else:
            # Toggle the individual item
            new_state = (Qt.CheckState.Unchecked
                         if item.checkState() == Qt.CheckState.Checked
                         else Qt.CheckState.Checked)
            item.setCheckState(new_state)
        self.update_select_all_state()
        self.update_display_text()
        self.selection_changed.emit()
        self.showPopup()

    def eventFilter(self, obj, event):
        if obj is self.lineEdit() and event.type() == QEvent.Type.MouseButtonPress:
            self.showPopup()
            return True
        return super().eventFilter(obj, event)

    def checkedItems(self):
        items = []
        for i in range(self.model().rowCount()):
            item = self.model().item(i)
            if item.text() == self.select_all_text:
                continue
            if item.checkState() == Qt.CheckState.Checked:
                items.append(item.text())
        return items

    def checkedIndices(self):
        """Return the integer indices of all checked items.

        The first selectable column resides at row 1 because row 0 is the
        "select all" entry.  If custom data has been set for an item it will be
        returned instead of the positional index, allowing callers to rely on
        stable identifiers.
        """
        indices = []
        for i in range(self.model().rowCount()):
            item = self.model().item(i)
            if item.text() == self.select_all_text:
                continue
            if item.checkState() == Qt.CheckState.Checked:
                data = item.data(Qt.ItemDataRole.UserRole)
                indices.append(data if data is not None else i - 1)
        return indices

    def update_display_text(self):
        checked = self.checkedItems()
        if self.show_count:
            total = 0
            for i in range(self.model().rowCount()):
                if self.model().item(i).text() != self.select_all_text:
                    total += 1
            self.lineEdit().setText(f"Выбрано {len(checked)} из {total}")
        else:
            self.lineEdit().setText(", ".join(checked))

    def set_all_checked(self, checked: bool):
        for i in range(self.model().rowCount()):
            item = self.model().item(i)
            if item.text() == self.select_all_text:
                continue
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self.update_select_all_state()
        self.update_display_text()
        self.selection_changed.emit()

    def update_select_all_state(self):
        select_all_item = None
        total = 0
        checked = 0
        for i in range(self.model().rowCount()):
            item = self.model().item(i)
            if item.text() == self.select_all_text:
                select_all_item = item
                continue
            total += 1
            if item.checkState() == Qt.CheckState.Checked:
                checked += 1
        if not select_all_item:
            return
        if checked == 0:
            select_all_item.setCheckState(Qt.CheckState.Unchecked)
        elif checked == total:
            select_all_item.setCheckState(Qt.CheckState.Checked)
        else:
            select_all_item.setCheckState(Qt.CheckState.PartiallyChecked)

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
        self.single_color = QColor('#FF0000')
        self.field_types = {}
        self.encoding = 'utf-8'
        self.manual_group_bounds = {}
        self.all_data = []
        self.all_headers = []
        self.all_field_types = {}
        self.selected_columns = []
        self.grouping_mode = 'numerical'
        self.group_opacity = 100
        self.current_header_combo = None # To keep track of the currently open QComboBox for header editing
        self.current_header_combo_column = -1 # Track which column the combo belongs to
        self.numerical_field_is_int = False

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


        file_group_box = QGroupBox("Исходный файл/Настройка вывода")
        file_group_box.setFont(bold_large_font)
        file_group_layout = QVBoxLayout()
        file_group_layout.setContentsMargins(10, 15, 10, 15)
        file_group_box.setLayout(file_group_layout)
        
        file_layout = QHBoxLayout()
        self.file_label = QLabel('Исходный файл (txt, csv, excel):')
        self.file_label.setStyleSheet(label_style)
        file_layout.addWidget(self.file_label)
        self.file_path_input = QLineEdit()
        self.file_path_input.setStyleSheet(lineedit_style)
        file_layout.addWidget(self.file_path_input)
        self.browse_button = QPushButton('Выбрать')
        self.browse_button.clicked.connect(self.browse_file)
        self.browse_button.setStyleSheet(button_style + button_hover_style)
        file_layout.addWidget(self.browse_button)
        file_group_layout.addLayout(file_layout)

        output_file_layout = QHBoxLayout()
        self.output_file_label = QLabel('Выходной KML файл:')
        self.output_file_label.setStyleSheet(label_style)
        output_file_layout.addWidget(self.output_file_label)
        self.output_file_path_input = QLineEdit()
        self.output_file_path_input.setStyleSheet(lineedit_style)
        output_file_layout.addWidget(self.output_file_path_input)
        self.browse_output_button = QPushButton('Выбрать')
        self.browse_output_button.clicked.connect(self.browse_output_file)
        self.browse_output_button.setStyleSheet(button_style + button_hover_style)
        output_file_layout.addWidget(self.browse_output_button)
        file_group_layout.addLayout(output_file_layout)

        sheet_layout = QHBoxLayout()
        self.sheet_label = QLabel('Лист Excel:')
        self.sheet_label.setStyleSheet(label_style)
        sheet_layout.addWidget(self.sheet_label)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setPlaceholderText('Выберите лист')
        self.sheet_combo.setStyleSheet(combobox_style)
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        self.sheet_combo.setEnabled(False)
        sheet_layout.addWidget(self.sheet_combo)
        file_group_layout.addLayout(sheet_layout)

        columns_layout = QHBoxLayout()
        self.columns_label = QLabel('Столбцы для работы:')
        self.columns_label.setStyleSheet(label_style)
        columns_layout.addWidget(self.columns_label)
        self.columns_combo = CheckableComboBox(show_count=True)
        self.columns_combo.setStyleSheet(combobox_style)
        self.columns_combo.lineEdit().setPlaceholderText('Выберите столбцы…')
        self.columns_combo.selection_changed.connect(self.on_columns_changed)
        self.columns_combo.setEnabled(False)
        columns_layout.addWidget(self.columns_combo)
        file_group_layout.addLayout(columns_layout)

        options_layout = QHBoxLayout()
        self.delimiter_label = QLabel('Разделитель:')
        self.delimiter_label.setStyleSheet(label_style)
        options_layout.addWidget(self.delimiter_label)
        self.delimiter_input = QLineEdit(';')
        self.delimiter_input.textChanged.connect(self.on_file_settings_changed)
        self.delimiter_input.setFixedWidth(30)
        self.delimiter_input.setStyleSheet(lineedit_style)
        options_layout.addWidget(self.delimiter_input)

        self.has_header_checkbox = QCheckBox('Наличик заголовка')
        self.has_header_checkbox.setChecked(True)
        self.has_header_checkbox.stateChanged.connect(self.on_file_settings_changed)
        self.has_header_checkbox.setStyleSheet(checkbox_style)
        options_layout.addWidget(self.has_header_checkbox)

        self.start_row_label = QLabel('Данные начинаются со строки:')
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
        self.encoding_label = QLabel('Кодировка файла:')
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
        
        coord_group_box = QGroupBox("Настройка системы координат/KML метки")
        coord_group_box.setFont(bold_large_font)
        coord_layout = QVBoxLayout()
        coord_layout.setContentsMargins(10, 15, 10, 15)
        coord_group_box.setLayout(coord_layout)
        self.coord_system_label = QLabel('Система координат:')
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
        self.wkt_field_label = QLabel('WKT поле:')
        self.wkt_field_label.setStyleSheet(label_style)
        self.wkt_field_combo = QComboBox()
        self.wkt_field_combo.setStyleSheet(combobox_style)
        self.wkt_field_layout.addWidget(self.wkt_field_label)
        self.wkt_field_layout.addWidget(self.wkt_field_combo)
        coord_layout.addLayout(self.wkt_field_layout)

        self.lon_lat_field_layout = QHBoxLayout()
        self.lon_field_label = QLabel('Longitude поле:')
        self.lon_field_label.setStyleSheet(label_style)
        self.lon_field_combo = QComboBox()
        self.lon_field_combo.setStyleSheet(combobox_style)
        self.lat_field_label = QLabel('Latitude поле:')
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
        
        self.add_label_button = QPushButton('Выбрать поле для label')
        self.add_label_button.clicked.connect(self.toggle_kml_label_field)
        self.add_label_button.setStyleSheet(button_style + button_hover_style)
        coord_layout.addWidget(self.add_label_button)

        self.kml_label_field_layout = QHBoxLayout()
        self.kml_label_field_label = QLabel('Поле для label:')
        self.kml_label_field_label.setStyleSheet(label_style)
        self.kml_label_field_combo = QComboBox()
        self.kml_label_field_combo.setStyleSheet(combobox_style)
        self.kml_label_field_layout.addWidget(self.kml_label_field_label)
        self.kml_label_field_layout.addWidget(self.kml_label_field_combo)
        coord_layout.addLayout(self.kml_label_field_layout)

        self.kml_label_field_label.setVisible(False)
        self.kml_label_field_combo.setVisible(False)

        # Description fields selection
        desc_layout = QHBoxLayout()
        self.description_fields_label = QLabel('Поля для описания:')
        self.description_fields_label.setStyleSheet(label_style)
        self.description_fields_combo = CheckableComboBox()
        self.description_fields_combo.setStyleSheet(combobox_style)
        desc_layout.addWidget(self.description_fields_label)
        desc_layout.addWidget(self.description_fields_combo)
        coord_layout.addLayout(desc_layout)

        self.use_custom_icon_checkbox = QCheckBox('Использовать пользовательскую иконку')
        self.use_custom_icon_checkbox.setChecked(False)
        self.use_custom_icon_checkbox.stateChanged.connect(self.toggle_custom_icon_input)
        self.use_custom_icon_checkbox.setStyleSheet(checkbox_style)
        coord_layout.addWidget(self.use_custom_icon_checkbox)
        self.icon_url_layout = QHBoxLayout()
        self.icon_url_label = QLabel('URL на иконку:')
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

        grouping_group_box = QGroupBox("Настройка группировки")
        grouping_group_box.setFont(bold_large_font)
        grouping_options_layout = QVBoxLayout()
        grouping_options_layout.setContentsMargins(10, 15, 10, 15)
        grouping_group_box.setLayout(grouping_options_layout)

        mode_layout = QHBoxLayout()
        self.numeric_mode_radio = QRadioButton('Группировка по диапазону значений')
        self.unique_mode_radio = QRadioButton('Групировка по уникальным значениям')
        self.single_mode_radio = QRadioButton('Единый цвет (без группировки)')
        self.numeric_mode_radio.setChecked(True)
        for rb in (self.numeric_mode_radio, self.unique_mode_radio, self.single_mode_radio):
            rb.setStyleSheet(radio_button_style)
        self.grouping_mode_group = QButtonGroup(self)
        self.grouping_mode_group.addButton(self.numeric_mode_radio)
        self.grouping_mode_group.addButton(self.unique_mode_radio)
        self.grouping_mode_group.addButton(self.single_mode_radio)
        self.grouping_mode_group.buttonClicked.connect(self.on_grouping_mode_changed)
        mode_layout.addWidget(self.numeric_mode_radio)
        mode_layout.addWidget(self.unique_mode_radio)
        mode_layout.addWidget(self.single_mode_radio)
        mode_layout.addStretch(1)
        grouping_options_layout.addLayout(mode_layout)
        numerical_field_selection_layout = QHBoxLayout()
        self.numerical_group_label = QLabel('Числовое поле для группировки:')
        self.numerical_group_label.setStyleSheet(label_style)
        numerical_field_selection_layout.addWidget(self.numerical_group_label)
        self.numerical_group_field_combo = QComboBox()
        self.numerical_group_field_combo.currentIndexChanged.connect(self.on_numerical_grouping_field_changed)
        self.numerical_group_field_combo.setStyleSheet(combobox_style)
        numerical_field_selection_layout.addWidget(self.numerical_group_field_combo)
        grouping_options_layout.addLayout(numerical_field_selection_layout)
        num_groups_layout = QHBoxLayout()
        self.num_groups_label = QLabel('Кол-во групп:')
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
        self.end_color_label = QLabel('Конечный цвет для градиента:')
        self.end_color_label.setStyleSheet(label_style)
        end_color_layout.addWidget(self.end_color_label)
        self.end_color_button = QPushButton()
        self.end_color_button.clicked.connect(self.pick_end_color)
        self.end_color_button.setFixedSize(20, 20)
        end_color_layout.addWidget(self.end_color_button)
        grouping_options_layout.addLayout(end_color_layout)

        single_color_layout = QHBoxLayout()
        self.single_color_label = QLabel('Цвет слоя:')
        self.single_color_label.setStyleSheet(label_style)
        single_color_layout.addWidget(self.single_color_label)
        self.single_color_button = QPushButton()
        self.single_color_button.clicked.connect(self.pick_single_color)
        self.single_color_button.setFixedSize(20, 20)
        single_color_layout.addWidget(self.single_color_button)
        grouping_options_layout.addLayout(single_color_layout)

        opacity_layout = QHBoxLayout()
        self.opacity_label = QLabel('Прозрачность (%):')
        self.opacity_label.setStyleSheet(label_style)
        opacity_layout.addWidget(self.opacity_label)
        self.opacity_spinbox = QSpinBox()
        self.opacity_spinbox.setMinimum(0)
        self.opacity_spinbox.setMaximum(100)
        self.opacity_spinbox.setValue(100)
        self.opacity_spinbox.setStyleSheet(spinbox_style)
        self.opacity_spinbox.valueChanged.connect(self.on_opacity_changed)
        opacity_layout.addWidget(self.opacity_spinbox)
        grouping_options_layout.addLayout(opacity_layout)
        self.update_end_color_button()
        self.update_single_color_button()
        self.numerical_color_display_layout = QVBoxLayout()
        self.numerical_color_label = QLabel('Группы значений и цвета:')
        self.numerical_color_label.setStyleSheet(label_style)
        self.numerical_color_display_layout.addWidget(self.numerical_color_label)
        grouping_options_layout.addLayout(self.numerical_color_display_layout)

        categorical_field_layout = QHBoxLayout()
        self.categorical_group_label = QLabel('Категориальное поле для группировки:')
        self.categorical_group_label.setStyleSheet(label_style)
        categorical_field_layout.addWidget(self.categorical_group_label)
        self.categorical_group_field_combo = QComboBox()
        self.categorical_group_field_combo.currentIndexChanged.connect(self.on_categorical_grouping_field_changed)
        self.categorical_group_field_combo.setStyleSheet(combobox_style)
        categorical_field_layout.addWidget(self.categorical_group_field_combo)
        grouping_options_layout.addLayout(categorical_field_layout)

        self.categorical_color_display_layout = QVBoxLayout()
        self.categorical_color_label = QLabel('Уникальные значения и цвета:')
        self.categorical_color_label.setStyleSheet(label_style)
        self.categorical_color_display_layout.addWidget(self.categorical_color_label)
        grouping_options_layout.addLayout(self.categorical_color_display_layout)

        # Hide categorical controls initially
        self.categorical_group_label.setVisible(False)
        self.categorical_group_field_combo.setVisible(False)
        self.categorical_color_label.setVisible(False)
        self.single_color_label.setVisible(False)
        self.single_color_button.setVisible(False)
        layout.addWidget(grouping_group_box)

        # Data filtering controls
        filter_group_box = QGroupBox("Фильтрация данных")
        filter_group_box.setFont(bold_large_font)
        filter_layout = QHBoxLayout()
        self.filter_label = QLabel('Формула фильтрации:')
        self.filter_label.setStyleSheet(label_style)
        filter_layout.addWidget(self.filter_label)
        self.filter_input = QLineEdit()
        self.filter_input.setStyleSheet(lineedit_style)
        self.filter_input.setPlaceholderText("Пример: Column == Value and Other > 5")
        filter_layout.addWidget(self.filter_input)
        self.apply_filter_button = QPushButton('Применить фильтр')
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

        self.generate_button = QPushButton('Создать KML файл')
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
        self.sheet_label.setEnabled(is_excel)
        if not is_excel:
            self.sheet_combo.clear()
        self.sheet_combo.setEnabled(is_excel and self.sheet_combo.count() > 0)

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
            if is_excel:
                self.load_sheet_names(file_name)
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

    def on_sheet_changed(self, _):
        file_path = self.file_path_input.text()
        if file_path:
            self.load_data(file_path)

    def load_sheet_names(self, file_path):
        self.sheet_combo.clear()
        try:
            engine = 'openpyxl' if file_path.lower().endswith(('.xlsx', '.xlsm')) else 'xlrd'
            xls = pd.ExcelFile(file_path, engine=engine)
            self.sheet_combo.addItems(xls.sheet_names)
            if xls.sheet_names:
                self.sheet_combo.setCurrentIndex(0)
            self.sheet_combo.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Не удалось получить список листов: {e}")
            self.sheet_combo.setEnabled(False)
            
    def load_data(self, file_path):
        """Загружает данные из файла с учетом выбранных параметров."""
        self.data = []
        self.filtered_data = []
        self.headers = []
        self.field_types = {}
        self.manual_group_bounds = {}
        self.all_data = []
        self.all_headers = []
        self.all_field_types = {}
        self.selected_columns = []
        self.columns_combo.clear()
        self.columns_combo.setEnabled(False)

        is_excel = file_path.lower().endswith(('.xlsx', '.xls', '.xlsm'))
        self.update_file_options_state(is_excel)

        has_header = self.has_header_checkbox.isChecked()
        start_row = self.start_row_spinbox.value() - 1

        try:
            if is_excel:
                header_row = 0 if has_header else None
                engine = 'openpyxl' if file_path.lower().endswith(('.xlsx', '.xlsm')) else 'xlrd'
                sheet_name = self.sheet_combo.currentText() if self.sheet_combo.currentText() else 0
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, skiprows=start_row, engine=engine)
                
                if not has_header:
                    df.columns = [f'Column {i}' for i in range(len(df.columns))]

                df = df.fillna('')
                self.headers = df.columns.astype(str).tolist()
                self.data = df.values.tolist()
                self._auto_cast_numeric()
                self.filtered_data = [row[:] for row in self.data]
                if hasattr(self, 'filter_input'):
                    self.filter_input.setText('')
            else:
                delimiter = self.delimiter_input.text()
                self.encoding = 'utf-8' if self.utf8_radio.isChecked() else 'cp1251'
                with open(file_path, 'r', encoding=self.encoding, errors='ignore') as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    all_lines = list(reader)
                    # Remove UTF-8 BOM from the first column if present so that
                    # the initial field is not dropped or mismatched when used
                    # as a label or description field.
                    for line in all_lines:
                        if line:
                            line[0] = line[0].lstrip('\ufeff')

                    data_start_index = start_row
                    if has_header:
                        if start_row < len(all_lines):
                            self.headers = all_lines[start_row]
                            data_start_index += 1
                    
                    if data_start_index < len(all_lines):
                        self.data = all_lines[data_start_index:]
                        self._auto_cast_numeric()
                        self.filtered_data = [row[:] for row in self.data]
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

            if not self.headers:
                QMessageBox.warning(self, "Warning", "В файле не обнаружены столбцы")
                self.all_data, self.all_headers, self.all_field_types, self.selected_columns = [], [], {}, []
                self.update_columns_combo()
                self.preview_data()
                self.update_field_combos()
                return

            inferred = self._infer_field_types(self.data, self.headers)
            for field, f_type in inferred.items():
                self.field_types[field] = f_type

            self.all_data = [row[:] for row in self.data]
            self.all_headers = self.headers[:]
            self.all_field_types = self.field_types.copy()
            # Track selected columns by their positional indices to avoid
            # dropping the first column or mis-handling duplicate names.
            self.selected_columns = list(range(len(self.headers)))
            self.update_columns_combo()

            self.preview_data()
            self.update_field_combos()
            if self.grouping_mode == 'numerical':
                self.on_numerical_grouping_field_changed()
            elif self.grouping_mode == 'categorical':
                self.on_categorical_grouping_field_changed()
            else:
                self.update_group_display()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ошибка загрузки файла: {e}\n\nДля файлов Excel убедитесь, что установлены 'pandas' и 'openpyxl'.")
            self.data, self.filtered_data, self.headers, self.field_types = [], [], [], {}
            self.all_data, self.all_headers, self.all_field_types, self.selected_columns = [], [], {}, []
            self.update_columns_combo()
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
        cat_group_idx = field_indices.get(self.categorical_group_field_combo.currentText(), -1)

        if use_wkt and wkt_idx == -1:
            QMessageBox.critical(self, "Error", "Выбранное поле WKT не найдено.")
            return
        if not use_wkt and (lon_idx == -1 or lat_idx == -1):
            QMessageBox.critical(self, "Error", "Выбранные поля долготы/широты не найдены.")
            return

        kml_folders = {}
        if self.grouping_mode == 'numerical':
            grouping_active = num_group_idx != -1 and self.groups
        elif self.grouping_mode == 'categorical':
            grouping_active = cat_group_idx != -1 and self.groups
        else:
            grouping_active = False
        if grouping_active:
            for group in self.groups:
                kml_folders[group['label']] = kml.newfolder(name=group['label'])

        try:
            for i, row in enumerate(self.filtered_data):
                if self.grouping_mode == 'numerical' and num_group_idx != -1:
                    if num_group_idx >= len(row) or str(row[num_group_idx]).strip() == '':
                        continue
                if self.grouping_mode == 'categorical' and cat_group_idx != -1:
                    if cat_group_idx >= len(row) or str(row[cat_group_idx]).strip() == '':
                        continue

                target_container = kml
                assigned_group = None

                if grouping_active:
                    if self.grouping_mode == 'numerical' and num_group_idx != -1 and num_group_idx < len(row):
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
                    elif self.grouping_mode == 'categorical' and cat_group_idx != -1 and cat_group_idx < len(row):
                        val = str(row[cat_group_idx])
                        if val in kml_folders:
                            target_container = kml_folders[val]
                            assigned_group = {'label': val}

                label_text = ''
                if self.kml_label_field_combo.isVisible() and label_idx != -1 and label_idx < len(row):
                    label_text = str(row[label_idx])

                kml_objects = []

                if use_wkt:
                    if wkt_idx != -1 and wkt_idx < len(row):
                        try:
                            geom = wkt.loads(str(row[wkt_idx]))
                            if geom.geom_type == 'Point':
                                kml_objects.append(target_container.newpoint(name=label_text, coords=[(geom.x, geom.y)]))
                            elif geom.geom_type == 'LineString':
                                kml_objects.append(target_container.newlinestring(name=label_text, coords=list(geom.coords)))
                            elif geom.geom_type == 'Polygon':
                                poly = target_container.newpolygon(
                                    name=label_text,
                                    outerboundaryis=list(geom.exterior.coords),
                                    innerboundaryis=[list(r.coords) for r in geom.interiors],
                                )
                                kml_objects.append(poly)
                                if label_text:
                                    pt = geom.representative_point()
                                    label_point = target_container.newpoint(
                                        name=label_text,
                                        coords=[(pt.x, pt.y)],
                                    )
                                    kml_objects.append(label_point)
                            elif geom.geom_type == 'MultiPolygon':
                                largest_poly = max(geom.geoms, key=lambda g: g.area)
                                label_pt = largest_poly.representative_point() if label_text else None
                                for poly_geom in geom.geoms:
                                    poly = target_container.newpolygon(
                                        name=label_text,
                                        outerboundaryis=list(poly_geom.exterior.coords),
                                        innerboundaryis=[list(r.coords) for r in poly_geom.interiors],
                                    )
                                    kml_objects.append(poly)
                                if label_text:
                                    label_point = target_container.newpoint(
                                        name=label_text,
                                        coords=[(label_pt.x, label_pt.y)],
                                    )
                                    kml_objects.append(label_point)
                        except Exception as e:
                            print(f"Row {i+1} WKT Error: {e}"); continue
                else:
                    if lon_idx != -1 and lat_idx != -1 and lon_idx < len(row) and lat_idx < len(row):
                        try:
                            lon = float(str(row[lon_idx]).replace(',', '.'))
                            lat = float(str(row[lat_idx]).replace(',', '.'))
                            kml_objects.append(target_container.newpoint(name=label_text, coords=[(lon, lat)]))
                        except (ValueError, TypeError):
                            print(f"Row {i+1} Lon/Lat Error"); continue

                if not kml_objects:
                    continue

                if assigned_group:
                    color = self.group_colors.get(assigned_group['label'])
                    if color:
                        alpha = int(255 * (self.group_opacity / 100))
                        fill_color = simplekml.Color.rgb(
                            color.red(), color.green(), color.blue(), alpha
                        )
                        line_color = simplekml.Color.rgb(
                            color.red(), color.green(), color.blue()
                        )

                        for obj in kml_objects:
                            if isinstance(obj, simplekml.Point):
                                obj.style.iconstyle.color = line_color
                            elif isinstance(obj, simplekml.LineString):
                                obj.style.linestyle.color = line_color
                            elif isinstance(obj, simplekml.Polygon):
                                obj.style.polystyle.color = fill_color
                                obj.style.linestyle.color = line_color
                elif self.grouping_mode == 'single':
                    color = self.single_color
                    alpha = int(255 * (self.group_opacity / 100))
                    fill_color = simplekml.Color.rgb(
                        color.red(), color.green(), color.blue(), alpha
                    )
                    line_color = simplekml.Color.rgb(
                        color.red(), color.green(), color.blue()
                    )

                    for obj in kml_objects:
                        if isinstance(obj, simplekml.Point):
                            obj.style.iconstyle.color = line_color
                        elif isinstance(obj, simplekml.LineString):
                            obj.style.linestyle.color = line_color
                        elif isinstance(obj, simplekml.Polygon):
                            obj.style.polystyle.color = fill_color
                            obj.style.linestyle.color = line_color

                # Build description snippet
                desc_fields = self.description_fields_combo.checkedItems()
                if desc_fields:
                    lines = []
                    for field in desc_fields:
                        idx = field_indices.get(field, -1)
                        value = ''
                        if idx != -1 and idx < len(row):
                            value = row[idx]
                        lines.append(f"<b>{field}</b>: {value}")
                    snippet_html = "<br>".join(lines)
                    for obj in kml_objects:
                        try:
                            obj.snippet = simplekml.Snippet(snippet_html, maxlines=len(lines))
                        except Exception:
                            obj.snippet = snippet_html

                        # Show description on click
                        obj.description = snippet_html



                for obj in kml_objects:
                    if isinstance(obj, simplekml.Point):
                        use_custom_icon = self.use_custom_icon_checkbox.isChecked()
                        custom_icon_url = self.icon_url_input.text()
                        if use_custom_icon and custom_icon_url:
                            obj.style.iconstyle.icon.href = custom_icon_url
                        else:
                            obj.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/paddle/wht-blank.png'  # Changed default icon URL
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

            self.field_types[self.headers[col_idx]] = "Int" if is_int_col else "Float"

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
                inferred_types[field_name] = 'Varchar'
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
                inferred_types[field_name] = 'Geometry'
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
                inferred_types[field_name] = 'Int' if is_integer else 'Float'
            else:
                inferred_types[field_name] = 'Varchar'
        
        return inferred_types


    def _format_range_value(self, value):
        """Format range boundary based on current field type."""
        return f"{int(value)}" if self.numerical_field_is_int else f"{value:.2f}"


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
        data_types = ['Auto', 'Int', 'Float', 'Varchar', 'Geometry'] 
        combo.addItems(data_types)
        combo.setStyleSheet("QComboBox { background-color: #DDDDDD; border: 1px solid #AAAAAA; padding: 1px; }")

        field_name = self.headers[column_index]
        current_type = self.field_types.get(field_name, 'Auto')
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
        if self.grouping_mode == 'numerical':
            self.on_numerical_grouping_field_changed()
        elif self.grouping_mode == 'categorical':
            self.on_categorical_grouping_field_changed()
        else:
            self.update_group_display()
        
    def _infer_field_types_for_column(self, column_index):
        """
        Infers the data type for a single column.
        Helper function for 'auto' type selection.
        """
        if not self.data or column_index >= len(self.headers):
            return 'Varchar' # Default if no data or invalid column

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
            return 'Varchar'

        for value in sample_values:
            try:
                geom = wkt.loads(value)
                if geom.geom_type in ['Point', 'LineString', 'Polygon', 'MultiPoint', 'MultiLineString', 'MultiPolygon', 'GeometryCollection']:
                    is_wkt = True
                    break
            except Exception:
                pass
        
        if is_wkt:
            return 'Geometry'

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
            return 'Int' if is_integer else 'Float'
        else:
            return 'Varchar'


    def update_field_combos(self):
        """
        Updates the available fields in all QComboBox widgets based on loaded headers
        and inferred field types.
        """
        if not self.data and not self.headers:
            for combo in [self.wkt_field_combo, self.lon_field_combo, self.lat_field_combo,
                          self.numerical_group_field_combo, self.categorical_group_field_combo,
                          self.kml_label_field_combo, self.description_fields_combo]:
                combo.clear()
            return

        num_columns = len(self.headers) if self.headers else (len(self.data[0]) if self.data else 0)
        all_fields = self.headers[:]
        
        numerical_fields = [field for field in all_fields if self.field_types.get(field) in ['Int', 'Float']]
        categorical_fields = [f for f in all_fields if f not in numerical_fields]

        combos = [self.wkt_field_combo, self.lon_field_combo, self.lat_field_combo,
                  self.numerical_group_field_combo, self.categorical_group_field_combo,
                  self.kml_label_field_combo]
        current_texts = [c.currentText() for c in combos]
        selected_desc = set(self.description_fields_combo.checkedItems())


        # Prevent signals from firing while repopulating combos
        for c in combos + [self.description_fields_combo]:
            c.blockSignals(True)


        for c in combos:
            c.clear()

        self.description_fields_combo.clear()
        for field in all_fields:
            self.description_fields_combo.addItem(field)
            if field in selected_desc:
                items = self.description_fields_combo.model().findItems(field)
                if items:
                    items[0].setCheckState(Qt.CheckState.Checked)

        for c in [self.wkt_field_combo, self.lon_field_combo, self.lat_field_combo, self.kml_label_field_combo]:
            c.addItems(all_fields)
        self.numerical_group_field_combo.addItems(numerical_fields)
        self.categorical_group_field_combo.addItems(categorical_fields)

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

        if current_texts[4] in categorical_fields:
            self.categorical_group_field_combo.setCurrentText(current_texts[4])
        elif categorical_fields:
            self.categorical_group_field_combo.setCurrentText(categorical_fields[0])
        else:
            self.categorical_group_field_combo.setCurrentText('')

        if current_texts[5] in all_fields:
            self.kml_label_field_combo.setCurrentText(current_texts[5])
        elif all_fields:
            self.kml_label_field_combo.setCurrentText(all_fields[0])

        self._auto_select_coord_fields(current_texts)

        # Re-enable signals now that combo boxes are populated
        for c in combos + [self.description_fields_combo]:
            c.blockSignals(False)


    def _auto_select_coord_fields(self, previous_texts):
        wkt_prev, lon_prev, lat_prev = previous_texts[:3]
        headers = self.headers

        wkt_candidate = None
        if wkt_prev not in headers:
            for field in headers:
                norm = normalize_field_name(field)
                if self.field_types.get(field) == 'Geometry' or norm in WKT_FIELD_NAMES:
                    wkt_candidate = field
                    break
            if wkt_candidate:
                self.wkt_field_combo.setCurrentText(wkt_candidate)
                self.wkt_radio.setChecked(True)
        else:
            wkt_candidate = wkt_prev

        if (lon_prev not in headers or lat_prev not in headers) and not wkt_candidate:
            lon_candidate = None
            lat_candidate = None
            for field in headers:
                norm = normalize_field_name(field)
                if not lon_candidate and norm in LON_FIELD_NAMES:
                    lon_candidate = field
                if not lat_candidate and norm in LAT_FIELD_NAMES:
                    lat_candidate = field
            if lon_candidate and lat_candidate:
                self.lon_field_combo.setCurrentText(lon_candidate)
                self.lat_field_combo.setCurrentText(lat_candidate)
                self.lonlat_radio.setChecked(True)


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
                    if t in ["Int", "Float"]:
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
                            if t in ["Int", "Float"] and col in df_numeric.columns:



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
        if self.grouping_mode == 'numerical':
            self.on_numerical_grouping_field_changed()
        elif self.grouping_mode == 'categorical':
            self.on_categorical_grouping_field_changed()
        else:
            self.update_group_display()


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
                    header_name = self.headers[j]
                    field_type = self.field_types.get(header_name)
                    display_text = str(item)
                    if field_type == 'geometry' and len(display_text) > 1000:
                        display_text = display_text[:1000]
                    self.data_table.setItem(i, j, QTableWidgetItem(display_text))

        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


    def update_columns_combo(self):
        self.columns_combo.blockSignals(True)
        self.columns_combo.clear()
        self.columns_combo.addItem(self.columns_combo.select_all_text)
        for idx, field in enumerate(self.all_headers):
            # Store the original column index in the item for later retrieval
            self.columns_combo.addItem(field, idx)
            item = self.columns_combo.model().item(idx + 1)
            if idx in self.selected_columns:
                item.setCheckState(Qt.CheckState.Checked)
        self.columns_combo.update_select_all_state()
        self.columns_combo.update_display_text()
        self.columns_combo.blockSignals(False)

        total = len(self.all_headers)
        enabled = total > 0
        self.columns_combo.setEnabled(enabled)

    def on_columns_changed(self):
        indices = sorted(self.columns_combo.checkedIndices())
        self.selected_columns = indices

        if not indices:
            self.headers = []
            self.data = []
            self.filtered_data = []
            self.field_types = {}
            if hasattr(self, 'filter_input'):
                self.filter_input.setText('')
            self.update_field_combos()
            self.preview_data()
            if self.grouping_mode == 'numerical':
                self.on_numerical_grouping_field_changed()
            elif self.grouping_mode == 'categorical':
                self.on_categorical_grouping_field_changed()
            else:
                self.update_group_display()
            return

        self.headers = [self.all_headers[i] for i in indices]
        self.data = [[row[i] for i in indices] for row in self.all_data]
        self.field_types = {h: self.all_field_types.get(h, 'auto') for h in self.headers}
        self.filtered_data = [row[:] for row in self.data]
        if hasattr(self, 'filter_input'):
            self.filter_input.setText('')
        self.update_field_combos()
        self.preview_data()
        if self.grouping_mode == 'numerical':
            self.on_numerical_grouping_field_changed()
        elif self.grouping_mode == 'categorical':
            self.on_categorical_grouping_field_changed()
        else:
            self.update_group_display()

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
            self.add_label_button.setText('Скрыть поле KML метки')
        else:
            self.add_label_button.setText('Выбрать поле для label')

    def toggle_custom_icon_input(self):
        """Показать или скрыть поле ввода URL иконки."""
        is_checked = self.use_custom_icon_checkbox.isChecked()
        self.icon_url_label.setVisible(is_checked)
        self.icon_url_input.setVisible(is_checked)

    def on_grouping_mode_changed(self):
        """Switch between numerical, categorical, and single-color modes."""
        if self.numeric_mode_radio.isChecked():
            self.grouping_mode = 'numerical'
        elif self.unique_mode_radio.isChecked():
            self.grouping_mode = 'categorical'
        else:
            self.grouping_mode = 'single'

        numerical = self.grouping_mode == 'numerical'
        categorical = self.grouping_mode == 'categorical'
        single = self.grouping_mode == 'single'

        for w in [self.numerical_group_label, self.numerical_group_field_combo,
                   self.num_groups_label, self.num_groups_spinbox,
                   self.end_color_label, self.end_color_button,
                   self.numerical_color_label]:
            w.setVisible(numerical)

        self.categorical_group_label.setVisible(categorical)
        self.categorical_group_field_combo.setVisible(categorical)
        self.categorical_color_label.setVisible(categorical)

        self.single_color_label.setVisible(single)
        self.single_color_button.setVisible(single)

        if self.grouping_mode == 'numerical':
            self.on_numerical_grouping_field_changed()
        elif self.grouping_mode == 'categorical':
            self.on_categorical_grouping_field_changed()
        else:
            self.update_group_display()
    
    def pick_end_color(self):
        """Открывает диалог выбора цвета для конечного цвета градиента."""
        color = QColorDialog.getColor(self.end_color, self, "Выбрать конечный цвет для градиента")
        if color.isValid():
            self.end_color = color
            self.update_end_color_button()
            if self.grouping_mode == 'numerical':
                self.on_numerical_grouping_field_changed()

    def pick_single_color(self):
        """Открывает диалог выбора цвета слоя."""
        color = QColorDialog.getColor(self.single_color, self, "Выбрать цвет слоя")
        if color.isValid():
            self.single_color = color
            self.update_single_color_button()

    def on_opacity_changed(self):
        """Update stored group opacity when the spinbox value changes."""
        self.group_opacity = self.opacity_spinbox.value()

    def update_end_color_button(self):
        """Обновляет цвет кнопки, отображающей конечный цвет."""
        self.end_color_button.setStyleSheet(f"background-color: {self.end_color.name()}; border: 1px solid #888888;")

    def update_single_color_button(self):
        """Обновляет цвет кнопки, отображающей выбранный цвет слоя."""
        self.single_color_button.setStyleSheet(f"background-color: {self.single_color.name()}; border: 1px solid #888888;")

    def on_categorical_grouping_field_changed(self):
        """Rebuild groups based on unique categorical values."""
        self.groups = []
        self.group_colors = {}

        selected_field = self.categorical_group_field_combo.currentText()
        if not selected_field or not self.filtered_data or selected_field not in self.headers:
            self.update_group_display()
            return

        col_index = self.headers.index(selected_field)
        values = []
        for row in self.filtered_data:
            if col_index < len(row):
                val = str(row[col_index]).strip()
                if val != '':
                    values.append(val)

        unique_vals = sorted(set(values))
        n = len(unique_vals)
        for i, val in enumerate(unique_vals):
            hue = (i * 360 / max(1, n)) / 360
            r, g, b = [int(c * 255) for c in colorsys.hsv_to_rgb(hue, 0.7, 1)]
            color = QColor(r, g, b)
            self.groups.append({'label': val, 'value': val, 'color': color})
            self.group_colors[val] = color

        self.update_group_display()

    def pick_category_color(self, index):
        """Allow manual selection of a category color."""
        current = self.groups[index]['color']
        color = QColorDialog.getColor(current, self, "Выбрать цвет категории")
        if color.isValid():
            self.groups[index]['color'] = color
            self.group_colors[self.groups[index]['label']] = color
            self.update_group_display()

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
        self.numerical_field_is_int = self.field_types.get(selected_field) == 'Int'
        numerical_values = []
        for row in self.filtered_data:
            if col_index < len(row):
                try:
                    val = str(row[col_index]).replace(',', '.')
                    numerical_values.append(int(float(val)) if self.numerical_field_is_int else float(val))
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

        bins = list(bins)
        if self.numerical_field_is_int:
            bins = [int(round(b)) for b in bins]

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

            if self.numerical_field_is_int:
                lower_bound = int(round(lower_bound))
                upper_bound = int(round(upper_bound))

            r = start_color.red() + i * (end_color.red() - start_color.red()) // (num_groups - 1) if num_groups > 1 else start_color.red()
            g = start_color.green() + i * (end_color.green() - start_color.green()) // (num_groups - 1) if num_groups > 1 else start_color.green()
            b = start_color.blue() + i * (end_color.blue() - start_color.blue()) // (num_groups - 1) if num_groups > 1 else start_color.blue()
            group_color = QColor(r, g, b)

            label = f"{self._format_range_value(lower_bound)} - {self._format_range_value(upper_bound)}"
            self.groups.append({
                'label': label,
                'range': [lower_bound, upper_bound],
                'color': group_color
            })
            self.group_colors[label] = group_color

        self.update_group_display()

    def update_group_display(self):
        """Update the displayed grouping information for the current mode."""
        num_layout = self.numerical_color_display_layout
        cat_layout = self.categorical_color_display_layout

        for layout in (num_layout, cat_layout):
            while layout.count() > 1:
                child = layout.takeAt(1)
                if child.widget():
                    child.widget().deleteLater()
                elif child.layout():
                    self.clear_layout(child.layout())

        if not self.groups:
            if self.grouping_mode != 'single':
                lbl = QLabel("Группы не определены или данные недоступны.")
                lbl.setStyleSheet("QLabel { color: #555555;margin-left: 10px; }")
                if self.grouping_mode == 'categorical':
                    cat_layout.addWidget(lbl)
                else:
                    num_layout.addWidget(lbl)
            return

        if self.grouping_mode == 'numerical':
            selected_field = self.numerical_group_field_combo.currentText()
            numerical_values = []
            if selected_field and self.filtered_data and self.headers and selected_field in self.headers:
                col_index = self.headers.index(selected_field)
                for row in self.filtered_data:
                    if col_index < len(row):
                        try:
                            val = str(row[col_index]).replace(',', '.')
                            numerical_values.append(int(float(val)) if self.numerical_field_is_int else float(val))
                        except (ValueError, TypeError):
                            pass

            for i, group in enumerate(self.groups):
                g_layout = QHBoxLayout()
                swatch = QLabel()
                swatch.setFixedSize(20, 20)
                swatch.setStyleSheet(f"background-color: {group['color'].name()}; border: 1px solid #888888;")
                g_layout.addWidget(swatch)

                lower_label = QLabel(f"{self._format_range_value(group['range'][0])} - ")
                lower_label.setStyleSheet("QLabel { color: #333333;}")
                g_layout.addWidget(lower_label)

                upper_input = QLineEdit(self._format_range_value(group['range'][1]))
                upper_input.setFixedWidth(80)
                upper_input.setStyleSheet("QLineEdit { background-color: #EEEEEE; border: 1px solid #CCCCCC; padding: 3px; }")

                if i == len(self.groups) - 1:
                    upper_input.setReadOnly(True)
                    upper_input.setStyleSheet("QLineEdit { background-color: #E0E0E0; border: 1px solid #CCCCCC; padding: 3px; color: #888888; }")
                else:
                    upper_input.editingFinished.connect(lambda idx=i, sender=upper_input: self.on_group_bound_edited(idx, sender))

                g_layout.addWidget(upper_input)

                item_count = 0
                lower, upper = group['range']
                for val in numerical_values:
                    if i == len(self.groups) - 1:
                        if lower <= val <= upper:
                            item_count += 1
                    else:
                        if lower <= val < upper:
                            item_count += 1

                count_label = QLabel(f" ({item_count} элементов)")
                count_label.setStyleSheet("QLabel { color: #555555;font-size: 9px; }")
                g_layout.addWidget(count_label)

                g_layout.addStretch(1)
                num_layout.addLayout(g_layout)
        elif self.grouping_mode == 'categorical':
            for i, group in enumerate(self.groups):
                g_layout = QHBoxLayout()
                swatch = QLabel()
                swatch.setFixedSize(20, 20)
                swatch.setStyleSheet(f"background-color: {group['color'].name()}; border: 1px solid #888888;")
                swatch.mousePressEvent = lambda e, idx=i: self.pick_category_color(idx)
                g_layout.addWidget(swatch)

                lbl = QLabel(group['label'])
                lbl.setStyleSheet("QLabel { color: #333333; }")
                g_layout.addWidget(lbl)
                g_layout.addStretch(1)
                cat_layout.addLayout(g_layout)

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
            sender.setText(self._format_range_value(self.groups[group_index]['range'][1]))
            return

        try:
            text_val = sender.text().replace(',', '.')
            new_value = int(float(text_val)) if self.numerical_field_is_int else float(text_val)
        except ValueError:
            QMessageBox.warning(self, "Неверный ввод", "Пожалуйста введите число.")
            sender.setText(self._format_range_value(self.groups[group_index]['range'][1]))
            return

        current_lower_bound = self.groups[group_index]['range'][0]
        if new_value <= current_lower_bound:
            QMessageBox.warning(self, "Неправильная граница",
                                f"Верхняя граница не может быть меньше нижней границы текущей группы ({self._format_range_value(current_lower_bound)}).")
            sender.setText(self._format_range_value(self.groups[group_index]['range'][1]))
            return

        if group_index + 1 < len(self.groups):
            next_upper_bound = self.groups[group_index + 1]['range'][1]
            if new_value >= next_upper_bound:
                QMessageBox.warning(self, "Неправильная граница",
                                    f"Верхняя граница не может быть больше верхней границы следующей группы ({self._format_range_value(next_upper_bound)}).")
                sender.setText(self._format_range_value(self.groups[group_index]['range'][1]))
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

            group['label'] = f"{self._format_range_value(lower_display)} - {self._format_range_value(upper_display)}"
            
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
