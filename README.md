# -*- coding: utf-8 -*-
import os
import sys
import pickle
import shutil
import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox,
    QTabWidget, QFormLayout, QHeaderView, QGroupBox, QComboBox, QCheckBox, QFileDialog, QDateEdit, QTextEdit
)
from PyQt5.QtCore import Qt, QTimer, QDate, QTime, QDateTime
from PyQt5.QtGui import QFont, QColor, QBrush, QIcon, QPixmap

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# --------- LOGGER FOR CHANGE HISTORY ----------
class ChangeLogger:
    def __init__(self, log_file="change_log.txt"):
        self.log_file = log_file
        if not os.path.exists(self.log_file):
            with open(self.log_file, "w", encoding="utf-8") as f:
                f.write("Fecha y hora\tEvento\n")

    def log(self, event: str):
        now = datetime.datetime.now()
        fecha_hora = now.strftime("%d/%m/%Y %H:%M:%S")
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(f"{fecha_hora}\t{event}\n")

    def read_log(self):
        if not os.path.exists(self.log_file):
            return ""
        with open(self.log_file, "r", encoding="utf-8") as f:
            return f.read()

class AppConfig:
    def __init__(self):
        self.config_file = 'app_config.dat'
        self.settings = {
            'auto_save': True,
            'dark_mode': False,
            'font_size': 12,
            'default_dept': 'General',
            'backup_path': ''
        }
        self.load_config()
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'rb') as f:
                    self.settings = pickle.load(f)
                if 'backup_path' not in self.settings:
                    self.settings['backup_path'] = ''
            except Exception:
                pass
    def save_config(self):
        try:
            with open(self.config_file, 'wb') as f:
                pickle.dump(self.settings, f)
            return True
        except Exception:
            return False
    def update_setting(self, key, value):
        self.load_config()
        self.settings[key] = value
        return self.save_config()

class EmployeeDatabase:
    def __init__(self, db_file='personal_db.dat'):
        self.db_file = db_file
        self.fields = [
            'nombre', 'apellidos', 'escala', 'empleo', 'destino',
            'fecha_incorporacion_unidad', 'fecha_incorporacion_fas', 'fecha_antiguedad_empleo',
            'fecha_nacimiento', 'dni', 'nº_telf_perosnal', 'nº_telf_extension',
            'correo_personal', 'correo_militar', 'baja', 'mision', 'hps'
        ]
        self.load_db()
        allowed = set(self.fields)
        for k in list(self.data.keys()):
            for field in list(self.data[k].keys()):
                if field not in allowed:
                    del self.data[k][field]
        for k in self.data:
            for f in self.fields:
                if f not in self.data[k]:
                    if f in ['baja', 'mision']:
                        self.data[k][f] = 'NO'
                    else:
                        self.data[k][f] = ''
        self.save_db()
    def load_db(self):
        self.data = {}
        if os.path.exists(self.db_file):
            try:
                with open(self.db_file, 'rb') as f:
                    obj = pickle.load(f)
                    if isinstance(obj, tuple) and len(obj) == 2:
                        fields, data = obj
                        self.data = {}
                        for k, v in data.items():
                            new_v = {}
                            for field in self.fields:
                                if field in v:
                                    new_v[field] = v[field]
                                else:
                                    new_v[field] = 'NO' if field in ['baja', 'mision'] else ''
                            self.data[k] = new_v
                    elif isinstance(obj, dict):
                        self.data = {}
                        for k, v in obj.items():
                            new_v = {}
                            for field in self.fields:
                                if field in v:
                                    new_v[field] = v[field]
                                else:
                                    new_v[field] = 'NO' if field in ['baja', 'mision'] else ''
                            self.data[k] = new_v
            except Exception:
                self.data = {}
    def save_db(self):
        try:
            with open(self.db_file, 'wb') as f:
                pickle.dump((self.fields, self.data), f)
            return True
        except Exception:
            return False
    def get_next_id(self):
        self.load_db()
        if not self.data:
            return "1"
        try:
            max_id = max([int(k) for k in self.data.keys() if k.isdigit()])
            return str(max_id + 1)
        except Exception:
            i = 1
            while str(i) in self.data:
                i += 1
            return str(i)
    def add_employee(self, employee_data):
        self.load_db()
        emp_id = self.get_next_id()
        self.data[emp_id] = employee_data
        return emp_id if self.save_db() else None
    def update_employee(self, employee_id, employee_data):
        self.load_db()
        if employee_id not in self.data:
            return False
        self.data[employee_id] = employee_data
        return self.save_db()
    def delete_employee(self, employee_id):
        self.load_db()
        if employee_id in self.data:
            del self.data[employee_id]
            return self.save_db()
        return False
    def get_employee(self, employee_id):
        self.load_db()
        return self.data.get(employee_id, None)
    def get_all_employees(self):
        self.load_db()
        return self.data
    def add_field(self, field_name, default_value=''):
        self.load_db()
        if field_name not in self.fields:
            self.fields.append(field_name)
            for emp in self.data.values():
                emp[field_name] = default_value
            return self.save_db()
        return False
    def update_field_name(self, old_name, new_name):
        self.load_db()
        if old_name in self.fields and new_name not in self.fields:
            index = self.fields.index(old_name)
            self.fields[index] = new_name
            for emp in self.data.values():
                if old_name in emp:
                    emp[new_name] = emp.pop(old_name)
            return self.save_db()
        return False
    def delete_field(self, field_name):
        self.load_db()
        if field_name in self.fields and field_name not in ['baja', 'mision']:
            self.fields.remove(field_name)
            for emp in self.data.values():
                if field_name in emp:
                    del emp[field_name]
            return self.save_db()
        return False

class ModernMainWindow(QMainWindow):
    ESCALA_OPTIONS = [
        "Oficiales", "Suboficiales", "Tropa", "Personal Laboral"
    ]
    EMPLEO_DICT = {
        "Oficiales": [
            "Coronel", "Teniente Coronel", "Comandante", "Capitán", "Teniente"
        ],
        "Suboficiales": [
            "Suboficial Mayor", "Subteniente", "Brigada", "Sargento Primero", "Sargento"
        ],
        "Tropa": [
            "Cabo Mayor", "Cabo Primero", "Cabo", "Soldado"
        ],
        "Personal Laboral": []
    }
    DESTINO_OPTIONS = [
        "Unidad A", "Unidad B", "Unidad C", "Otra"
    ]
    def __init__(self):
        super().__init__()
        self.db = EmployeeDatabase()
        self.config = AppConfig()
        self.logger = ChangeLogger()  # --- NUEVO: logger
        self.db.load_db()
        self.config.load_config()
        self.current_employee_id = None
        self.unsaved_changes = False
        self.backup_timer = QTimer(self)
        self.backup_timer.timeout.connect(self.automatic_backup)
        self.backup_timer.start(15 * 60 * 1000)
        self.last_backup_file = None
        self.form_tab = None
        self.field_inputs = {}
        self.form_widget = None
        self.form_layout = None
        self.setup_ui()
        self.load_employees()
        self.apply_settings()
        self.automatic_backup()

    def log_event(self, text):
        self.logger.log(text)
        if hasattr(self, "change_log_text"):
            self.change_log_text.setPlainText(self.logger.read_log())

    def load_all_data_and_reload(self):
        self.db.load_db()
        self.config.load_config()
        self.apply_settings()
        self.load_employees()
        self.build_search_filters()
        self.update_field_combobox()
        self.fields_list.setText(", ".join([f.capitalize() for f in self.db.fields]))
        self.search_table.setColumnCount(len(self.db.fields))
        self.search_table.setHorizontalHeaderLabels([f.capitalize() for f in self.db.fields])
        self.table.setColumnCount(len(self.db.fields))
        self.table.setHorizontalHeaderLabels([f.capitalize() for f in self.db.fields])

    def setup_ui(self):
        self.setWindowTitle("Gestión de Personal - Ejército del Aire y del Espacio")
        self.setGeometry(100, 100, 1100, 750)
        logo_path = os.path.join(os.path.dirname(__file__), "icons/logo.png")
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        # HEADER
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        if os.path.exists(logo_path):
            logo_pix = QPixmap(logo_path).scaled(90, 90, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(logo_pix)
        header_text = QLabel(
            "<span style='font-size:28px;font-weight:bold;color:#183a6b'>Gestión de Personal</span><br>"
            "<span style='font-size:16px; color:#183a6b;'>Ejército del Aire y del Espacio</span>"
        )
        header_text.setStyleSheet("padding-left:20px;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(header_text)
        header_layout.addStretch()
        # Fecha y hora actual del sistema en el encabezado
        self.datetime_label = QLabel()
        self.datetime_label.setStyleSheet("font-size:18px;color:#183a6b;padding-right:20px;")
        header_layout.addWidget(self.datetime_label, alignment=Qt.AlignRight | Qt.AlignVCenter)
        # Timer para actualizar fecha y hora cada segundo
        self.datetime_timer = QTimer(self)
        self.datetime_timer.timeout.connect(self.update_datetime_label)
        self.datetime_timer.start(1000)
        self.update_datetime_label()
        main_layout.addLayout(header_layout)
        self.tab_widget = QTabWidget()
        icon_dir = os.path.join(os.path.dirname(__file__), "icons")
        self.icons = {
            "empleados": QIcon(os.path.join(icon_dir, "icon_empleados.png")),
            "formulario": QIcon(os.path.join(icon_dir, "icon_formulario.png")),
            "busqueda": QIcon(os.path.join(icon_dir, "icon_busqueda.png")),
            "campos": QIcon(os.path.join(icon_dir, "icon_campos.png")),
            "configuracion": QIcon(os.path.join(icon_dir, "icon_configuracion.png")),
        }
        main_layout.addWidget(self.tab_widget)
        self.setup_employee_list_tab()
        self.setup_employee_form_tab(init=True)
        self.setup_search_tab()
        self.setup_fields_management_tab()
        self.setup_settings_tab()
        self.apply_styles()
    def apply_settings(self):
        self.config.load_config()
        font = QFont()
        font.setPointSize(self.config.settings['font_size'])
        self.setFont(font)
        if self.config.settings['dark_mode']:
            dark_style = """
                QMainWindow, QDialog {
                    background-color: #2d2d2d;
                }
                QWidget {
                    color: #f0f0f0;
                }
                QGroupBox {
                    background-color: #3d3d3d;
                    border-color: #555;
                }
                QTableWidget {
                    background-color: #3d3d3d;
                    color: #f0f0f0;
                    gridline-color: #555;
                }
                QHeaderView::section {
                    background-color: #4d4d4d;
                    color: #f0f0f0;
                }
                QLineEdit, QComboBox {
                    background-color: #4d4d4d;
                    color: #f0f0f0;
                    border: 1px solid #555;
                }
                QPushButton {
                    background-color: #555;
                    color: white;
                    border: 1px solid #666;
                }
                QPushButton.primary {
                    background-color: #0069d9;
                }
                QPushButton.danger {
                    background-color: #c82333;
                }
                QPushButton.secondary {
                    background-color: #5a6268;
                }
            """
            self.setStyleSheet(self.styleSheet() + dark_style)
        else:
            self.apply_styles()
    def update_datetime_label(self):
        now = QDateTime.currentDateTime()
        meses = [
            "enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
        ]
        dias = [
            "lunes", "martes", "miércoles", "jueves",
            "viernes", "sábado", "domingo"
        ]
        dia_semana = dias[now.date().dayOfWeek()-1]
        dia = now.date().day()
        mes = meses[now.date().month()-1]
        anio = now.date().year()
        hora = now.time().toString("HH:mm:ss")
        texto = f"<b>{dia_semana.capitalize()}, {dia} de {mes} de {anio}</b> {hora}"
        self.datetime_label.setText(texto)

    def automatic_backup(self):
        backup_path = self.config.settings.get('backup_path', '')
        if not backup_path:
            return
        src_db = self.db.db_file
        if not os.path.exists(src_db):
            return
        if os.path.isfile(backup_path):
            try:
                os.remove(backup_path)
            except Exception:
                pass
        try:
            shutil.copy2(src_db, backup_path)
            self.last_backup_file = backup_path
        except Exception:
            pass

    def apply_styles(self):
        style = """
            QMainWindow {
                background-color: #f8fafc;
                font-family: 'Segoe UI', 'Arial', 'sans-serif';
                font-size: 13px;
            }
            QLabel, QGroupBox, QTabWidget, QTableWidget, QTabBar::tab {
                color: #183a6b;
            }
            QTabWidget::pane {
                border: 1px solid #183a6b;
                padding: 5px;
            }
            QTabBar::tab {
                padding: 8px 18px;
                background: #e9f0fa;
                border: 1px solid #bcd0ee;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected {
                background: #fff;
                border-bottom: 1px solid #fff;
                color: #183a6b;
            }
            QPushButton {
                padding: 6px 14px;
                border-radius: 6px;
                border: 1px solid #bcd0ee;
                background: #f3f7fc;
                color: #183a6b;
                font-weight: bold;
            }
            QPushButton.primary {
                background-color: #183a6b;
                color: white;
            }
            QPushButton.danger {
                background-color: #c82333;
                color: white;
            }
            QPushButton.secondary {
                background-color: #5a6268;
                color: white;
            }
            QTableWidget {
                border: 1px solid #bcd0ee;
                gridline-color: #e9f0fa;
                background: #fff;
            }
            QHeaderView::section {
                background-color: #dde8fa;
                padding: 8px;
                border: none;
                color: #183a6b;
            }
            QLineEdit, QComboBox {
                padding: 5px;
                border: 1px solid #bcd0ee;
                border-radius: 4px;
                min-height: 30px;
                background: #f9fbfc;
                color: #183a6b;
            }
            QGroupBox {
                border: 1px solid #bcd0ee;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 15px;
                background: #f3f7fc;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """
        self.setStyleSheet(style)

    def setup_employee_list_tab(self):
        tab = QWidget()
        self.tab_widget.addTab(tab, self.icons["empleados"], "Empleados")
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(5, 5, 5, 5)
        toolbar = QWidget()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        self.refresh_btn = QPushButton(self.icons["empleados"], "Actualizar")
        self.refresh_btn.setProperty("class", "primary")
        self.refresh_btn.clicked.connect(self.load_all_data_and_reload)
        self.add_btn = QPushButton(self.icons["formulario"], "Nuevo")
        self.add_btn.setProperty("class", "primary")
        self.add_btn.clicked.connect(self.new_employee)
        self.delete_btn = QPushButton(self.icons["campos"], "Eliminar")
        self.delete_btn.setProperty("class", "danger")
        self.delete_btn.clicked.connect(self.delete_selected_employee)
        self.modify_btn = QPushButton(self.icons["formulario"], "Modificar")
        self.modify_btn.setProperty("class", "secondary")
        self.modify_btn.clicked.connect(self.modify_selected_employee)
        toolbar_layout.addWidget(self.refresh_btn)
        toolbar_layout.addWidget(self.add_btn)
        toolbar_layout.addWidget(self.delete_btn)
        toolbar_layout.addWidget(self.modify_btn)
        toolbar_layout.addStretch()
        layout.addWidget(toolbar)
        self.table = QTableWidget()
        self.table.setColumnCount(len(self.db.fields))
        headers = [f.capitalize() for f in self.db.fields]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)
        layout.addWidget(self.table)

    def setup_employee_form_tab(self, init=False):
        if self.form_tab is not None and not init:
            index = self.tab_widget.indexOf(self.form_tab)
            if index != -1:
                self.tab_widget.removeTab(index)
        tab = QWidget()
        self.form_tab = tab
        if init:
            self.tab_widget.addTab(tab, self.icons["formulario"], "Formulario")
        else:
            self.tab_widget.insertTab(1, tab, self.icons["formulario"], "Formulario")
            self.tab_widget.setCurrentWidget(tab)
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(5, 5, 5, 5)
        form_container = QWidget()
        form_layout = QVBoxLayout(form_container)
        self.form_widget = QWidget()
        self.form_layout = QFormLayout(self.form_widget)
        self.form_layout.setLabelAlignment(Qt.AlignRight)
        self.form_layout.setFormAlignment(Qt.AlignCenter)
        self.form_layout.setSpacing(10)
        self.field_inputs = {}
        for field in self.db.fields:
            if field == "escala":
                escala_combo = QComboBox()
                escala_combo.addItems(self.ESCALA_OPTIONS)
                escala_combo.currentTextChanged.connect(self.sync_escala_empleo)
                self.field_inputs[field] = escala_combo
                self.form_layout.addRow("Escala:", escala_combo)
            elif field == "empleo":
                empleo_combo = QComboBox()
                self.field_inputs[field] = empleo_combo
                self.form_layout.addRow("Empleo:", empleo_combo)
            elif field == "baja" or field == "mision":
                checkbox = QCheckBox("Sí")
                self.field_inputs[field] = checkbox
                self.form_layout.addRow(QLabel(f"{field.upper()}:"), checkbox)
            elif field == "destino":
                combo = QComboBox()
                combo.addItems(self.DESTINO_OPTIONS)
                self.field_inputs[field] = combo
                self.form_layout.addRow("Destino:", combo)
            elif field in [
                "fecha_incorporacion_unidad", "fecha_incorporacion_fas",
                "fecha_antiguedad_empleo", "fecha_nacimiento", "hps"]:
                cal = QDateEdit()
                cal.setCalendarPopup(True)
                cal.setDisplayFormat("dd/MM/yyyy")
                cal.setDate(QDate.currentDate())
                self.field_inputs[field] = cal
                label = field.replace("_", " ").capitalize() + ":"
                self.form_layout.addRow(QLabel(label), cal)
            else:
                self.field_inputs[field] = QLineEdit()
                self.field_inputs[field].setPlaceholderText(f"Ingrese {field}")
                self.form_layout.addRow(f"{field.capitalize()}:", self.field_inputs[field])
        form_layout.addWidget(self.form_widget)
        layout.addWidget(form_container)
        btn_layout = QHBoxLayout()
        self.save_btn = QPushButton(self.icons["formulario"], "Guardar")
        self.save_btn.setProperty("class", "primary")
        self.save_btn.clicked.connect(self.save_employee)
        self.save_exit_btn = QPushButton(self.icons["formulario"], "Guardar y Salir")
        self.save_exit_btn.setProperty("class", "secondary")
        self.save_exit_btn.clicked.connect(self.save_and_exit)
        self.clear_btn = QPushButton(self.icons["campos"], "Limpiar")
        self.clear_btn.clicked.connect(self.clear_form)
        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.save_exit_btn)
        btn_layout.addWidget(self.clear_btn)
        layout.addLayout(btn_layout)
        self.sync_escala_empleo()

    def sync_escala_empleo(self):
        escala_selected = self.field_inputs["escala"].currentText()
        empleo_combo = self.field_inputs["empleo"]
        empleo_combo.clear()
        empleo_combo.addItems(self.EMPLEO_DICT.get(escala_selected, []))
        empleo_combo.setEnabled(bool(self.EMPLEO_DICT.get(escala_selected, [])))

    def save_and_exit(self):
        if self.save_employee(silent=True):
            self.close()
        else:
            QMessageBox.warning(self, "Error", "No se pudo guardar los cambios")

    def rebuild_form_fields(self):
        self.setup_employee_form_tab()

    def load_employee_to_form(self, emp_id):
        self.db.load_db()
        employee = self.db.get_employee(emp_id)
        if not employee:
            return
        self.current_employee_id = emp_id
        for field in self.db.fields:
            if field in self.field_inputs:
                if field == "baja" or field == "mision":
                    self.field_inputs[field].setChecked(employee.get(field, "NO") == "SI")
                elif field == "escala":
                    idx = self.field_inputs[field].findText(employee.get(field, ""), Qt.MatchExactly)
                    if idx >= 0:
                        self.field_inputs[field].setCurrentIndex(idx)
                    else:
                        self.field_inputs[field].setCurrentIndex(0)
                    self.sync_escala_empleo()
                elif field == "empleo":
                    if self.field_inputs["empleo"].isEnabled():
                        idx = self.field_inputs[field].findText(employee.get(field, ""), Qt.MatchExactly)
                        if idx >= 0:
                            self.field_inputs[field].setCurrentIndex(idx)
                elif field == "destino":
                    idx = self.field_inputs[field].findText(employee.get(field, ""), Qt.MatchExactly)
                    if idx >= 0:
                        self.field_inputs[field].setCurrentIndex(idx)
                elif field in [
                    "fecha_incorporacion_unidad", "fecha_incorporacion_fas",
                    "fecha_antiguedad_empleo", "fecha_nacimiento", "hps"]:
                    try:
                        value = employee.get(field, "")
                        if value:
                            date = QDate.fromString(value, "dd/MM/yyyy")
                            if date.isValid():
                                self.field_inputs[field].setDate(date)
                    except:
                        pass
                else:
                    self.field_inputs[field].setText(str(employee.get(field, '')))
        self.unsaved_changes = False

    def save_employee(self, silent=False):
        employee_data = {}
        for field in self.db.fields:
            if field == "baja" or field == "mision":
                employee_data[field] = "SI" if self.field_inputs[field].isChecked() else "NO"
            elif field == "escala":
                employee_data[field] = self.field_inputs[field].currentText()
            elif field == "empleo":
                empleo_combo = self.field_inputs[field]
                employee_data[field] = empleo_combo.currentText() if empleo_combo.isEnabled() else ""
            elif field == "destino":
                employee_data[field] = self.field_inputs[field].currentText()
            elif field in [
                "fecha_incorporacion_unidad", "fecha_incorporacion_fas",
                "fecha_antiguedad_empleo", "fecha_nacimiento", "hps"]:
                employee_data[field] = self.field_inputs[field].date().toString("dd/MM/yyyy")
            else:
                employee_data[field] = self.field_inputs[field].text().strip()

        # Detect action for log
        action = None
        if self.current_employee_id is None:
            action = "Alta de empleado: " + " ".join([employee_data.get("nombre", ""), employee_data.get("apellidos", "")])
        else:
            prev = self.db.get_employee(self.current_employee_id)
            if prev:
                cambios = []
                for k in self.db.fields:
                    if str(prev.get(k, "")) != str(employee_data.get(k, "")):
                        cambios.append(f"{k}: '{prev.get(k, '')}' -> '{employee_data.get(k, '')}'")
                if cambios:
                    action = f"Modificación de empleado {self.current_employee_id}: " + "; ".join(cambios)

        self.db.load_db()
        ok = False
        if self.current_employee_id is None:
            emp_id = self.db.add_employee(employee_data)
            if emp_id:
                ok = True
                if not silent:
                    QMessageBox.information(self, "Éxito", f"Empleado añadido correctamente")
                self.current_employee_id = emp_id
                self.unsaved_changes = False
                self.clear_form()
                self.load_all_data_and_reload()
        else:
            if self.db.update_employee(self.current_employee_id, employee_data):
                ok = True
                if not silent:
                    QMessageBox.information(self, "Éxito", "Empleado actualizado correctamente")
                self.unsaved_changes = False
                self.load_all_data_and_reload()
        if ok and action:
            self.log_event(action)
        return ok

    def clear_form(self):
        self.current_employee_id = None
        for field in self.db.fields:
            if field in self.field_inputs:
                if isinstance(self.field_inputs[field], QLineEdit):
                    self.field_inputs[field].clear()
                elif isinstance(self.field_inputs[field], QCheckBox):
                    self.field_inputs[field].setChecked(False)
                elif isinstance(self.field_inputs[field], QComboBox):
                    self.field_inputs[field].setCurrentIndex(0)
                elif isinstance(self.field_inputs[field], QDateEdit):
                    self.field_inputs[field].setDate(QDate.currentDate())
        self.sync_escala_empleo()

    def setup_search_tab(self):
        tab = QWidget()
        self.tab_widget.addTab(tab, self.icons["busqueda"], "Búsqueda")
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(5, 5, 5, 5)
        self.filter_container = QWidget()
        filter_layout = QVBoxLayout(self.filter_container)
        self.filter_group = QGroupBox("Filtros de Búsqueda")
        self.group_layout = QFormLayout(self.filter_group)
        self.group_layout.setLabelAlignment(Qt.AlignRight)
        self.group_layout.setVerticalSpacing(10)
        self.group_layout.setHorizontalSpacing(15)
        self.search_filters = {}
        self.build_search_filters()
        filter_layout.addWidget(self.filter_group)
        filter_layout.addStretch()
        search_btn = QPushButton(self.icons["busqueda"], "Buscar")
        search_btn.setProperty("class", "primary")
        search_btn.clicked.connect(self.perform_search)
        export_btn = QPushButton(self.icons["campos"], "Exportar a Excel")
        export_btn.setProperty("class", "secondary")
        export_btn.clicked.connect(self.export_search_results_to_excel)
        btns = QHBoxLayout()
        btns.addWidget(search_btn)
        btns.addWidget(export_btn)
        btns.addStretch()
        results_group = QGroupBox("Resultados")
        results_layout = QVBoxLayout(results_group)
        self.search_table = QTableWidget()
        self.search_table.setColumnCount(len(self.db.fields))
        self.search_table.setHorizontalHeaderLabels([f.capitalize() for f in self.db.fields])
        self.search_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.search_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.search_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.search_table.verticalHeader().setVisible(False)
        self.search_table.setSortingEnabled(True)
        results_layout.addWidget(self.search_table)
        layout.addWidget(self.filter_container)
        layout.addLayout(btns)
        layout.addWidget(results_group)

    def build_search_filters(self):
        self.search_filters = {}
        while self.group_layout.rowCount():
            self.group_layout.removeRow(0)
        for field in self.db.fields:
            if field == "baja" or field == "mision":
                checkbox = QCheckBox("Sí")
                self.group_layout.addRow(f"{field.upper()}:", checkbox)
                self.search_filters[field] = checkbox
            elif field == "escala":
                combo = QComboBox()
                combo.addItem("")
                combo.addItems(self.ESCALA_OPTIONS)
                self.group_layout.addRow("Escala:", combo)
                self.search_filters[field] = combo
            elif field == "empleo":
                combo = QComboBox()
                combo.addItem("")
                self.group_layout.addRow("Empleo:", combo)
                self.search_filters[field] = combo
            elif field == "destino":
                combo = QComboBox()
                combo.addItem("")
                combo.addItems(self.DESTINO_OPTIONS)
                self.group_layout.addRow("Destino:", combo)
                self.search_filters[field] = combo
            elif field in [
                "fecha_incorporacion_unidad", "fecha_incorporacion_fas",
                "fecha_antiguedad_empleo", "fecha_nacimiento", "hps"]:
                cal = QDateEdit()
                cal.setCalendarPopup(True)
                cal.setDisplayFormat("dd/MM/yyyy")
                cal.setDate(QDate.currentDate())
                self.group_layout.addRow(QLabel(field.replace("_", " ").capitalize() + ":"), cal)
                self.search_filters[field] = cal
            else:
                input_field = QLineEdit()
                input_field.setPlaceholderText(f"Buscar por {field}")
                self.group_layout.addRow(f"{field.capitalize()}:", input_field)
                self.search_filters[field] = input_field

        if "escala" in self.search_filters and "empleo" in self.search_filters:
            self.search_filters["escala"].currentTextChanged.connect(self.sync_escala_empleo_search)
        self.sync_escala_empleo_search()

    def sync_escala_empleo_search(self):
        escala_selected = self.search_filters["escala"].currentText()
        empleo_combo = self.search_filters["empleo"]
        empleo_combo.blockSignals(True)
        empleo_combo.clear()
        empleo_combo.addItem("")
        empleo_combo.addItems(self.EMPLEO_DICT.get(escala_selected, []))
        empleo_combo.setEnabled(bool(self.EMPLEO_DICT.get(escala_selected, [])))
        empleo_combo.blockSignals(False)

    def perform_search(self):
        self.load_all_data_and_reload()
        filters = {}
        for field, widget in self.search_filters.items():
            if isinstance(widget, QCheckBox):
                value = "SI" if widget.isChecked() else ""
                if value:
                    filters[field] = value
            elif isinstance(widget, QComboBox):
                value = widget.currentText()
                if value:
                    filters[field] = value
            elif isinstance(widget, QDateEdit):
                if widget.date() != QDate.currentDate():
                    filters[field] = widget.date().toString("dd/MM/yyyy")
            else:
                value = widget.text().strip()
                if value:
                    filters[field] = value
        results = []
        for emp_id, emp_data in self.db.get_all_employees().items():
            match = True
            for field, value in filters.items():
                if isinstance(self.search_filters[field], QDateEdit):
                    if emp_data.get(field, "") != value:
                        match = False
                elif isinstance(self.search_filters[field], QCheckBox):
                    if emp_data.get(field, "NO") != value:
                        match = False
                elif isinstance(self.search_filters[field], QComboBox):
                    if emp_data.get(field, "") != value:
                        match = False
                else:
                    if value.lower() not in str(emp_data.get(field, '')).lower():
                        match = False
            if match:
                results.append(emp_data)
        self.search_table.setRowCount(len(results))
        for row, emp_data in enumerate(results):
            color_row = None
            if emp_data.get('baja', 'NO') == 'SI':
                color_row = QBrush(QColor(255, 125, 125))
            elif emp_data.get('mision', 'NO') == 'SI':
                color_row = QBrush(QColor(180, 220, 255))
            for col, field in enumerate(self.db.fields):
                value = str(emp_data.get(field, ''))
                item = QTableWidgetItem(value)
                if field == "hps" and value:
                    try:
                        hps_date = datetime.datetime.strptime(value, "%d/%m/%Y").date()
                        if (QDate.currentDate().toPyDate() - hps_date).days < 365:
                            item.setText(f"{value}  !")
                            item.setForeground(QColor("red"))
                            font = item.font()
                            font.setBold(True)
                            item.setFont(font)
                    except Exception:
                        pass
                if color_row is not None:
                    item.setBackground(color_row)
                self.search_table.setItem(row, col, item)
        self.search_table.resizeColumnsToContents()

    def export_search_results_to_excel(self):
        row_count = self.search_table.rowCount()
        col_count = self.search_table.columnCount()
        if row_count == 0:
            QMessageBox.information(self, "Sin resultados", "No hay datos en la tabla de búsqueda para exportar.")
            return

        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar como Excel",
            "resultados_busqueda.xlsx",
            "Archivos Excel (*.xlsx)"
        )
        if not filename:
            return

        wb = Workbook()
        ws = wb.active

        headers = []
        for col in range(col_count):
            header = self.search_table.horizontalHeaderItem(col)
            headers.append(header.text() if header else "")
        ws.append(headers)

        for row in range(row_count):
            row_data = []
            for col in range(col_count):
                item = self.search_table.item(row, col)
                row_data.append(item.text() if item else "")
            ws.append(row_data)

        for col in range(1, col_count + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18

        try:
            wb.save(filename)
            QMessageBox.information(self, "Exportación exitosa", f"Resultados exportados a:\n{filename}")
        except Exception as e:
            QMessageBox.critical(self, "Error al exportar", f"No se pudo exportar:\n{e}")

    def setup_fields_management_tab(self):
        tab = QWidget()
        self.tab_widget.addTab(tab, self.icons["campos"], "Campos")
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(5, 5, 5, 5)
        edit_group = QGroupBox("Modificar Campos Existentes")
        edit_layout = QVBoxLayout(edit_group)
        self.field_combobox = QComboBox()
        self.update_field_combobox()
        self.new_name_input = QLineEdit()
        self.new_name_input.setPlaceholderText("Nuevo nombre para el campo")
        update_btn = QPushButton(self.icons["formulario"], "Actualizar Nombre")
        update_btn.setProperty("class", "primary")
        update_btn.clicked.connect(self.update_field_name)
        delete_btn = QPushButton(self.icons["campos"], "Eliminar Campo")
        delete_btn.setProperty("class", "danger")
        delete_btn.clicked.connect(self.delete_field)
        edit_layout.addWidget(QLabel("Campo a modificar:"))
        edit_layout.addWidget(self.field_combobox)
        edit_layout.addWidget(QLabel("Nuevo nombre:"))
        edit_layout.addWidget(self.new_name_input)
        edit_layout.addWidget(update_btn)
        edit_layout.addWidget(delete_btn)
        edit_layout.addStretch()
        add_group = QGroupBox("Añadir Nuevo Campo")
        add_layout = QFormLayout(add_group)
        self.new_field_name_input = QLineEdit()
        self.new_field_name_input.setPlaceholderText("Nombre del nuevo campo")
        self.default_value_input = QLineEdit()
        self.default_value_input.setPlaceholderText("Valor por defecto")
        add_btn = QPushButton(self.icons["campos"], "Añadir Campo")
        add_btn.setProperty("class", "primary")
        add_btn.clicked.connect(self.add_custom_field)
        add_layout.addRow("Nombre:", self.new_field_name_input)
        add_layout.addRow("Valor por defecto:", self.default_value_input)
        add_layout.addRow(add_btn)
        self.fields_list_label = QLabel("Campos actuales:")
        self.fields_list_label.setStyleSheet("font-weight: bold;")
        self.fields_list = QLabel(", ".join([f.capitalize() for f in self.db.fields]))
        self.fields_list.setWordWrap(True)
        layout.addWidget(edit_group)
        layout.addWidget(add_group)
        layout.addWidget(self.fields_list_label)
        layout.addWidget(self.fields_list)
        layout.addStretch()

    def update_field_combobox(self):
        self.db.load_db()
        self.field_combobox.clear()
        self.field_combobox.addItems([f.capitalize() for f in self.db.fields])

    def add_custom_field(self):
        field_name = self.new_field_name_input.text().strip().lower()
        if not field_name:
            QMessageBox.warning(self, "Error", "El nombre del campo no puede estar vacío")
            return
        if not all(c.isalnum() or c == '_' for c in field_name):
            QMessageBox.warning(self, "Error", "El nombre del campo solo puede contener letras, números y guiones bajos")
            return
        if field_name in ['baja', 'mision']:
            QMessageBox.warning(self, "Error", "No puede añadir un campo llamado BAJA o MISION")
            return
        default_value = self.default_value_input.text().strip()
        if self.db.add_field(field_name, default_value):
            self.db.save_db()
            self.log_event(f"Añadido campo nuevo: {field_name}")
            self.load_all_data_and_reload()
            QMessageBox.information(self, "Éxito", f"Campo '{field_name}' añadido correctamente")
            self.new_field_name_input.clear()
            self.default_value_input.clear()
        else:
            QMessageBox.warning(self, "Error", "El campo ya existe")

    def update_field_name(self):
        if self.field_combobox.count() == 0:
            QMessageBox.warning(self, "Error", "No hay campos para modificar")
            return
        old_name = self.db.fields[self.field_combobox.currentIndex()]
        new_name = self.new_name_input.text().strip().lower()
        if not new_name:
            QMessageBox.warning(self, "Error", "El nuevo nombre no puede estar vacío")
            return
        if not all(c.isalnum() or c == '_' for c in new_name):
            QMessageBox.warning(self, "Error", "El nuevo nombre solo puede contener letras, números y guiones bajos")
            return
        if new_name in self.db.fields or new_name in ['baja', 'mision']:
            QMessageBox.warning(self, "Error", "El nombre del campo ya existe o es reservado")
            return
        if self.db.update_field_name(old_name, new_name):
            self.db.save_db()
            self.log_event(f"Renombrado campo: {old_name} -> {new_name}")
            self.load_all_data_and_reload()
            QMessageBox.information(self, "Éxito", f"Campo '{old_name}' renombrado a '{new_name}'")
            self.new_name_input.clear()
        else:
            QMessageBox.warning(self, "Error", "No se pudo actualizar el campo")

    def delete_field(self):
        if self.field_combobox.count() == 0:
            QMessageBox.warning(self, "Error", "No hay campos para eliminar")
            return
        field_name = self.db.fields[self.field_combobox.currentIndex()]
        if field_name in ['baja', 'mision']:
            QMessageBox.warning(self, "Error", "No se puede eliminar un campo obligatorio o reservado")
            return
        reply = QMessageBox.question(
            self, 'Confirmar',
            f'¿Está seguro que desea eliminar el campo "{field_name}"? Esta acción no se puede deshacer.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            if self.db.delete_field(field_name):
                self.db.save_db()
                self.log_event(f"Eliminado campo: {field_name}")
                self.load_all_data_and_reload()
                QMessageBox.information(self, "Éxito", f"Campo '{field_name}' eliminado correctamente")
                self.new_name_input.clear()
            else:
                QMessageBox.warning(self, "Error", "No se pudo eliminar el campo")

    def update_ui_after_field_change(self):
        self.fields_list.setText(", ".join([f.capitalize() for f in self.db.fields]))
        self.update_field_combobox()
        headers = [f.capitalize() for f in self.db.fields]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.load_all_data_and_reload()
        self.rebuild_form_fields()
        self.build_search_filters()
        self.search_table.setColumnCount(len(headers))
        self.search_table.setHorizontalHeaderLabels(headers)
        self.db.save_db()

    def setup_settings_tab(self):
        tab = QWidget()
        self.tab_widget.addTab(tab, self.icons["configuracion"], "Configuración")
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(5, 5, 5, 5)
        settings_group = QGroupBox("Configuración de la Aplicación")
        settings_layout = QFormLayout(settings_group)
        self.auto_save_check = QCheckBox("Guardado automático")
        self.auto_save_check.setChecked(self.config.settings['auto_save'])
        self.auto_save_check.stateChanged.connect(lambda: self.update_config('auto_save', self.auto_save_check.isChecked()))
        self.dark_mode_check = QCheckBox("Modo oscuro")
        self.dark_mode_check.setChecked(self.config.settings['dark_mode'])
        self.dark_mode_check.stateChanged.connect(lambda: self.update_config('dark_mode', self.dark_mode_check.isChecked()))
        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(['10', '12', '14', '16'])
        self.font_size_combo.setCurrentText(str(self.config.settings['font_size']))
        self.font_size_combo.currentTextChanged.connect(lambda: self.update_config('font_size', int(self.font_size_combo.currentText())))
        self.default_dept_input = QLineEdit()
        self.default_dept_input.setText(self.config.settings['default_dept'])
        self.default_dept_input.textChanged.connect(lambda: self.update_config('default_dept', self.default_dept_input.text()))
        self.backup_path_line = QLineEdit()
        self.backup_path_line.setReadOnly(True)
        self.backup_path_line.setText(self.config.settings.get('backup_path', ''))
        self.backup_path_btn = QPushButton("Elegir ruta de copia")
        self.backup_path_btn.clicked.connect(self.choose_backup_path)
        backup_layout = QHBoxLayout()
        backup_layout.addWidget(self.backup_path_line)
        backup_layout.addWidget(self.backup_path_btn)
        settings_layout.addRow("Ruta copia seguridad:", backup_layout)
        settings_layout.addRow("Guardado automático:", self.auto_save_check)
        settings_layout.addRow("Modo oscuro:", self.dark_mode_check)
        settings_layout.addRow("Tamaño de fuente:", self.font_size_combo)
        settings_layout.addRow("Departamento por defecto:", self.default_dept_input)
        layout.addWidget(settings_group)

        # --------------- REGISTRO DE CAMBIOS ---------------
        changelog_group = QGroupBox("Registro de Cambios")
        changelog_layout = QVBoxLayout(changelog_group)
        self.change_log_text = QTextEdit()
        self.change_log_text.setReadOnly(True)
        self.change_log_text.setStyleSheet("font-family:monospace;")
        self.change_log_text.setPlainText(self.logger.read_log())
        changelog_layout.addWidget(self.change_log_text)
        layout.addWidget(changelog_group)
        # ----------------------------------------------------
        layout.addStretch()

    def choose_backup_path(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Elegir ruta para la copia de seguridad",
            "personal_db_backup.dat",
            "Archivos de datos (*.dat)"
        )
        if path:
            self.config.update_setting('backup_path', path)
            self.backup_path_line.setText(path)
            self.log_event(f"Cambio de ruta de copia de seguridad: {path}")
            self.automatic_backup()

    def update_config(self, key, value):
        self.config.load_config()
        changed = False
        old_value = self.config.settings.get(key)
        if self.config.update_setting(key, value):
            self.apply_settings()
            self.config.save_config()
            self.load_all_data_and_reload()
            changed = True
        if changed and old_value != value:
            self.log_event(f"Cambio de configuración: {key} de '{old_value}' a '{value}'")
        return changed

    def load_employees(self):
        self.db.load_db()
        employees = self.db.get_all_employees()
        self.table.setRowCount(len(employees))
        for row, (emp_id, emp_data) in enumerate(employees.items()):
            color_row = None
            if emp_data.get('baja', 'NO') == 'SI':
                color_row = QBrush(QColor(255, 125, 125))
            elif emp_data.get('mision', 'NO') == 'SI':
                color_row = QBrush(QColor(180, 220, 255))
            for col, field in enumerate(self.db.fields):
                value = str(emp_data.get(field, ''))
                item = QTableWidgetItem(value)
                if field == "hps" and value:
                    try:
                        hps_date = datetime.datetime.strptime(value, "%d/%m/%Y").date()
                        if (QDate.currentDate().toPyDate() - hps_date).days < 365:
                            item.setText(f"{value}  !")
                            item.setForeground(QColor("red"))
                            font = item.font()
                            font.setBold(True)
                            item.setFont(font)
                    except Exception:
                        pass
                if color_row is not None:
                    item.setBackground(color_row)
                self.table.setItem(row, col, item)
        self.table.resizeColumnsToContents()

    def new_employee(self):
        self.current_employee_id = None
        self.clear_form()
        self.tab_widget.setCurrentWidget(self.form_tab)
        first_field = next(iter(self.field_inputs.values()))
        if isinstance(first_field, QLineEdit):
            first_field.setFocus()
        self.sync_escala_empleo()
    def delete_selected_employee(self):
        selected = self.table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Advertencia", "Por favor seleccione un empleado")
            return
        row = self.table.currentRow()
        if row == -1:
            QMessageBox.warning(self, "Advertencia", "Por favor seleccione un empleado")
            return
        employees = list(self.db.get_all_employees().items())
        if row < 0 or row >= len(employees):
            QMessageBox.warning(self, "Error", "No se pudo identificar el empleado para eliminar")
            return
        emp_id, emp_data = employees[row]
        reply = QMessageBox.question(
            self, 'Confirmar',
            f'¿Está seguro que desea eliminar a este empleado?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            if self.db.delete_employee(emp_id):
                self.log_event(f"Eliminado empleado: ID {emp_id} - {emp_data.get('nombre', '')} {emp_data.get('apellidos', '')}")
                QMessageBox.information(self, "Éxito", "Empleado eliminado correctamente")
                self.load_all_data_and_reload()
                if self.current_employee_id == emp_id:
                    self.clear_form()
            else:
                QMessageBox.warning(self, "Error", "No se pudo eliminar el empleado")
    def modify_selected_employee(self):
        selected = self.table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Advertencia", "Por favor seleccione un empleado")
            return
        row = self.table.currentRow()
        if row == -1:
            QMessageBox.warning(self, "Advertencia", "Por favor seleccione un empleado")
            return
        employees = list(self.db.get_all_employees().items())
        if row < 0 or row >= len(employees):
            QMessageBox.warning(self, "Error", "No se pudo identificar el empleado para modificar")
            return
        emp_id, _ = employees[row]
        self.setup_employee_form_tab()  # reconstruir campos sin duplicar pestaña
        self.load_employee_to_form(emp_id)
        self.tab_widget.setCurrentWidget(self.form_tab)
        self.sync_escala_empleo()

    def closeEvent(self, event):
        self.db.save_db()
        self.config.save_config()
        event.accept()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        os.chdir(sys._MEIPASS)
    app = QApplication(sys.argv)
    palette = app.palette()
    palette.setColor(palette.Window, QColor(240, 240, 240))
    palette.setColor(palette.WindowText, Qt.black)
    palette.setColor(palette.Base, QColor(255, 255, 255))
    palette.setColor(palette.Text, Qt.black)
    app.setPalette(palette)
    window = ModernMainWindow()
    window.show()
    sys.exit(app.exec_())
