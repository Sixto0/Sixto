import os
import sys
import pickle
import shutil
import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QTabWidget, QFormLayout, QHeaderView, QGroupBox,
    QComboBox, QCheckBox, QFileDialog, QDateEdit, QTextEdit
)
from PyQt5.QtCore import Qt, QTimer, QDate, QDateTime
from PyQt5.QtGui import QFont, QColor, QBrush, QIcon, QPixmap
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# --- Utility Classes ---
class ChangeLogger:
    def __init__(self, log_file="change_log.txt"):
        self.log_file = log_file
        if not os.path.exists(self.log_file):
            with open(self.log_file, "w", encoding="utf-8") as f:
                f.write("Fecha y hora\tEvento\n")

    def log(self, event: str):
        fecha_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(f"{fecha_hora}\t{event}\n")

    def read_log(self):
        if not os.path.exists(self.log_file):
            return ""
        with open(self.log_file, "r", encoding="utf-8") as f:
            return f.read()

class AppConfig:
    DEFAULTS = {
        'auto_save': True, 'dark_mode': False, 'font_size': 12,
        'default_dept': 'General', 'backup_path': ''
    }
    def __init__(self, config_file='app_config.dat'):
        self.config_file = config_file
        self.settings = self.DEFAULTS.copy()
        self.load_config()
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'rb') as f:
                    self.settings.update(pickle.load(f))
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
        self.settings[key] = value
        return self.save_config()

class EmployeeDatabase:
    BASE_FIELDS = [
        'nombre', 'apellidos', 'escala', 'empleo', 'destino',
        'fecha_incorporacion_unidad', 'fecha_incorporacion_fas',
        'fecha_antiguedad_empleo', 'fecha_nacimiento', 'dni',
        'nº_telf_perosnal', 'nº_telf_extension', 'correo_personal',
        'correo_militar', 'baja', 'mision', 'hps'
    ]
    def __init__(self, db_file='personal_db.dat'):
        self.db_file = db_file
        self.fields = self.BASE_FIELDS.copy()
        self.data = {}
        self.load_db()
        self._normalize_data()

    def _normalize_data(self):
        for emp in self.data.values():
            for field in self.fields:
                if field not in emp:
                    emp[field] = 'NO' if field in ['baja', 'mision'] else ''
        for emp in self.data.values():
            for key in list(emp.keys()):
                if key not in self.fields:
                    emp.pop(key)

    def load_db(self):
        if os.path.exists(self.db_file):
            try:
                with open(self.db_file, 'rb') as f:
                    obj = pickle.load(f)
                    self.fields, self.data = (obj if isinstance(obj, tuple) else (self.BASE_FIELDS, obj))
            except Exception:
                self.data = {}
        else:
            self.data = {}
        self._normalize_data()

    def save_db(self):
        try:
            with open(self.db_file, 'wb') as f:
                pickle.dump((self.fields, self.data), f)
            return True
        except Exception:
            return False

    def get_next_id(self):
        if not self.data: return "1"
        try:
            return str(max(int(k) for k in self.data if k.isdigit()) + 1)
        except Exception:
            i = 1
            while str(i) in self.data: i += 1
            return str(i)

    def add_employee(self, employee_data):
        emp_id = self.get_next_id()
        self.data[emp_id] = employee_data
        return emp_id if self.save_db() else None

    def update_employee(self, emp_id, employee_data):
        if emp_id not in self.data: return False
        self.data[emp_id] = employee_data
        return self.save_db()

    def delete_employee(self, emp_id):
        if emp_id in self.data:
            self.data.pop(emp_id)
            return self.save_db()
        return False

    def get_employee(self, emp_id):
        return self.data.get(emp_id)

    def get_all_employees(self):
        return self.data

    def add_field(self, field_name, default_value=''):
        if field_name not in self.fields:
            self.fields.append(field_name)
            for emp in self.data.values():
                emp[field_name] = default_value
            return self.save_db()
        return False

    def update_field_name(self, old, new):
        if old in self.fields and new not in self.fields:
            idx = self.fields.index(old)
            self.fields[idx] = new
            for emp in self.data.values():
                if old in emp:
                    emp[new] = emp.pop(old)
            return self.save_db()
        return False

    def delete_field(self, field):
        if field in self.fields and field not in ['baja', 'mision']:
            self.fields.remove(field)
            for emp in self.data.values():
                emp.pop(field, None)
            return self.save_db()
        return False

# --- Main Application ---
class ModernMainWindow(QMainWindow):
    ESCALA_OPTIONS = [
        "Oficiales", "Suboficiales", "Tropa", "Personal Laboral"
    ]
    EMPLEO_DICT = {
        "Oficiales": ["Coronel", "Teniente Coronel", "Comandante", "Capitán", "Teniente"],
        "Suboficiales": ["Suboficial Mayor", "Subteniente", "Brigada", "Sargento Primero", "Sargento"],
        "Tropa": ["Cabo Mayor", "Cabo Primero", "Cabo", "Soldado"],
        "Personal Laboral": []
    }
    DESTINO_OPTIONS = ["Unidad A", "Unidad B", "Unidad C", "Otra"]

    def __init__(self):
        super().__init__()
        self.db = EmployeeDatabase()
        self.config = AppConfig()
        self.logger = ChangeLogger()
        self.current_employee_id = None
        self.unsaved_changes = False
        self.last_backup_file = None

        self._setup_timers()
        self._setup_ui()
        self._reload_all()

    def _setup_timers(self):
        self.backup_timer = QTimer(self)
        self.backup_timer.timeout.connect(self.automatic_backup)
        self.backup_timer.start(15 * 60 * 1000)
        self.datetime_timer = QTimer(self)
        self.datetime_timer.timeout.connect(self.update_datetime_label)
        self.datetime_timer.start(1000)

    def apply_settings(self):
        font = QFont()
        font.setPointSize(self.config.settings.get('font_size', 12))
        self.setFont(font)
        # Opcional: modo oscuro
        if self.config.settings.get('dark_mode', False):
            dark_style = """
                QMainWindow, QDialog { background-color: #2d2d2d; }
                QWidget { color: #f0f0f0; }
            """
            self.setStyleSheet(self.styleSheet() + dark_style)
        else:
            self.apply_styles()

    # ---- UI ----
    def _setup_ui(self):
        self.setWindowTitle("Gestión de Personal - Ejército del Aire y del Espacio")
        self.setGeometry(100, 100, 1100, 750)
        self._set_icon()
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        self.datetime_label = QLabel()
        main_layout.addLayout(self._build_header())
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        self.icons = self._load_icons()
        self._add_tabs()
        self.apply_styles()

    def _set_icon(self):
        logo_path = os.path.join(os.path.dirname(__file__), "icons/logo.png")
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))

    def _build_header(self):
        h = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), "icons/logo.png")
        if os.path.exists(logo_path):
            logo_pix = QPixmap(logo_path).scaled(90, 90, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(logo_pix)
        header_text = QLabel(
            "<span style='font-size:28px;font-weight:bold;color:#183a6b'>Gestión de Personal</span><br>"
            "<span style='font-size:16px; color:#183a6b;'>Ejército del Aire y del Espacio</span>"
        )
        header_text.setStyleSheet("padding-left:20px;")
        h.addWidget(logo_label)
        h.addWidget(header_text)
        h.addStretch()
        self.datetime_label.setStyleSheet("font-size:18px;color:#183a6b;padding-right:20px;")
        h.addWidget(self.datetime_label, alignment=Qt.AlignRight | Qt.AlignVCenter)
        self.update_datetime_label()
        return h

    def _load_icons(self):
        icon_dir = os.path.join(os.path.dirname(__file__), "icons")
        return {
            "empleados": QIcon(os.path.join(icon_dir, "icon_empleados.png")),
            "formulario": QIcon(os.path.join(icon_dir, "icon_formulario.png")),
            "busqueda": QIcon(os.path.join(icon_dir, "icon_busqueda.png")),
            "campos": QIcon(os.path.join(icon_dir, "icon_campos.png")),
            "configuracion": QIcon(os.path.join(icon_dir, "icon_configuracion.png")),
        }

    def _add_tabs(self):
        self._setup_employee_list_tab()
        self._setup_employee_form_tab(init=True)
        self._setup_search_tab()
        self._setup_fields_management_tab()
        self._setup_settings_tab()

    def _reload_all(self):
        self.db.load_db()
        self.config.load_config()
        self.apply_settings()
        self.load_employees()
        self._rebuild_search_filters()
        self._update_field_combobox()
        self.fields_list.setText(", ".join([f.capitalize() for f in self.db.fields]))
        self.search_table.setColumnCount(len(self.db.fields))
        self.search_table.setHorizontalHeaderLabels([f.capitalize() for f in self.db.fields])
        self.table.setColumnCount(len(self.db.fields))
        self.table.setHorizontalHeaderLabels([f.capitalize() for f in self.db.fields])

    # ---- Employee Table Tab ----
    def _setup_employee_list_tab(self):
        tab = QWidget()
        self.tab_widget.addTab(tab, self.icons["empleados"], "Empleados")
        layout = QVBoxLayout(tab)
        toolbar = QWidget()
        toolbar_layout = QHBoxLayout(toolbar)
        for text, icon, slot, cls in [
            ("Actualizar", "empleados", self._reload_all, "primary"),
            ("Nuevo", "formulario", self.new_employee, "primary"),
            ("Eliminar", "campos", self.delete_selected_employee, "danger"),
            ("Modificar", "formulario", self.modify_selected_employee, "secondary")
        ]:
            btn = QPushButton(self.icons[icon], text)
            btn.setProperty("class", cls)
            btn.clicked.connect(slot)
            toolbar_layout.addWidget(btn)
        toolbar_layout.addStretch()
        layout.addWidget(toolbar)
        self.table = QTableWidget()
        self.table.setColumnCount(len(self.db.fields))
        self.table.setHorizontalHeaderLabels([f.capitalize() for f in self.db.fields])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)
        layout.addWidget(self.table)

    # ---- Employee Form Tab ----
    def _setup_employee_form_tab(self, init=False):
        if hasattr(self, "form_tab") and self.form_tab and not init:
            idx = self.tab_widget.indexOf(self.form_tab)
            if idx != -1: self.tab_widget.removeTab(idx)
        tab = QWidget()
        self.form_tab = tab
        if init:
            self.tab_widget.addTab(tab, self.icons["formulario"], "Formulario")
        else:
            self.tab_widget.insertTab(1, tab, self.icons["formulario"], "Formulario")
            self.tab_widget.setCurrentWidget(tab)
        layout = QVBoxLayout(tab)
        form_container = QWidget()
        form_layout = QVBoxLayout(form_container)
        self.form_widget = QWidget()
        self.form_layout = QFormLayout(self.form_widget)
        self.form_layout.setLabelAlignment(Qt.AlignRight)
        self.form_layout.setFormAlignment(Qt.AlignCenter)
        self.form_layout.setSpacing(10)
        self.field_inputs = {}
        for field in self.db.fields:
            self.field_inputs[field] = self._create_field_widget(field)
            label = field.replace("_", " ").capitalize() + ":"
            self.form_layout.addRow(QLabel(label), self.field_inputs[field])
        form_layout.addWidget(self.form_widget)
        layout.addWidget(form_container)
        btn_layout = QHBoxLayout()
        for text, icon, slot, cls in [
            ("Guardar", "formulario", self.save_employee, "primary"),
            ("Guardar y Salir", "formulario", self.save_and_exit, "secondary"),
            ("Limpiar", "campos", self.clear_form, "")
        ]:
            btn = QPushButton(self.icons[icon], text)
            btn.setProperty("class", cls)
            btn.clicked.connect(slot)
            btn_layout.addWidget(btn)
        layout.addLayout(btn_layout)
        self._sync_escala_empleo()

    def _create_field_widget(self, field):
        if field == "escala":
            cb = QComboBox()
            cb.addItems(self.ESCALA_OPTIONS)
            cb.currentTextChanged.connect(self._sync_escala_empleo)
            return cb
        if field == "empleo":
            return QComboBox()
        if field in ["baja", "mision"]:
            return QCheckBox("Sí")
        if field == "destino":
            cb = QComboBox()
            cb.addItems(self.DESTINO_OPTIONS)
            return cb
        if field in [
            "fecha_incorporacion_unidad", "fecha_incorporacion_fas",
            "fecha_antiguedad_empleo", "fecha_nacimiento", "hps"]:
            cal = QDateEdit()
            cal.setCalendarPopup(True)
            cal.setDisplayFormat("dd/MM/yyyy")
            cal.setDate(QDate.currentDate())
            return cal
        le = QLineEdit()
        le.setPlaceholderText(f"Ingrese {field}")
        return le

    def _sync_escala_empleo(self):
        escala = self.field_inputs.get("escala")
        empleo_combo = self.field_inputs.get("empleo")
        if escala and empleo_combo:
            empleo_combo.clear()
            empleo_combo.addItems(self.EMPLEO_DICT.get(escala.currentText(), []))
            empleo_combo.setEnabled(bool(self.EMPLEO_DICT.get(escala.currentText(), [])))

    def save_and_exit(self):
        if self.save_employee(silent=True): self.close()
        else: QMessageBox.warning(self, "Error", "No se pudo guardar los cambios")

    def clear_form(self):
        self.current_employee_id = None
        for field, widget in self.field_inputs.items():
            if isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, QCheckBox):
                widget.setChecked(False)
            elif isinstance(widget, QComboBox):
                widget.setCurrentIndex(0)
            elif isinstance(widget, QDateEdit):
                widget.setDate(QDate.currentDate())
        self._sync_escala_empleo()

    # ---- Search Tab ----
    def _setup_search_tab(self):
        tab = QWidget()
        self.tab_widget.addTab(tab, self.icons["busqueda"], "Búsqueda")
        layout = QVBoxLayout(tab)
        self.filter_container = QWidget()
        self.filter_group = QGroupBox("Filtros de Búsqueda")
        self.group_layout = QFormLayout(self.filter_group)
        self.search_filters = {}
        self._rebuild_search_filters()
        filter_layout = QVBoxLayout(self.filter_container)
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

    def _rebuild_search_filters(self):
        while self.group_layout.rowCount():
            self.group_layout.removeRow(0)
        self.search_filters.clear()
        for field in self.db.fields:
            self.search_filters[field] = self._create_field_widget(field)
            self.group_layout.addRow(f"{field.capitalize()}:", self.search_filters[field])
        if "escala" in self.search_filters and "empleo" in self.search_filters:
            self.search_filters["escala"].currentTextChanged.connect(self._sync_escala_empleo_search)
        self._sync_escala_empleo_search()

    def _sync_escala_empleo_search(self):
        escala = self.search_filters.get("escala")
        empleo_combo = self.search_filters.get("empleo")
        if escala and empleo_combo:
            empleo_combo.blockSignals(True)
            empleo_combo.clear()
            empleo_combo.addItem("")
            empleo_combo.addItems(self.EMPLEO_DICT.get(escala.currentText(), []))
            empleo_combo.setEnabled(bool(self.EMPLEO_DICT.get(escala.currentText(), [])))
            empleo_combo.blockSignals(False)

    def perform_search(self):
        self._reload_all()
        filters = {}
        for field, widget in self.search_filters.items():
            if isinstance(widget, QCheckBox):
                value = "SI" if widget.isChecked() else ""
                if value: filters[field] = value
            elif isinstance(widget, QComboBox):
                value = widget.currentText()
                if value: filters[field] = value
            elif isinstance(widget, QDateEdit):
                if widget.date() != QDate.currentDate():
                    filters[field] = widget.date().toString("dd/MM/yyyy")
            else:
                value = widget.text().strip()
                if value: filters[field] = value
        results = []
        for emp_data in self.db.get_all_employees().values():
            if all(
                (isinstance(self.search_filters[field], QDateEdit) and emp_data.get(field, "") == value)
                or (isinstance(self.search_filters[field], QCheckBox) and emp_data.get(field, "NO") == value)
                or (isinstance(self.search_filters[field], QComboBox) and emp_data.get(field, "") == value)
                or (value.lower() in str(emp_data.get(field, '')).lower())
                for field, value in filters.items()
            ):
                results.append(emp_data)
        self.search_table.setRowCount(len(results))
        for row, emp_data in enumerate(results):
            color_row = QBrush(QColor(255, 125, 125)) if emp_data.get('baja', 'NO') == 'SI' else \
                        QBrush(QColor(180, 220, 255)) if emp_data.get('mision', 'NO') == 'SI' else None
            for col, field in enumerate(self.db.fields):
                value = str(emp_data.get(field, ''))
                item = QTableWidgetItem(value)
                if color_row: item.setBackground(color_row)
                self.search_table.setItem(row, col, item)
        self.search_table.resizeColumnsToContents()

    def export_search_results_to_excel(self):
        row_count, col_count = self.search_table.rowCount(), self.search_table.columnCount()
        if row_count == 0:
            QMessageBox.information(self, "Sin resultados", "No hay datos para exportar.")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self, "Guardar como Excel", "resultados_busqueda.xlsx", "Archivos Excel (*.xlsx)"
        )
        if not filename: return
        wb = Workbook(); ws = wb.active
        headers = [self.search_table.horizontalHeaderItem(col).text() for col in range(col_count)]
        ws.append(headers)
        for row in range(row_count):
            ws.append([self.search_table.item(row, col).text() if self.search_table.item(row, col) else "" for col in range(col_count)])
        for col in range(1, col_count + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18
        try:
            wb.save(filename)
            QMessageBox.information(self, "Exportación exitosa", f"Resultados exportados a:\n{filename}")
        except Exception as e:
            QMessageBox.critical(self, "Error al exportar", f"No se pudo exportar:\n{e}")

    # ---- Fields Management Tab ----
    def _setup_fields_management_tab(self):
        tab = QWidget()
        self.tab_widget.addTab(tab, self.icons["campos"], "Campos")
        layout = QVBoxLayout(tab)
        edit_group = QGroupBox("Modificar Campos Existentes")
        edit_layout = QVBoxLayout(edit_group)
        self.field_combobox = QComboBox()
        self._update_field_combobox()
        self.new_name_input = QLineEdit()
        update_btn = QPushButton(self.icons["formulario"], "Actualizar Nombre")
        update_btn.setProperty("class", "primary")
        update_btn.clicked.connect(self.update_field_name)
        delete_btn = QPushButton(self.icons["campos"], "Eliminar Campo")
        delete_btn.setProperty("class", "danger")
        delete_btn.clicked.connect(self.delete_field)
        for w in [("Campo a modificar:", self.field_combobox), ("Nuevo nombre:", self.new_name_input), (None, update_btn), (None, delete_btn)]:
            edit_layout.addWidget(QLabel(w[0]) if w[0] else w[1])
        edit_layout.addStretch()
        add_group = QGroupBox("Añadir Nuevo Campo")
        add_layout = QFormLayout(add_group)
        self.new_field_name_input, self.default_value_input = QLineEdit(), QLineEdit()
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
        for w in [edit_group, add_group, self.fields_list_label, self.fields_list]:
            layout.addWidget(w)
        layout.addStretch()

    def _update_field_combobox(self):
        self.db.load_db()
        self.field_combobox.clear()
        self.field_combobox.addItems([f.capitalize() for f in self.db.fields])

    def add_custom_field(self):
        field_name = self.new_field_name_input.text().strip().lower()
        if not field_name or not all(c.isalnum() or c == '_' for c in field_name) or field_name in ['baja', 'mision']:
            QMessageBox.warning(self, "Error", "Nombre inválido")
            return
        default_value = self.default_value_input.text().strip()
        if self.db.add_field(field_name, default_value):
            self.db.save_db()
            self.logger.log(f"Añadido campo nuevo: {field_name}")
            self._reload_all()
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
        if not new_name or not all(c.isalnum() or c == '_' for c in new_name) or new_name in self.db.fields or new_name in ['baja', 'mision']:
            QMessageBox.warning(self, "Error", "Nombre inválido")
            return
        if self.db.update_field_name(old_name, new_name):
            self.db.save_db()
            self.logger.log(f"Renombrado campo: {old_name} -> {new_name}")
            self._reload_all()
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
        if QMessageBox.question(self, 'Confirmar',
            f'¿Está seguro que desea eliminar el campo "{field_name}"?',
            QMessageBox.Yes|QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
            if self.db.delete_field(field_name):
                self.db.save_db()
                self.logger.log(f"Eliminado campo: {field_name}")
                self._reload_all()
                QMessageBox.information(self, "Éxito", f"Campo '{field_name}' eliminado correctamente")
                self.new_name_input.clear()
            else:
                QMessageBox.warning(self, "Error", "No se pudo eliminar el campo")

    # ---- Settings Tab ----
    def _setup_settings_tab(self):
        tab = QWidget()
        self.tab_widget.addTab(tab, self.icons["configuracion"], "Configuración")
        layout = QVBoxLayout(tab)
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
        backup_layout = QHBoxLayout(); backup_layout.addWidget(self.backup_path_line); backup_layout.addWidget(self.backup_path_btn)
        settings_layout.addRow("Ruta copia seguridad:", backup_layout)
        settings_layout.addRow("Guardado automático:", self.auto_save_check)
        settings_layout.addRow("Modo oscuro:", self.dark_mode_check)
        settings_layout.addRow("Tamaño de fuente:", self.font_size_combo)
        settings_layout.addRow("Departamento por defecto:", self.default_dept_input)
        layout.addWidget(settings_group)
        # Changelog
        changelog_group = QGroupBox("Registro de Cambios")
        changelog_layout = QVBoxLayout(changelog_group)
        self.change_log_text = QTextEdit()
        self.change_log_text.setReadOnly(True)
        self.change_log_text.setStyleSheet("font-family:monospace;")
        self.change_log_text.setPlainText(self.logger.read_log())
        changelog_layout.addWidget(self.change_log_text)
        layout.addWidget(changelog_group)
        layout.addStretch()

    # ---- Utility/UI Methods ----
    def update_datetime_label(self):
        now = QDateTime.currentDateTime()
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
        dia_semana = dias[now.date().dayOfWeek()-1]
        dia, mes, anio = now.date().day(), meses[now.date().month()-1], now.date().year()
        hora = now.time().toString("HH:mm:ss")
        self.datetime_label.setText(f"<b>{dia_semana.capitalize()}, {dia} de {mes} de {anio}</b> {hora}")

    def automatic_backup(self):
        backup_path = self.config.settings.get('backup_path', '')
        if not backup_path or not os.path.exists(self.db.db_file): return
        if os.path.isfile(backup_path):
            try: os.remove(backup_path)
            except Exception: pass
        try:
            shutil.copy2(self.db.db_file, backup_path)
            self.last_backup_file = backup_path
        except Exception: pass

    def apply_styles(self):
        style = """
            QMainWindow { background-color: #f8fafc; font-family: 'Segoe UI', 'Arial', 'sans-serif'; font-size: 13px;}
            QLabel, QGroupBox, QTabWidget, QTableWidget, QTabBar::tab { color: #183a6b; }
            QTabWidget::pane { border: 1px solid #183a6b; padding: 5px;}
            QTabBar::tab { padding: 8px 18px; background: #e9f0fa; border: 1px solid #bcd0ee; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px;}
            QTabBar::tab:selected { background: #fff; border-bottom: 1px solid #fff; color: #183a6b;}
            QPushButton { padding: 6px 14px; border-radius: 6px; border: 1px solid #bcd0ee; background: #f3f7fc; color: #183a6b; font-weight: bold;}
            QPushButton.primary { background-color: #183a6b; color: white; }
            QPushButton.danger { background-color: #c82333; color: white; }
            QPushButton.secondary { background-color: #5a6268; color: white; }
            QTableWidget { border: 1px solid #bcd0ee; gridline-color: #e9f0fa; background: #fff; }
            QHeaderView::section { background-color: #dde8fa; padding: 8px; border: none; color: #183a6b; }
            QLineEdit, QComboBox { padding: 5px; border: 1px solid #bcd0ee; border-radius: 4px; min-height: 30px; background: #f9fbfc; color: #183a6b;}
            QGroupBox { border: 1px solid #bcd0ee; border-radius: 6px; margin-top: 10px; padding-top: 15px; background: #f3f7fc;}
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px;}
        """
        self.setStyleSheet(style)

    def update_config(self, key, value):
        old_value = self.config.settings.get(key)
        changed = self.config.update_setting(key, value)
        if changed and old_value != value:
            self.logger.log(f"Cambio de configuración: {key} de '{old_value}' a '{value}'")
            self.apply_settings()
            self._reload_all()
        return changed

    def choose_backup_path(self):
        path, _ = QFileDialog.getSaveFileName(self, "Elegir ruta para la copia de seguridad", "personal_db_backup.dat", "Archivos de datos (*.dat)")
        if path:
            self.config.update_setting('backup_path', path)
            self.backup_path_line.setText(path)
            self.logger.log(f"Cambio de ruta de copia de seguridad: {path}")
            self.automatic_backup()

    def load_employees(self):
        self.db.load_db()
        employees = self.db.get_all_employees()
        self.table.setRowCount(len(employees))
        for row, (emp_id, emp_data) in enumerate(employees.items()):
            color_row = QBrush(QColor(255, 125, 125)) if emp_data.get('baja', 'NO') == 'SI' else \
                        QBrush(QColor(180, 220, 255)) if emp_data.get('mision', 'NO') == 'SI' else None
            for col, field in enumerate(self.db.fields):
                value = str(emp_data.get(field, ''))
                item = QTableWidgetItem(value)
                if color_row: item.setBackground(color_row)
                self.table.setItem(row, col, item)
        self.table.resizeColumnsToContents()

    def new_employee(self):
        self.current_employee_id = None
        self.clear_form()
        self.tab_widget.setCurrentWidget(self.form_tab)
        first_field = next(iter(self.field_inputs.values()))
        if isinstance(first_field, QLineEdit): first_field.setFocus()
        self._sync_escala_empleo()

    def delete_selected_employee(self):
        row = self.table.currentRow()
        if row == -1:
            QMessageBox.warning(self, "Advertencia", "Seleccione un empleado")
            return
        emp_id = list(self.db.get_all_employees().keys())[row]
        if QMessageBox.question(self, 'Confirmar', '¿Está seguro que desea eliminar a este empleado?',
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
            if self.db.delete_employee(emp_id):
                self.logger.log(f"Eliminado empleado: ID {emp_id}")
                QMessageBox.information(self, "Éxito", "Empleado eliminado correctamente")
                self._reload_all()
                if self.current_employee_id == emp_id: self.clear_form()
            else:
                QMessageBox.warning(self, "Error", "No se pudo eliminar el empleado")

    def modify_selected_employee(self):
        row = self.table.currentRow()
        if row == -1:
            QMessageBox.warning(self, "Advertencia", "Seleccione un empleado")
            return
        emp_id = list(self.db.get_all_employees().keys())[row]
        self._setup_employee_form_tab()
        self._load_employee_to_form(emp_id)
        self.tab_widget.setCurrentWidget(self.form_tab)
        self._sync_escala_empleo()

    def _load_employee_to_form(self, emp_id):
        employee = self.db.get_employee(emp_id)
        if not employee: return
        self.current_employee_id = emp_id
        for field, widget in self.field_inputs.items():
            value = employee.get(field, "")
            if isinstance(widget, QCheckBox):
                widget.setChecked(value == "SI")
            elif isinstance(widget, QComboBox):
                idx = widget.findText(value, Qt.MatchExactly)
                widget.setCurrentIndex(idx if idx >= 0 else 0)
                if field == "empleo": self._sync_escala_empleo()
            elif isinstance(widget, QDateEdit):
                date = QDate.fromString(value, "dd/MM/yyyy")
                if date.isValid(): widget.setDate(date)
            else:
                widget.setText(value)
        self.unsaved_changes = False

    def save_employee(self, silent=False):
        employee_data = {}
        for field, widget in self.field_inputs.items():
            if isinstance(widget, QCheckBox):
                employee_data[field] = "SI" if widget.isChecked() else "NO"
            elif isinstance(widget, QComboBox):
                employee_data[field] = widget.currentText()
            elif isinstance(widget, QDateEdit):
                employee_data[field] = widget.date().toString("dd/MM/yyyy")
            else:
                employee_data[field] = widget.text().strip()
        action = None
        if self.current_employee_id is None:
            action = "Alta de empleado: " + " ".join([employee_data.get("nombre", ""), employee_data.get("apellidos", "")])
        else:
            prev = self.db.get_employee(self.current_employee_id)
            if prev:
                cambios = [f"{k}: '{prev.get(k, '')}' -> '{employee_data.get(k, '')}'" for k in self.db.fields if str(prev.get(k, "")) != str(employee_data.get(k, ""))]
                if cambios:
                    action = f"Modificación de empleado {self.current_employee_id}: " + "; ".join(cambios)
        ok = False
        if self.current_employee_id is None:
            emp_id = self.db.add_employee(employee_data)
            if emp_id:
                ok = True
                if not silent: QMessageBox.information(self, "Éxito", f"Empleado añadido correctamente")
                self.current_employee_id = emp_id
                self.unsaved_changes = False
                self.clear_form()
                self._reload_all()
        else:
            if self.db.update_employee(self.current_employee_id, employee_data):
                ok = True
                if not silent: QMessageBox.information(self, "Éxito", "Empleado actualizado correctamente")
                self.unsaved_changes = False
                self._reload_all()
        if ok and action: self.logger.log(action)
        return ok

    def closeEvent(self, event):
        self.db.save_db()
        self.config.save_config()
        event.accept()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False): os.chdir(sys._MEIPASS)
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
