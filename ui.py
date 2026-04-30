import json
import os
import shutil
import sys
from datetime import datetime, timedelta
from copy import deepcopy
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTextEdit,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from calendar_math import count_workdays, format_workday_label
from data_loader import (
    load_holidays,
    load_progress,
    load_rooms,
    load_settings,
    load_staffing,
    validate_workbook,
)
from exporter import (
    create_input_template,
    export_result_workbook,
    export_work_instructions_workbook,
    ensure_cleaning_methods_sheet,
    format_cleaning_method_preview,
    load_cleaning_methods,
)
from models import StaffingDay
from services.scenario_service import ScenarioInput, run_scenario
from timeblock_generator import format_time_blocks_for_text

CONFIG_FILE = Path(__file__).resolve().parent / "config.json"


def load_app_config() -> dict:
    if not CONFIG_FILE.exists():
        return {}

    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_app_config(data: dict) -> None:
    CONFIG_FILE.write_text(
        json.dumps(data, indent=2),
        encoding="utf-8",
    )


def get_last_workbook_file() -> str:
    config = load_app_config()
    value = config.get("last_workbook_file", "")
    return str(value).strip()


def set_last_workbook_file(file_path: str) -> None:
    config = load_app_config()
    config["last_workbook_file"] = str(file_path).strip()
    save_app_config(config)

def fmt_hours(value: float) -> str:
    return f"{value:.2f}"


def yes_no(value: bool) -> str:
    return "Yes" if value else "No"


def parse_int_field(value: str, default: int = 0, field_name: str = "number") -> int:
    raw = str(value or "").strip()
    if raw == "":
        return default
    try:
        return int(float(raw))
    except ValueError as exc:
        raise ValueError(f"{field_name} must be a whole number. Got: {raw}") from exc


def parse_float_field(value: str, default: float = 0.0, field_name: str = "number") -> float:
    raw = str(value or "").strip()
    if raw == "":
        return default
    try:
        return float(raw)
    except ValueError as exc:
        raise ValueError(f"{field_name} must be a number. Got: {raw}") from exc


def calculate_workdays(
    start_date_str: str,
    end_date_str: str,
    include_weekends: bool,
    paid_holidays: int = 0,
    holidays: set | None = None,
) -> int:
    holiday_dates = holidays or set()
    workdays = count_workdays(
        start_date_str,
        end_date_str,
        work_on_weekends=include_weekends,
        holidays=holiday_dates,
    )

    # Backward compatibility for workbooks that do not have a Holidays sheet yet.
    # Once holiday dates are present, exact dates are the source of truth.
    if not holiday_dates and paid_holidays > 0:
        workdays -= paid_holidays

    return max(workdays, 0)


def get_dark_stylesheet() -> str:
    return """
    QWidget {
        font-size: 13px;
        color: #ececf1;
    }

    QMainWindow {
        background: #0f0f0f;
    }

    QLabel {
        color: #ececf1;
    }

    QTabWidget::pane {
        border: none;
        background: #0f0f0f;
    }

    QTabBar::tab {
        background: #1b1b1b;
        border: 1px solid #2f2f2f;
        padding: 8px 14px;
        margin-right: 4px;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        font-weight: 600;
        color: #c8c8c8;
    }

    QTabBar::tab:selected {
        background: #2a2a2a;
        color: #ffffff;
    }

    #sidePanel {
        background: #171717;
        border: 1px solid #2f2f2f;
        border-radius: 12px;
    }

    #panelTitle {
        font-size: 20px;
        font-weight: 700;
        color: #ffffff;
    }

    #pageTitle {
        font-size: 24px;
        font-weight: 700;
        color: #ffffff;
    }

    #statusChip {
        background: #242424;
        color: #ececf1;
        border: 1px solid #3a3a3a;
        border-radius: 10px;
        padding: 6px 10px;
        font-weight: 600;
    }

    #summaryCard {
        background: #171717;
        border: 1px solid #2f2f2f;
        border-radius: 12px;
    }

    #summaryCardTitle {
        color: #a3a3a3;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
    }

    #summaryCardValue {
        color: #ffffff;
        font-size: 20px;
        font-weight: 700;
    }

    QPushButton#primaryButton {
        background: #E6B800;
        color: #000000;
        font-weight: 800;
        padding: 8px 14px;
        border-radius: 8px;
        border: 1px solid #E6B800;
    }

    QPushButton#primaryButton:hover {
        background: #F2C94C;
        border: 1px solid #F2C94C;
        color: #000000;
    }

    QPushButton#primaryButton:pressed {
        background: #C99A00;
        border: 1px solid #C99A00;
        color: #000000;
    }

    QGroupBox {
        background: #171717;
        color: #ffffff;
        border: 1px solid #2f2f2f;
        border-radius: 12px;
        margin-top: 10px;
        font-weight: 700;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px 0 6px;
        color: #d4d4d4;
    }

    QPlainTextEdit, QTextEdit, QLineEdit, QTableWidget, QComboBox {
        background: #101010;
        color: #ececf1;
        border: 1px solid #333333;
        border-radius: 8px;
        selection-background-color: #444444;
        selection-color: #ffffff;
    }

    QComboBox {
        padding: 4px 8px;
    }

    QComboBox QAbstractItemView {
        background: #101010;
        color: #ececf1;
        border: 1px solid #333333;
        selection-background-color: #444444;
    }

    QHeaderView::section {
        background: #222222;
        color: #ffffff;
        padding: 6px;
        border: none;
        font-weight: 600;
    }

    QTableWidget::item:selected {
        background: #2f2f2f;
        color: #ffffff;
    }

    QPushButton {
        min-height: 28px;
        padding: 8px 12px;
        border: 1px solid #3a3a3a;
        border-radius: 8px;
        background: #242424;
        color: #ececf1;
        font-weight: 600;
    }

    QPushButton:hover {
        background: #303030;
    }

    QPushButton:pressed {
        background: #1c1c1c;
    }

    QCheckBox {
        color: #ececf1;
    }

    QScrollBar:vertical, QScrollBar:horizontal {
        background: #101010;
        border: none;
        margin: 0px;
    }

    QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
        background: #3a3a3a;
        border-radius: 4px;
        min-height: 24px;
        min-width: 24px;
    }

    QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
        background: #4a4a4a;
    }

    QMessageBox {
        background-color: #171717;
    }

    QMessageBox QLabel {
        color: #ececf1;
        min-width: 320px;
    }

    QMessageBox QPushButton {
        min-width: 80px;
        padding: 6px 12px;
        background: #242424;
        color: #ececf1;
        border: 1px solid #3a3a3a;
        border-radius: 8px;
    }

    #mutedLabel {
        color: #a3a3a3;
    }
    """


def get_light_stylesheet() -> str:
    return """
    QWidget {
        font-size: 13px;
        color: #202123;
    }

    QMainWindow {
        background: #f7f7f8;
    }

    QLabel {
        color: #202123;
    }

    QTabWidget::pane {
        border: none;
        background: #f7f7f8;
    }

    QTabBar::tab {
        background: #ececf1;
        border: 1px solid #d9d9e3;
        padding: 8px 14px;
        margin-right: 4px;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        font-weight: 600;
        color: #353740;
    }

    QTabBar::tab:selected {
        background: #ffffff;
        color: #202123;
    }

    #sidePanel {
        background: #ffffff;
        border: 1px solid #d9d9e3;
        border-radius: 12px;
    }

    #panelTitle {
        font-size: 20px;
        font-weight: 700;
        color: #202123;
    }

    #pageTitle {
        font-size: 24px;
        font-weight: 700;
        color: #202123;
    }

    #statusChip {
        background: #ececf1;
        color: #202123;
        border: 1px solid #d9d9e3;
        border-radius: 10px;
        padding: 6px 10px;
        font-weight: 600;
    }

    #summaryCard {
        background: #ffffff;
        border: 1px solid #d9d9e3;
        border-radius: 12px;
    }

    #summaryCardTitle {
        color: #6e6e80;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
    }

    #summaryCardValue {
        color: #202123;
        font-size: 20px;
        font-weight: 700;
    }

    QPushButton#primaryButton {
        background: #E6B800;
        color: #000000;
        font-weight: 800;
        padding: 8px 14px;
        border-radius: 8px;
        border: 1px solid #E6B800;
    }

    QPushButton#primaryButton:hover {
        background: #F2C94C;
        border: 1px solid #F2C94C;
        color: #000000;
    }

    QPushButton#primaryButton:pressed {
        background: #C99A00;
        border: 1px solid #C99A00;
        color: #000000;
    }

    QGroupBox {
        background: #ffffff;
        color: #202123;
        border: 1px solid #d9d9e3;
        border-radius: 12px;
        margin-top: 10px;
        font-weight: 700;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px 0 6px;
        color: #353740;
    }

    QPlainTextEdit, QTextEdit, QLineEdit, QTableWidget, QComboBox {
        background: #ffffff;
        color: #202123;
        border: 1px solid #d9d9e3;
        border-radius: 8px;
        selection-background-color: #ececf1;
        selection-color: #202123;
    }

    QComboBox {
        padding: 4px 8px;
    }

    QComboBox QAbstractItemView {
        background: #ffffff;
        color: #202123;
        border: 1px solid #d9d9e3;
        selection-background-color: #ececf1;
    }

    QHeaderView::section {
        background: #ececf1;
        color: #202123;
        padding: 6px;
        border: none;
        font-weight: 600;
    }

    QTableWidget::item:selected {
        background: #ececf1;
        color: #202123;
    }

    QPushButton {
        min-height: 28px;
        padding: 8px 12px;
        border: 1px solid #d9d9e3;
        border-radius: 8px;
        background: #ffffff;
        color: #202123;
        font-weight: 600;
    }

    QPushButton:hover {
        background: #f1f1f3;
    }

    QPushButton:pressed {
        background: #e5e5ea;
    }

    QCheckBox {
        color: #202123;
    }

    QScrollBar:vertical, QScrollBar:horizontal {
        background: #f7f7f8;
        border: none;
        margin: 0px;
    }

    QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
        background: #c8c8d0;
        border-radius: 4px;
        min-height: 24px;
        min-width: 24px;
    }

    QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
        background: #acacb8;
    }

    QMessageBox {
        background-color: #ffffff;
    }

    QMessageBox QLabel {
        color: #202123;
        min-width: 320px;
    }

    QMessageBox QPushButton {
        min-width: 80px;
        padding: 6px 12px;
        background: #ffffff;
        color: #202123;
        border: 1px solid #d9d9e3;
        border-radius: 8px;
    }

    #mutedLabel {
        color: #6e6e80;
    }
    """


class SummaryCard(QFrame):
    def __init__(self, title: str, value: str = "", parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setObjectName("summaryCard")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(4)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("summaryCardTitle")

        self.value_label = QLabel(value)
        self.value_label.setObjectName("summaryCardValue")
        self.value_label.setWordWrap(True)
        self.setMinimumHeight(140)
        self.setMaximumHeight(180)

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)

    def set_value(self, value: str) -> None:
        self.value_label.setText(value)


class SchedulerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Summer Scheduler")
        self.resize(1600, 950)

        self.result = None
        self.settings = None
        self.rooms = []
        self.schools = []
        self.staffing_days = []
        self.progress_entries = []
        self.holidays = set()

        self.current_theme = "dark"
        self.staffing_overrides = []
        self.summary_reveal_timer = QTimer(self)
        self.summary_reveal_timer.timeout.connect(self._reveal_next_summary_step)

        self.summary_reveal_steps = []
        self.summary_reveal_index = 0
        

        self._build_ui()
        self._apply_theme()
        self._load_startup_workbook_file()
        self._load_defaults_into_form()

    def _load_startup_workbook_file(self) -> None:
        last_file = get_last_workbook_file()
        if not last_file:
            return

        file_path = Path(last_file)
        if not file_path.exists() or not file_path.is_file():
            return

        resolved = str(file_path.resolve())
        self.workbook_path_edit.setText(resolved)
        self.template_path_edit.setText(resolved)

    def _build_ui(self) -> None:
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.run_tab = self._build_run_tab()
        self.schedule_tab = self._build_schedule_tab()
        self.export_tab = self._build_export_tab()
        self.data_tab = self._build_data_tab()

        self.tabs.addTab(self.run_tab, "Run")
        self.tabs.addTab(self.schedule_tab, "Schedule")
        self.tabs.addTab(self.export_tab, "Export")
        self.tabs.addTab(self.data_tab, "Data")

    def _build_run_tab(self) -> QWidget:
        root = QWidget()
        outer_layout = QHBoxLayout(root)
        outer_layout.setContentsMargins(12, 12, 12, 12)
        outer_layout.setSpacing(12)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        outer_layout.addWidget(splitter)

        # --- LEFT PANEL (SCROLLABLE) ---
        left_panel = self._build_left_panel()
        left_panel.setMinimumWidth(490)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setMinimumWidth(505)
        left_scroll.setWidget(left_panel)

        # --- RIGHT PANEL ---
        right_panel = self._build_run_right_panel()
        right_panel.setMinimumWidth(900)

        splitter.addWidget(left_scroll)
        splitter.addWidget(right_panel)

        # Give the right panel more stretch, but keep left usable
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([520, 1100])

        return root

    def _build_left_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("sidePanel")

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        title_row = QHBoxLayout()
        title = QLabel("Scenario Controls")
        title.setObjectName("panelTitle")
        title_row.addWidget(title)
        title_row.addStretch(1)

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Dark", "Light"])
        self.theme_combo.setCurrentText("Dark")
        self.theme_combo.setMinimumWidth(96)
        self.theme_combo.setMinimumHeight(34)
        self.theme_combo.currentTextChanged.connect(self._on_theme_changed)
        title_row.addWidget(QLabel("Theme"))
        title_row.addWidget(self.theme_combo)

        layout.addLayout(title_row)

        workbook_group = QGroupBox("Workbook File")
        workbook_layout = QVBoxLayout(workbook_group)

        workbook_row = QHBoxLayout()
        self.workbook_path_edit = QLineEdit("")
        self.workbook_path_edit.setMinimumHeight(38)
        self.browse_workbook_button = QPushButton("Browse")
        self.browse_workbook_button.setMinimumHeight(38)
        self.browse_workbook_button.clicked.connect(self._browse_workbook)
        workbook_row.addWidget(self.workbook_path_edit)
        workbook_row.addWidget(self.browse_workbook_button)

        workbook_layout.addLayout(workbook_row)

        workbook_actions_row_1 = QHBoxLayout()
        workbook_actions_row_1.setSpacing(8)

        self.make_template_button = QPushButton("Create Blank Template")
        self.make_template_button.clicked.connect(self._create_template_from_ui)
        self.make_template_button.setMinimumHeight(38)

        self.save_copy_button = QPushButton("Save Working Copy")
        self.save_copy_button.clicked.connect(self._save_working_copy)
        self.save_copy_button.setMinimumHeight(38)

        workbook_actions_row_1.addWidget(self.make_template_button)
        workbook_actions_row_1.addWidget(self.save_copy_button)

        workbook_actions_row_2 = QHBoxLayout()
        workbook_actions_row_2.setSpacing(8)

        self.reload_button = QPushButton("Reload Workbook")
        self.reload_button.clicked.connect(self._load_defaults_into_form)
        self.reload_button.setMinimumHeight(38)
        workbook_actions_row_2.addWidget(self.reload_button)

        workbook_layout.addLayout(workbook_actions_row_1)
        workbook_layout.addLayout(workbook_actions_row_2)
        layout.addWidget(workbook_group)

        button_row = QHBoxLayout()

        self.run_button = QPushButton("Run Scheduler")
        self.run_button.setObjectName("primaryButton")
        self.run_button.clicked.connect(self.run_scheduler_from_ui)

        button_row.addWidget(self.run_button)
        layout.addLayout(button_row)

        run_group = QGroupBox("Run Settings")
        run_form = QFormLayout(run_group)

        self.schedule_name_edit = QLineEdit()
        self.schedule_start_date_edit = QLineEdit()
        self.schedule_start_date_edit.setPlaceholderText("YYYY-MM-DD")
        self.current_day_edit = QLineEdit()
        self.target_end_day_edit = QLineEdit()
        self.target_end_date_edit = QLineEdit()
        self.target_end_date_edit.setPlaceholderText("YYYY-MM-DD")

        self.paid_holidays_edit = QLineEdit("0")
        self.work_on_weekends_check = QCheckBox("Count weekends as workdays")

        self.include_deep_clean_check = QCheckBox("Include Deep Clean")
        self.include_strip_check = QCheckBox("Include Strip")
        self.include_wax_check = QCheckBox("Include Wax")
        self.include_carpet_check = QCheckBox("Enable Carpet Cleaning Globally")
        self.include_exterior_check = QCheckBox("Enable Exterior Cleaning Globally")

        self.include_carpet_check.toggled.connect(self._sync_carpet_toggle_state)

        self.general_can_do_carpet_check = QCheckBox(
            "Allow general crew to do carpet work (override)"
        )
        self.general_can_do_carpet_check.setChecked(True)
        self.general_can_do_carpet_check.setEnabled(True)
        self.general_can_do_carpet_check.setToolTip(
            "Runtime override. This changes the current run without editing the workbook. Only matters when carpet cleaning is globally enabled."
        )

        run_form.addRow("Schedule Name", self.schedule_name_edit)
        run_form.addRow("Start Date", self.schedule_start_date_edit)
        run_form.addRow("Current Day", self.current_day_edit)
        run_form.addRow("Target End Date", self.target_end_date_edit)
        run_form.addRow("Legacy Paid Holidays", self.paid_holidays_edit)
        run_form.addRow("", self.work_on_weekends_check)
        run_form.addRow("", self.include_deep_clean_check)
        run_form.addRow("", self.include_strip_check)
        run_form.addRow("", self.include_wax_check)
        run_form.addRow("", self.include_carpet_check)
        run_form.addRow("", self.include_exterior_check)
        run_form.addRow("", self.general_can_do_carpet_check)

        layout.addWidget(run_group)

        staffing_group = QGroupBox("Staffing Overrides")
        staffing_layout = QVBoxLayout(staffing_group)

        staffing_top = QHBoxLayout()
        self.override_mode_combo = QComboBox()
        self.override_mode_combo.addItems(["Global", "Weekly", "Daily"])
        self.override_mode_combo.currentTextChanged.connect(self._refresh_override_mode_label)

        self.override_anchor_label = QLabel("Week #")
        self.override_anchor_edit = QLineEdit("1")
        self.override_length_label = QLabel("Span")
        self.override_length_edit = QLineEdit("1")

        staffing_top.addWidget(QLabel("Mode"))
        staffing_top.addWidget(self.override_mode_combo)
        staffing_top.addWidget(self.override_anchor_label)
        staffing_top.addWidget(self.override_anchor_edit)
        staffing_top.addWidget(self.override_length_label)
        staffing_top.addWidget(self.override_length_edit)

        staffing_layout.addLayout(staffing_top)

        staffing_form = QFormLayout()
        self.override_cleaning_staff_edit = QLineEdit("4")
        self.override_carpet_staff_edit = QLineEdit("0")
        self.override_outside_help_edit = QLineEdit("0")
        self.override_absences_edit = QLineEdit("0")

        staffing_form.addRow("Cleaning Staff", self.override_cleaning_staff_edit)
        staffing_form.addRow("Carpet Staff", self.override_carpet_staff_edit)
        staffing_form.addRow("Outside Help", self.override_outside_help_edit)
        staffing_form.addRow("Absences", self.override_absences_edit)

        staffing_layout.addLayout(staffing_form)

        staffing_button_row = QHBoxLayout()
        self.apply_override_button = QPushButton("Apply Override")
        self.apply_override_button.clicked.connect(self._apply_staffing_override)

        self.clear_overrides_button = QPushButton("Clear Overrides")
        self.clear_overrides_button.clicked.connect(self._clear_staffing_overrides)

        staffing_button_row.addWidget(self.apply_override_button)
        staffing_button_row.addWidget(self.clear_overrides_button)
        staffing_layout.addLayout(staffing_button_row)

        self.override_preview_text = QPlainTextEdit()
        self.override_preview_text.setReadOnly(True)
        self.override_preview_text.setMaximumHeight(110)
        self.override_preview_text.setPlainText("No staffing overrides yet.")
        staffing_layout.addWidget(self.override_preview_text)

        layout.addWidget(staffing_group)

        time_group = QGroupBox("Daily Time Model")
        time_form = QFormLayout(time_group)

        self.shift_hours_edit = QLineEdit()
        self.lunch_hours_edit = QLineEdit()
        self.break_hours_edit = QLineEdit()
        self.setup_hours_edit = QLineEdit()
        self.cleanup_hours_edit = QLineEdit()
        self.productive_hours_edit = QLineEdit()

        time_form.addRow("Shift Hours", self.shift_hours_edit)
        time_form.addRow("Lunch Hours", self.lunch_hours_edit)
        time_form.addRow("Break Hours", self.break_hours_edit)
        time_form.addRow("Setup Hours", self.setup_hours_edit)
        time_form.addRow("Cleanup Hours", self.cleanup_hours_edit)
        time_form.addRow("Productive Hours", self.productive_hours_edit)

        layout.addWidget(time_group)

        info_group = QGroupBox("What This Tab Does")
        info_layout = QVBoxLayout(info_group)

        notes = QLabel(
            "Run tab = scenario controls + fast answer.\n\n"
            "The app now reads from one workbook file containing district data,\n"
            "planning assumptions and current run input.\n\n"
            "Room Scope task toggles are exception-based: blank means ON/included, FALSE skips.\n"
            "Task work only runs when the global setting and room-level toggle are both on.\n\n"
            "Use Global, Weekly or Daily mode, apply an override, then run."
        )
        notes.setWordWrap(True)
        notes.setObjectName("mutedLabel")

        info_layout.addWidget(notes)
        layout.addWidget(info_group)
        layout.addStretch(1)

        self._refresh_override_mode_label()
        panel.adjustSize()
        return panel

    def _build_run_right_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        title_row = QHBoxLayout()

        page_title = QLabel("Run Overview")
        page_title.setObjectName("pageTitle")

        self.status_chip = QLabel("Ready")
        self.status_chip.setObjectName("statusChip")
        self.status_chip.setFixedHeight(44)
        self.status_chip.setFixedWidth(110)
        self.status_chip.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_row.addWidget(page_title)
        title_row.addStretch(1)
        title_row.addWidget(self.status_chip)

        layout.addLayout(title_row)

        cards_row = QHBoxLayout()
        cards_row.setSpacing(12)

        self.finish_day_card = SummaryCard("Projected Finish", "-")
        self.deadline_card = SummaryCard("Deadline Met", "-")
        self.backlog_card = SummaryCard("Remaining Backlog", "-")
        self.recommendation_card = SummaryCard("Recommendation Status", "-")

        cards_row.addWidget(self.finish_day_card)
        cards_row.addWidget(self.deadline_card)
        cards_row.addWidget(self.backlog_card)
        cards_row.addWidget(self.recommendation_card)

        layout.addLayout(cards_row)

        summary_splitter = QSplitter(Qt.Orientation.Horizontal)
        summary_splitter.setChildrenCollapsible(False)

        cleaning_group = QGroupBox("Cleaning Summary")
        cleaning_layout = QVBoxLayout(cleaning_group)

        self.cleaning_summary_text = QTextEdit()
        self.cleaning_summary_text.setReadOnly(True)
        self.cleaning_summary_text.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.cleaning_summary_text.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.cleaning_summary_text.setMinimumHeight(220)

        cleaning_layout.addWidget(self.cleaning_summary_text)

        carpet_group = QGroupBox("Carpet Summary")
        carpet_layout = QVBoxLayout(carpet_group)

        self.carpet_summary_text = QTextEdit()
        self.carpet_summary_text.setReadOnly(True)
        self.carpet_summary_text.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.carpet_summary_text.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.carpet_summary_text.setMinimumHeight(220)

        carpet_layout.addWidget(self.carpet_summary_text)

        summary_splitter.addWidget(cleaning_group)
        summary_splitter.addWidget(carpet_group)
        summary_splitter.setSizes([700, 500])
        summary_splitter.setMinimumHeight(260)
        summary_splitter.setMaximumHeight(360)

        layout.addWidget(summary_splitter, 1)

        return panel

    def _build_schedule_tab(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        title = QLabel("Schedule Breakdown")
        title.setObjectName("pageTitle")
        layout.addWidget(title)

        detail_splitter = QSplitter(Qt.Orientation.Vertical)
        detail_splitter.setChildrenCollapsible(False)

        # Top: full-width day-by-day schedule table.
        top_half = QWidget()
        top_layout = QVBoxLayout(top_half)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(12)

        days_group = QGroupBox("Day-by-Day Schedule")
        days_layout = QVBoxLayout(days_group)

        self.days_table = QTableWidget(0, 10)
        self.days_table.setHorizontalHeaderLabels(
            [
                "Day",
                "Date",
                "Active School",
                "General Staff",
                "Carpet Staff",
                "Total Capacity",
                "Used",
                "Unused",
                "Work Items",
                "Day Note",
            ]
        )
        self.days_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.days_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.days_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.days_table.verticalHeader().setVisible(False)
        self.days_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.days_table.setWordWrap(False)
        self.days_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.days_table.horizontalHeader().setSectionResizeMode(
            9, QHeaderView.ResizeMode.Interactive
        )
        self.days_table.itemSelectionChanged.connect(self._populate_day_detail)

        days_layout.addWidget(self.days_table)
        top_layout.addWidget(days_group)

        # Bottom: left side is the day outlook; right side stacks work log and details.
        bottom_splitter = QSplitter(Qt.Orientation.Horizontal)
        bottom_splitter.setChildrenCollapsible(False)

        outlook_group = QGroupBox("Day Outlook")
        outlook_layout = QVBoxLayout(outlook_group)

        self.day_outlook_text = QPlainTextEdit()
        self.day_outlook_text.setReadOnly(True)
        self.day_outlook_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.day_outlook_text.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.day_outlook_text.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        outlook_layout.addWidget(self.day_outlook_text)

        right_stack = QSplitter(Qt.Orientation.Vertical)
        right_stack.setChildrenCollapsible(False)

        worklog_group = QGroupBox("Selected Day Work Log")
        worklog_layout = QVBoxLayout(worklog_group)

        self.worklog_table = QTableWidget(0, 8)
        self.worklog_table.setHorizontalHeaderLabels(
            [
                "Crew",
                "School",
                "Building",
                "Zone",
                "Room",
                "Phase",
                "Hours",
                "Note",
            ]
        )
        self.worklog_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.worklog_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.worklog_table.verticalHeader().setVisible(False)
        self.worklog_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.worklog_table.setWordWrap(False)
        self.worklog_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.worklog_table.horizontalHeader().setSectionResizeMode(
            7, QHeaderView.ResizeMode.Interactive
        )

        worklog_layout.addWidget(self.worklog_table)
        right_stack.addWidget(worklog_group)

        detail_group = QGroupBox("Selected Day Details")
        detail_layout = QVBoxLayout(detail_group)

        self.day_detail_text = QPlainTextEdit()
        self.day_detail_text.setReadOnly(True)
        self.day_detail_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.day_detail_text.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.day_detail_text.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        detail_layout.addWidget(self.day_detail_text)
        right_stack.addWidget(detail_group)
        right_stack.setSizes([280, 160])

        bottom_splitter.addWidget(outlook_group)
        bottom_splitter.addWidget(right_stack)
        bottom_splitter.setSizes([900, 600])

        detail_splitter.addWidget(top_half)
        detail_splitter.addWidget(bottom_splitter)
        detail_splitter.setSizes([380, 500])

        layout.addWidget(detail_splitter)
        return panel

    def _build_export_tab(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        title = QLabel("Export")
        title.setObjectName("pageTitle")
        layout.addWidget(title)

        export_splitter = QSplitter(Qt.Orientation.Horizontal)
        export_splitter.setChildrenCollapsible(False)
        layout.addWidget(export_splitter, 1)

        # LEFT: export controls
        controls_panel = QWidget()
        controls_layout = QVBoxLayout(controls_panel)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(12)

        schedule_group = QGroupBox("Schedule Export")
        schedule_layout = QVBoxLayout(schedule_group)

        self.export_path_edit = QLineEdit("output")
        self.export_path_edit.setVisible(False)

        self.export_button = QPushButton("Export Result Workbook")
        self.export_button.clicked.connect(self._export_result_workbook)
        self.export_button.setMinimumHeight(40)
        schedule_layout.addWidget(self.export_button)

        self.export_status_text = QPlainTextEdit()
        self.export_status_text.setReadOnly(True)
        self.export_status_text.setMaximumHeight(170)
        self.export_status_text.setPlainText(
            "No export yet.\n\nRun the scheduler first, then export the workbook."
        )
        schedule_layout.addWidget(self.export_status_text)

        controls_layout.addWidget(schedule_group)

        instructions_group = QGroupBox("Work Instructions")
        instructions_layout = QVBoxLayout(instructions_group)

        self.instruction_method_combo = QComboBox()
        self.instruction_method_combo.currentTextChanged.connect(self._select_instruction_method)
        instructions_layout.addWidget(QLabel("Cleaning Method"))
        instructions_layout.addWidget(self.instruction_method_combo)

        instruction_button_row_1 = QHBoxLayout()
        self.refresh_instructions_button = QPushButton("Refresh Preview")
        self.refresh_instructions_button.clicked.connect(self._load_instruction_methods_from_workbook)

        self.edit_instructions_button = QPushButton("Edit in Workbook")
        self.edit_instructions_button.clicked.connect(self._open_workbook_for_instruction_edit)

        self.add_methods_sheet_button = QPushButton("Add Starter Sheet")
        self.add_methods_sheet_button.clicked.connect(self._add_starter_cleaning_methods_sheet)

        instruction_button_row_1.addWidget(self.refresh_instructions_button)
        instruction_button_row_1.addWidget(self.edit_instructions_button)
        instruction_button_row_1.addWidget(self.add_methods_sheet_button)
        instructions_layout.addLayout(instruction_button_row_1)

        instruction_button_row_2 = QHBoxLayout()
        self.export_selected_instructions_button = QPushButton("Export Selected")
        self.export_selected_instructions_button.clicked.connect(self._export_selected_work_instructions)

        self.export_all_instructions_button = QPushButton("Export All")
        self.export_all_instructions_button.clicked.connect(self._export_all_work_instructions)

        instruction_button_row_2.addWidget(self.export_selected_instructions_button)
        instruction_button_row_2.addWidget(self.export_all_instructions_button)
        instructions_layout.addLayout(instruction_button_row_2)

        self.instruction_status_text = QPlainTextEdit()
        self.instruction_status_text.setReadOnly(True)
        self.instruction_status_text.setMaximumHeight(170)
        self.instruction_status_text.setPlainText(
            "Use the Cleaning Methods sheet in the workbook to manage printable step sheets.\n\n"
            "Refresh Preview after editing the workbook."
        )
        instructions_layout.addWidget(self.instruction_status_text)

        controls_layout.addWidget(instructions_group)
        controls_layout.addStretch(1)

        # RIGHT: preview
        preview_group = QGroupBox("Work Instructions Preview")
        preview_layout = QVBoxLayout(preview_group)

        self.instructions_preview_tabs = QTabWidget()
        preview_layout.addWidget(self.instructions_preview_tabs)

        export_splitter.addWidget(controls_panel)
        export_splitter.addWidget(preview_group)
        export_splitter.setSizes([420, 1100])
        export_splitter.setStretchFactor(0, 0)
        export_splitter.setStretchFactor(1, 1)

        self.cleaning_methods = {}
        self._load_instruction_methods_from_workbook(show_errors=False)

        return panel


    def _load_instruction_methods_from_workbook(self, show_errors: bool = True) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
            self.cleaning_methods = load_cleaning_methods(workbook_path)
        except Exception as exc:
            self.cleaning_methods = {}
            if show_errors:
                self._show_error("Instructions Error", str(exc))

        self.instruction_method_combo.blockSignals(True)
        self.instruction_method_combo.clear()
        self.instructions_preview_tabs.clear()

        if not self.cleaning_methods:
            self.instruction_method_combo.addItem("No Cleaning Methods found")
            empty_preview = QPlainTextEdit()
            empty_preview.setReadOnly(True)
            empty_preview.setPlainText(
                "No Cleaning Methods sheet was found, or it has no printable rows.\n\n"
                "Create a new blank template or add a Cleaning Methods sheet to your current workbook."
            )
            self.instructions_preview_tabs.addTab(empty_preview, "Preview")
            self.instruction_status_text.setPlainText(
                "No Cleaning Methods found.\n\n"
                "The new template includes a Cleaning Methods sheet with starter step sheets."
            )
            self.instruction_method_combo.blockSignals(False)
            return

        for method_name, steps in self.cleaning_methods.items():
            self.instruction_method_combo.addItem(method_name)

            preview = QPlainTextEdit()
            preview.setReadOnly(True)
            preview.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
            preview.setPlainText(format_cleaning_method_preview(method_name, steps))
            self.instructions_preview_tabs.addTab(preview, method_name[:28])

        self.instruction_method_combo.blockSignals(False)
        self.instruction_status_text.setPlainText(
            f"Loaded {len(self.cleaning_methods)} cleaning method(s).\n\n"
            "Select a method to preview it. Use Edit in Workbook to make changes, then Refresh Preview."
        )

    def _select_instruction_method(self, method_name: str) -> None:
        if not hasattr(self, "instructions_preview_tabs"):
            return
        if not hasattr(self, "cleaning_methods") or method_name not in self.cleaning_methods:
            return

        index = list(self.cleaning_methods.keys()).index(method_name)
        if index >= 0:
            self.instructions_preview_tabs.setCurrentIndex(index)

    def _add_starter_cleaning_methods_sheet(self) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
            ensure_cleaning_methods_sheet(workbook_path)
            self._load_instruction_methods_from_workbook(show_errors=False)
            self.instruction_status_text.setPlainText(
                "Starter Cleaning Methods sheet added or already present.\n\n"
                "Use Edit in Workbook to customize the steps."
            )
        except Exception as exc:
            self._show_error("Cleaning Methods Error", str(exc))

    def _open_workbook_for_instruction_edit(self) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
            os.startfile(workbook_path)
            self.instruction_status_text.setPlainText(
                "Workbook opened for editing.\n\n"
                "Edit the Cleaning Methods sheet, save the workbook, then click Refresh Preview."
            )
        except Exception as exc:
            self._show_error("Open Workbook Error", str(exc))

    def _export_selected_work_instructions(self) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
            method_name = self.instruction_method_combo.currentText().strip()

            if not method_name or method_name not in getattr(self, "cleaning_methods", {}):
                raise ValueError("Select a valid cleaning method first.")

            suggested = str(Path.cwd() / f"{method_name} Work Instructions.xlsx")
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Selected Work Instructions",
                suggested,
                "Excel Workbooks (*.xlsx)",
            )
            if not file_path:
                return

            exported_path = export_work_instructions_workbook(
                workbook_path=workbook_path,
                file_path=file_path,
                selected_method=method_name,
            )
            self.instruction_status_text.setPlainText(f"Export complete:\n{exported_path}")
            self._show_info("Instructions Exported", exported_path)
        except Exception as exc:
            self.instruction_status_text.setPlainText(str(exc))
            self._show_error("Instructions Export Error", str(exc))

    def _export_all_work_instructions(self) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
            suggested = str(Path.cwd() / "Work Instructions.xlsx")
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save All Work Instructions",
                suggested,
                "Excel Workbooks (*.xlsx)",
            )
            if not file_path:
                return

            exported_path = export_work_instructions_workbook(
                workbook_path=workbook_path,
                file_path=file_path,
                selected_method=None,
            )
            self.instruction_status_text.setPlainText(f"Export complete:\n{exported_path}")
            self._show_info("Instructions Exported", exported_path)
        except Exception as exc:
            self.instruction_status_text.setPlainText(str(exc))
            self._show_error("Instructions Export Error", str(exc))


    def _build_data_tab(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        title = QLabel("Data / Workbook")
        title.setObjectName("pageTitle")
        layout.addWidget(title)

        template_group = QGroupBox("Template Tools")
        template_layout = QVBoxLayout(template_group)

        template_path_row = QHBoxLayout()
        self.template_path_edit = QLineEdit("")
        template_path_row.addWidget(self.template_path_edit)

        template_button_row = QHBoxLayout()
        template_button_row.setSpacing(8)

        self.template_button = QPushButton("Create Blank Template")
        self.template_button.clicked.connect(self._create_template_from_ui)
        self.template_button.setMinimumHeight(38)

        self.data_save_copy_button = QPushButton("Save Working Copy")
        self.data_save_copy_button.clicked.connect(self._save_working_copy)
        self.data_save_copy_button.setMinimumHeight(38)

        template_button_row.addWidget(self.template_button)
        template_button_row.addWidget(self.data_save_copy_button)

        template_layout.addLayout(template_path_row)
        template_layout.addLayout(template_button_row)
        layout.addWidget(template_group)

        workbook_group = QGroupBox("Workbook Validation")
        workbook_layout = QVBoxLayout(workbook_group)

        self.validate_button = QPushButton("Validate Current Workbook")
        self.validate_button.clicked.connect(self._validate_current_workbook)
        workbook_layout.addWidget(self.validate_button)

        self.data_status_text = QPlainTextEdit()
        self.data_status_text.setReadOnly(True)
        self.data_status_text.setPlainText(
            "This app now uses one workbook file:\n\n"
            "- Summer Scheduler Workbook.xlsx\n\n"
            "Use Create Blank Template for a clean starter file, or Save Working Copy to copy your current real workbook.\n\n"
            "Room Scope tips:\n- Available Day can be left blank and defaults to Day 1.\n- Leave task toggles blank to include by default. Type FALSE only for rooms or tasks you want to skip.\n\n"
            "Global settings control whole task types. Room Scope toggles control room-level exceptions. Both must be on for the task to run.\n\n"
            "Holidays tip: add a Holidays sheet with Date, Observed Date and Counts As Non-Workday. If Observed Date is blank, Saturday holidays observe Friday and Sunday holidays observe Monday."
        )
        workbook_layout.addWidget(self.data_status_text)

        layout.addWidget(workbook_group)
        layout.addStretch(1)
        return panel

    def _apply_theme(self) -> None:
        if self.current_theme == "dark":
            self.setStyleSheet(get_dark_stylesheet())
        else:
            self.setStyleSheet(get_light_stylesheet())

    def _on_theme_changed(self, value: str) -> None:
        self.current_theme = "dark" if value.lower() == "dark" else "light"
        self._apply_theme()

        # Re-render summary HTML so inline text colors match the selected theme.
        # Without this, switching from dark to light mode can leave old white text
        # inside white summary panels.
        if self.result and self.settings:
            self._populate_summary()

    def _show_error(self, title: str, message: str) -> None:
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec()

    def _show_info(self, title: str, message: str) -> None:
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec()

    def _browse_workbook(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Summer Scheduler Workbook",
            str(Path.cwd()),
            "Excel Workbooks (*.xlsx *.xlsm)",
        )
        if file_path:
            resolved = str(Path(file_path).resolve())
            self.workbook_path_edit.setText(resolved)
            self.template_path_edit.setText(resolved)
            set_last_workbook_file(resolved)
            self._validate_current_workbook()
            if hasattr(self, "instruction_method_combo"):
                self._load_instruction_methods_from_workbook(show_errors=False)

    def _sync_carpet_toggle_state(self) -> None:
        carpet_on = self.include_carpet_check.isChecked()
        self.general_can_do_carpet_check.setEnabled(True)
        if carpet_on:
            self.general_can_do_carpet_check.setToolTip(
                "Runtime override. Checked means the general crew may use leftover capacity for carpet work."
            )
        else:
            self.general_can_do_carpet_check.setToolTip(
                "Carpet is globally disabled, so this override will not affect the current run until carpet is enabled."
            )

    def _refresh_override_mode_label(self) -> None:
        mode = self.override_mode_combo.currentText()

        if mode == "Global":
            self.override_anchor_label.setVisible(False)
            self.override_anchor_edit.setVisible(False)
            self.override_length_label.setVisible(False)
            self.override_length_edit.setVisible(False)

            self.override_anchor_edit.setText("1")
            self.override_length_edit.setText("1")

        elif mode == "Weekly":
            self.override_anchor_label.setVisible(True)
            self.override_anchor_edit.setVisible(True)
            self.override_length_label.setVisible(True)
            self.override_length_edit.setVisible(True)

            self.override_anchor_label.setText("Start Week")
            self.override_length_label.setText("Duration (weeks)")

            if not self.override_anchor_edit.text().strip():
                self.override_anchor_edit.setText("1")
            if not self.override_length_edit.text().strip():
                self.override_length_edit.setText("1")

        else:
            self.override_anchor_label.setVisible(True)
            self.override_anchor_edit.setVisible(True)
            self.override_length_label.setVisible(True)
            self.override_length_edit.setVisible(True)

            self.override_anchor_label.setText("Start Day")
            self.override_length_label.setText("Duration (days)")

            if not self.override_anchor_edit.text().strip():
                self.override_anchor_edit.setText("1")
            if not self.override_length_edit.text().strip():
                self.override_length_edit.setText("1")

    def _resolve_workbook_path(self) -> str:
        raw_path = self.workbook_path_edit.text().strip()

        if not raw_path:
            raise ValueError("No workbook file selected.")

        path_obj = Path(raw_path)

        if not path_obj.exists():
            raise ValueError(f"Workbook file does not exist:\n{path_obj}")

        if not path_obj.is_file():
            raise ValueError("Select the Excel workbook file, not a folder.")

        if path_obj.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Workbook must be an .xlsx or .xlsm file.")

        return str(path_obj.resolve())

    def _load_defaults_into_form(self) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
            errors = validate_workbook(workbook_path)
            if errors:
                raise ValueError("\n".join(errors))

            settings = load_settings(workbook_path)
            self.holidays = load_holidays(workbook_path)
        except Exception as exc:
            self._show_error("Load Error", f"Could not load workbook defaults.\n\n{exc}")
            return

        config = load_app_config()
        last_start_date = str(config.get("last_schedule_start_date", "") or "").strip()
        last_target_end_date = str(config.get("last_target_end_date", "") or "").strip()

        self.schedule_name_edit.setText(settings.schedule_name)
        self.schedule_start_date_edit.setText(last_start_date or settings.schedule_start_date)
        self.current_day_edit.setText(str(settings.current_day))
        self.target_end_date_edit.setText(last_target_end_date or settings.target_end_date)
        self.paid_holidays_edit.setText(str(settings.paid_holidays_in_range))
        self.work_on_weekends_check.setChecked(settings.work_on_weekends)

        self.include_deep_clean_check.setChecked(settings.include_deep_clean)
        self.include_strip_check.setChecked(settings.include_strip)
        self.include_wax_check.setChecked(settings.include_wax)
        self.include_carpet_check.setChecked(settings.include_carpet)
        self.include_exterior_check.setChecked(settings.include_exterior)

        self.shift_hours_edit.setText(str(settings.scheduled_shift_hours_per_day))
        self.lunch_hours_edit.setText(str(settings.lunch_hours_per_day))
        self.break_hours_edit.setText(str(settings.break_hours_per_day))
        self.setup_hours_edit.setText(str(settings.setup_hours_per_day))
        self.cleanup_hours_edit.setText(str(settings.cleanup_hours_per_day))
        self.productive_hours_edit.setText(str(settings.productive_hours_per_staff_per_day))

        self._sync_carpet_toggle_state()
        if hasattr(self, "instruction_method_combo"):
            self._load_instruction_methods_from_workbook(show_errors=False)
        self.data_status_text.setPlainText(f"Workbook loaded:\n{workbook_path}")

    def _apply_form_overrides_to_settings(self) -> None:
        self.settings.schedule_name = (
            self.schedule_name_edit.text().strip() or self.settings.schedule_name
        )
        self.settings.schedule_start_date = (
            self.schedule_start_date_edit.text().strip() or self.settings.schedule_start_date
        )
        self.settings.target_end_date = (
            self.target_end_date_edit.text().strip() or self.settings.target_end_date
        )

        config = load_app_config()
        config["last_schedule_start_date"] = self.settings.schedule_start_date
        config["last_target_end_date"] = self.settings.target_end_date
        save_app_config(config)

        self.settings.paid_holidays_in_range = parse_int_field(
            self.paid_holidays_edit.text(),
            default=self.settings.paid_holidays_in_range,
            field_name="Paid Holidays",
        )
        self.settings.current_day = parse_int_field(
            self.current_day_edit.text(),
            default=self.settings.current_day,
            field_name="Current Day",
        )
        self.settings.work_on_weekends = self.work_on_weekends_check.isChecked()

        self.settings.include_deep_clean = self.include_deep_clean_check.isChecked()
        self.settings.include_strip = self.include_strip_check.isChecked()
        self.settings.include_wax = self.include_wax_check.isChecked()
        self.settings.include_carpet = self.include_carpet_check.isChecked()
        self.settings.include_exterior = self.include_exterior_check.isChecked()
        self.settings.general_crew_can_do_carpet = self.general_can_do_carpet_check.isChecked()

        self.settings.scheduled_shift_hours_per_day = parse_float_field(
            self.shift_hours_edit.text(),
            default=self.settings.scheduled_shift_hours_per_day,
            field_name="Shift Hours",
        )
        self.settings.lunch_hours_per_day = parse_float_field(
            self.lunch_hours_edit.text(),
            default=self.settings.lunch_hours_per_day,
            field_name="Lunch Hours",
        )
        self.settings.break_hours_per_day = parse_float_field(
            self.break_hours_edit.text(),
            default=self.settings.break_hours_per_day,
            field_name="Break Hours",
        )
        self.settings.setup_hours_per_day = parse_float_field(
            self.setup_hours_edit.text(),
            default=self.settings.setup_hours_per_day,
            field_name="Setup Hours",
        )
        self.settings.cleanup_hours_per_day = parse_float_field(
            self.cleanup_hours_edit.text(),
            default=self.settings.cleanup_hours_per_day,
            field_name="Cleanup Hours",
        )
        self.settings.productive_hours_per_staff_per_day = parse_float_field(
            self.productive_hours_edit.text(),
            default=self.settings.productive_hours_per_staff_per_day,
            field_name="Productive Hours",
        )

        if self.settings.target_end_date:
            self.settings.target_end_day = calculate_workdays(
                self.settings.schedule_start_date,
                self.settings.target_end_date,
                self.settings.work_on_weekends,
                self.settings.paid_holidays_in_range,
                holidays=self.holidays,
            )

        self.settings.validate_or_normalize()

    def _render_override_preview(self) -> None:
        if not self.staffing_overrides:
            self.override_preview_text.setPlainText("No staffing overrides yet.")
            return

        lines = []
        for index, item in enumerate(self.staffing_overrides, start=1):
            lines.append(
                f'{index}. {item["label"]} | Cleaning {item["cleaning_staff"]} | '
                f'Carpet {item["carpet_staff"]} | Outside {item["outside_help"]} | '
                f'Absences {item["absences"]}'
            )

        self.override_preview_text.setPlainText("\n".join(lines))

    def _apply_staffing_override(self) -> None:
        try:
            mode = self.override_mode_combo.currentText()
            start_date = self.schedule_start_date_edit.text().strip()
            target_end_date = self.target_end_date_edit.text().strip()
            paid_holidays = parse_int_field(self.paid_holidays_edit.text(), default=0, field_name="Paid Holidays")
            if not target_end_date:
                raise ValueError("Target End Date is required before applying a staffing override.")
            target_end_day = calculate_workdays(
                start_date,
                target_end_date,
                self.work_on_weekends_check.isChecked(),
                paid_holidays,
                holidays=self.holidays,
            )

            cleaning_staff = parse_int_field(self.override_cleaning_staff_edit.text(), default=0, field_name="Cleaning Staff")
            carpet_staff = parse_int_field(self.override_carpet_staff_edit.text(), default=0, field_name="Carpet Staff")
            outside_help = parse_int_field(self.override_outside_help_edit.text(), default=0, field_name="Outside Help")
            absences = parse_int_field(self.override_absences_edit.text(), default=0, field_name="Absences")

            if mode == "Global":
                day_start = 1
                day_end = target_end_day
                label = f"Global Days 1-{target_end_day}"
            elif mode == "Weekly":
                week_num = parse_int_field(self.override_anchor_edit.text(), default=1, field_name="Start Week")
                week_span = parse_int_field(self.override_length_edit.text(), default=1, field_name="Duration (weeks)")
                day_start = ((week_num - 1) * 5) + 1
                day_end = day_start + (week_span * 5) - 1
                label = f"Week {week_num} for {week_span} week(s) -> Days {day_start}-{day_end}"
            else:
                start_day = parse_int_field(self.override_anchor_edit.text(), default=1, field_name="Start Day")
                day_span = parse_int_field(self.override_length_edit.text(), default=1, field_name="Duration (days)")
                day_start = start_day
                day_end = start_day + day_span - 1
                label = f"Days {day_start}-{day_end}"

            self.staffing_overrides.append(
                {
                    "mode": mode,
                    "day_start": day_start,
                    "day_end": day_end,
                    "cleaning_staff": cleaning_staff,
                    "carpet_staff": carpet_staff,
                    "outside_help": outside_help,
                    "absences": absences,
                    "label": label,
                }
            )
            self._render_override_preview()

        except Exception as exc:
            self._show_error("Override Error", str(exc))

    def _clear_staffing_overrides(self) -> None:
        self.staffing_overrides = []
        self._render_override_preview()

    def _build_effective_staffing_maps(self):
        cleaning_staff_by_day = {}
        carpet_staff_by_day = {}
        outside_help_by_day = {}
        absences_by_day = {}

        max_day = self.settings.target_end_day

        if self.staffing_days:
            max_day = max(max_day, max(item.day for item in self.staffing_days))

        if self.staffing_overrides:
            max_day = max(max_day, max(ov["day_end"] for ov in self.staffing_overrides))

        for day in range(1, max_day + 1):
            matching = next((item for item in self.staffing_days if item.day == day), None)

            if matching:
                cleaning_staff_by_day[day] = matching.general_crew_staff()
                carpet_staff_by_day[day] = matching.carpet_crew_staff()
                outside_help_by_day[day] = matching.temporary_help
                absences_by_day[day] = matching.absences
            else:
                cleaning_staff_by_day[day] = 0
                carpet_staff_by_day[day] = 0
                outside_help_by_day[day] = 0
                absences_by_day[day] = 0

        for override in self.staffing_overrides:
            for day in range(override["day_start"], override["day_end"] + 1):
                cleaning_staff_by_day[day] = override["cleaning_staff"]
                carpet_staff_by_day[day] = override["carpet_staff"]
                outside_help_by_day[day] = override["outside_help"]
                absences_by_day[day] = override["absences"]

        return {
            "cleaning_staff_by_day": cleaning_staff_by_day,
            "carpet_staff_by_day": carpet_staff_by_day,
            "outside_help_by_day": outside_help_by_day,
            "absences_by_day": absences_by_day,
        }

    def run_scheduler_from_ui(self) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
            errors = validate_workbook(workbook_path)
            if errors:
                raise ValueError("\n".join(errors))

            self.settings = load_settings(workbook_path)
            self.rooms, self.schools = load_rooms(workbook_path)
            self.staffing_days = load_staffing(workbook_path)
            self.progress_entries = load_progress(workbook_path)
            self.holidays = load_holidays(workbook_path)

            self._apply_form_overrides_to_settings()
            staffing_maps = self._build_effective_staffing_maps()

            scenario = ScenarioInput(
                settings=self.settings,
                rooms=self.rooms,
                progress_entries=self.progress_entries,
                cleaning_staff_by_day=staffing_maps["cleaning_staff_by_day"],
                carpet_staff_by_day=staffing_maps["carpet_staff_by_day"],
                outside_help_by_day=staffing_maps["outside_help_by_day"],
                absences_by_day=staffing_maps["absences_by_day"],
            )

            self.status_chip.setText("Running...")
            self.tabs.setCurrentWidget(self.run_tab)
            self._reset_run_overview_for_animation()

            self.result = run_scenario(scenario)

            self._populate_summary()

        except Exception as exc:
            self.status_chip.setText("Run failed")
            error_text = f"{type(exc).__name__}: {exc}"
            self.cleaning_summary_text.setPlainText(error_text)
            self.carpet_summary_text.setPlainText(error_text)
            self.day_detail_text.setPlainText(error_text)
            self.day_outlook_text.setPlainText(error_text)
            self.days_table.setRowCount(0)
            self.worklog_table.setRowCount(0)
            self.export_status_text.setPlainText(error_text)
            self._show_error("Scheduler Error", error_text)

    def _day_has_exterior_work(self, day) -> bool:
        return any((item.phase_name or "") == "Exterior" for item in day.work_log)

    def _day_has_blocked_later_interior_work(self, day) -> bool:
        """Return True when this day still has same-site interior work blocked for later."""
        if not self.result:
            return False

        touched_sites = {
            item.school_name
            for item in day.work_log
            if item.school_name and item.school_name != "MULTI-SCHOOL"
        }
        if day.active_school_name and "|" not in day.active_school_name:
            touched_sites.add(day.active_school_name)

        if not touched_sites:
            return False

        ignored_phases = {"Exterior", "Carpet", "Site Transition", "Transition / Logistics"}
        for task in self.result.task_items:
            if task.school_name not in touched_sites:
                continue
            if (task.phase_name or "") in ignored_phases:
                continue
            if getattr(task, "available_day", 1) > day.day:
                return True
        return False

    def _day_access_notice(self, day) -> str:
        notes = []
        if "ACCESS DOWNTIME WARNING" in (day.status_note or ""):
            notes.append(
                "Access downtime: crew has more than 1 hour of unused time because rooms are restricted. Consider assigning alternate work."
            )

        if self._day_has_exterior_work(day) and self._day_has_blocked_later_interior_work(day):
            notes.append(
                "Access constraint workaround: exterior work is being used while interior rooms are blocked. Check this day before handing off assignments."
            )

        return " ".join(notes)

    def _build_access_alert_block(self) -> str:
        if not self.result or not self.settings:
            return ""

        alert_days = []
        for day in self.result.days:
            notice = self._day_access_notice(day)
            if not notice:
                continue
            label = format_workday_label(
                self.settings.schedule_start_date,
                day.day,
                self.settings.work_on_weekends,
                holidays=self.holidays,
            )
            alert_days.append((label, notice))

        if not alert_days:
            return ""

        display_items = alert_days[:8]
        more_text = ""
        if len(alert_days) > len(display_items):
            more_text = f" plus {len(alert_days) - len(display_items)} more day(s)"

        rows = "".join(
            f"<li><b>{label}</b>: {notice}</li>"
            for label, notice in display_items
        )

        if self.current_theme == "dark":
            body_bg = "#141414"
            body_text = "#f5f5f5"
            border = "#E6B800"
            muted = "#d4d4d4"
        else:
            body_bg = "#fffdf4"
            body_text = "#111827"
            border = "#E6B800"
            muted = "#374151"

        return f"""
        <div style="
            background:{body_bg};
            color:{body_text};
            border:1px solid {border};
            border-radius:10px;
            overflow:hidden;
            margin:6px 0;
        ">
            <div style="
                background:#E6B800;
                color:#000000;
                padding:8px 12px;
                font-size:15px;
                font-weight:800;
            ">Schedule Attention Needed</div>
            <div style="padding:10px 12px; font-weight:700;">
                <div style="margin-bottom:6px;">Hey, look at this before exporting or handing this schedule to staff.</div>
                <ul style="margin-top:6px; margin-bottom:6px; padding-left:20px;">{rows}</ul>
                <div>{more_text}</div>
                <div style="margin-top:6px; color:{muted};">Check the affected Day Outlook and Day Note. These are usually caused by restricted rooms, summer school access, or exterior work being used as filler while interior rooms are blocked.</div>
            </div>
        </div>
        """

    def _populate_summary(self) -> None:
        if not self.result or not self.settings:
            return

        rec = self.result.recommendation

        finish_label = format_workday_label(
            self.settings.schedule_start_date,
            self.result.finish_day,
            self.settings.work_on_weekends,
            holidays=self.holidays,
        )
        target_label = format_workday_label(
            self.settings.schedule_start_date,
            self.result.target_end_day,
            self.settings.work_on_weekends,
            holidays=self.holidays,
        )
        current_label = format_workday_label(
            self.settings.schedule_start_date,
            self.result.current_day,
            self.settings.work_on_weekends,
            holidays=self.holidays,
        )

        target_end_date = self.target_end_date_edit.text().strip()
        paid_holidays = parse_int_field(self.paid_holidays_edit.text(), default=0, field_name="Paid Holidays")

        total_workdays = None
        if target_end_date:
            try:
                total_workdays = calculate_workdays(
                    self.settings.schedule_start_date,
                    target_end_date,
                    self.settings.work_on_weekends,
                    paid_holidays,
                    holidays=self.holidays,
                )
            except Exception:
                total_workdays = None

        access_warning_block = self._build_access_alert_block()

        cleaning_blocks = [
            self._summary_table_html(
                "Schedule",
                [
                    ("Schedule Name:", self.result.schedule_name),
                    ("Start Date:", self.settings.schedule_start_date),
                    ("Weekends Count As Workdays:", yes_no(self.settings.work_on_weekends)),
                    ("Current Position:", current_label),
                    ("Target End (Day):", target_label),
                    ("Target End (Date):", target_end_date or "—"),
                    ("Total Workdays Available:", str(total_workdays) if total_workdays is not None else "—"),
                    ("Holiday Dates Loaded:", ", ".join(sorted(d.isoformat() for d in self.holidays)) if self.holidays else "None"),
                    ("Legacy Paid Holidays:", str(paid_holidays)),
                    ("Projected Finish:", finish_label),
                    ("Deadline Met:", yes_no(self.result.met_deadline)),
                ],
            ),
            self._summary_table_html(
                "Workload",
                [
                    ("Total Planned Hours:", fmt_hours(self.result.total_planned_hours)),
                    ("Completed Before Rerun:", fmt_hours(self.result.completed_hours_before_run)),
                    ("Remaining At Start:", fmt_hours(self.result.remaining_hours_at_start)),
                    ("Total Used Hours:", fmt_hours(self.result.total_used_hours)),
                    ("Remaining Backlog:", fmt_hours(self.result.remaining_backlog_hours)),
                ],
            ),
            self._summary_table_html(
                "Recommendation",
                [
                    ("Status:", rec.status_label),
                    ("Bottleneck:", rec.bottleneck_type),
                    ("Available-Now Backlog At Start:", fmt_hours(rec.available_backlog_hours)),
                    ("Blocked-Later Backlog At Start:", fmt_hours(rec.blocked_backlog_hours)),
                    ("Total Backlog At Start:", fmt_hours(rec.total_backlog_hours)),
                    ("Capacity To Deadline:", fmt_hours(rec.capacity_to_deadline_hours)),
                    ("Extra Staff-Days Needed:", fmt_hours(rec.extra_staff_days_needed)),
                    ("Action:", rec.recommended_action),
                ],
            ),
            self._summary_table_html(
                "Daily Time Model",
                [
                    ("Shift Hours:", fmt_hours(self.settings.scheduled_shift_hours_per_day)),
                    ("Lunch Hours:", fmt_hours(self.settings.lunch_hours_per_day)),
                    ("Break Hours:", fmt_hours(self.settings.break_hours_per_day)),
                    ("Setup Hours:", fmt_hours(self.settings.setup_hours_per_day)),
                    ("Cleanup Hours:", fmt_hours(self.settings.cleanup_hours_per_day)),
                    ("Productive Hours:", fmt_hours(self.settings.productive_hours_per_staff_per_day)),
                ],
            ),
        ]

        if access_warning_block:
            cleaning_blocks.insert(1, access_warning_block)

        carpet_blocks = [
            self._summary_table_html(
                "Carpet Summary",
                [
                    ("Carpet Included:", yes_no(self.settings.include_carpet)),
                    (
                        "General Crew Allowed To Do Carpet:",
                        yes_no(self.general_can_do_carpet_check.isChecked()),
                    ),
                    ("UI Staffing Overrides Loaded:", str(len(self.staffing_overrides))),
                ],
            ),
            self._summary_table_html(
                "Status",
                [
                    ("Note:", "Carpet-specific summary formatting is coming next."),
                ],
            ),
        ]

        self.summary_reveal_steps = [
            ("card", "finish", finish_label, 180),
            ("card", "deadline", yes_no(self.result.met_deadline), 180),
            ("card", "backlog", f"{fmt_hours(self.result.remaining_backlog_hours)} hrs", 220),
            ("card", "recommendation", rec.status_label, 280),
        ]

        for block in cleaning_blocks:
            self.summary_reveal_steps.append(("cleaning", block, None, 240))

        for block in carpet_blocks:
            self.summary_reveal_steps.append(("carpet", block, None, 240))

        self.summary_reveal_steps.append(("finalize", None, None, 0))
        self.summary_reveal_index = 0
        self.summary_reveal_timer.start(100)

    def _reset_run_overview_for_animation(self) -> None:
        self.finish_day_card.set_value("-")
        self.deadline_card.set_value("-")
        self.backlog_card.set_value("-")
        self.recommendation_card.set_value("-")

        self.cleaning_summary_text.clear()
        self.carpet_summary_text.clear()

        self.days_table.setRowCount(0)
        self.worklog_table.setRowCount(0)
        self.day_detail_text.clear()
        self.day_outlook_text.clear()

    def _append_html_block(self, widget, html_block: str) -> None:
        current = widget.toHtml()
        if widget.toPlainText().strip():
            widget.setHtml(current + "<br><br>" + html_block)
        else:
            widget.setHtml(html_block)

    def _summary_table_html(self, title: str, rows: list[tuple[str, str]]) -> str:
        if self.current_theme == "dark":
            label_color = "#f9fafb"
            value_color = "#e5e7eb"
            title_color = "#f9fafb"
        else:
            label_color = "#111827"
            value_color = "#374151"
            title_color = "#111827"

        row_html = "".join(
            f"""
            <tr>
                <td style="
                    font-weight:700;
                    white-space:nowrap;
                    padding:0 22px 6px 0;
                    vertical-align:top;
                    color:{label_color};
                ">{label}</td>
                <td style="
                    padding:0 0 6px 0;
                    vertical-align:top;
                    color:{value_color};
                ">{value}</td>
            </tr>
            """
            for label, value in rows
        )

        return f"""
        <div style="color:{value_color};">
            <div style="
                font-weight:700;
                text-decoration:underline;
                margin-bottom:10px;
                color:{title_color};
            ">{title}</div>
            <table style="width:100%; border-collapse:collapse;">
                {row_html}
            </table>
        </div>
        """

    def _reveal_next_summary_step(self) -> None:
        if self.summary_reveal_index >= len(self.summary_reveal_steps):
            self.summary_reveal_timer.stop()
            return

        step = self.summary_reveal_steps[self.summary_reveal_index]
        step_type = step[0]
        payload = step[1]
        extra = step[2]
        next_delay = step[3]

        if step_type == "card":
            if payload == "finish":
                self.finish_day_card.set_value(extra)
            elif payload == "deadline":
                self.deadline_card.set_value(extra)
            elif payload == "backlog":
                self.backlog_card.set_value(extra)
            elif payload == "recommendation":
                self.recommendation_card.set_value(extra)

        elif step_type == "cleaning":
            self._append_html_block(self.cleaning_summary_text, payload)

        elif step_type == "carpet":
            self._append_html_block(self.carpet_summary_text, payload)

        elif step_type == "finalize":
            self._populate_days_table()
            self.status_chip.setText("Run complete")
            self.export_status_text.setPlainText(
                "Run complete.\n\nGo to Export tab to write an Excel report workbook."
            )
            self.summary_reveal_timer.stop()

        self.summary_reveal_index += 1

        if self.summary_reveal_timer.isActive():
            self.summary_reveal_timer.start(next_delay)

    def _make_table_item(self, value: str) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        return item

    def _populate_days_table(self) -> None:
        self.days_table.setRowCount(0)
        self.worklog_table.setRowCount(0)
        self.day_detail_text.clear()
        self.day_outlook_text.clear()

        if not self.result or not self.settings:
            return

        days = self.result.days
        self.days_table.setRowCount(len(days))

        for row, day in enumerate(days):
            day_label = format_workday_label(
                self.settings.schedule_start_date,
                day.day,
                self.settings.work_on_weekends,
                holidays=self.holidays,
            )
            date_only = day_label.split(" - ", 1)[1] if " - " in day_label else day_label

            self.days_table.setItem(row, 0, self._make_table_item(str(day.day)))
            self.days_table.setItem(row, 1, self._make_table_item(date_only))
            self.days_table.setItem(row, 2, self._make_table_item(day.active_school_name or ""))
            self.days_table.setItem(row, 3, self._make_table_item(str(day.general_staff)))
            self.days_table.setItem(row, 4, self._make_table_item(str(day.carpet_staff)))
            self.days_table.setItem(row, 5, self._make_table_item(fmt_hours(day.daily_capacity)))
            self.days_table.setItem(row, 6, self._make_table_item(fmt_hours(day.used_capacity)))
            self.days_table.setItem(row, 7, self._make_table_item(fmt_hours(day.unused_capacity)))
            self.days_table.setItem(row, 8, self._make_table_item(str(len(day.work_log))))
            day_note = day.status_note or ""
            access_notice = self._day_access_notice(day)
            if access_notice:
                day_note = f"{day_note} {access_notice}".strip()
            self.days_table.setItem(row, 9, self._make_table_item(day_note))

        if days:
            self.days_table.selectRow(0)
            self._populate_day_detail()

    def _populate_day_detail(self) -> None:
        if not self.result or not self.settings:
            return

        selected_rows = self.days_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        day_index = selected_rows[0].row()
        day = self.result.days[day_index]

        self.worklog_table.setRowCount(len(day.work_log))

        for row, item in enumerate(day.work_log):
            self.worklog_table.setItem(row, 0, self._make_table_item(item.crew_type))
            self.worklog_table.setItem(row, 1, self._make_table_item(item.school_name))
            self.worklog_table.setItem(row, 2, self._make_table_item(item.building_name))
            self.worklog_table.setItem(row, 3, self._make_table_item(item.zone_name))
            self.worklog_table.setItem(row, 4, self._make_table_item(item.room_name))
            self.worklog_table.setItem(row, 5, self._make_table_item(item.phase_name))
            self.worklog_table.setItem(row, 6, self._make_table_item(fmt_hours(item.hours_done)))
            self.worklog_table.setItem(row, 7, self._make_table_item(item.note or ""))

        day_label = format_workday_label(
            self.settings.schedule_start_date,
            day.day,
            self.settings.work_on_weekends,
            holidays=self.holidays,
        )

        access_notice = self._day_access_notice(day)
        full_day_note = day.status_note or "None"
        if access_notice:
            full_day_note = f"{full_day_note} {access_notice}".strip()

        detail_lines = [
            day_label,
            f"Active School: {day.active_school_name or 'None'}",
            f"Day Note: {full_day_note}",
            "",
            f"Effective Staff: {day.effective_staff}",
            f"General Staff: {day.general_staff}",
            f"Carpet Staff: {day.carpet_staff}",
            "",
            f"General Capacity: {fmt_hours(day.general_capacity)} hrs",
            f"Carpet Capacity: {fmt_hours(day.carpet_capacity)} hrs",
            f"Daily Capacity: {fmt_hours(day.daily_capacity)} hrs",
            "",
            f"General Used: {fmt_hours(day.general_used_capacity)} hrs",
            f"Carpet Used: {fmt_hours(day.carpet_used_capacity)} hrs",
            f"Total Used: {fmt_hours(day.used_capacity)} hrs",
            "",
            f"General Unused: {fmt_hours(day.general_unused_capacity)} hrs",
            f"Carpet Unused: {fmt_hours(day.carpet_unused_capacity)} hrs",
            f"Total Unused: {fmt_hours(day.unused_capacity)} hrs",
            "",
            f"Work Items: {len(day.work_log)}",
        ]

        self.day_detail_text.setPlainText("\n".join(detail_lines))
        outlook_text = format_time_blocks_for_text(day, self.settings)
        if access_notice:
            outlook_text = f"⚠ {access_notice}\n\n{outlook_text}"
        self.day_outlook_text.setPlainText(outlook_text)

    def _export_result_workbook(self) -> None:
        if not self.result or not self.settings:
            self._show_error("Nothing To Export", "Run the scheduler first.")
            return

        try:
            suggested = str(Path.cwd() / "Full Schedule Results.xlsx")
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Schedule Results Workbook",
                suggested,
                "Excel Workbooks (*.xlsx)",
            )
            if not file_path:
                return

            destination = Path(file_path).resolve()
            if destination.suffix.lower() != ".xlsx":
                destination = destination.with_suffix(".xlsx")

            exported_path = export_result_workbook(
                self.result,
                self.settings,
                str(destination),
                holidays=self.holidays,
            )

            self.export_path_edit.setText(str(destination.parent))
            self.export_status_text.setPlainText(
                f"Export complete:\n{exported_path}"
            )
            self._show_info("Export Complete", exported_path)
        except Exception as exc:
            self.export_status_text.setPlainText(str(exc))
            self._show_error("Export Error", str(exc))

    def _create_template_from_ui(self) -> None:
        try:
            suggested = self.template_path_edit.text().strip()
            if not suggested:
                suggested = str(Path.cwd() / "Summer Scheduler Workbook.xlsx")

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Summer Scheduler Workbook Template",
                suggested,
                "Excel Workbooks (*.xlsx)",
            )
            if not file_path:
                return

            created = create_input_template(file_path)
            created_path = Path(created).resolve()

            self.workbook_path_edit.setText(str(created_path))
            self.template_path_edit.setText(str(created_path))
            set_last_workbook_file(str(created_path))

            self.data_status_text.setPlainText(
                "Workbook created:\n\n"
                f"- {created_path.name}\n\n"
                f"File:\n{created_path}"
            )

            self._show_info(
                "Workbook Created",
                "Created workbook:\n\n"
                f"{created_path.name}\n\n"
                f"File:\n{created_path}"
            )
        except Exception as exc:
            self._show_error("Template Error", str(exc))

    def _save_working_copy(self) -> None:
        try:
            source_path = Path(self._resolve_workbook_path()).resolve()

            suggested_name = source_path.with_name(
                f"{source_path.stem} - Working Copy{source_path.suffix}"
            )

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Working Copy",
                str(suggested_name),
                "Excel Workbooks (*.xlsx *.xlsm)",
            )
            if not file_path:
                return

            destination_path = Path(file_path).resolve()
            if destination_path.suffix.lower() not in {".xlsx", ".xlsm"}:
                destination_path = destination_path.with_suffix(source_path.suffix)

            if destination_path == source_path:
                raise ValueError(
                    "Choose a different file name or location. Save Working Copy should not overwrite the current workbook."
                )

            shutil.copy2(source_path, destination_path)

            self.workbook_path_edit.setText(str(destination_path))
            self.template_path_edit.setText(str(destination_path))
            set_last_workbook_file(str(destination_path))

            self.data_status_text.setPlainText(
                "Working copy saved and selected:\n\n"
                f"File:\n{destination_path}"
            )

            self._show_info(
                "Working Copy Saved",
                "Saved and selected working copy:\n\n"
                f"{destination_path}",
            )
        except Exception as exc:
            self._show_error("Save Working Copy Error", str(exc))

    def _validate_current_workbook(self) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
        except Exception as exc:
            self.data_status_text.setPlainText(f"Workbook validation failed:\n\n- {exc}")
            return

        errors = validate_workbook(workbook_path)
        file_path = Path(workbook_path).resolve()

        expected_sheets = [
            "Sites",
            "Rooms",
            "Setup",
            "Run Settings",
            "Room Scope",
            "Staffing",
            "Progress",
            "Holidays (optional but recommended)",
        ]

        if errors:
            self.data_status_text.setPlainText(
                "Workbook validation failed:\n\n"
                + "\n".join(f"- {e}" for e in errors)
                + "\n\nExpected workbook:\n- Summer Scheduler Workbook.xlsx"
                + "\n\nExpected sheets:\n"
                + "\n".join(f"- {name}" for name in expected_sheets)
                + f"\n\nChecked file:\n{file_path}"
            )
        else:
            self.data_status_text.setPlainText(
                "Workbook looks valid:\n\n"
                + "\n".join(f"- {name}" for name in expected_sheets)
                + f"\n\nFile:\n{file_path}"
            )


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Summer Scheduler")

    font = QFont("Segoe UI", 10)
    app.setFont(font)

    window = SchedulerWindow()
    window.showMaximized()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
