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

from calendar_math import format_workday_label
from data_loader import (
    load_progress,
    load_rooms,
    load_settings,
    load_staffing,
    validate_workbook,
)
from exporter import create_input_template, export_result_workbook
from models import StaffingDay
from services.scenario_service import ScenarioInput, run_scenario
from timeblock_generator import format_time_blocks_for_text


def fmt_hours(value: float) -> str:
    return f"{value:.2f}"


def yes_no(value: bool) -> str:
    return "Yes" if value else "No"

def calculate_workdays(
    start_date_str: str,
    end_date_str: str,
    include_weekends: bool,
    paid_holidays: int = 0,
) -> int:
    start_date = datetime.strptime(start_date_str.strip(), "%Y-%m-%d").date()
    end_date = datetime.strptime(end_date_str.strip(), "%Y-%m-%d").date()

    if end_date < start_date:
        return 0

    count = 0
    current = start_date

    while current <= end_date:
        if include_weekends or current.weekday() < 5:
            count += 1
        current += timedelta(days=1)

    count -= paid_holidays
    return max(count, 0)


def get_dark_stylesheet() -> str:
    return """
    QWidget {
        font-size: 13px;
        color: #e5e7eb;
    }

    QMainWindow {
        background: #111827;
    }

    QLabel {
        color: #e5e7eb;
    }

    QTabWidget::pane {
        border: none;
        background: #111827;
    }

    QTabBar::tab {
        background: #1f2937;
        border: 1px solid #374151;
        padding: 8px 14px;
        margin-right: 4px;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        font-weight: 600;
        color: #d1d5db;
    }

    QTabBar::tab:selected {
        background: #243041;
        color: #ffffff;
    }

    #sidePanel {
        background: #1f2937;
        border: 1px solid #374151;
        border-radius: 12px;
    }

    #panelTitle {
        font-size: 20px;
        font-weight: 700;
        color: #f9fafb;
    }

    #pageTitle {
        font-size: 24px;
        font-weight: 700;
        color: #f9fafb;
    }

    #statusChip {
        background: #1e3a5f;
        color: #dbeafe;
        border: 1px solid #31537c;
        border-radius: 10px;
        padding: 6px 10px;
        font-weight: 600;
    }

    #summaryCard {
        background: #1f2937;
        border: 1px solid #374151;
        border-radius: 12px;
    }

    #summaryCardTitle {
        color: #9ca3af;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
    }

    #summaryCardValue {
        color: #f9fafb;
        font-size: 20px;
        font-weight: 700;
    }

    #primaryButton {
        background: #2563eb;
        color: white;
        font-weight: 700;
        padding: 8px 14px;
        border-radius: 8px;
        border: 1px solid #2563eb;
    }

    #primaryButton:hover {
        background: #1d4ed8;
    }

    QGroupBox {
        background: #1f2937;
        color: #f9fafb;
        border: 1px solid #374151;
        border-radius: 12px;
        margin-top: 10px;
        font-weight: 700;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px 0 6px;
        color: #d1d5db;
    }

    QPlainTextEdit, QTextEdit, QLineEdit, QTableWidget, QComboBox {
        background: #111827;
        color: #f9fafb;
        border: 1px solid #4b5563;
        border-radius: 8px;
        selection-background-color: #1d4ed8;
        selection-color: #ffffff;
    }

    QComboBox {
        padding: 4px 8px;
    }

    QComboBox QAbstractItemView {
        background: #111827;
        color: #f9fafb;
        border: 1px solid #4b5563;
        selection-background-color: #1d4ed8;
    }

    QHeaderView::section {
        background: #0f172a;
        color: #ffffff;
        padding: 6px;
        border: none;
        font-weight: 600;
    }

    QPushButton {
        padding: 8px 12px;
        border: 1px solid #4b5563;
        border-radius: 8px;
        background: #1f2937;
        color: #f9fafb;
        font-weight: 600;
    }

    QPushButton:hover {
        background: #273548;
    }

    QCheckBox {
        color: #e5e7eb;
    }

    QMessageBox {
        background-color: #1f2937;
    }

    QMessageBox QLabel {
        color: #f9fafb;
        min-width: 320px;
    }

    QMessageBox QPushButton {
        min-width: 80px;
        padding: 6px 12px;
        background: #111827;
        color: #f9fafb;
        border: 1px solid #4b5563;
        border-radius: 8px;
    }

    #mutedLabel {
        color: #9ca3af;
    }
    """


def get_light_stylesheet() -> str:
    return """
    QWidget {
        font-size: 13px;
        color: #1f2937;
    }

    QMainWindow {
        background: #f4f6f8;
    }

    QLabel {
        color: #1f2937;
    }

    QTabWidget::pane {
        border: none;
        background: #f4f6f8;
    }

    QTabBar::tab {
        background: #e8edf5;
        border: 1px solid #cfd6df;
        padding: 8px 14px;
        margin-right: 4px;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        font-weight: 600;
    }

    QTabBar::tab:selected {
        background: #ffffff;
        color: #111827;
    }

    #sidePanel {
        background: #ffffff;
        border: 1px solid #d9dee5;
        border-radius: 12px;
    }

    #panelTitle {
        font-size: 20px;
        font-weight: 700;
        color: #111827;
    }

    #pageTitle {
        font-size: 24px;
        font-weight: 700;
        color: #111827;
    }

    #statusChip {
        background: #e8eefc;
        color: #1f3a6d;
        border: 1px solid #cdd9f7;
        border-radius: 10px;
        padding: 6px 10px;
        font-weight: 600;
    }

    #summaryCard {
        background: #ffffff;
        border: 1px solid #d9dee5;
        border-radius: 12px;
    }

    #summaryCardTitle {
        color: #5a6473;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
    }

    #summaryCardValue {
        color: #111827;
        font-size: 20px;
        font-weight: 700;
    }

    #primaryButton {
        background: #1f6feb;
        color: white;
        font-weight: 700;
        padding: 8px 14px;
        border-radius: 8px;
        border: 1px solid #1f6feb;
    }

    #primaryButton:hover {
        background: #1a62cf;
    }

    QGroupBox {
        background: #ffffff;
        color: #111827;
        border: 1px solid #d9dee5;
        border-radius: 12px;
        margin-top: 10px;
        font-weight: 700;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 6px 0 6px;
        color: #374151;
    }

    QPlainTextEdit, QTextEdit, QLineEdit, QTableWidget, QComboBox {
        background: #ffffff;
        color: #111827;
        border: 1px solid #cfd6df;
        border-radius: 8px;
        selection-background-color: #dbeafe;
        selection-color: #111827;
    }

    QComboBox {
        padding: 4px 8px;
    }

    QComboBox QAbstractItemView {
        background: #ffffff;
        color: #111827;
        border: 1px solid #cfd6df;
        selection-background-color: #dbeafe;
    }

    QHeaderView::section {
        background: #374151;
        color: #ffffff;
        padding: 6px;
        border: none;
        font-weight: 600;
    }

    QPushButton {
        padding: 8px 12px;
        border: 1px solid #c7cfda;
        border-radius: 8px;
        background: #ffffff;
        color: #111827;
        font-weight: 600;
    }

    QPushButton:hover {
        background: #f5f7fb;
    }

    QCheckBox {
        color: #1f2937;
    }

    QMessageBox {
        background-color: #ffffff;
    }

    QMessageBox QLabel {
        color: #111827;
        min-width: 320px;
    }

    QMessageBox QPushButton {
        min-width: 80px;
        padding: 6px 12px;
        background: #ffffff;
        color: #111827;
        border: 1px solid #c7cfda;
        border-radius: 8px;
    }

    #mutedLabel {
        color: #5a6473;
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

        self.current_theme = "dark"
        self.staffing_overrides = []
        self.summary_reveal_timer = QTimer(self)
        self.summary_reveal_timer.timeout.connect(self._reveal_next_summary_step)

        self.summary_reveal_steps = []
        self.summary_reveal_index = 0
        

        self._build_ui()
        self._apply_theme()
        self._load_defaults_into_form()

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
        left_panel.setMinimumWidth(400)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setWidget(left_panel)

        # --- RIGHT PANEL ---
        right_panel = self._build_run_right_panel()
        right_panel.setMinimumWidth(900)

        splitter.addWidget(left_scroll)
        splitter.addWidget(right_panel)

        # Give the right panel more stretch, but keep left usable
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([430, 1150])

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

        workbook_group = QGroupBox("Workbook Set")
        workbook_layout = QVBoxLayout(workbook_group)

        workbook_row = QHBoxLayout()
        self.workbook_path_edit = QLineEdit("data")
        self.browse_workbook_button = QPushButton("Browse")
        self.browse_workbook_button.clicked.connect(self._browse_workbook)
        workbook_row.addWidget(self.workbook_path_edit)
        workbook_row.addWidget(self.browse_workbook_button)

        workbook_layout.addLayout(workbook_row)

        template_row = QHBoxLayout()
        self.make_template_button = QPushButton("Download Fresh Workbook Set")
        self.make_template_button.clicked.connect(self._create_template_from_ui)
        self.reload_button = QPushButton("Reload Workbook Set")
        self.reload_button.clicked.connect(self._load_defaults_into_form)
        template_row.addWidget(self.make_template_button)
        template_row.addWidget(self.reload_button)

        workbook_layout.addLayout(template_row)
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
        self.current_day_edit = QLineEdit()
        self.target_end_day_edit = QLineEdit()
        self.target_end_date_edit = QLineEdit()
        self.target_end_date_edit.setPlaceholderText("YYYY-MM-DD")

        self.paid_holidays_edit = QLineEdit("0")
        self.work_on_weekends_check = QCheckBox("Count weekends as workdays")

        self.include_deep_clean_check = QCheckBox("Include Deep Clean")
        self.include_strip_check = QCheckBox("Include Strip")
        self.include_wax_check = QCheckBox("Include Wax")
        self.include_carpet_check = QCheckBox("Include Carpet")
        self.include_exterior_check = QCheckBox("Include Exterior")

        self.include_carpet_check.toggled.connect(self._sync_carpet_toggle_state)

        self.general_can_do_carpet_check = QCheckBox(
            "General crew allowed to do carpet work"
        )
        self.general_can_do_carpet_check.setChecked(True)
        self.general_can_do_carpet_check.setEnabled(False)
        self.general_can_do_carpet_check.setToolTip(
            "Coming next. Engine support is not wired yet."
        )

        run_form.addRow("Schedule Name", self.schedule_name_edit)
        run_form.addRow("Start Date", self.schedule_start_date_edit)
        run_form.addRow("Current Day", self.current_day_edit)
        run_form.addRow("Target End Day", self.target_end_day_edit)
        run_form.addRow("Target End Date", self.target_end_date_edit)
        run_form.addRow("Paid Holidays", self.paid_holidays_edit)
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
            "The app now reads from a workbook set folder containing district data,\n"
            "planning assumptions and current run input.\n\n"
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
        self.days_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.days_table.horizontalHeader().setSectionResizeMode(
            9, QHeaderView.ResizeMode.Stretch
        )
        self.days_table.itemSelectionChanged.connect(self._populate_day_detail)

        days_layout.addWidget(self.days_table)
        top_layout.addWidget(days_group)

        bottom_half = QWidget()
        bottom_layout = QVBoxLayout(bottom_half)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(12)

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
        self.worklog_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.worklog_table.horizontalHeader().setSectionResizeMode(
            7, QHeaderView.ResizeMode.Stretch
        )

        worklog_layout.addWidget(self.worklog_table)
        bottom_layout.addWidget(worklog_group)

        detail_group = QGroupBox("Selected Day Details")
        detail_layout = QVBoxLayout(detail_group)

        self.day_detail_text = QPlainTextEdit()
        self.day_detail_text.setReadOnly(True)
        self.day_detail_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)

        detail_layout.addWidget(self.day_detail_text)
        bottom_layout.addWidget(detail_group)

        outlook_group = QGroupBox("Day Outlook")
        outlook_layout = QVBoxLayout(outlook_group)

        self.day_outlook_text = QPlainTextEdit()
        self.day_outlook_text.setReadOnly(True)
        self.day_outlook_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)

        outlook_layout.addWidget(self.day_outlook_text)
        bottom_layout.addWidget(outlook_group)

        detail_splitter.addWidget(top_half)
        detail_splitter.addWidget(bottom_half)
        detail_splitter.setSizes([380, 430])

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

        export_group = QGroupBox("Export Workbook")
        export_layout = QVBoxLayout(export_group)

        export_row = QHBoxLayout()
        self.export_path_edit = QLineEdit("output")
        self.export_button = QPushButton("Export Result Workbook")
        self.export_button.clicked.connect(self._export_result_workbook)
        export_row.addWidget(QLabel("Output Folder"))
        export_row.addWidget(self.export_path_edit)
        export_row.addWidget(self.export_button)

        export_layout.addLayout(export_row)
        layout.addWidget(export_group)

        preview_group = QGroupBox("Export Status")
        preview_layout = QVBoxLayout(preview_group)

        self.export_status_text = QPlainTextEdit()
        self.export_status_text.setReadOnly(True)
        self.export_status_text.setPlainText(
            "No export yet.\n\nRun the scheduler first, then export the workbook."
        )

        preview_layout.addWidget(self.export_status_text)
        layout.addWidget(preview_group)

        layout.addStretch(1)
        return panel

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

        template_row = QHBoxLayout()
        self.template_path_edit = QLineEdit("data")
        self.template_button = QPushButton("Download Fresh Workbook Set")
        self.template_button.clicked.connect(self._create_template_from_ui)
        template_row.addWidget(self.template_path_edit)
        template_row.addWidget(self.template_button)

        template_layout.addLayout(template_row)
        layout.addWidget(template_group)

        workbook_group = QGroupBox("Workbook Validation")
        workbook_layout = QVBoxLayout(workbook_group)

        self.validate_button = QPushButton("Validate Current Workbook")
        self.validate_button.clicked.connect(self._validate_current_workbook)
        workbook_layout.addWidget(self.validate_button)

        self.data_status_text = QPlainTextEdit()
        self.data_status_text.setReadOnly(True)
        self.data_status_text.setPlainText(
            "This app now uses a 3-workbook input set:\n\n"
            "- District Facility Data.xlsx\n"
            "- Cleaning Planning Assumptions.xlsx\n"
            "- Summer Scheduler Run Input.xlsx\n\n"
            "Use Download Fresh Workbook Set to create clean starter files."
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
        folder_path = QFileDialog.getExistingDirectory(
        self,
        "Select Workbook Set Folder",
        str(Path.cwd()),
    )
        if folder_path:
            self.workbook_path_edit.setText(folder_path)
            self.template_path_edit.setText(folder_path)
            self._validate_current_workbook()

    def _sync_carpet_toggle_state(self) -> None:
        carpet_on = self.include_carpet_check.isChecked()
        self.general_can_do_carpet_check.setEnabled(carpet_on)

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
            raw_path = "data"

        base_dir = Path(__file__).resolve().parent
        path_obj = Path(raw_path)

        if path_obj.is_absolute():
            return str(path_obj)

        return str((base_dir / path_obj).resolve())

    def _load_defaults_into_form(self) -> None:
        try:
            workbook_path = self._resolve_workbook_path()
            errors = validate_workbook(workbook_path)
            if errors:
                raise ValueError("\n".join(errors))

            settings = load_settings(workbook_path)
        except Exception as exc:
            self._show_error("Load Error", f"Could not load workbook defaults.\n\n{exc}")
            return

        self.schedule_name_edit.setText(settings.schedule_name)
        self.schedule_start_date_edit.setText(settings.schedule_start_date)
        self.current_day_edit.setText(str(settings.current_day))
        self.target_end_day_edit.setText(str(settings.target_end_day))
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
        self.data_status_text.setPlainText(f"Workbook loaded:\n{workbook_path}")

    def _apply_form_overrides_to_settings(self) -> None:
        self.settings.schedule_name = (
            self.schedule_name_edit.text().strip() or self.settings.schedule_name
        )
        self.settings.schedule_start_date = (
            self.schedule_start_date_edit.text().strip() or self.settings.schedule_start_date
        )
        self.settings.current_day = int(self.current_day_edit.text().strip())
        self.settings.target_end_day = int(self.target_end_day_edit.text().strip())
        self.settings.work_on_weekends = self.work_on_weekends_check.isChecked()

        self.settings.include_deep_clean = self.include_deep_clean_check.isChecked()
        self.settings.include_strip = self.include_strip_check.isChecked()
        self.settings.include_wax = self.include_wax_check.isChecked()
        self.settings.include_carpet = self.include_carpet_check.isChecked()
        self.settings.include_exterior = self.include_exterior_check.isChecked()

        self.settings.scheduled_shift_hours_per_day = float(
            self.shift_hours_edit.text().strip()
        )
        self.settings.lunch_hours_per_day = float(self.lunch_hours_edit.text().strip())
        self.settings.break_hours_per_day = float(self.break_hours_edit.text().strip())
        self.settings.setup_hours_per_day = float(self.setup_hours_edit.text().strip())
        self.settings.cleanup_hours_per_day = float(
            self.cleanup_hours_edit.text().strip()
        )
        self.settings.productive_hours_per_staff_per_day = float(
            self.productive_hours_edit.text().strip()
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
            target_end_day = int(self.target_end_day_edit.text().strip())

            cleaning_staff = int(self.override_cleaning_staff_edit.text().strip())
            carpet_staff = int(self.override_carpet_staff_edit.text().strip())
            outside_help = int(self.override_outside_help_edit.text().strip())
            absences = int(self.override_absences_edit.text().strip())

            if mode == "Global":
                day_start = 1
                day_end = target_end_day
                label = f"Global Days 1-{target_end_day}"
            elif mode == "Weekly":
                week_num = int(self.override_anchor_edit.text().strip())
                week_span = int(self.override_length_edit.text().strip())
                day_start = ((week_num - 1) * 5) + 1
                day_end = day_start + (week_span * 5) - 1
                label = f"Week {week_num} for {week_span} week(s) -> Days {day_start}-{day_end}"
            else:
                start_day = int(self.override_anchor_edit.text().strip())
                day_span = int(self.override_length_edit.text().strip())
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

    def _populate_summary(self) -> None:
        if not self.result or not self.settings:
            return

        rec = self.result.recommendation

        finish_label = format_workday_label(
            self.settings.schedule_start_date,
            self.result.finish_day,
            self.settings.work_on_weekends,
        )
        target_label = format_workday_label(
            self.settings.schedule_start_date,
            self.result.target_end_day,
            self.settings.work_on_weekends,
        )
        current_label = format_workday_label(
            self.settings.schedule_start_date,
            self.result.current_day,
            self.settings.work_on_weekends,
        )

        target_end_date = self.target_end_date_edit.text().strip()
        paid_holidays = int(self.paid_holidays_edit.text() or "0")

        total_workdays = None
        if target_end_date:
            try:
                total_workdays = calculate_workdays(
                    self.settings.schedule_start_date,
                    target_end_date,
                    self.settings.work_on_weekends,
                    paid_holidays,
                )
            except Exception:
                total_workdays = None

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
                    ("Paid Holidays:", str(paid_holidays)),
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
            self.days_table.setItem(row, 9, self._make_table_item(day.status_note))

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
        )

        detail_lines = [
            day_label,
            f"Active School: {day.active_school_name or 'None'}",
            f"Day Note: {day.status_note or 'None'}",
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
        self.day_outlook_text.setPlainText(format_time_blocks_for_text(day, self.settings))

    def _export_result_workbook(self) -> None:
        if not self.result or not self.settings:
            self._show_error("Nothing To Export", "Run the scheduler first.")
            return

        try:
            output_folder = self.export_path_edit.text().strip() or "output"
            file_path = export_result_workbook(self.result, self.settings, output_folder)
            self.export_status_text.setPlainText(
                f"Export complete:\n{file_path}"
            )
            self._show_info("Export Complete", file_path)
        except Exception as exc:
            self.export_status_text.setPlainText(str(exc))
            self._show_error("Export Error", str(exc))

    def _create_template_from_ui(self) -> None:
        try:
            folder_path = self.template_path_edit.text().strip() or "data"
            created = create_input_template(folder_path)

            base_dir = Path(created).resolve().parent
            district_file = base_dir / "District Facility Data.xlsx"
            assumptions_file = base_dir / "Cleaning Planning Assumptions.xlsx"
            run_input_file = base_dir / "Summer Scheduler Run Input.xlsx"

            self.workbook_path_edit.setText(str(base_dir))
            self.template_path_edit.setText(str(base_dir))

            self.data_status_text.setPlainText(
                "Workbook set created:\n\n"
                f"- {district_file.name}\n"
                f"- {assumptions_file.name}\n"
                f"- {run_input_file.name}\n\n"
                f"Folder:\n{base_dir}"
            )

            self._show_info(
                "Workbook Set Created",
                "Created these files:\n\n"
                f"{district_file.name}\n"
                f"{assumptions_file.name}\n"
                f"{run_input_file.name}\n\n"
                f"Folder:\n{base_dir}"
            )
        except Exception as exc:
            self._show_error("Template Error", str(exc))

    def _validate_current_workbook(self) -> None:
        workbook_path = self._resolve_workbook_path()
        errors = validate_workbook(workbook_path)

        base_dir = Path(workbook_path).resolve()
        if base_dir.is_file():
            base_dir = base_dir.parent

        expected_files = [
            "District Facility Data.xlsx",
            "Cleaning Planning Assumptions.xlsx",
            "Summer Scheduler Run Input.xlsx",
        ]

        if errors:
            self.data_status_text.setPlainText(
                "Workbook set validation failed:\n\n"
                + "\n".join(f"- {e}" for e in errors)
                + "\n\nExpected files:\n"
                + "\n".join(f"- {name}" for name in expected_files)
                + f"\n\nChecked folder:\n{base_dir}"
            )
        else:
            self.data_status_text.setPlainText(
                "Workbook set looks valid:\n\n"
                + "\n".join(f"- {name}" for name in expected_files)
                + f"\n\nFolder:\n{base_dir}"
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