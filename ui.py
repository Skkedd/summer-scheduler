import sys
from pathlib import Path

from PySide6.QtCore import Qt
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
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from data_loader import (
    load_progress,
    load_rooms,
    load_settings,
    load_staffing,
    validate_workbook,
)
from exporter import create_input_template, export_result_workbook
from scheduler import run_scheduler


def fmt_hours(value: float) -> str:
    return f"{value:.2f}"


def yes_no(value: bool) -> str:
    return "Yes" if value else "No"


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

    QPlainTextEdit, QLineEdit, QTableWidget, QComboBox {
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

    QPlainTextEdit, QLineEdit, QTableWidget, QComboBox {
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

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)

    def set_value(self, value: str) -> None:
        self.value_label.setText(value)


class SchedulerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Summer Scheduler")
        self.resize(1500, 900)

        self.result = None
        self.settings = None
        self.rooms = []
        self.schools = []
        self.staffing_days = []
        self.progress_entries = []

        self.current_theme = "dark"

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
        outer_layout.addWidget(splitter)

        left_panel = self._build_left_panel()
        right_panel = self._build_run_right_panel()

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([360, 1100])

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
        self.theme_combo.currentTextChanged.connect(self._on_theme_changed)
        title_row.addWidget(QLabel("Theme"))
        title_row.addWidget(self.theme_combo)

        layout.addLayout(title_row)

        workbook_group = QGroupBox("Workbook")
        workbook_layout = QVBoxLayout(workbook_group)

        workbook_row = QHBoxLayout()
        self.workbook_path_edit = QLineEdit("data/summer_scheduler_template.xlsx")
        self.browse_workbook_button = QPushButton("Browse")
        self.browse_workbook_button.clicked.connect(self._browse_workbook)
        workbook_row.addWidget(self.workbook_path_edit)
        workbook_row.addWidget(self.browse_workbook_button)

        workbook_layout.addLayout(workbook_row)

        template_row = QHBoxLayout()
        self.make_template_button = QPushButton("Create Template")
        self.make_template_button.clicked.connect(self._create_template_from_ui)
        self.reload_button = QPushButton("Reload Workbook Defaults")
        self.reload_button.clicked.connect(self._load_defaults_into_form)
        template_row.addWidget(self.make_template_button)
        template_row.addWidget(self.reload_button)

        workbook_layout.addLayout(template_row)
        layout.addWidget(workbook_group)

        run_group = QGroupBox("Run Settings")
        run_form = QFormLayout(run_group)

        self.schedule_name_edit = QLineEdit()
        self.current_day_edit = QLineEdit()
        self.target_end_day_edit = QLineEdit()

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
        run_form.addRow("Current Day", self.current_day_edit)
        run_form.addRow("Target End Day", self.target_end_day_edit)
        run_form.addRow("", self.include_deep_clean_check)
        run_form.addRow("", self.include_strip_check)
        run_form.addRow("", self.include_wax_check)
        run_form.addRow("", self.include_carpet_check)
        run_form.addRow("", self.include_exterior_check)
        run_form.addRow("", self.general_can_do_carpet_check)

        layout.addWidget(run_group)

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

        button_row = QHBoxLayout()

        self.run_button = QPushButton("Run Scheduler")
        self.run_button.setObjectName("primaryButton")
        self.run_button.clicked.connect(self.run_scheduler_from_ui)

        button_row.addWidget(self.run_button)
        layout.addLayout(button_row)

        info_group = QGroupBox("What This Tab Does")
        info_layout = QVBoxLayout(info_group)

        notes = QLabel(
            "Run tab = scenario controls + fast answer.\n\n"
            "If the schedule misses, you should be able to see that here without digging.\n"
            "Schedule details live on the Schedule tab.\n"
            "Exports live on the Export tab.\n"
            "Workbook/template tools live on the Data tab."
        )
        notes.setWordWrap(True)
        notes.setObjectName("mutedLabel")

        info_layout.addWidget(notes)
        layout.addWidget(info_group)
        layout.addStretch(1)

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

        title_row.addWidget(page_title)
        title_row.addStretch(1)
        title_row.addWidget(self.status_chip)

        layout.addLayout(title_row)

        cards_row = QHBoxLayout()
        cards_row.setSpacing(12)

        self.finish_day_card = SummaryCard("Projected Finish Day", "-")
        self.deadline_card = SummaryCard("Deadline Met", "-")
        self.backlog_card = SummaryCard("Remaining Backlog", "-")
        self.recommendation_card = SummaryCard("Recommendation Status", "-")

        cards_row.addWidget(self.finish_day_card)
        cards_row.addWidget(self.deadline_card)
        cards_row.addWidget(self.backlog_card)
        cards_row.addWidget(self.recommendation_card)

        layout.addLayout(cards_row)

        summary_group = QGroupBox("Run Summary")
        summary_layout = QVBoxLayout(summary_group)

        self.summary_text = QPlainTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)

        summary_layout.addWidget(self.summary_text)
        layout.addWidget(summary_group)

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

        self.days_table = QTableWidget(0, 9)
        self.days_table.setHorizontalHeaderLabels(
            [
                "Day",
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
            8, QHeaderView.ResizeMode.Stretch
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

        detail_splitter.addWidget(top_half)
        detail_splitter.addWidget(bottom_half)
        detail_splitter.setSizes([420, 320])

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
        self.template_path_edit = QLineEdit("data/summer_scheduler_template.xlsx")
        self.template_button = QPushButton("Create Fresh Template")
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
            "Expected workbook sheets:\n"
            "- Setup\n"
            "- Rooms\n"
            "- Staffing\n"
            "- Progress (optional)\n\n"
            "This app now reads Excel workbooks only."
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
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Scheduler Workbook",
            str(Path.cwd()),
            "Excel Workbook (*.xlsx *.xlsm)",
        )
        if file_path:
            self.workbook_path_edit.setText(file_path)
            self._validate_current_workbook()

    def _sync_carpet_toggle_state(self) -> None:
        carpet_on = self.include_carpet_check.isChecked()
        self.general_can_do_carpet_check.setEnabled(carpet_on)

    def _resolve_workbook_path(self) -> str:
        raw_path = self.workbook_path_edit.text().strip()
        if not raw_path:
            raw_path = "data/summer_scheduler_template.xlsx"

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
        self.current_day_edit.setText(str(settings.current_day))
        self.target_end_day_edit.setText(str(settings.target_end_day))

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
        self.settings.current_day = int(self.current_day_edit.text().strip())
        self.settings.target_end_day = int(self.target_end_day_edit.text().strip())

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

            self.result = run_scheduler(
                rooms=self.rooms,
                settings=self.settings,
                staffing_days=self.staffing_days,
                progress_entries=self.progress_entries,
            )

            self._populate_summary()
            self._populate_days_table()
            self.status_chip.setText("Run complete")
            self.export_status_text.setPlainText(
                "Run complete.\n\nGo to Export tab to write an Excel report workbook."
            )
            self.tabs.setCurrentWidget(self.run_tab)

        except Exception as exc:
            self.status_chip.setText("Run failed")
            error_text = f"{type(exc).__name__}: {exc}"
            self.summary_text.setPlainText(error_text)
            self.day_detail_text.setPlainText(error_text)
            self.days_table.setRowCount(0)
            self.worklog_table.setRowCount(0)
            self.export_status_text.setPlainText(error_text)
            self._show_error("Scheduler Error", error_text)

    def _populate_summary(self) -> None:
        if not self.result:
            return

        rec = self.result.recommendation

        self.finish_day_card.set_value(f"Day {self.result.finish_day}")
        self.deadline_card.set_value(yes_no(self.result.met_deadline))
        self.backlog_card.set_value(f"{fmt_hours(self.result.remaining_backlog_hours)} hrs")
        self.recommendation_card.set_value(rec.status_label)

        summary_lines = [
            f"Schedule Name: {self.result.schedule_name}",
            f"Current Day: {self.result.current_day}",
            f"Target End Day: {self.result.target_end_day}",
            f"Projected Finish Day: {self.result.finish_day}",
            f"Deadline Met: {yes_no(self.result.met_deadline)}",
            "",
            f"Total Planned Hours: {fmt_hours(self.result.total_planned_hours)}",
            f"Completed Before Rerun: {fmt_hours(self.result.completed_hours_before_run)}",
            f"Remaining At Start: {fmt_hours(self.result.remaining_hours_at_start)}",
            f"Total Used Hours: {fmt_hours(self.result.total_used_hours)}",
            f"Remaining Backlog: {fmt_hours(self.result.remaining_backlog_hours)}",
            "",
            "Recommendation",
            f"Status: {rec.status_label}",
            f"Bottleneck: {rec.bottleneck_type}",
            f"Available-Now Backlog At Start: {fmt_hours(rec.available_backlog_hours)}",
            f"Blocked-Later Backlog At Start: {fmt_hours(rec.blocked_backlog_hours)}",
            f"Total Backlog At Start: {fmt_hours(rec.total_backlog_hours)}",
            f"Capacity To Deadline: {fmt_hours(rec.capacity_to_deadline_hours)}",
            f"Extra Staff-Days Needed: {fmt_hours(rec.extra_staff_days_needed)}",
            f"Action: {rec.recommended_action}",
            "",
            "Daily Time Model",
            f"Shift Hours: {fmt_hours(self.settings.scheduled_shift_hours_per_day)}",
            f"Lunch Hours: {fmt_hours(self.settings.lunch_hours_per_day)}",
            f"Break Hours: {fmt_hours(self.settings.break_hours_per_day)}",
            f"Setup Hours: {fmt_hours(self.settings.setup_hours_per_day)}",
            f"Cleanup Hours: {fmt_hours(self.settings.cleanup_hours_per_day)}",
            f"Productive Hours: {fmt_hours(self.settings.productive_hours_per_staff_per_day)}",
        ]

        self.summary_text.setPlainText("\n".join(summary_lines))

    def _make_table_item(self, value: str) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        return item

    def _populate_days_table(self) -> None:
        self.days_table.setRowCount(0)
        self.worklog_table.setRowCount(0)
        self.day_detail_text.clear()

        if not self.result:
            return

        days = self.result.days
        self.days_table.setRowCount(len(days))

        for row, day in enumerate(days):
            self.days_table.setItem(row, 0, self._make_table_item(str(day.day)))
            self.days_table.setItem(
                row, 1, self._make_table_item(day.active_school_name or "")
            )
            self.days_table.setItem(row, 2, self._make_table_item(str(day.general_staff)))
            self.days_table.setItem(row, 3, self._make_table_item(str(day.carpet_staff)))
            self.days_table.setItem(
                row, 4, self._make_table_item(fmt_hours(day.daily_capacity))
            )
            self.days_table.setItem(
                row, 5, self._make_table_item(fmt_hours(day.used_capacity))
            )
            self.days_table.setItem(
                row, 6, self._make_table_item(fmt_hours(day.unused_capacity))
            )
            self.days_table.setItem(
                row, 7, self._make_table_item(str(len(day.work_log)))
            )
            self.days_table.setItem(row, 8, self._make_table_item(day.status_note))

        if days:
            self.days_table.selectRow(0)
            self._populate_day_detail()

    def _populate_day_detail(self) -> None:
        if not self.result:
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
            self.worklog_table.setItem(
                row, 6, self._make_table_item(fmt_hours(item.hours_done))
            )
            self.worklog_table.setItem(row, 7, self._make_table_item(item.note or ""))

        detail_lines = [
            f"Day {day.day}",
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

    def _export_result_workbook(self) -> None:
        if not self.result:
            self._show_error("Nothing To Export", "Run the scheduler first.")
            return

        try:
            output_folder = self.export_path_edit.text().strip() or "output"
            file_path = export_result_workbook(self.result, output_folder)
            self.export_status_text.setPlainText(
                f"Export complete:\n{file_path}"
            )
            self._show_info("Export Complete", file_path)
        except Exception as exc:
            self.export_status_text.setPlainText(str(exc))
            self._show_error("Export Error", str(exc))

    def _create_template_from_ui(self) -> None:
        try:
            file_path = self.template_path_edit.text().strip() or "data/summer_scheduler_template.xlsx"
            created = create_input_template(file_path)
            self.workbook_path_edit.setText(created)
            self.data_status_text.setPlainText(f"Template created:\n{created}")
            self._show_info("Template Created", created)
        except Exception as exc:
            self._show_error("Template Error", str(exc))

    def _validate_current_workbook(self) -> None:
        workbook_path = self._resolve_workbook_path()
        errors = validate_workbook(workbook_path)

        if errors:
            self.data_status_text.setPlainText(
                "Workbook validation failed:\n\n" + "\n".join(f"- {e}" for e in errors)
            )
        else:
            self.data_status_text.setPlainText(
                f"Workbook looks valid:\n{workbook_path}"
            )


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Summer Scheduler")

    font = QFont("Segoe UI", 10)
    app.setFont(font)

    window = SchedulerWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()