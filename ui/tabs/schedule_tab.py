from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel


class ScheduleTab(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Schedule View (coming next phase)"))

        self.setLayout(layout)