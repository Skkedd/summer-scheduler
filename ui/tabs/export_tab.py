from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel


class ExportTab(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Export Tools (coming next phase)"))

        self.setLayout(layout)