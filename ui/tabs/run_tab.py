from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QLabel, QSpinBox, QHBoxLayout
)

from services.scenario_service import ScenarioInput, run_scenario


class RunTab(QWidget):
    def __init__(self, get_rooms_callback):
        super().__init__()

        self.get_rooms = get_rooms_callback

        layout = QVBoxLayout()

        # --- STAFF INPUT ---
        staff_row = QHBoxLayout()

        self.staff_input = QSpinBox()
        self.staff_input.setMinimum(1)
        self.staff_input.setMaximum(100)

        staff_row.addWidget(QLabel("Cleaning Staff:"))
        staff_row.addWidget(self.staff_input)

        layout.addLayout(staff_row)

        # --- RUN BUTTON ---
        self.run_button = QPushButton("Run Schedule")
        self.run_button.clicked.connect(self.run_clicked)
        layout.addWidget(self.run_button)

        # --- RESULT SUMMARY ---
        self.result_label = QLabel("No run yet")
        layout.addWidget(self.result_label)

        self.setLayout(layout)

    def run_clicked(self):
        rooms = self.get_rooms()

        staff = self.staff_input.value()

        # simple staffing map (can expand later)
        staff_by_day = {day: staff for day in range(1, 60)}

        scenario = ScenarioInput(
            rooms=rooms,
            start_day=1,
            end_day=20,
            cleaning_staff_by_day=staff_by_day,
            run_cleaning=True,
            run_carpet=False,
        )

        result = run_scenario(scenario)

        self.result_label.setText(
            f"Finish Day: {result.projected_finish_day} | Deadline Met: {result.deadline_met}"
        )