from scheduler import run_schedule  # or whatever your entry function is called


class ScenarioInput:
    def __init__(
        self,
        rooms,
        start_day,
        end_day,
        cleaning_staff_by_day,
        carpet_staff_by_day=None,
        run_cleaning=True,
        run_carpet=False,
    ):
        self.rooms = rooms
        self.start_day = start_day
        self.end_day = end_day
        self.cleaning_staff_by_day = cleaning_staff_by_day
        self.carpet_staff_by_day = carpet_staff_by_day or {}
        self.run_cleaning = run_cleaning
        self.run_carpet = run_carpet


def run_scenario(scenario: ScenarioInput):
    """
    Single entry point for UI → engine.
    This is the ONLY thing UI should call.
    """

    result = run_schedule(
        rooms=scenario.rooms,
        start_day=scenario.start_day,
        end_day=scenario.end_day,
        cleaning_staff_by_day=scenario.cleaning_staff_by_day,
        carpet_staff_by_day=scenario.carpet_staff_by_day,
        run_cleaning=scenario.run_cleaning,
        run_carpet=scenario.run_carpet,
    )

    return result