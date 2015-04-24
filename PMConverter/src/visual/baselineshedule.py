__author__ = 'Eveline'
from visual.visualization import Visualization


class BaselineSchedule(Visualization):
    """
    type gantt chart
    bereik? sheet: baseline schedule, kolom B
            baseline timing

    """

    def __init__(self):
        self.title = "Baseline Schedule"
        self.description = ""
        self.parameters = None