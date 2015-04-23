__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class CostValueMetrics(Visualization):
    """
    enkel voor tracking overview!
    data: workpackage name, AC, EV, PV
    instelling: x-as: per period of per datum
    """

    def __init__(self):
        self.title = "AC, EV, PV"
        self.description = ""
        self.parameters = {"x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}