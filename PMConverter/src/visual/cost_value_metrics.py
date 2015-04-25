__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class CostValueMetrics(Visualization):
    """
    :var x_axis: XAxis enum

    data: workpackage name, AC, EV, PV
    """

    def __init__(self):
        self.title = "AC, EV, PV"
        self.description = ""
        self.parameters = {"x_axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}