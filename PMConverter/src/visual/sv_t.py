__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class SvT(Visualization):
    """
    :var x-axis
    """

    def __init__(self):
        self.title = "SV(t)"
        self.description = ""
        self.parameters = {"x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}