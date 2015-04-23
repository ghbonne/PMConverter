__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class CPI(Visualization):
    """
    :var threshold
    :var x-axis
    

    cpi apart met threshold
    """

    def __init__(self):
        self.title = "CPI"
        self.description = ""
        self.parameters = {"threshold": True,
                           "x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}