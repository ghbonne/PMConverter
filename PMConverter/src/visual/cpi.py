__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class CPI(Visualization):
    """
    enkel voor tracking overview!
    cpi apart met threshold
    instelling: threshold,
                tracking period of datum
    :var threshold
    :var x-axis
    """

    def __init__(self):
        self.title = "CPI"
        self.description = ""
        self.parameters = {"threshold": True,
                           "x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}