__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class SpiT(Visualization):
    """

    :var threshold
    :var x-axis: type XAxis
    """

    def __init__(self):
        self.title = "SPI(t)"
        self.description = ""
        self.parameters = {"threshold": True,
                           "x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}