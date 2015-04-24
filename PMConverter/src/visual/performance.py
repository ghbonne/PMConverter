__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class Performance(Visualization):

    """
    :var x-axis

    CPI/SPI
    """

    def __init__(self):
        self.title = "CPI,SPI"
        self.description = ""
        self.parameters = {"x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}