__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class SpiTvsPfactor(Visualization):
    """
    :var x_axis: XAxis enum

    ranges aanpassen!!
    """

    def __init__(self):
        self.title = "SPI(t), p-factor"
        self.description = ""
        self.parameters = {"x_axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}