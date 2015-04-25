__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis


class SpiT(Visualization):
    """

    :var threshold: bool
    :var thresholdValues: tuple of floats, [0] indicating the starting threshold and [1] indicating the ending threshold
    :var x_axis: XAxis enum

    :var parameters: dict, the present keys indicate which parameters should be available for the user
    """

    def __init__(self):
        self.title = "SPI(t)"
        self.description = ""
        self.parameters = {"threshold": True,
                           "x_axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}