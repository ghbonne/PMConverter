__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import DataType


class CV(Visualization):
    """
    instelling: absoluut of relatief
    """

    def __init__(self):
        self.title = "CV"
        self.description = ""
        self.parameters = {"data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}
