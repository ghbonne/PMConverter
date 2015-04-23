__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import DataType


class ResourceDistribution(Visualization):

    """
    data, name + total cost
    instelling: absoluut of relatief
    """

    def __init__(self):
        self.title = "Resource costs"
        self.description = ""
        self.parameters = {'data_type': [DataType.ABSOLUTE, DataType.RELATIVE]}
