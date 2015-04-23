__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import DataType, LevelOfDetail


class ActualCost(Visualization):
    """
    zie actual duration
    """

    def __init__(self):
        self.title = "PC vs AC"
        self.description = ""
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES],
                           "data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}