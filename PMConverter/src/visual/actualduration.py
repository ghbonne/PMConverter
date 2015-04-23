__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import LevelOfDetail, DataType


class ActualDuration(Visualization):
    """
    instelling: level of detail,
                relatief of absoluut (indien relatief, extra berekening nodig)
    data: baseline duration, actual duration, percentage completed
    """

    def __init__(self):
        self.title = "PD vs AD"
        self.description = ""
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES],
                           "data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}