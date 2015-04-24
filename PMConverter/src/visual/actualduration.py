__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import LevelOfDetail, DataType


class ActualDuration(Visualization):
    """
    :var level_of_detail
    :var data_type

    data: baseline duration, actual duration, percentage completed
    """

    def __init__(self):
        self.title = "PD vs AD"
        self.description = ""
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES],
                           "data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}

    def draw(self, workbook, worksheet, project_object):
        if not self.level_of_detail:
            raise Exception("Please first set var level_of_detail")
        if not self.data_type:
            raise Exception("Please first set var data_type")

