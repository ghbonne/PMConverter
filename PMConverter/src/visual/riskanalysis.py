__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import LevelOfDetail


class RiskAnalysis(Visualization):

    """
    instelling: workpackage or activites (level of detail)
    data: naam, baseline duration, optimistic, most probable, pessimistic

    """

    def __init__(self):
        self.title = "Risk analysis"
        self.description = ""
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES]}