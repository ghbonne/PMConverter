__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import LevelOfDetail
from visual.charts.barchart import BarChart
from convert.XLSXparser import XLSXParser


class RiskAnalysis(Visualization):

    """
    :var level_of_detail

    data: naam, baseline duration, optimistic, most probable, pessimistic
    """

    def __init__(self):
        self.title = "Risk analysis"
        self.description = ""
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES]}

    def draw(self, workbook, worksheet, project_object):
        if not self.level_of_detail:
            raise Exception("Please first set var level_of_detail")

        activities = project_object.activities
        names = "=("
        optimistic = "=("
        most_probable = "=("
        pessimistic = "=("
        height = 0  # calculate 20 pixels per element in barchart

        i = 0
        start = 3
        while i < len(activities):
            if len(activities[i].wbs_id) == 2:  # work package
                if self.level_of_detail == LevelOfDetail.WORK_PACKAGES:
                    names += "'Risk Analysis'!$B$" + str(start+i) + ","
                    optimistic += "'Risk Analysis'!$E$" + str(start+i) + ","
                    most_probable += "'Risk Analysis'!$F$" + str(start+i) + ","
                    pessimistic += "'Risk Analysis'!$G$" + str(start+i) + ","
                    height += 20
                i += 1
            elif len(activities[i].wbs_id) == 3:  # activity level
                if self.level_of_detail == LevelOfDetail.ACTIVITIES:
                    names += "'Risk Analysis'!$B$" + str(start+i) + ","
                    optimistic += "'Risk Analysis'!$E$" + str(start+i) + ","
                    most_probable += "'Risk Analysis'!$F$" + str(start+i) + ","
                    pessimistic += "'Risk Analysis'!$G$" + str(start+i) + ","
                    height += 20
                i += 1
            else:  # 1 = project level, >3 not supported yet
                i += 1

        # remove last ';' and add ')'
        names = names[:-1] + ")"
        optimistic = optimistic[:-1] + ")"
        most_probable = most_probable[:-1] + ")"
        pessimistic = pessimistic[:-1] + ")"

        data_series = [
            ["Optimistic",
             names,
             optimistic
             ],
            ["Most probable",
             names,
             most_probable
             ],
            ["Pessimistic",
             names,
             pessimistic
             ]
        ]

        chart = BarChart(self.title, ["Hours", self.level_of_detail.value], data_series)

        options = {'height': height}
        chart.draw(workbook, worksheet, 'I1', None, options)