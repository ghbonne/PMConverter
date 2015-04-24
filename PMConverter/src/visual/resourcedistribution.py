__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import DataType
from visual.charts.piechart import PieChart


class ResourceDistribution(Visualization):

    """
    :var data_type

    data, name + total cost
    """

    def __init__(self):
        self.title = "Resource costs"
        self.description = ""
        self.parameters = {'data_type': [DataType.ABSOLUTE, DataType.RELATIVE]}

    def draw(self, workbook, worksheet):
        if not self.data_type:
            raise Exception("Please first set var data_type")

        data_series2 = [
            ['Resources', 2, 1, 13, 1],
            ['Resources', 2, 7, 13, 7]
        ]

        if self.data_type == DataType.RELATIVE:
            relative = True
        else:
            relative = False

        chart2 = PieChart('Resources', data_series2, relative)

        chart2.draw(workbook, worksheet, 'I1')