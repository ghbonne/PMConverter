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
        self.data_type = None

    def draw(self, workbook, worksheet, project_object):
        if not self.data_type:
            raise Exception("Please first set var data_type")

        res_size = len(project_object.resources)

        data_series = [
            ['Resources', 2, 1, (1+res_size), 1],
            ['Resources', 2, 7, (1+res_size), 7]
        ]

        if self.data_type == DataType.RELATIVE:
            relative = True
        else:
            relative = False

        chart2 = PieChart('Resources', data_series, relative)
        size = {'height': 150 + res_size*20}
        options = {'x_offset': 25, 'y_offset': 10}
        chart2.draw(workbook, worksheet, 'I1', options, size)
