__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis
from visual.charts.linechart import LineChart

class CPI(Visualization):

    """
    Implements drawings for cpi chart (type = Line chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user

    Settings:
    :var x_axis: XAxis, x-axis of the chart can be expressed in status dates or in tracking periods
    :var threshold: bool
    :var thresholdValues: tuple of floats, [0] indicating the starting threshold and [1] indicating the ending threshold
    """

    def __init__(self):
        self.title = "CPI"
        self.description = "Line graph showing the cost performance index over the different tracking periods or on an absolute time scale, with the ability to add a threshold line."
        self.parameters = {"threshold": True,
                           "x_axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}
        self.x_axis = None
        self.threshold = None
        self.thresholdValues = None

    def draw(self, workbook, worksheet, project_object):
        if not self.x_axis:
            raise Exception("Please first set var x_axis")

        self.calculate_values(workbook, worksheet, project_object)

        chartsheet = workbook.add_worksheet(self.title)

        # number tracking periods
        tp_size = len(project_object.tracking_periods)

        # values for x_axis
        if self.x_axis == XAxis.TRACKING_PERIOD:
            names = ['Tracking Overview', 2, 0, (1+tp_size), 0]
        elif self.x_axis == XAxis.DATE:
            names = ['Tracking Overview', 2, 2, (1+tp_size), 2]

        if self.threshold:
            self.calculate_threshold(workbook, worksheet, tp_size)
            data_series = [
                ["CPI",
                 names,
                 ['Tracking Overview', 2, 36, (1+tp_size), 36]
                 ],
                ["threshold " + str(self.thresholdValues),
                  names,
                  ['Tracking Overview', 2, 37, (1+tp_size), 37],
                ]
            ]
        else:
            data_series = [
                ["CPI",
                 names,
                 ['Tracking Overview', 2, 36, (1+tp_size), 36]
                 ],
            ]

        labels = ["", ""]
        chart = LineChart(self.title, labels, data_series)

        size = {'width': 750, 'height': 500}
        chart.draw(workbook, chartsheet, 'A1', None, size)

    """
    Private methods
    """
    def calculate_threshold(self, workbook, worksheet, tp_size):
        """
        Calculate the values for the treshold
        :param workbook:
        :param worksheet:
        :param tp_size:
        :return:
        """
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        worksheet.write('AL2', 'CPI threshold', header)

        start = 2
        if self.thresholdValues[0] == self.thresholdValues[1] or tp_size <= 1:
            for i in range(0, tp_size):
                worksheet.write(start + i, 37, self.thresholdValues[0], calculation)
        else:
            value = (self.thresholdValues[1] - self.thresholdValues[0])/(tp_size - 1)
            for i in range(0, tp_size):
                worksheet.write(start + i, 37, self.thresholdValues[0] + (i * value), calculation)

    def calculate_values(self, workbook, worksheet, project_object):
        """

        :param workbook: Workbook
        :param worksheet: Worksheet
        :param project_object: ProjectObject
        """
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        worksheet.write('AK2', 'CPI', header)

        counter = 2

        for tp in project_object.tracking_periods:
            value = tp.cpi
            worksheet.write_number(counter, 36, value, calculation)
            counter += 1

