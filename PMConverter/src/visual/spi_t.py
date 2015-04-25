__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis
from visual.charts.linechart import LineChart


class SpiT(Visualization):
    """

    :var threshold
    :var x-axis: type XAxis
    """

    def __init__(self):
        self.title = "SPI(t)"
        self.description = ""
        self.parameters = {"threshold": True,
                           "x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}
        self.x_axis = None
        self.threshold = None

    def draw(self, workbook, worksheet, project_object):
        #todo: threshold
        if not self.x_axis:
            raise Exception("Please first set var x_axis")

        chartsheet = workbook.add_worksheet(self.title)

        # number tracking periods
        tp_size = len(project_object.tracking_periods)

        if self.threshold and self.threshold != (0, 0):
            self.calculate_threshold(workbook, worksheet, tp_size)

        # values for x_axis
        if self.x_axis == XAxis.TRACKING_PERIOD:
            names = ['Tracking Overview', 2, 0, (1+tp_size), 0]
        elif self.x_axis == XAxis.DATE:
            names = ['Tracking Overview', 2, 2, (1+tp_size), 2]

        data_series = [
            ["SPI(t)",
             names,
             ['Tracking Overview', 2, 12, (1+tp_size), 12]
             ],
        ]

        labels = ["", ""]
        chart = LineChart(self.title, labels, data_series)

        size = {'width': 750, 'height': 500}
        chart.draw(workbook, chartsheet, 'A1', None, size)

    def calculate_threshold(self, workbook, worksheet, tp_size):
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        worksheet.write('AD4', 'SPI(t) threshold', header)

        start = 2
        if self.threshold[0] == self.threshold[1]:
            for i in range(0, tp_size):
                worksheet.write(start + i, 29, self.threshold[0], calculation)
        else:
            if self.threshold[0] > self.threshold[1]:
                value = (self.threshold[0] - self.threshold[1])/tp_size
            else:
                value = (self.threshold[1] - self.threshold[0])/tp_size
            for i in range(0, tp_size):
                worksheet.write(start + i, 29, value, calculation)
