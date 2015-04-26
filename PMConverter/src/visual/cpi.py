__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis, ExcelVersion
from visual.charts.linechart import LineChart

class CPI(Visualization):

    """
    Implements drawings for cost-value metrics chart (AC, EV, PV) (type = Line chart)

    :var x_axis: XAxis, x-axis of the chart can be expressed in status dates or in tracking periods
    :var threshold: tuple (x,y), a linear threshold defined by a line through the values x and y
    """

    def __init__(self):
        self.title = "CPI"
        self.description = ""
        self.parameters = {"threshold": True,
                           "x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}
        self.x_axis = None
        self.threshold = None
        self.support = [ExcelVersion.EXTENDED, ExcelVersion.BASIC]

    def draw(self, workbook, worksheet, project_object, excel_version):
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

        if self.threshold and self.threshold != (0, 0):
            self.calculate_threshold(workbook, worksheet, tp_size)
            data_series = [
                ["CPI",
                 names,
                 ['Tracking Overview', 2, 36, (1+tp_size), 36]
                 ],
                ["threshold " + str(self.threshold),
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
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        worksheet.write('AL2', 'CPI threshold', header)

        start = 2
        if self.threshold[0] == self.threshold[1]:
            for i in range(0, tp_size):
                worksheet.write(start + i, 37, self.threshold[0], calculation)
        else:
            if self.threshold[0] > self.threshold[1]:
                value = (self.threshold[0] - self.threshold[1])/(tp_size - 1)
            else:
                value = (self.threshold[1] - self.threshold[0])/(tp_size - 1)
            for i in range(0, tp_size):
                worksheet.write(start + i, 37, self.threshold[0] + (i * value), calculation)

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
            worksheet.write_number(counter, 36, tp.cpi/100, calculation)
            counter += 1