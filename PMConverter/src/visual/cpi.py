__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis, ExcelVersion
from visual.charts.linechart import LineChart

class CPI(Visualization):

    """
    Implements drawings for cost-value metrics chart (AC, EV, PV) (type = Line chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user
    :var supported: list of ExcelVersion, containing the version that are supported

    Settings:
    :var x_axis: XAxis, x-axis of the chart can be expressed in status dates or in tracking periods
    :var threshold: bool
    :var thresholdValues: tuple of floats, [0] indicating the starting threshold and [1] indicating the ending threshold
    """

    def __init__(self):
        self.title = "CPI"
        self.description = ""
        self.parameters = {"threshold": True,
                           "x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}
        self.x_axis = None
        self.threshold = None
        self.support = [ExcelVersion.EXTENDED, ExcelVersion.BASIC]

    def draw(self, workbook, worksheet, project_object, excel_version): #todo probleem met percenten
        if not self.x_axis:
            raise Exception("Please first set var x_axis")

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
                 ['Tracking Overview', 2, 10, (1+tp_size), 10]
                 ],
                ["threshold " + str(self.threshold),
                  names,
                  ['Tracking Overview', 2, 34, (1+tp_size), 34],
                ]
            ]
        else:
            data_series = [
                ["CPI",
                 names,
                 ['Tracking Overview', 2, 10, (1+tp_size), 10]
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

        worksheet.write('AI2', 'CPI threshold', header)

        start = 2
        if self.threshold[0] == self.threshold[1]:
            for i in range(0, tp_size):
                worksheet.write(start + i, 34, self.threshold[0], calculation)
        else:
            if self.threshold[0] > self.threshold[1]:
                value = (self.threshold[0] - self.threshold[1])/(tp_size - 1)
            else:
                value = (self.threshold[1] - self.threshold[0])/(tp_size - 1)
            for i in range(0, tp_size):
                worksheet.write(start + i, 34, self.threshold[0] + (i * value), calculation)
