__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis, ExcelVersion
from visual.charts.linechart import LineChart


class CostValueMetrics(Visualization):
    """
    Implements drawings for cost-value metrics chart (AC, EV, PV) (type = Line chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user
    :var supported: list of ExcelVersion, containing the version that are supported

    Settings:
    :var x_axis: XAxis, x-axis of the chart can be expressed in status dates or in tracking periods
    """

    def __init__(self):
        self.title = "AC, EV, PV"
        self.description = ""
        self.parameters = {"x-axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}
        self.x_axis = None
        self.support = [ExcelVersion.EXTENDED, ExcelVersion.BASIC]

    def draw(self, workbook, worksheet, project_object, excel_version):
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

        data_series = [
            ["AC",
             names,
             ['Tracking Overview', 2, 5, (1+tp_size), 5]
             ],
            ["EV",
             names,
             ['Tracking Overview', 2, 4, (1+tp_size), 4]
             ],
            ["PV",
             names,
             ['Tracking Overview', 2, 3, (1+tp_size), 3]
             ]
        ]

        labels = ["", ""]
        chart = LineChart(self.title, labels, data_series)

        size = {'width': 750, 'height': 500}
        chart.draw(workbook, chartsheet, 'A1', None, size)