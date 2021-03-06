__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis
from visual.charts.linechart import LineChart


class CV(Visualization):
    """
    Implements drawings for CV chart (type = Line chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user

    Settings:
    :var x_axis: XAxis, x-axis of the chart can be expressed in status dates or in tracking periods
    """

    def __init__(self):
        self.title = "CV"
        self.description = "Line graph showing the cost variance over the different tracking periods or on an absolute time scale."
        self.parameters = {"x_axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}
        self.x_axis = None

    def draw(self, workbook, worksheet, project_object):
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
            ["CV",
             names,
             ['Tracking Overview', 2, 9, (1+tp_size), 9]
             ],
        ]

        labels = ["", ""]
        chart = LineChart("Cost Variance (CV)", labels, data_series)

        size = {'width': 750, 'height': 500}
        chart.draw(workbook, chartsheet, 'A1', None, size)
