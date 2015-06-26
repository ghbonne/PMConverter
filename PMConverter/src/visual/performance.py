__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis, ExcelVersion
from visual.charts.linechart import LineChart


class Performance(Visualization):

    """
    Implements drawings for performance chart (CPI,SPI(t)) (type = Line chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user
    :var supported: list of ExcelVersion, containing the version that are supported

    Settings:
    :var x_axis: XAxis, x-axis of the chart can be expressed in status dates or in tracking periods
    """

    def __init__(self):
        self.title = "CPI,SPI"
        self.description = "A line graph showing the Cost Performance Index and Schedule Performance Index of the project, based on the available tracking periods."
        self.parameters = {"x_axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}
        self.x_axis = None
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

        data_series = [
            ["CPI",
             names,
             ['Tracking Overview', 2, 32, (1+tp_size), 32]
             ],
            ["SPI(t)",
             names,
             ['Tracking Overview', 2, 38, (1+tp_size), 38]
             ],
        ]

        labels = ["", ""]
        chart = LineChart(self.title, labels, data_series)

        size = {'width': 750, 'height': 500}
        chart.draw(workbook, chartsheet, 'A1', None, size)

    """
    Private methods
    """
    def calculate_values(self, workbook, worksheet, project_object):
        """

        :param workbook: Workbook
        :param worksheet: Worksheet
        :param project_object: ProjectObject
        """
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        worksheet.write('AF2', 'SPI', header)
        worksheet.write('AG2', 'CPI', header)

        counter = 2

        for tp in project_object.tracking_periods:
            worksheet.write_number(counter, 31, tp.spi, calculation)
            worksheet.write_number(counter, 32, tp.cpi, calculation)
            counter += 1
