__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import XAxis
from visual.charts.linechart import LineChart


class SpiTvsPfactor(Visualization):
    """
    Implements drawings for SPI(t) (type = Line chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user

    Settings:
    :var x_axis: XAxis, x-axis of the chart can be expressed in status dates or in tracking periods
    """

    def __init__(self):
        self.title = "SPI, SPI(t), p-factor"
        self.description = "A line graph showing the Schedule Performance Index, the Schedule Performance Index as (earned schedule / actual duration) and the p-factor of the project, based on the available tracking periods."
        self.parameters = {"x_axis": [XAxis.TRACKING_PERIOD, XAxis.DATE]}
        self.x_axis = None

    def draw(self, workbook, worksheet, project_object):
        #todo: axis ranges aanpassen?
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
            ["SPI(t)",
             names,
             ['Tracking Overview', 2, 31, (1+tp_size), 31]
             ],
            ["SPI",
             names,
             ['Tracking Overview', 2, 33, (1+tp_size), 33]
             ],
            ["p-factor",
             names,
             ['Tracking Overview', 2, 34, (1+tp_size), 34]
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

        worksheet.write('AF2', 'SPI(t)', header)
        worksheet.write('AH2', 'SPI', header)
        worksheet.write('AI2', 'p-factor', header)

        counter = 2

        for tp in project_object.tracking_periods:
            worksheet.write_number(counter, 31, tp.spi_t, calculation)
            worksheet.write_number(counter, 33, tp.spi, calculation)
            worksheet.write_number(counter, 34, tp.p_factor, calculation)
            counter += 1

