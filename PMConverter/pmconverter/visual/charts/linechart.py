__author__ = 'Project Management Group 8 - 2015 Ghent University'

from pmconverter.visual.charts import TwoDimChart


class LineChart(TwoDimChart):

    """
    :var title: title of the chart
    :var labels: labels on the x and y-axis. Format: [x-axis,y-axis]
    :var data_series: data
        Format: [
                    [heading1:String, #first line
                        [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int], #x-values
                        [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int]  #y-values
                    ],
                    [heading2:String,
                        [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int], #x-values
                        [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int]  #y-values
                    ],
                    ...
                ]
    """

    def __init__(self, title, labels, data_series):
        self.title = title
        self.labels = labels
        self.data_series = data_series
        super().__init__("line")

    def draw(self, workbook, worksheet, position, options=None, size=None):
        """
        add a linechart visualisation to the Excel-file
        :param workbook: xlsxworkbook
        :return:
        """
        super().draw(workbook, worksheet, position, options, size)

