__author__ = 'Eveline'

from enum import Enum

from visual.charts.twoDimVisualization import TwoDimChart


class ColumnSubType(Enum):
    stacked = "stacked"
    percent_stacked = "percent_stacked"


class ColumnChart(TwoDimChart):

    """
    :var title: title of the chart
    :var labels: labels on the x and y-axis. Format: [x-axis,y-axis]
    :var data_series: data
        Format: [
                    [heading1:String, #first dataset
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

    def __init__(self, title, labels, data_series, subtype=None):
        self.title = title
        self.labels = labels
        self.data_series = data_series
        if subtype and not isinstance(subtype, ColumnSubType):
            raise TypeError("Subtype has to be of type ENUM SubType")
        super().__init__("column", subtype)

    def draw(self, workbook):
        """
        add a linechart visualisation to the Excel-file
        :param workbook: xlsxworkbook
        :return:
        """
        super().draw(workbook)
