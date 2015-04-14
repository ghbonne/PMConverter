__author__ = 'Eveline'

from visual.twoDimVisualization import TwoDimVisualization
from enum import Enum


class ColumnSubType(Enum):
    stacked = "stacked"
    percent_stacked = "percent_stacked"


class ColumnChart(TwoDimVisualization):

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

    def visualize(self, workbook):
        """
        add a linechart visualisation to the Excel-file
        :param workbook: xlsxworkbook
        :return:
        """
        super().visualize(workbook)
