__author__ = 'Eveline'

from visual.twoDimVisualization import TwoDimVisualization
from enum import Enum


class BarSubType(Enum):
    stacked = "stacked"
    percent_stacked = "percent_stacked"


class BarChart(TwoDimVisualization):

    def __init__(self, title, labels, data_series, subtype=None):
        self.title = title
        self.labels = labels
        self.data_series = data_series
        if subtype and not isinstance(subtype, BarSubType):
            raise TypeError("Subtype has to be of type ENUM SubType")
        super().__init__("bar", subtype)

    def visualize(self, workbook):
        """
        add a linechart visualisation to the Excel-file
        :param workbook: xlsxworkbook
        :return:
        """
        super().visualize(workbook)