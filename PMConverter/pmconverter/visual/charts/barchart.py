__author__ = 'Eveline'

from enum import Enum

from pmconverter.visual.charts import TwoDimChart


class BarSubType(Enum):
    stacked = "stacked"
    percent_stacked = "percent_stacked"


class BarChart(TwoDimChart):

    def __init__(self, title, labels, data_series, subtype=None):
        self.title = title
        self.labels = labels
        self.data_series = data_series
        if subtype and not isinstance(subtype, BarSubType):
            raise TypeError("Subtype has to be of type ENUM SubType")
        super().__init__("bar", subtype)

    def draw(self, workbook, worksheet, position, options=None, size=None):
        """
        add a linechart visualisation to the Excel-file
        :param workbook: xlsxworkbook
        :return:
        """
        super().draw(workbook, worksheet, position, options, size)