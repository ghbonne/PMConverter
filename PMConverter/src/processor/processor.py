__author__ = 'PM Group 8'

from src.convert.XLSXparser import XLSXParser
from src.convert.XMLparser import XMLParser
from src.visual.linechart import LineChart
from src.visual.piechart import PieChart


class Processor(object):

    def __init__(self):
        self.visualizations = []
        self.file_parsers = []

    def convert(self, parser_from, parser_to, file_path_from, file_path_to):
        pass

    def create_all_file_parsers(self):
        self.file_parsers.append(XLSXParser())
        self.file_parsers.append(XMLParser())

    def create_all_visualizations(self):
        self.visualizations.append(LineChart())
        self.visualizations.append(PieChart())