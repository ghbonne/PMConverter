__author__ = 'PM Group 8'

from convert.XLSXparser import XLSXParser
from convert.XMLparser import XMLParser
from visual.charts.linechart import LineChart
from visual.charts.piechart import PieChart
from visual.baselineshedule import BaselineSchedule
from visual.actualcost import ActualCost
from visual.resourcedistribution import ResourceDistribution
from visual.riskanalysis import RiskAnalysis
from visual.actualduration import ActualDuration
from visual.cpi import CPI
from visual.cost_value_metrics import CostValueMetrics
from visual.sv_t import SvT
from visual.spi_t import SpiT
from visual.spi_t_p_factor import SpiTvsPfactor
from visual.performance import Performance
from visual.budget import CV


class Processor(object):

    def __init__(self):
        self.visualizations = [BaselineSchedule(), "--",
           ResourceDistribution(), "--",
           RiskAnalysis(), "--",
           ActualDuration(), ActualCost(), "--",
           CostValueMetrics(), Performance(), SpiTvsPfactor(), SvT(), CV(), CPI(), SpiT()]
        self.file_parsers = []

    def convert(self, parser_from, parser_to, file_path_from, file_path_to):
        pass

    def create_all_file_parsers(self):
        self.file_parsers.append(XLSXParser())
        self.file_parsers.append(XMLParser())

    def create_all_visualizations(self):
        self.visualizations.append(LineChart())
        self.visualizations.append(PieChart())