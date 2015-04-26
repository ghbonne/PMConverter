__author__ = 'PM Group 8'
import os
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
        self.visualizations = ["Baseline schedule's visualisations", BaselineSchedule(),
                               "Resources' visualisations", ResourceDistribution(),
                               "Risk analysis' visualisations", RiskAnalysis(),
                               "Tracking periods' visualisations", ActualDuration(), ActualCost(),
                               "Tracking overview's visualisations", CostValueMetrics(), Performance(), SpiTvsPfactor(), SvT(), CV(), CPI(), SpiT()]
        self.file_parsers = []

    def convert(self, parser_from, parser_to, file_path_from, visualisations=[], extended=0):
        # Parse to project object
        project_object = None
        if parser_from == "protrack":
            #todo: call Alexanders code
            # save result in project object
            pass
        elif parser_from == "Excel":
            xlsx_parser = XLSXParser()
            project_object = xlsx_parser.to_schedule_object(os.path.join(os.path.dirname(__file__), file_path_from))

        # Parse from project object
        file_path_to = "" #todo: build file_path_to, based on file_path_from?
        if parser_to == "Excel":
            xlsx_parser = XLSXParser()
            workbook = xlsx_parser.from_schedule_object(project_object, file_path_to)#todo: how do the parser know if it is basic or extended version?
            if visualisations:
                for worksheet in workbook.worksheets():
                    if worksheet.get_name() == "Resources":
                        pass
                    if worksheet.get_name() == "Risk Analysis":
                        pass
                    if "TP" in worksheet.get_name():
                        pass
                    if worksheet.get_name() == "Baseline Schedule":
                        pass
                    if worksheet.get_name() == "Tracking Overview":
                        pass
                try:
                    workbook.close()
                except Exception:
                    #todo: back to gui: "Can't write to file_path_to, please first close Excel and try again"
                    pass
                os.system("start excel.exe output/test2.xlsx")#todo: works only on windows, maybe not neccesary?

    def create_all_file_parsers(self):
        self.file_parsers.append(XLSXParser())
        self.file_parsers.append(XMLParser())
