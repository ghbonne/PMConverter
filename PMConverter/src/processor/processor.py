__author__ = 'PM Group 8'
import os
import copy
from convert.XLSXparser import XLSXParser
from convert.XMLparser import XMLParser
from visual.visualization import Visualization
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
from visual.enums import ExcelVersion


class Processor(object):

    def __init__(self):
        self.visualizations = ["Baseline schedule's visualisations", BaselineSchedule(),
                               "Resources' visualisations", ResourceDistribution(),
                               "Risk analysis' visualisations", RiskAnalysis(),
                               "Tracking periods' visualisations", ActualDuration(), ActualCost(),
                               "Tracking overview's visualisations", CostValueMetrics(), Performance(), SpiTvsPfactor(), SvT(), CV(), CPI(), SpiT()]
        self.file_parsers = []

    def convert(self, parser_from, parser_to, file_path_from, visualisations={}, extended=False):
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
                    if worksheet.get_name() == "Baseline Schedule":
                        for visualisation in visualisations["Baseline schedule's visualisations"]:
                            visualisation.draw(workbook, worksheet, project_object)
                    if worksheet.get_name() == "Resources":
                        for visualisation in visualisations["Resources' visualisations"]:
                            visualisation.draw(workbook, worksheet, project_object)
                    if worksheet.get_name() == "Risk Analysis":
                        for visualisation in visualisations["Risk analysis' visualisations"]:
                            visualisation.draw(workbook, worksheet, project_object)
                    if "TP" in worksheet.get_name():
                        for visualisation in visualisations["Tracking periods' visualisations"]:
                            visualisation.draw(workbook, worksheet, project_object)
                    if worksheet.get_name() == "Tracking Overview":
                        for visualisation in visualisations["Tracking overview's visualisations"]:
                            visualisation.draw(workbook, worksheet, project_object)
                try:
                    workbook.close()
                except Exception:
                    #todo: back to gui: "Can't write to file_path_to, please first close Excel and try again"
                    pass
                os.system("start excel.exe output/test2.xlsx")#todo: works only on windows, maybe not neccesary?

    def create_all_file_parsers(self):
        self.file_parsers.append(XLSXParser())
        self.file_parsers.append(XMLParser())

    def get_supported_visualisations(self, excel_version):
        supported_visualisations = []
        # build new list
        for item in self.visualizations:
            if isinstance(item, Visualization):
                if excel_version in item.support:
                    supported_visualisations.append(item)
            else:
                supported_visualisations.append(item)
        # check for empty headers
        supported_copy = copy.deepcopy(supported_visualisations)
        previous_str = isinstance(supported_copy[0], str)
        previous_item = None
        for i in range(1, len(supported_copy)):  #skip first element
            item = supported_copy[i]
            if isinstance(item, str):
                if previous_str:
                    supported_visualisations.remove(previous_item)
                previous_str = 1
                previous_item = item
            else:
                previous_str = 0
                previous_item = None
        return supported_visualisations











