__author__ = 'Eveline and ghbonne'
__license__ = "GPL"

import os
import copy
import re
import sys
import ntpath
from PyQt4.QtCore import *
import time
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
import traceback



class Processor(QThread):

    def __init__(self, parent = None):
        super(Processor, self).__init__(parent)
        self.visualizations = ["Baseline schedule's visualisations", BaselineSchedule(),
                               "Resources' visualisations", ResourceDistribution(),
                               "Risk analysis' visualisations", RiskAnalysis(),
                               "Tracking periods' visualisations", ActualDuration(), ActualCost(),
                               "Tracking overview's visualisations", CostValueMetrics(), Performance(), SpiTvsPfactor(), CV(), SvT(), CPI(), SpiT()]
        self.file_parsers = []
        self.inputFiletypes = {"Excel": ".xlsx", "ProTrack": ".p2x"}

        # init signals to GUI here:
        self.conversionSucceeded = SIGNAL("conversionSucceeded(QString)")
        self.conversionFailedErrorMessage = SIGNAL("conversionFailedErrorMessage(QString)")

        # conversion settings
        self.parser_from = ""
        self.parser_to = ""
        self.file_path_from = ""
        self.wantedVisualisations = {}

    def run(self):
        " Main running function of processor"

        inputFilePath, inputfileExtension = os.path.splitext(self.file_path_from)
        inputfilename = ntpath.basename(inputFilePath) + inputfileExtension
            
        # Parse to project object
        project_object = None
        if self.parser_from == "ProTrack":
            try:
                xml_parser = XMLParser()
                project_object = xml_parser.to_schedule_object(self.file_path_from)
            except:
                self.emit(self.conversionFailedErrorMessage, "Failed to parse {0}\nException of type {1} occurred.\nValue of exception = {2}\n\n {3}".format(inputfilename, 
                                                                            sys.exc_info()[0], sys.exc_info()[1] if sys.exc_info()[1] is not None else "", traceback.format_exc()))
                return
            
        elif self.parser_from == "Excel":
            try:
                xlsx_parser = XLSXParser()
                project_object = xlsx_parser.to_schedule_object(self.file_path_from)
            except:
                self.emit(self.conversionFailedErrorMessage, "Failed to parse {0}\nException of type {1} occurred.\nValue of exception = {2}\n\n {3}".format(inputfilename, 
                                                                                            sys.exc_info()[0], sys.exc_info()[1] if sys.exc_info()[1] is not None else "", traceback.format_exc()))
                return

        # Parse from project object
        file_path_to = inputFilePath + "_converted_" + time.strftime("%d_%H%M%S") + self.inputFiletypes[self.parser_to]
        if self.parser_to == "Excel":
            try:
                xlsx_parser = XLSXParser()
                workbook = xlsx_parser.from_schedule_object(project_object, file_path_to, self.excel_version)
                if self.wantedVisualisations:
                    current_tp = 0
                    for worksheet in workbook.worksheets():
                        if worksheet.get_name() == "Baseline Schedule":
                            for visualisation in self.wantedVisualisations["Baseline schedule's visualisations"]:
                                visualisation.draw(workbook, worksheet, project_object, self.excel_version)
                        if worksheet.get_name() == "Resources":
                            for visualisation in self.wantedVisualisations["Resources' visualisations"]:
                                visualisation.draw(workbook, worksheet, project_object, self.excel_version)
                        if worksheet.get_name() == "Risk Analysis":
                            for visualisation in self.wantedVisualisations["Risk analysis' visualisations"]:
                                visualisation.draw(workbook, worksheet, project_object, self.excel_version)
                        if "TP" in worksheet.get_name():
                            #tp = int(re.search(r'\d+', worksheet.get_name()).group())
                            for visualisation in self.wantedVisualisations["Tracking periods' visualisations"]:
                                visualisation.tp = current_tp
                                visualisation.draw(workbook, worksheet, project_object, self.excel_version)
                            current_tp += 1
                        if worksheet.get_name() == "Tracking Overview":
                            for visualisation in self.wantedVisualisations["Tracking overview's visualisations"]:
                                visualisation.draw(workbook, worksheet, project_object, self.excel_version)
                #try:
                    workbook.close()
                #except Exception:
                #    #todo: back to gui: "Can't write to file_path_to, please first close Excel and try again"
                #    pass
                #os.system("start excel.exe output/test2.xlsx")#todo: works only on windows, maybe not neccesary?
            except:
                self.emit(self.conversionFailedErrorMessage, "Failed to convert {0} to Excel\nException of type {1} occurred.\nValue of exception = {2}\n {3}".format(inputfilename, 
                                                    sys.exc_info()[0], sys.exc_info()[1] if sys.exc_info()[1] is not None else "", traceback.format_exc()))
                return
        elif self.parser_to == "ProTrack":
            try:
                xml_parser = XMLParser()
                parsingSuccessful = xml_parser.from_schedule_object(project_object, file_path_to)
            except:
                self.emit(self.conversionFailedErrorMessage, "Failed to convert {0} to ProTrack\nException of type {1} occurred.\nValue of exception = {2}\n\n {3}".format(inputfilename, 
                                                                                            sys.exc_info()[0], sys.exc_info()[1] if sys.exc_info()[1] is not None else "", traceback.format_exc()))
                return
            
        # conversion successful:
        self.emit(self.conversionSucceeded, file_path_to)
        return   

    def setConversionSettings(self, parser_from, parser_to, file_path_from, wantedVisualisations={}):
        self.parser_from = parser_from
        self.parser_to = parser_to
        self.file_path_from = file_path_from
        self.wantedVisualisations = wantedVisualisations
        self.excel_version = excel_version
    
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











