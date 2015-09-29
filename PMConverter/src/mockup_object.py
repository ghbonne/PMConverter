import datetime
import os
import re
import sys, traceback
from convert.XLSXparser import XLSXParser
from convert.XMLparser import XMLParser
from visual.resourcedistribution import ResourceDistribution
from visual.riskanalysis import RiskAnalysis
from visual.enums import DataType, LevelOfDetail, XAxis
from visual.actualduration import ActualDuration
from visual.actualcost import ActualCost
from visual.baselineshedule import BaselineSchedule
from visual.cost_value_metrics import CostValueMetrics
from visual.performance import Performance
from visual.spi_t_p_factor import SpiTvsPfactor
from visual.sv_t import SvT
from visual.budget import CV
from visual.cpi import CPI
from visual.spi_t import SpiT

__author__ = 'gilles'

dir = "../data_analyse"
"""
for file_name in os.listdir(os.path.join(os.path.dirname(__file__), dir)):
    if file_name.endswith(".p2x"):
        print(file_name)
        input_path = dir + "/" + file_name
        output_path = dir + "/output/" + file_name[:-3] + "xlsx"

        #make new instances of parsers
        xlsx_parser = XLSXParser()
        xml_parser = XMLParser()

        print("Parsing xml to project object")
        try:
            po = xml_parser.to_schedule_object(input_path)

            print("Generating excel output")
            # Write the projectobject we just processed to a file
            workbook = xlsx_parser.from_schedule_object(po, output_path)

        except:
            print("FAILED")
            print("Unhandled Exception occurred of type: {0}".format(sys.exc_info()[0]))
            print("Unhandled Exception value = {0}".format(sys.exc_info()[1] if sys.exc_info()[1] is not None else "None"))
            exc_type, exc_obj, tb = sys.exc_info()
            frame = tb.tb_frame
            linenr = tb.tb_lineno
            filename = frame.f_code.co_filename
            print("EXCEPTION in {0} on line {1}\n".format(filename, linenr))
            traceback.print_exc()
            continue

        print("visualisations")
        #addvisualissations
        for worksheet in workbook.worksheets():
            if worksheet.get_name() == "Resources":
                v1 = ResourceDistribution()
                v1.data_type = DataType.RELATIVE
                v1.draw(workbook, worksheet,po)
            if worksheet.get_name() == "Risk Analysis":
                v2 = RiskAnalysis()
                v2.data_type = DataType.ABSOLUTE
                v2.draw(workbook, worksheet,po)
            if "TP" in worksheet.get_name():
                tp = int(re.search(r'\d+', worksheet.get_name()).group())
                v3 = ActualDuration()
                v3.level_of_detail = LevelOfDetail.WORK_PACKAGES
                v3.data_type = DataType.ABSOLUTE
                v3.tp = (tp-1)
                v3.draw(workbook, worksheet,po)
                v4 = ActualCost()
                v4.level_of_detail = LevelOfDetail.WORK_PACKAGES
                v4.data_type = DataType.ABSOLUTE
                v4.tp =(tp-1)
                v4.draw(workbook, worksheet,po)
            if worksheet.get_name() == "Baseline Schedule":
                v5 = BaselineSchedule()
                v5.draw(workbook, worksheet,po)
            if worksheet.get_name() == "Tracking Overview":
                v6 = CostValueMetrics()
                v6.x_axis = XAxis.TRACKING_PERIOD
                v6.draw(workbook, worksheet,po)
                v7 = Performance()
                v7.x_axis = XAxis.TRACKING_PERIOD
                v7.draw(workbook, worksheet,po)
                v8 = SpiTvsPfactor()
                v8.x_axis = XAxis.TRACKING_PERIOD
                v8.draw(workbook, worksheet,po)
                v9 = SvT()
                v9.x_axis = XAxis.TRACKING_PERIOD
                v9.draw(workbook, worksheet,po)
                v10 = CV()
                v10.x_axis = XAxis.TRACKING_PERIOD
                v10.draw(workbook, worksheet,po)
                v11 = CPI()
                v11.x_axis = XAxis.TRACKING_PERIOD
                v11.threshold = True
                v11.thresholdValues = (0.2, 0.5)
                v11.draw(workbook, worksheet,po)
                v12 = SpiT()
                v12.x_axis = XAxis.TRACKING_PERIOD
                v12.threshold = True
                v12.thresholdValues = (0.1, 0.7)
                v12.draw(workbook, worksheet,po)

        workbook.close()
"""
dir = "../data_analyse/output"
for file_name in os.listdir(os.path.join(os.path.dirname(__file__), dir)):
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        print(file_name)
        input_path = input_path = dir + "/" + file_name
        output_path = dir + "/excel_output/" + file_name[:-5] + "_excel.xlsx"
        xml_output_path = dir + "/excel_output/" + file_name[:-4] + "_xml.p2x"
        xlsx_parser = XLSXParser()
        xml_parser = XMLParser()

        print("Parsing excel to project object")
        po = xlsx_parser.to_schedule_object(input_path)
        print(po.activities[3].name)
        print(po.activities[3].baseline_schedule.start)
        print(po.activities[3].baseline_schedule.end)

        print("Generating excel output")
        workbook = xlsx_parser.from_schedule_object(po, output_path)
        print("Generating XML output")
        xml_parser.from_schedule_object(po, xml_output_path)

        workbook.close()
    print("--DONE--\n")



