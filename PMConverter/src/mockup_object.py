import datetime
import os
import re
from convert.XLSXparser import XLSXParser
from objects.activity import Activity
from objects.baselineschedule import BaselineScheduleRecord
from objects.projectobject import ProjectObject
from objects.resource import Resource
from objects.riskanalysisdistribution import RiskAnalysisDistribution
from visual.charts.piechart import PieChart
from visual.resourcedistribution import ResourceDistribution
from visual.riskanalysis import RiskAnalysis
from visual.enums import DataType, LevelOfDetail, XAxis, ExcelVersion
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

xlsx_parser = XLSXParser()

# Some tests for writing to a XLSX File
# TODO: refactor to test class

# TODO: multiple resources!
# TODO: where are the last 2 rows?

"""
res1 = Resource(1, name="Programmer", resource_type="Renewable", cost_unit=100.0)
res2 = Resource(1, name="Tester", resource_type="Renewable", cost_unit=75.0)
bsr1 = BaselineScheduleRecord(datetime.datetime.now(), datetime.datetime.now() + datetime.timedelta(days=5, hours=0),
                              datetime.timedelta(days=5, hours=0), 1000, 0, 1000)
bsr2 = BaselineScheduleRecord(datetime.datetime.now(), datetime.datetime.now() + datetime.timedelta(days=10, hours=0),
                              datetime.timedelta(days=10, hours=0), 10000, 10, 100)
ra1 = RiskAnalysisDistribution("manual", "absolute", 402, 480, 812)
ra2 = RiskAnalysisDistribution("manual", "absolute", 402, 480, 812)
act1 = Activity(1, name="App Dev", wbs_id=(1,), resources=[(res1, 10), (res2, 5)], baseline_schedule=bsr1,
                risk_analysis=ra1)
act2 = Activity(2, name="Testing", wbs_id=(1, 1,), resources=[(res2, 50)], baseline_schedule=bsr2,
                predecessors=[(1, "FS", 0,)], risk_analysis=ra2)
act1.successors = [(2, "FS", 0,)]
project_object = ProjectObject(name="PMConverter", activities=[act1, act2], resources=[res1, res2])
print(project_object.__dict__)

xlsx_parser.from_schedule_object(project_object, os.path.join(os.path.dirname(__file__), "test.xlsx"))
"""
# Some tests for reading from a XLSX File
# TODO: refactor to test class

print("Parsing the extended input sheet")
po = xlsx_parser.to_schedule_object(os.path.join(os.path.dirname(__file__),
                                            "../administration/2_Project data input sheet_extended_checked.xlsx"))

print("Generating output")
# Write the file we just processed to a file
excel_version = ExcelVersion.EXTENDED
workbook = xlsx_parser.from_schedule_object(po, "output/testTrackingOveriew_v2.xlsx", excel_version)


for worksheet in workbook.worksheets():
    if worksheet.get_name() == "Resources":
        v1 = ResourceDistribution()
        v1.data_type = DataType.RELATIVE
        v1.draw(workbook, worksheet,po,excel_version)
    if worksheet.get_name() == "Risk Analysis":
        v2 = RiskAnalysis()
        v2.data_type = DataType.ABSOLUTE
        v2.draw(workbook, worksheet,po,excel_version)
    if "TP" in worksheet.get_name():
        tp = int(re.search(r'\d+', worksheet.get_name()).group())
        v3 = ActualDuration()
        v3.level_of_detail = LevelOfDetail.WORK_PACKAGES
        v3.data_type = DataType.ABSOLUTE
        v3.tp = (tp-1)
        v3.draw(workbook, worksheet,po,excel_version)
        v4 = ActualCost()
        v4.level_of_detail = LevelOfDetail.WORK_PACKAGES
        v4.data_type = DataType.ABSOLUTE
        v4.tp =(tp-1)
        v4.draw(workbook, worksheet,po,excel_version)
    if worksheet.get_name() == "Baseline Schedule":
        v5 = BaselineSchedule()
        v5.draw(workbook, worksheet,po,excel_version)
    if worksheet.get_name() == "Tracking Overview":
        v6 = CostValueMetrics()
        v6.x_axis = XAxis.TRACKING_PERIOD
        v6.draw(workbook, worksheet,po,excel_version)
        v7 = Performance()
        v7.x_axis = XAxis.TRACKING_PERIOD
        v7.draw(workbook, worksheet,po,excel_version)
        v8 = SpiTvsPfactor()
        v8.x_axis = XAxis.TRACKING_PERIOD
        v8.draw(workbook, worksheet,po,excel_version)
        v9 = SvT()
        v9.x_axis = XAxis.TRACKING_PERIOD
        v9.draw(workbook, worksheet,po,excel_version)
        v10 = CV()
        v10.x_axis = XAxis.TRACKING_PERIOD
        v10.draw(workbook, worksheet,po,excel_version)
        v11 = CPI()
        v11.x_axis = XAxis.TRACKING_PERIOD
        v11.threshold = True
        v11.thresholdValues = (0.2, 0.5)
        v11.draw(workbook, worksheet,po,excel_version)
        v12 = SpiT()
        v12.x_axis = XAxis.TRACKING_PERIOD
        v12.threshold = True
        v12.thresholdValues = (0.1, 0.7)
        v12.draw(workbook, worksheet,po,excel_version)


workbook.close()
os.system("start excel.exe output/testTrackingOveriew_v2.xlsx")



