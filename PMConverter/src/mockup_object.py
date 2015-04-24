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
from visual.enums import DataType, LevelOfDetail
from visual.actualduration import ActualDuration
from visual.actualcost import ActualCost
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
                                            "../administration/2_Project data input sheet_extended.xlsx"))

# Write the file we just processed to a file
workbook = xlsx_parser.from_schedule_object(po, "output/test2.xlsx")

for worksheet in workbook.worksheets():
    if worksheet.get_name() == "Resources":
        vis = ResourceDistribution()
        vis.data_type = DataType.RELATIVE
        vis.draw(workbook, worksheet)
    if worksheet.get_name() == "Risk Analysis":
        v2 = RiskAnalysis()
        v2.level_of_detail = LevelOfDetail.ACTIVITIES
        v2.draw(workbook, worksheet, po)
    if "TP" in worksheet.get_name():
        tp = int(re.search(r'\d+', worksheet.get_name()).group())
        v3 = ActualDuration()
        v3.level_of_detail = LevelOfDetail.ACTIVITIES
        v3.data_type = DataType.RELATIVE
        v3.draw(workbook, worksheet, po, tp-1)
        v4 = ActualCost()
        v4.level_of_detail = LevelOfDetail.ACTIVITIES
        v4.data_type = DataType.RELATIVE
        v4.draw(workbook, worksheet, po, tp-1)


workbook.close()
os.system("start excel.exe output/test2.xlsx")


