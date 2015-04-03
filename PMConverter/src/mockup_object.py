import datetime
import os
from convert.XLSXparser import XLSXParser
from object.Activity import Activity
from object.baselineschedule import BaselineScheduleRecord
from object.projectobject import ProjectObject
from object.resource import Resource
from object.riskanalysisdistribution import RiskAnalysisDistribution

__author__ = 'gilles'

xlsx_parser = XLSXParser()

# Some tests for writing to a XLSX File
# TODO: refactor to test class

res1 = Resource(1, name="Programmer", resource_type="Renewable", cost_unit=100.0)
res2 = Resource(1, name="Tester", resource_type="Renewable", cost_unit=75.0)
bsr1 = BaselineScheduleRecord(datetime.datetime.now(), datetime.datetime.now() + datetime.timedelta(days=5), 1000, 0,
                             1000)
bsr2 = BaselineScheduleRecord(datetime.datetime.now(), datetime.datetime.now() + datetime.timedelta(days=10), 10000, 10,
                             100)
ra1 = RiskAnalysisDistribution("Manual", "Absolute", 402, 480, 812)
ra2 = RiskAnalysisDistribution("Manual", "Absolute", 402, 480, 812)
act1 = Activity(1, name="App Dev", wbs_id=(1,), resources=[(res1, 10), (res2, 5)], baseline_schedule=bsr1,
                risk_analysis=ra1)
act2 = Activity(2, name="Testing", wbs_id=(1, 1,), resources=[(res2, 50)], baseline_schedule=bsr2,
                predecessors=[(act1, "FS", 0,)], risk_analysis=ra2)
act1.successors = [(act2, "FS", 0,)]
project_object = ProjectObject(name="PMConverter", activities=[act1, act2], resources=[res1, res2])
print(project_object.__dict__)

xlsx_parser.from_schedule_object(project_object, os.path.join(os.path.dirname(__file__), "test.xlsx"))

# Some tests for reading from a XLSX File
# TODO: refactor to test class

xlsx_parser.to_schedule_object(os.path.join(os.path.dirname(__file__),
                                            "../administration/2_Project data input sheet_extended.xlsx"))

