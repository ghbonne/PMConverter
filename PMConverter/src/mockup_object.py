import datetime
from convert.XMLparser import XMLParser
from object.Activity import Activity
from object.baselineschedule import BaselineScheduleRecord
from object.projectobject import ProjectObject
from object.resource import Resource
from object.riskanalysisdistribution import RiskAnalysisDistribution

__author__ = 'gilles'

res1 = Resource(1, name="Programmer", resource_type="Renewable", cost_unit=100.0)
res2 = Resource(1, name="Tester", resource_type="Renewable", cost_unit=75.0)
bsr = BaselineScheduleRecord(datetime.datetime.now(), datetime.datetime.now(), 1000, 0, 1000)
ra = RiskAnalysisDistribution("Manual", "Absolute", 402, 480, 812)
act1 = Activity(1, name="App Dev", wbs_id=(1), resources=[(res1, 10), (res2, 5)], baseline_schedule=bsr,
                risk_analysis=ra)
project_object = ProjectObject(name="testing", activities=[act1], resources=[res1, res2])
print(project_object.__dict__)

xml_parser = XMLParser()
xml_parser.from_schedule_object(project_object, "test.xlsx")