from object.Activity import Activity
from object.baselineschedule import BaselineScheduleRecord
from object.resource import Resource

__author__ = 'gilles'

res1 = Resource(1, name="Programmer", resource_type="Renewable", cost_unit=100.0)
res2 = Resource(1, name="Tester", resource_type="Renewable", cost_unit=75.0)
bsr = BaselineScheduleRecord("01/01/2001", "15/01/2001", 1000, 0, 1000)
act1 = Activity(1, activity_name="App Dev", wbs_id=(1), resources=[(res1, 10), (res2, 5)], baseline_schedule=bsr)