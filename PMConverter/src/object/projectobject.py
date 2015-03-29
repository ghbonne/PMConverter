__author__ = 'PM Group 8'

from object.baselineschedule import BaselineScheduleRecord


class ProjectObject(object):
    """
    Object that holds all information about a project needed to perform baseline scheduling with risk analysis
    and project control

    :var name: String
    :var baseline_schedule: BaselineScheduleRecord
    :var baseline_costs: dictionary containing "Fixed Cost", "Cost/Hour", "Variable Cost" and "Total Cost" if nonzero
    :var activities: list of Activity
    :var tracking_periods: list of TrackingPeriod
    :var resources: list of Resource
    """

    def __init__(self, name="", baseline_schedule=BaselineScheduleRecord(), baseline_costs={}, activities=[],
                 tracking_periods=[], resources=[]):
        self.name = name
        self._baselineSchedule = baseline_schedule
        self._baselineCosts = baseline_costs
        self._activityList = activities
        self._trackingPeriodsList = tracking_periods
        self._resourcesList = resources
 
