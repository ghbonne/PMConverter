__author__ = 'PM Group 8'
from datetime import datetime, timedelta


class ProjectObject(object):
    "Object that holds all information about a project needed to perform baseline scheduling with risk analysis and project control"
    # instance variables:
    # _baselineSchedule: BaselineScheduleRecord
    # _baselineCosts: dictionary containing fields "Fixed Cost", "Cost/Hour", "Variable Cost" and "Total Cost" if nonezero
    # _activityList: List of activities in this project
    # _trackingPeriodsList: List of all tracking periods in this project
    # _resourcesList: List of all defined resources in this project
    # class variables:
    ProjectName = ""

    def __init__(self, name="", baselineSchedule=BaselineScheduleRecord(), baselineCosts={}, activityList=list(), trackingPeriodsList=list(), resourcesList=list()):
        ProjectObject.ProjectName = name
        self._baselineSchedule = baselineSchedule
        self._baselineCosts = baselineCosts
        self._activityList = activityList
        self._trackingPeriodsList = trackingPeriodsList
        self._resourcesList = resourcesList
 
