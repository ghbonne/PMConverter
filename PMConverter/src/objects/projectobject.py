from objects.agenda import Agenda
from objects.activity import Activity
from objects.trackingperiod import TrackingPeriod
from objects.resource import Resource

__author__ = 'PM Group 8'


class ProjectObject(object):
    """
    Object that holds all information about a project needed to perform baseline scheduling with risk analysis
    and project control

    :var name: String
    :var activities: list of Activity
    :var tracking_periods: list of TrackingPeriod
    :var resources: list of Resource
    :var calendar: calendar
    """

    def __init__(self, name="", activities=[], tracking_periods=[], resources=[], agenda=Agenda(), type_check = True):
        if type_check:
            if not isinstance(name, str):
                raise  TypeError('name should be a string!')
            if not isinstance(activities, list) or not all(isinstance(element, Activity) for element in activities):
                raise TypeError('activities should be a list of Activity objects!')
            if not isinstance(tracking_periods, list) or not all(isinstance(element, TrackingPeriod) for element in tracking_periods):
                raise TypeError('tracking_periods should be a list of TrackingPeriod objects!')
            if not isinstance(resources, list) or not all(isinstance(element, Resource) for element in resources):
                raise TypeError('resources should be a list of Resource objects!')
            if not isinstance(agenda, Agenda):
                raise TypeError('agend should be of type Agenda')
        self.name = name
        self.activities = activities
        self.tracking_periods = tracking_periods
        self.resources = resources
        self.agenda = agenda
