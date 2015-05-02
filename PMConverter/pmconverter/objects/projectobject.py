__author__ = 'Project Management Group 8 - 2015 Ghent University'

from pmconverter.objects import Agenda, Activity, TrackingPeriod, Resource


class ProjectObject(object):
    """
    Object that holds all information about a project needed to perform baseline scheduling with risk analysis
    and project control

    :var name: String
    :var activities: list of Activity
    :var tracking_periods: list of TrackingPeriod
    :var resources: list of Resource
    :var agenda: Agenda
    """

    def __init__(self, name="", activities=None, tracking_periods=None, resources=None, agenda=None, type_check = True):
        # avoid mutable default parameters!
        if activities is None: activities = []
        if tracking_periods is None: tracking_periods = []
        if resources is None: resources = []
        if agenda is None: agenda = Agenda()

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
    
    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented