from calendar import Calendar

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

    def __init__(self, name="", activities=[], tracking_periods=[], resources=[], calendar=Calendar()):
        # TODO: Typechecking?
        self.name = name
        self.activities = activities
        self.tracking_periods = tracking_periods
        self.resources = resources
 
