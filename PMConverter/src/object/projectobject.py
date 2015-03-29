__author__ = 'PM Group 8'


class ProjectObject(object):
    """
    Object that holds all information about a project needed to perform baseline scheduling with risk analysis
    and project control

    :var name: String
    :var activities: list of Activity
    :var tracking_periods: list of TrackingPeriod
    :var resources: list of Resource
    """

    def __init__(self, name="", activities=[], tracking_periods=[], resources=[]):
        self.name = name
        self._activityList = activities
        self._trackingPeriodsList = tracking_periods
        self._resourcesList = resources
 
