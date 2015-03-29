__author__ = 'PM Group 8'

from object.baselineschedule import BaselineScheduleRecord
from object.activitytracking import ActivityTrackingRecord


class Activity(object):
    """
    A project consists of multiple activities

    :var activity_id: int
    :var name: String
    :var wbs_id: tuple of ints, the id in the work breakdown structure (e.g (1, 2, 2))
    :var predecessors: list of tuples (Activity, Relation, lag), relations are (FS (Finish-Start), FF, SS, SF)
    :var successors: list of tuples (Activity, Relation, lag)
    :var resources: list of Resource
    :var baseline_schedule: BaseLineScheduleRecord
    :var activity_tracking: ActivityTrackingRecord
    """

    def __init__(self, activity_id, name="", wbs_id=(), predecessors = [()],
                 successors = [()], resources=[], baseline_schedule=BaselineScheduleRecord(),
                 activity_tracking = ActivityTrackingRecord()):
        self.activity_id = activity_id
        self.name = name
        self.wbs_id = wbs_id
        self.predecessors = predecessors
        self.successors = successors
        self.resources = resources
        self.baseline_schedule = baseline_schedule
        self.activity_tracking = activity_tracking


