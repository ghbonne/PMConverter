from object.resource import Resource

__author__ = 'PM Group 8'

from object.baselineschedule import BaselineScheduleRecord
from object.activitytracking import ActivityTrackingRecord
from object.riskanalysisdistribution import RiskAnalysisDistribution


class Activity(object):
    """
    A project consists of multiple activities

    :var activity_id: int
    :var name: String
    :var wbs_id: tuple of ints, the id in the work breakdown structure (e.g (1, 2, 2))
    :var predecessors: list of tuples (Activity, Relation, lag), relations are (FS (Finish-Start), FF, SS, SF)
    :var successors: list of tuples (Activity, Relation, lag)
    :var resources: list of tuples (Resource, demand)
    :var baseline_schedule: BaseLineScheduleRecord
    :var risk_analysis: RiskAnalysisDistribution
    :var activity_tracking: ActivityTrackingRecord
    """

    def __init__(self, activity_id, name="", wbs_id=(), predecessors = [()],
                 successors = [()], resources=[()], baseline_schedule=BaselineScheduleRecord(),
                 risk_analysis=RiskAnalysisDistribution(), activity_tracking = ActivityTrackingRecord()):
        if not isinstance(activity_id, int):
            raise TypeError('activity_id should be a number!')

        if not isinstance(name, str):
            raise TypeError('name should be a string!')

        if not isinstance(wbs_id, tuple):
            raise TypeError('wbs_id should be a tuple!')

        # TODO: check if predecessors[:][0] is Activity, [:][1] = FS, FF, SS or SF and [:][2] = int
        if not isinstance(predecessors, list) or not all(isinstance(element, tuple) for element in predecessors):
            raise TypeError('predecessors should be a list with tuples (activity: Activity, '
                            'relation: [FS, FF, SS, SF], lag: int)!')

        if not isinstance(successors, list) or not all(isinstance(element, tuple) for element in successors):
            raise TypeError('successors should be a list with tuples (activity: Activity, '
                            'relation: [FS, FF, SS, SF], lag: int)!')

        if(not isinstance(resources, list) or not all(isinstance(element, tuple) for element in resources)
           or not all(isinstance(element[0], Resource) for element in resources)
           or not all(isinstance(element[1], int) for element in resources)):
            raise TypeError('resources should be a list with tuples (resource: Resource, demand: int)!')

        if not isinstance(baseline_schedule, BaselineScheduleRecord):
            raise TypeError('baseline_schedule must be a BaselineScheduleRecord object!')

        if not isinstance(risk_analysis, RiskAnalysisDistribution):
            raise TypeError('risk_analysis must be a RiskAnalysisDistribution object!')

        if not isinstance(activity_tracking, ActivityTrackingRecord):
            raise TypeError('activity_tracking must be a ActivityTrackingRecord object!')
        self.activity_id = activity_id
        self.name = name
        self.wbs_id = wbs_id
        self.predecessors = predecessors
        self.successors = successors
        self.resources = resources
        self.baseline_schedule = baseline_schedule
        self.risk_analysis = risk_analysis
        self.activity_tracking = activity_tracking


