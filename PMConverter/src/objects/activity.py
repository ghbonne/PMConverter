from objects.resource import Resource

__author__ = 'PM Group 8'

from objects.baselineschedule import BaselineScheduleRecord
from objects.riskanalysisdistribution import RiskAnalysisDistribution


class Activity(object):
    """
    A project consists of multiple activities.

    :var activity_id: int
    :var name: String
    :var wbs_id: tuple of ints, the id in the work breakdown structure (e.g (1, 2, 2))
    :var predecessors: list of tuples (activity_id, Relation, lag), relations are (FS (Finish-Start), FF, SS, SF), lag is an int expressed in workinghours (can be positive or negative)
    :var successors: list of tuples (activity_id, Relation, lag) lag is an int expressed in workinghours (can be positive or negative)
    :var resources: list of tuples (Resource, demand)
    :var resource_cost: float
    :var baseline_schedule: BaseLineScheduleRecord
    :var risk_analysis: RiskAnalysisDistribution
    """

    def __init__(self, activity_id, name="", wbs_id=(), predecessors=None, successors=None, resources=None, resource_cost=0.0,
                 baseline_schedule=None, risk_analysis=None,
                 type_check = True):
        """
        Initialize an Activity. The data types of the parameters must be the same as the properties of an Activity.

        :raises TypeError: one of the parameters is not the right type.
        """
        # avoid mutable default parameters!
        if predecessors is None: predecessors = []
        if successors is None: successors = []
        if resources is None: resources = []
        if baseline_schedule is None: baseline_schedule =  BaselineScheduleRecord()
        if risk_analysis is None: risk_analysis = RiskAnalysisDistribution()

        if type_check:
            if not isinstance(activity_id, int):
                raise TypeError('activity_id should be a number!')

            if not isinstance(name, str):
                raise TypeError('name should be a string!')

            if not isinstance(wbs_id, tuple) or not all(isinstance(element, int) for element in wbs_id):
                raise TypeError('wbs_id should be a tuple of ints!')

            if not isinstance(predecessors, list) or not all(isinstance(element, tuple) for element in predecessors):
                raise TypeError('predecessors should be a list with tuples!')

            if(len(predecessors) > 0 and
                (not all(len(element) for element in predecessors)
                 or not all(isinstance(element[0], int) for element in predecessors)
                 or not all(element[1] in ["FS", "FF", "SS", "SF"] for element in predecessors)
                 or not all(isinstance(element[2], int) for element in predecessors))):
                raise TypeError('predecessors should be a list with tuples (activity: int, '
                                'relation: [FS, FF, SS, SF], lag: int)!')

            if not isinstance(successors, list) or not all(isinstance(element, tuple) for element in successors):
                raise TypeError('successors should be a list with tuples!')

            if(len(successors) > 0 and
                (not all(len(element) for element in successors)
                 or not all(isinstance(element[0], int) for element in successors)
                 or not all(element[1] in ["FS", "FF", "SS", "SF"] for element in successors)
                 or not all(isinstance(element[2], int) for element in successors))):
                raise TypeError('successors should be a list with tuples (activity: int, '
                                'relation: [FS, FF, SS, SF], lag: int)!')

            if(not isinstance(resources, list) or not all(isinstance(element, tuple) for element in resources)
               or not all(isinstance(element[0], Resource) for element in resources)
               or not all(isinstance(element[1], int) for element in resources)):
                raise TypeError('resources should be a list with tuples (resource: Resource, demand: int)!')

            if not isinstance(resource_cost, float):
                raise TypeError('resource_cost must be a float!')

            if not isinstance(baseline_schedule, BaselineScheduleRecord):
                raise TypeError('baseline_schedule must be a BaselineScheduleRecord objects!')

            if not isinstance(risk_analysis, RiskAnalysisDistribution) and risk_analysis is not None:
                raise TypeError('risk_analysis must be a RiskAnalysisDistribution objects!')

        self.activity_id = activity_id
        self.name = name
        self.wbs_id = wbs_id
        self.predecessors = predecessors
        self.successors = successors
        self.resources = resources
        self.resource_cost = resource_cost
        self.baseline_schedule = baseline_schedule
        self.risk_analysis = risk_analysis

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented
