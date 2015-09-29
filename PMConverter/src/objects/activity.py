from objects.resource import Resource

__author__ = 'Project management group 8, Ghent University 2015'

from objects.baselineschedule import BaselineScheduleRecord
from objects.riskanalysisdistribution import RiskAnalysisDistribution
from datetime import datetime


class Activity(object):
    """
    A project consists of multiple activities.

    :var activity_id: int
    :var name: String
    :var wbs_id: tuple of ints, the id in the work breakdown structure (e.g (1, 2, 2))
    :var predecessors: list of tuples (activity_id, Relation, lag), relations are (FS (Finish-Start), FF, SS, SF), lag is an int expressed in workinghours (can be positive or negative)
    :var successors: list of tuples (activity_id, Relation, lag) lag is an int expressed in workinghours (can be positive or negative)
    :var resources: list of tuples (Resource, demand, fixed_consumable_assignment), (Resource, float or int, bool)
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
               or not all(isinstance(element[1], float) for element in resources)
               or not all(isinstance(element[2], bool) for element in resources)):
                raise TypeError('resources should be a list with tuples (resource: Resource, demand: float, fixed assignment: bool)!')

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

    @staticmethod
    def is_not_lowest_level_activity(activity, activities):
        # Decide whether an activity is not of the lowest level or not.
        if not activity.wbs_id:
            # wbs_id is None or empty
            if activity.baseline_schedule.var_cost is not None:
                return False
            else:
                return True
        else:
            for _activity in activities:
                if _activity is not activity and len(activity.wbs_id) < len(_activity.wbs_id):
                    if activity.wbs_id[:] == _activity.wbs_id[:len(activity.wbs_id)]:
                        # _activity found with a wbs below activity => activity is an activityGroup
                        return True
            return False

    @staticmethod
    def generate_activityGroups_to_childActivities_dict(activitiesOnly, activityGroups):
        """This function links activityGroup id's to activity id's
        :returns: dict with each activityGroup_id as key and a list of all sub acitivityId's as value
        """
        activityGroup_to_childActivities_dict = {}
        
        for activityGroup in activityGroups:
            activityGroupId = activityGroup.activity_id
            activityGroup_wbs_level = len(activityGroup.wbs_id)
            activityGroup_to_childActivities_dict[activityGroupId] = [x.activity_id for x in activitiesOnly if x.wbs_id > activityGroup.wbs_id and x.wbs_id[:activityGroup_wbs_level] == activityGroup.wbs_id]
        return activityGroup_to_childActivities_dict

    @staticmethod
    def update_activityGroups_aggregated_values(activityGroups_list, activity_dict, activityGroup_to_childActivities_dict, agenda):
        "This function calculates the aggregated values of activityGroups"

        for activityGroup in activityGroups_list:
            childActivityIds = activityGroup_to_childActivities_dict[activityGroup.activity_id]
            # clear activityGroup values: (necessary for non propper given activityGroups via Activity.check_lists_activities_groups)
            activityGroup.baseline_schedule.fixed_cost = 0
            activityGroup.baseline_schedule.total_cost = 0
            activityGroup.baseline_schedule.hourly_cost = 0
            activityGroup.baseline_schedule.var_cost = None # identifies an activityGroup
            activityGroup.successors = []
            activityGroup.predecessors = []
            activityGroup.resources = []
            activityGroup.resource_cost = 0

            earliestStart = datetime.max
            latestFinish = datetime.min

            for childActivityId in childActivityIds:
                childActivity = activity_dict[childActivityId]

                if childActivity.baseline_schedule.start < earliestStart:
                    earliestStart = childActivity.baseline_schedule.start
                if childActivity.baseline_schedule.end > latestFinish:
                    latestFinish = childActivity.baseline_schedule.end

                activityGroup.baseline_schedule.fixed_cost += childActivity.baseline_schedule.fixed_cost
                activityGroup.baseline_schedule.total_cost += childActivity.baseline_schedule.total_cost

            # calculate activityGroup duration:
            if earliestStart < datetime.max and latestFinish > datetime.min:
                activityGroup.baseline_schedule.duration = agenda.get_time_between(earliestStart, latestFinish)
                activityGroup.baseline_schedule.start = earliestStart
                activityGroup.baseline_schedule.end = latestFinish

    @staticmethod
    def check_lists_activities_groups(activities_dict, activityGroups_dict):
        """
        Checks if all activityGroups have subactivities, else exception.
        Checks if all activities with subactivities are added to the activityGroups_dict
        """
        found_activityGroup_ids = []
        for row, activity in activities_dict.values():
            current_wbs = activity.wbs_id
            current_wbs_level = len(current_wbs)
            current_id = activity.activity_id
            # check if subactivities found based on wbs:
            for row2, activity2 in activities_dict.values():
                if activity2.wbs_id > current_wbs and activity2.wbs_id[:current_wbs_level] == current_wbs:
                    # subactivity found:
                    found_activityGroup_ids.append(current_id)
                    # check if already present in the activityGroups_dict:
                    if current_id not in activityGroups_dict:
                        # if not yet present: add to groups dict
                        print("Activity:check_lists_activities_groups: An activityGroup (id = {0}) found which was not present in the activityGroups_dict!".format(current_id))
                        activityGroups_dict[current_id] = activity
                    break
        # check if no activityGroups without subactivities:
        for activityGroup_id in activityGroups_dict.keys():
            if activityGroup_id not in found_activityGroup_ids:
                raise Exception("Activity: An activityGroup (id = {0}) was found which has no subactivities!".format(activityGroup_id))
        return    

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented
