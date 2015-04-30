__author__ = 'PM Group 8'

from datetime import datetime, timedelta,date
import xml.etree.ElementTree as ET
from objects.activity import Activity
from objects.baselineschedule import BaselineScheduleRecord
from objects.resource import Resource
from objects.riskanalysisdistribution import RiskAnalysisDistribution
from objects.activitytracking import ActivityTrackingRecord
from objects.trackingperiod import TrackingPeriod
from objects.resource import ResourceType
from objects.riskanalysisdistribution import DistributionType
from objects.riskanalysisdistribution import ManualDistributionUnit
from objects.riskanalysisdistribution import StandardDistributionUnit
from objects.agenda import Agenda
from objects.projectobject import ProjectObject
from convert.fileparser import FileParser
import math
import ast
from exceptions import XMLParseError




class XMLParser(FileParser):

    def __init__(self):
        super(XMLParser, self).__init__()

    def getdate(self,datestring="", dateformat=""):
            if dateformat == "d/MM/yyyy h:mm" or "d-M-yyyy h:m":
                if len(datestring) == 12:
                    day=int(datestring[:2])
                    month=int(datestring[2:4])
                    year=int(datestring[4:8])
                    hour=int(datestring[8:10])
                    minute=int(datestring[10:12])
                    return datetime(year, month, day, hour, minute)
                elif len(datestring) == 8:
                    day=int(datestring[:2])
                    month=int(datestring[2:4])
                    year=int(datestring[4:8])
                    return datetime(year, month, day)
                else:
                    #raise XMLParseError("getdate: datestring length {0} is not equal to 8 or 12. datestring = {1}".format(len(datestring), datestring))  #TODO: datestring="0" gets here
                    return datetime.max
            elif dateformat == "MM/d/yyyy h:mm":
                if len(datestring) == 12:
                    month=int(datestring[:2])
                    day=int(datestring[2:4])
                    year=int(datestring[4:8])
                    hour=int(datestring[8:10])
                    minute=int(datestring[10:12])
                    return datetime(year, month, day, hour, minute)
                elif len(datestring) == 8:
                    month=int(datestring[:2])
                    day=int(datestring[2:4])
                    year=int(datestring[4:8])
                    return datetime(year, month, day)
                else:
                    #raise XMLParseError("getdate: datestring length {0} is not equal to 8 or 12. datestring = {1}".format(len(datestring), datestring))
                    return datetime.max
            else:
                raise XMLParseError("getdate: unexpected dateformat: {0}".format(dateformat))


    def get_date_string(self,date=None,dateformat=""):
        "This funciton converts a datetime to a string in the given format"
        # avoid mutable default parameters!
        if date is None: date = datetime.min

        if date=='0':
            return "0"
        year=date.year
        month=date.month
        day=date.day
        hour=date.hour
        minute=date.minute
        year_str=str(year)
        if month < 10:
            month_str="0"
            month_str+=str(month)
        else:
            month_str=str(month)
        if day < 10:
            day_str="0"
            day_str+=str(day)
        else:
            day_str=str(day)
        if hour < 10:
            hour_str="0"
            hour_str+=str(hour)
        else:
            hour_str=str(hour)
        if minute < 10:
            minute_str="0"
            minute_str+=str(minute)
        else:
            minute_str=str(minute)
        if dateformat == "d/MM/yyyy h:mm":
            return day_str+month_str+year_str+hour_str+minute_str
        elif dateformat == "MM/d/yyyy h:mm":
            return month_str+day_str+year_str+hour_str+minute_str
        else:
            raise XMLParseError("get_date_string: Dateformat undefined: {0}".format(dateformat))

    def set_wbs_read_activities_and_groups(self, parent_wbs, children_nodes, activity_dict, activityGroup_dict, parentActivityGroup_id, activityGroup_to_childActivities_dict):
        """
        This function sets the wbs id tuple of the given children outlineList nodes and adds it to their corresponding activity or activityGroup.
        :param parent_wbs: tuple, identifies the wbs id of the parent activityGroup
        :param children_nodes: list of Elements, below the parent activityGroup
        :param activity_dict: dict, activity_Id: Activity
        :param activityGroup_dict: dict: activityGroup_Id: Activity
        :param parentActivityGroup_id: int, id of the parent acitivityGroup
        :param activityGroup_to_childActivities_dict: dict, activityGroup_id: list of Activivities, links an activityGroup to all its child low-level activities

        :returns: list of all low-level activity id's below the parent_wbs
        """
        # make starting wbs from parent
        current_wbs_list = list(parent_wbs).append(1)
        current_child_activityIds = []
        for child_node in children_nodes:
            # determine type of child:
            nodeType_node = child_node.find("Type")
            nodeType = 0 if nodeType_node is None else int(nodeType_node.text)
            if nodeType == 1:
                # this node corresponds with an activityGroup:
                activityGroupId = int(child_node.find("Data").text)
                if activityGroupId in activityGroup_dict:
                    activityGroup_dict[activityGroupId].wbs_id = tuple(current_wbs_list)
                    grandChildrenNode = child_node.find("List")
                    if grandChildrenNode is not None:
                        # call recursive function to add its children:
                        grandchild_activityIds = self.set_wbs_read_activities_and_groups(tuple(current_wbs_list), list(grandChildrenNode), activity_dict, activityGroup_dict)
                        # save children ids of this activityGroup:
                        activityGroup_to_childActivities_dict[activityGroupId] = grandchild_activityIds

                        # also add grandchildren to the children's list of this parent
                        current_child_activityIds.extend(grandchild_activityIds)
                    else:
                        # activityGroup has no children:
                        print("XMLparser:set_wbs_read_activities_and_groups: Found an activityGroup without child activities!") # Warning
                        activityGroup_to_childActivities_dict[activityGroupId] = []

                    # increment wbs_id for next child
                    current_wbs_list[-1] += 1
                else:
                    raise XMLParseError("set_wbs_read_activities_and_groups: Could not find activityGroup with id = {0} in read activityGroups.".format(activityGroupId))
            elif nodeType == 2:
                # this node corresponds with an activity:
                activityId = int(child_node.find("Data").text)
                if activityId in activity_dict:
                    activity_dict[activityId].wbs_id = tuple(current_wbs_list)
                    # add child id to list of parent:
                    current_child_activityIds.append(activityId)

                    # increment wbs_id for next child
                    current_wbs_list[-1] += 1
                else:
                    raise XMLParseError("set_wbs_read_activities_and_groups: Could not find activity with id = {0} in read activities.".format(activityId))
            else:
                # not a valid node
                continue
        return current_child_activityIds

    def update_activities_aggregated_costs(self, activities_list, agenda):
        "This function updates the given activities their total_cost and resource_cost fields"
        for activity in activities_list:
            resource_cost = 0
            for resourceTuple in activity.resources:
                resource = resourceTuple[0]
                if resource.resource_type == ResourceType.CONSUMABLE:
                    ## only add once the cost for its use!
                    # check if fixed resource assignment:
                    if resourceTuple[2]:
                        # fixed resource assignment => variable cost is not multiplied by activity duration:
                        resource_cost += resource.cost_use + resourceTuple[1] * resource.cost_unit
                    else:
                        # non fixed resource assignment:
                        resource_cost += resource.cost_use + resourceTuple[1] * resource.cost_unit * actualDuration_hours
                else:
                    #resource type is renewable:
                    #add cost_use and variable cost:
                    resource_cost += resourceTuple[1] * (resource.cost_use + resource.cost_unit * actualDuration_hours)
            #endFor adding resource costs

            activity.resource_cost = resource_cost

            activityDuration_hours = activity.baseline_schedule.duration.days * agenda.get_working_hours_in_a_day() + round(activity.baseline_schedule.duration.seconds / 3600)

            # calculate total activity cost:
            activity.baseline_schedule.total_cost = activity.resource_cost + activity.baseline_schedule.fixed_cost + activity.baseline_schedule.hourly_cost * activityDuration_hours

    def update_activityGroups_aggregated_values(self, activityGroups_list, activity_dict, activityGroup_to_childActivities_dict, agenda):
        "This function calculates the aggregated values of activityGroups"

        for activityGroup in activityGroups_list:
            childActivityIds = activityGroup_to_childActivities_dict[activityGroup.activity_id]

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
            activityGroup.baseline_schedule.duration = agenda.get_time_between(earliestStart, latestFinish)
                    

    def process_activityTrackingRecord_Node(self, activity, activityTrackingRecord_node, statusdate_datetime, agenda):
        """"
        This function processes an acitivityTrackingRecord node from ProTrack p2x file.
        :returns: ActivityTrackingRecord
        """
        # Read Data
        actualStart = self.getdate(tracking_activity.find('ActualStart').text, dateformat)
        actualDuration_hours = int(tracking_activity.find('ActualDuration').text)
        actualDuration = agenda.get_workingDuration_timedelta(duration_hours=actualDuration_hours)
        actualCostDev = float(tracking_activity.find('ActualCostDev').text)
        remainingDuration_hours = int(tracking_activity.find('RemainingDuration').text)
        remainingDuration = agenda.get_workingDuration_timedelta(duration_hours=remainingDuration_hours)
        remainingCostDev = float(tracking_activity.find('RemainingCostDev').text)
        percentageComplete = float(tracking_activity.find('PercentageComplete').text)*100
        if percentageComplete >= 100:
            trackingStatus='Finished'
        elif abs(percentageComplete) < 1e-5: # compare float with 0
            trackingStatus='Not Started'
        else:
            trackingStatus='Started'

        # derive remaining data fields:
        #planned_actual_cost = 0
        #planned_remaining_cost = 0
        #actual_cost = 0
        #remaining_cost = 0
        #earned_value = 0
        #planned_value = 0

        # Calculate PAC and PRC:
        if actualDuration_hours > 0:
            # activity started or already finished: PAC depends on real actual duration!
            planned_actual_cost = activity.baseline_schedule.fixed_cost + actualDuration_hours * activity.baseline_schedule.hourly_cost
            # no fixed starting costs anymore for PRC:
            planned_remaining_cost = remainingDuration_hours * activity.baseline_schedule.hourly_cost
            # add costs of used resources:
            for resourceTuple in activity.resources:
                resource = resourceTuple[0]
                if resource.resource_type == ResourceType.CONSUMABLE:
                    ## only add once the cost for its use!
                    # check if fixed resource assignment:
                    if resourceTuple[2]:
                        # fixed resource assignment => variable cost is not multiplied by activity duration:
                        planned_actual_cost += resource.cost_use + resourceTuple[1] * resource.cost_unit
                        # NOTE: expects fixed resource assignment to take place at start of the activity => no contribution to PRC
                    else:
                        # non fixed resource assignment:
                        planned_actual_cost += resource.cost_use + resourceTuple[1] * resource.cost_unit * actualDuration_hours
                        planned_remaining_cost += resourceTuple[1] * resource.cost_unit * remainingDuration_hours
                else:
                    #resource type is renewable:
                    #add cost_use and variable cost:
                    planned_actual_cost += resourceTuple[1] * (resource.cost_use + resource.cost_unit * actualDuration_hours)
                    planned_remaining_cost += resourceTuple[1] * resource.cost_unit * remainingDuration_hours
            #endFor adding resource costs

        else:
            # activity not yet started:
            planned_actual_cost = 0
            planned_remaining_cost = activity.baseline_schedule.total_cost

        actual_cost = planned_actual_cost + actualCostDev
        remaining_cost = planned_remaining_cost + remainingCostDev

        # Calculate EV:
        if remainingDuration_hours == 0:
            # activity finished
            earned_value = activity.baseline_schedule.total_cost
        elif actualDuration_hours == 0:
            # activity not yet started:
            earned_value = 0
        else:
            # activity running:
            earned_value = activity.baseline_schedule.fixed_cost + percentageComplete * (activity.baseline_schedule.var_cost + activity.resource_cost)

        # Calculate PV:
        if statusdate_datetime >= activity.baseline_schedule.end:
            # activity should be finished according to baselinschedule
            planned_value = activity.baseline_schedule.total_cost
        elif statusdate_datetime < activity.baseline_schedule.start:
            # activity is not yet started according to baselineschedule
            planned_value = 0
        else:
            # activity is running
            activityRunningDuration = agenda.get_time_between(activity.baseline_schedule.start, statusdate_datetime)
            activityRunningDuration_workingHours = activityRunningDuration.days * agenda.get_working_hours_in_a_day() + activityRunningDuration.seconds / 3600

            planned_value = activity.baseline_schedule.fixed_cost + activityRunningDuration_workingHours * activity.baseline_schedule.hourly_cost
                    
            # add costs of resources until now:
            for resourceTuple in activityTrackingRecord.activity.resources:
                resource = resourceTuple[0]
                if resource.resource_type == ResourceType.CONSUMABLE:
                    ## only add once the cost for its use!
                    # check if fixed resource assignment:
                    if resourceTuple[2]:
                        # fixed resource assignment => variable cost is not multiplied by activity duration:
                        planned_value += resource.cost_use + resourceTuple[1] * resource.cost_unit
                    else:
                        # non fixed resource assignment:
                        planned_value += resource.cost_use + resourceTuple[1] * resource.cost_unit * activityRunningDuration_workingHours
                else:
                    #resource type is renewable:
                    #add cost_use and variable cost:
                    planned_value += resourceTuple[1] * (resource.cost_use + resource.cost_unit * activityRunningDuration_workingHours)
            #endFor adding resource costs
        #endIf calculating PV
        
        return ActivityTrackingRecord(activity, actualStart, actualDuration, planned_actual_cost, planned_remaining_cost, remainingDuration, actualCostDev,
                                                        remainingCostDev, actual_cost, remaining_cost, int(round(percentageComplete)), trackingStatus, earned_value, planned_value, True)

    def construct_activityGroup_trackingRecord(self, activityGroup, childActivityIds, currentTrackingPeriod_records_dict):
        """This function constructs an aggregated trackingRecord of an activityGroup
        :returns: ActivityTrackingRecord of aggregated activityGroup
        """

        total_earned_value = 0.
        total_planned_value = 0.

        if childActivityIds:
            for childActivityId in childActivityIds:
                childActivityTrackingRecord = currentTrackingPeriod_records_dict[childActivityId]
                total_earned_value += childActivityTrackingRecord.earned_value
                total_planned_value += childActivityTrackingRecord.planned_value

        percentage_completed = int(round(total_earned_value / activityGroup.baseline_schedule.total_cost))
        return ActivityTrackingRecord(activity= activityGroup, percentage_completed= percentage_completed, earned_value= total_earned_value, planned_value= total_planned_value)

    def to_schedule_object(self, file_path_input):
        tree = ET.parse(file_path_input)
        root = tree.getroot()

        # read dicts from ProTrack file:
        activity_dict = []
        activityGroup_dict = {}
        activityGroup_to_childActivities_dict = {} # links activityGroup ids to all child low-level activity id's
        res_dict = {}  # resources dict
        trackingPeriodsList = []
        project_agenda = Agenda()

        ## Project name
        project_name = root.find("NAME").text

        ## Dateformat
        dateformat = ""
        for settings in root.findall('Settings'):
            dateformat=settings.find('DateTimeFormat').text

        ## Create Agenda: Working hours, Working days, Holidays
        for agenda in root.findall('Agenda'):
            # working hours
            for nonworkinghour in agenda.findall('NonWorkingHours'):
                for hours in nonworkinghour.findall('Hour'):
                    hour=int(hours.text)
                    project_agenda.set_non_working_hour(hour)
            # working days
            for nonworkingday in agenda.findall('NonWorkingDays'):
                for days in nonworkingday.findall('Day'):
                    day=int(days.text)
                    project_agenda.set_non_working_day(day)
            # Holidays
            for holidays in agenda.findall('Holidays'):
                for holiday in holidays.findall('Holiday'):
                    holiday=holiday.text[:8]
                    project_agenda.set_holiday(holiday)

        ###### Resources ######
        ## Resources (Definition): create dict of all resources in project:
        for resourcesNode in root.findall('Resources'):
            for resourceNode in resourcesNode.findall('Resource'):
                res_ID = int(resourceNode.find('FIELD0').text)
                name = resourceNode.find('FIELD1').text
                res_type = ResourceType.RENEWABLE if int(resourceNode.find('FIELD769').text) else ResourceType.CONSUMABLE
                cost_per_use = float(resourceNode.find('FIELD770').text)
                cost_per_unit = float(resourceNode.find('FIELD771').text)
                total_resource_cost = ast.literal_eval(resourceNode.find('FIELD776').text)
                availability_default = ast.literal_eval(resourceNode.find('FIELD780').text)
                res_dict[res_ID] = Resource(res_ID, name, res_type, availability_default, cost_per_use, cost_per_unit, total_resource_cost, type_check = True)
                

        ### Risk Analysis ###
        distribution_dict =  {} # dict with possible distributions and their ID as key
        #Standard distributions
        distribution_dict[1] = RiskAnalysisDistribution(distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.NO_RISK, optimistic_duration=99,
                         probable_duration=100, pessimistic_duration=101)
        distribution_dict[2] = RiskAnalysisDistribution(distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SYMMETRIC, optimistic_duration=80,
                         probable_duration=100, pessimistic_duration=120)
        distribution_dict[3] = RiskAnalysisDistribution(distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SKEWED_LEFT, optimistic_duration=80,
                         probable_duration=110, pessimistic_duration=120)
        distribution_dict[4] = RiskAnalysisDistribution(distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SKEWED_RIGHT, optimistic_duration=80,
                         probable_duration=90, pessimistic_duration=120)

        for distributionsNode in root.findall('SensitivityDistributions'):
            # in a distributions node, loop over all its children nodes starting with a tag "TProTrackSensitivityDistribution"
            for distribution_node in list(distributionsNode):
                if distribution_node.tag.startswith("TProTrackSensitivityDistribution"):
                    # found sensitivity distribution node:
                    distributionID = int(distribution_node.find("UniqueID").text)
                    distribution_units = ManualDistributionUnit.RELATIVE if int(distribution_node.find("Style").text) else ManualDistributionUnit.ABSOLUTE
                    distributionPointsNode = distribution_node.find("Distribution")
                    # extract X and Y points:
                    distributionXPoints = distributionPointsNode.findall("X")
                    distributionYPoints = distributionPointsNode.findall("Y")
                    if len(distributionXPoints) != 3 or len(distributionYPoints) != 3 or not (distributionYPoints == [0, 100, 0]):
                        # unsupported distribution type: don't fail conversion by this
                        print("XMLparser:to_schedule_object: Only sensitivity distributions defined by 3 points are supported!\n Not with {0} points.".format(len(distributionXPoints)))
                        continue
                    # valid 3 point distribution here:
                    distribution_dict[distributionID] = RiskAnalysisDistribution(distribution_type=DistributionType.MANUAL, distribution_units=distribution_units,
                                                                  optimistic_duration=distributionXPoints[0],probable_duration=distributionXPoints[1], pessimistic_duration=distributionXPoints[2])


        ### Read all activities and store them in list ###
        for activitiesNode in root.findall('Activities'):  # findall because ProTrack constains 1 dummy "activities" node
            for activityNode in activitiesNode.findall('Activity'):
                activityID = int(activityNode.find('UniqueID').text)
                activityName = activityNode.find("Name").text
                usedDistributionId = int(activityNode.find("Distribution").text)
                # Note: the same distribution object is used for different activities pointing to the same sensitivity distribution
                activity_distribution = RiskAnalysisDistribution(DistributionType.MANUAL, ManualDistributionUnit.RELATIVE, 99,100,101,False) if usedDistributionId not in distribution_dict \
                                        else distribution_dict[usedDistributionId]

                ##BaseLineSchedule
                #Baseline duration (workinghours)
                baselineDuration_hours = int(activityNode.find('BaseLineDuration').text)
                baselineDuration = project_agenda.get_workingDuration_timedelta(duration_hours=BaselineDuration_hours)
                #BaseLineStart
                baseLineStart = self.getdate(activityNode.find('BaseLineStart').text, dateformat)
                #FixedBaselineCost
                baselineFixedCost = float(activityNode.find('FixedBaselineCost').text)
                #BaselineCostByUnit
                baselineCostByUnit = float(activityNode.find('BaselineCostByUnit').text)

                # calculate derived fields:
                baseline_enddate = project_agenda.get_end_date(BaseLineStart, baselineDuration.days, round(baselineDuration.seconds / 3600))
                baseline_var_cost = baselineDuration_hours * baselineCostByUnit
                # activity_total_cost != baselineFixedCost + baseline_var_cost, because resource cost should be incorporated! => calculate it afterwards
                activity_baselineScheduleRecord = BaselineScheduleRecord(baseLineStart, baseline_enddate, baselineDuration, baselineFixedCost, baselineCostByUnit, baseline_var_cost, 0, True)
                # add new activity to the dict of all read activities
                activity_dict[activity_ID] = Activity(activity_id= activity_ID, name= activityName, baseline_schedule= activity_baselineScheduleRecord, risk_analysis= activity_distribution, type_check= True)


        ### Read activitygroups and store them: ###
        for activityGroupsNode in root.findall("ActivityGroups"):
            for activityGroupNode in activityGroupsNode.findall("ActivityGroup"):
                #UniqueID
                activityID = int(activityGroupNode.find('UniqueID').text)
                #Name
                activityName = activityGroupNode.find('Name').text
                # add activity group to dict:
                activityGroup_dict[activityID] = Activity(activity_id= activityID, name= activityName)

        ### Read the activities outline tree for wbs and filter all activities and groups not in that tree ###
        outlineListNode = root.find("OutlineList").find("List")
        outlineListChildren = list(outlineListNode)
        if len(outlineListChildren) > 0:
            # add project activityGroup root:
            activityGroup_dict[0] = Activity(activity_id= 0, name= project_name, wbs_id= (1,))
            # Call recursive function here to process the children of an activityGroup:
            child_activity_Ids = self.set_wbs_read_activities_and_groups((1,), list(grandChildrenNode), activity_dict, activityGroup_dict, 0, activityGroup_to_childActivities_dict)
            activityGroup_to_childActivities_dict[0] = child_activity_Ids
            

        else:
            print("XMLparser:to_schedule_�object: Activity outline list is empty!")
            activity_dict = {}
            activityGroup_dict = {}

        # filter out any activityGroup or activity that has no wbs_id:
        activity_dict = {key: value for key, value in activity_dict.items() if value.wbs_id} 
        activityGroup_dict = {key: value for key, value in activityGroup_dict.items() if value.wbs_id}


        ### Read activity relations and add them directly to the corresponding activities ###
        lagTypeConversion_dict = {0: "SS", 1: "SF", 2:"FS", 3: "FF"}

        for relationsNode in root.findall('Relations'):
            for relationNode in relationsNode.findall('Relation'):
                predecessor = int(relationNode.find('FromTask').text)
                successor = int(relationNode.find('ToTask').text)
                lag = int(relationNode.find('Lag').text)
                lagKind = int(relationNode.find('LagKind').text)
                lagType = int(relationNode.find('LagType').text)
                
                if lagType not in lagTypeConversion_dict:
                    raise XMLParseError("to_schedule_object:Reading relations: Unexpected lagType, expects lagType in {0} but found {1}".format(list(lagTypeConversion_dict.keys()), lagType))

                lagType_str = lagTypeConversion_dict[lagType]
                if lagKind != 0:
                    print("XMLparser:to_schedule_object: Reading relations: Expects LagKing = 0 but found a relation with LagKind = {0} for relation id = {1}".format(lagKind, relationNode.find("UniqueID")))
                    lagType_str = "undefinedLagKind"

                successorTuple = (successor, lagType_str, lag)
                predecessorTuple = (predecessor, lagType_str, lag)

                # add tuples to the resource lists of the specified activities if both activities can be found:
                if predecessor in activity_dict and successor in activity_dict:
                    activity_dict[predecessor].successors.append(successorTuple)
                    activity_dict[successor].predecessors.append(predecessorTuple)
                #else:
                #    # found loose relation
                #    pass


        ### Read the activity resource assignments and add them directly to the corresponding activities:
        for resource_assignmentsNode in root.findall('ResourceAssignments'):
            for resource_assignmentNode in resource_assignmentsNode.findall('ResourceAssignment'):
                res_id = int(resource_assignmentNode.find('FIELD1025').text)
                activity_ID = int(resource_assignmentNode.find('FIELD1024').text)
                res_demand = ast.literal_eval(resource_assignmentNode.find('FIELD1026').text)
                res_fixed_assignment = False if int(resource_assignmentNode.find('FIELD1027').text) == 0 else True  # fixed assignment of consumable resources

                # look if both resource_id and activity_ID are found before adding them
                if res_id in res_dict and activity_ID in activity_dict:
                    # both Id's are found => add resource assignment
                    resource = res_dict[res_id]
                    # check if consumable resource if to add fixed assignment
                    res_assignment = (resource, res_demand, False if resource.resource_type != ResourceType.CONSUMABLE else res_fixed_assignment)
                    activity_dict[activity_ID].resources.append(res_assignment)
                #else:
                #    # resource assignment to a non-existing acitivity or resource: ignore
                #    pass

        # calculate and add total cost for each activity, resource_cost of activity
        self.update_activities_aggregated_costs(list(activity_dict.values()), project_agenda)
        self.update_activityGroups_aggregated_values(list(activityGroup_dict.values()), activity_dict, activityGroup_to_childActivities_dict, project_agenda)

        # check total resource cost with read value
        #DEBUG: #TODO

        ### Read trackinglist:
        for trackingListNode in root.findall("TrackingList"):
            trackingPeriodHeaderNode_list = trackingListNode.findall("TrackingPeriod")
            if not trackingPeriodHeaderNode_list:
                # no tracking periods found
                continue
            # tracking periods found:
            trackingPeriodNode_list = trackingListNode.findall("TProTrackActivities-1")

            if len(trackingPeriodHeaderNode_list) != len(trackingPeriodNode_list):
                raise XMLParseError("to_schedule_object:Reading TrackingList: Inconsistent number of TrackingPeriod nodes w.r.t. TProTrackActivities-1 nodes: {0} != {1}".format(
                    len(trackingPeriodHeaderNode_list), len(trackingPeriodNode_list)))

            # process each tracking period:  
            for i in range(0, len(trackingPeriodHeaderNode_list)):
                # read tracking period header first:
                name = trackingPeriodHeaderNode_list[i].find('Name').text
                statusdate = trackingPeriodHeaderNode_list[i].find('EndDate').text
                statusdate_datetime = self.getdate(statusdate, dateformat)

                # first, make a TrackingPeriod object
                currentTrackingPeriod = TrackingPeriod(tracking_period_name= name, tracking_period_statusdate= statusdate_datetime)

                # next, process all activity tracking records:
                activityTrackingHeader_nodes = trackingPeriodNode_list[i].findall("Activity")
                activityTrackingRecord_nodes = trackingPeriodNode_list[i].findall("TProTrackActivityTracking-1")
                if len(activityTrackingHeader_nodes) != len(activityTrackingRecord_nodes):
                    raise XMLParseError("to_schedule_object:Reading TrackingList: Inconsistent number of Activity nodes w.r.t. TProTrackActivityTracking-1 nodes in TrackingPeriod {0}: {1} != {2}".format( name,
                        len(activityTrackingHeader_nodes), len(activityTrackingRecord_nodes)))

                currentTrackingPeriod_records_dict = {}  # dict to link an activityId with its trackingRecord in the current trackingPeriod
                # process all activity tracking records:
                for j in range(0, len(activityTrackingHeader_nodes)):
                    activityID = int(activityTrackingHeader_nodes[j].text)

                    if activityID in activity_dict:
                        # valid activity to track:
                        activity_record = self.process_activityTrackingRecord_Node(activity_dict[activityID], activityTrackingRecord_nodes[j], statusdate_datetime, project_agenda)
                        currentTrackingPeriod_records_dict[activityID] = activity_record

                    elif activityID in activityGroup_dict:
                        raise NotImplementedError("""XMLparser:to_schedule_object: Tracking records for activityGroups is a feature currently not implemented in PMConverter. 
                                                TrackingRecord found for activityGroupid = {0}""".format(activityID))
                    #else:
                    #    # invalid tracking record.
                    #    # could be of an acitivity we filtered out because it is not present in outlinelist.
                    #    pass

                #endFor activityTrackingRecords

                # calculate aggregated values for activityGroups and add them to currentTrackingPeriod:
                for activityGroupId in activityGroup_dict.keys():
                    childActivityIds = activityGroup_to_childActivities_dict[activityGroupId]
                    # construct the activityGroup trackingRecord:
                    activityGroup_trackingRecord = self.construct_activityGroup_trackingRecord(activityGroup_dict[activityGroupId], childActivityIds, currentTrackingPeriod_records_dict)
                    currentTrackingPeriod.tracking_period_records.append(activityGroup_trackingRecord)

                # add all activity records to trackingPeriod:
                currentTrackingPeriod.tracking_period_records.extend(list(currentTrackingPeriod_records_dict.values()))

                # sort currentTrackingPeriod on wbs:
                currentTrackingPeriod.tracking_period_records = sorted(currentTrackingPeriod.tracking_period_records, key=lambda activityTrackingRecord: activityTrackingRecord.activity.wbs_id)

                # add trackingperiod to list
                trackingPeriodsList.append(currentTrackingPeriod)
            #endFor processing trackingPeriods
        #endFor processing trackingLists

        # sort tracking periods on status_date
        trackingPeriodsList = sorted(trackingPeriodsList, key= lambda trackingPeriod: trackingPeriod.tracking_period_statusdate)

        # combine activities and activityGroups in 1 list:
        activities_list = list(activity_dict.values())
        activities_list.extend(list(activityGroup_dict.values()))

        # sort resources:
        resources_list = sorted(list(res_dict.values()), key= lambda resource: resource.resource_id)

        # sort activities_list on wbs:
        activities_list = sorted(activities_list, key= lambda activity: activity.wbs_id)

        ## Make project object
        return ProjectObject(project_name, activities_list, trackingPeriodsList, resources_list, project_agenda)

    @staticmethod
    def xml_escape(text):
        "Replaces undesired characters by escaped variant"
        xml_escape_table = {"&": "&amp;"}
        return "".join(xml_escape_table.get(c,c) for c in text)

    def from_schedule_object(self, project_object, file_path_output="output.xml"):

        ### Sort activities based on WBS
        activity_list_wbs=sorted(project_object.activities, key=lambda x: x.wbs_id)
        project_object.activities=activity_list_wbs

        file = open(file_path_output, 'w')

        file.write('<Project>')

        ### Project name ###
        file.write('<NAME>')
        project_name=project_object.activities[0].name
        file.write(project_name)
        file.write('</NAME>')

        ### Projectinfo (unimportant) ###
        file.write("<ProjectInfo><LastSavedBy>PMConverter</LastSavedBy><Name>ProjectInfo</Name><SavedWithMayorBuild>3</SavedWithMayorBuild><SavedWithMinorBuild>0</SavedWithMinorBuild>")
        file.write("<SavedWithVersion>0</SavedWithVersion><UniqueID>-1</UniqueID><UserID>0</UserID></ProjectInfo>")

        ### Settings (Somewhat important) ###
        ## A lot of this info can (and should probably) deleted
        file.write("<Settings><AbsProjectBuffer>311220071300</AbsProjectBuffer><ActionEndThreshold>100</ActionEndThreshold><ActionStartThreshold>60</ActionStartThreshold>")
        file.write("<ActiveSensResult>1</ActiveSensResult><ActiveTrackingPeriod>8</ActiveTrackingPeriod><AllocationMethod>0</AllocationMethod><AutomaticBuffer>0</AutomaticBuffer>")
        file.write("<ConnectResourceBars>0</ConnectResourceBars><ConstraintHardness>3</ConstraintHardness><CurrencyPrecision>2</CurrencyPrecision><CurrencySymbol></CurrencySymbol>")
        file.write("<CurrencySymbolPosition>1</CurrencySymbolPosition>")
        ### Dateformat ###
        ## TODO; Read Datetimeformat
        dateformat="d/MM/yyyy h:mm"
        file.write('<DateTimeFormat>')
        file.write(dateformat)
        file.write("</DateTimeFormat>")

        file.write("<DefaultRowBuffer>50</DefaultRowBuffer><DrawRelations>1</DrawRelations><DrawShadow>1</DrawShadow><DurationFormat>1</DurationFormat><DurationLevels>2</DurationLevels>")
        file.write("<ESSLSSFloat>0</ESSLSSFloat><GanttStartDate>080220070800</GanttStartDate><GanttZoomLevel>0.0127138157894736</GanttZoomLevel><GroupFilter>0</GroupFilter><HideGraphMarks>0</HideGraphMarks>")
        file.write("<Name>Settings</Name><PlanningEndThreshold>60</PlanningEndThreshold><PlanningStartThreshold>20</PlanningStartThreshold><PlanningUnit>1</PlanningUnit>")
        file.write("<ResAllocation1Color>12632256</ResAllocation1Color><ResAllocation2Color>8421504</ResAllocation2Color><ResAvailableColor>15780518</ResAvailableColor>")
        file.write("<ResourceChartEndDate>220520071000</ResourceChartEndDate><ResourceChartStartDate>060320060000</ResourceChartStartDate><ResOverAllocationColor>255</ResOverAllocationColor>")
        file.write("<ShowCanEditResultsInHelp>1</ShowCanEditResultsInHelp><ShowCriticalPath>1</ShowCriticalPath><ShowInputModelInfoInHelp>1</ShowInputModelInfoInHelp>")
        file.write("<SyncGanttAndResourceChart>0</SyncGanttAndResourceChart><UniqueID>-1</UniqueID><UseResourceScheduling>0</UseResourceScheduling><UserID>0</UserID><ViewDateTimeAsUnits>0</ViewDateTimeAsUnits>")
        file.write('</Settings>')

        ## Defaults: Obsolete? ###
        file.write("<Defaults><DefaultCostPerUnit>50</DefaultCostPerUnit><DefaultDisplayDurationType>0</DefaultDisplayDurationType><DefaultDistributionType>2</DefaultDistributionType>")
        file.write("<DefaultDurationInput>0</DefaultDurationInput><DefaultLagTime>0</DefaultLagTime><DefaultNumberOfSimulationRuns>100</DefaultNumberOfSimulationRuns>")
        file.write("<DefaultNumberOfTrackingPeriodsGeneration>20</DefaultNumberOfTrackingPeriodsGeneration><DefaultNumberOfTrackingPeriodsSimulation>50</DefaultNumberOfTrackingPeriodsSimulation>")
        file.write("<DefaultRelationType>2</DefaultRelationType><DefaultResourceRenewable>1</DefaultResourceRenewable><DefaultSimulationType>0</DefaultSimulationType><DefaultStartPage>start.html</DefaultStartPage>")
        file.write("<DefaultTaskDuration>10</DefaultTaskDuration><DefaultTrackingPeriodOffset>50</DefaultTrackingPeriodOffset><DefaultWorkingDaysPerWeek>5</DefaultWorkingDaysPerWeek>")
        file.write("<DefaultWorkingHoursPerDay>8</DefaultWorkingHoursPerDay><Name>Defaults</Name><UniqueID>-1</UniqueID><UserID>0</UserID></Defaults>")

        ### Agenda ###
        #Startdate
        file.write("<Agenda>")
        file.write("<StartDate>")
        startdate=self.get_date_string(project_object.activities[0].baseline_schedule.start, dateformat)
        file.write(startdate)
        file.write("</StartDate>")
        #Non workinghours
        file.write("<NonWorkingHours>")
        for i in range(0,24):
            if project_object.agenda.working_hours[i] == False:
                file.write("<Hour>"+str(i)+"</Hour>")
        file.write("</NonWorkingHours>")
        #Non working days
        file.write("<NonWorkingDays>")
        for i in range(0,7):
            if project_object.agenda.working_days[i] == False:
                file.write("<Day>"+str(i)+"</Day>")
        file.write("</NonWorkingDays>")
        #Holidays
        for holiday in project_object.agenda.holidays:
            file.write("<Holidays>")
            file.write(str(holiday)+"0000")
            file.write("</Holidays>")
        file.write("</Agenda>")

        ### Activities ###
        file.write("""<Activities><Name>Activities</Name><UniqueID>-1</UniqueID><UserID>0</UserID></Activities><Activities>""")


        for activity in project_object.activities:
            if len(activity.wbs_id) == 3:
                # No Activity groups
                file.write("<Activity>")
                # hourly cost
                file.write("<BaselineCostByUnit>")
                hourly_cost=activity.baseline_schedule.hourly_cost
                file.write(str(hourly_cost))
                file.write("</BaselineCostByUnit>")
                # Duration (hours)
                file.write("<BaseLineDuration>")
                duration=str(activity.baseline_schedule.duration.days*8)
                file.write(duration)
                file.write("</BaseLineDuration>")
                # BaselineStart
                file.write("<BaseLineStart>")
                BaselineStartString=self.get_date_string(activity.baseline_schedule.start,dateformat)
                file.write(BaselineStartString)
                file.write("</BaseLineStart>")
                # Constraints (Unimportant)
                file.write("""<Constraints><Direction>0</Direction><DueDateEnd>010101000000</DueDateEnd><DueDateStart>010101000000</DueDateStart><LockedTimeEnd>010101000000</LockedTimeEnd>""")
                file.write("<LockedTimeStart>010101000000</LockedTimeStart><Name/><ReadyTimeEnd>010101000000</ReadyTimeEnd><ReadyTimeStart>010101000000</ReadyTimeStart>")
                file.write("<UniqueID>-1</UniqueID><UserID>0</UserID></Constraints>")
                # Distribution
                file.write("<Distribution>")
                distr=str(activity.risk_analysis.distr_id)
                file.write(distr)
                file.write("</Distribution>")
                # DurationCPMunits = Baselineduration?
                file.write("<DurationCPMUnits>")
                file.write(duration)
                file.write("</DurationCPMUnits>")
                # Fixed Cost
                file.write("<FixedBaselineCost>")
                fixed_cost=str(activity.baseline_schedule.fixed_cost)
                file.write(fixed_cost)
                file.write("</FixedBaselineCost>")
                # Milestone (Unimportant)
                file.write("<IsMilestone>0</IsMilestone>")
                # Activity name
                file.write("<Name>")
                name=str(XMLParser.xml_escape(activity.name)).lstrip(' ')
                file.write(name)
                file.write("</Name>")
                # ?
                file.write("<StartCPMUnits>0</StartCPMUnits>")
                # Activity ID
                file.write("<UniqueID>")
                activity_id=str(activity.activity_id)
                file.write(activity_id)
                file.write("</UniqueID>")
                # Unimportant
                file.write("<UserID>0</UserID>")
                file.write("</Activity>")
        file.write("</Activities>")



        ### Relations ###
        file.write("""<Relations><Name>Relations</Name><UniqueID>-1</UniqueID><UserID>0</UserID></Relations><Relations>""")
        # Start with ID = 5? (Shouldn't matter anyway)
        unique_resource_id_counter=4
        for activity in project_object.activities:
            if len(activity.successors) !=0:
                for successor in activity.successors:
                    unique_resource_id_counter+=1
                    file.write("<Relation>")
                    # From
                    file.write("<FromTask>")
                    fromtask_id=str(activity.activity_id)
                    file.write(fromtask_id)
                    file.write("</FromTask>")
                    # Lag
                    file.write("<Lag>")
                    lag=str(successor[2])
                    file.write(lag)
                    file.write("</Lag>")
                    # LagKind/Type
                    if successor[1] == "FS":
                        lagKind="0"
                        lagType="2"
                    elif successor[1] == "FF":
                        lagKind="0"
                        lagType="3"
                    elif successor[1] == "SS":
                        lagKind="0"
                        lagType="0"
                    elif successor[1] == "SF":
                        lagKind="0"
                        lagType="1"
                    else:
                        raise "Lag Undefined"
                    file.write("<LagKind>")
                    file.write(lagKind)
                    file.write("</LagKind><LagType>")
                    file.write(lagType)
                    file.write("</LagType>")
                    # Name (Unimportant)
                    file.write("<Name>Relation</Name>")
                    # To
                    file.write("<ToTask>")
                    totask_id=str(successor[0])
                    file.write(totask_id)
                    file.write("</ToTask>")
                    # Unique Relation ID
                    file.write("<UniqueID>")
                    file.write(str(unique_resource_id_counter))
                    file.write("</UniqueID>")
                    # UserID (Unimportant)
                    file.write("<UserID>0</UserID>")
                    file.write("</Relation>")
        file.write("</Relations>")

        ### Activity Groups ###
        file.write("""<ActivityGroups><Name>ActivityGroups</Name><UniqueID>-1</UniqueID><UserID>0</UserID></ActivityGroups><ActivityGroups>""")
        for activitygroup in project_object.activities:
            if len(activitygroup.wbs_id) == 2:
                # Only Activity groups
                file.write("<ActivityGroup>")
                # Unimportant
                file.write("<Expanded>1</Expanded>")
                # Name
                file.write("<Name>")
                name=str(XMLParser.xml_escape(activitygroup.name)).lstrip(' ')
                file.write(name)
                file.write("</Name>")
                # ID
                file.write("<UniqueID>")
                activity_id=str(activitygroup.activity_id)
                file.write(activity_id)
                file.write("</UniqueID>")
                # Unimportant
                file.write("<UserID>0</UserID>")
                file.write("</ActivityGroup>")
        file.write("</ActivityGroups>")

        ### Outline list ###
        file.write("<OutlineList><List>")
        for i in range(1,len(project_object.activities)):
            # Activity group
            if len(project_object.activities[i].wbs_id) == 2:
                k=1
                file.write("<Child><Type>1</Type><Data>")
                # ID
                id = str(project_object.activities[i].activity_id)
                file.write(id)
                file.write("</Data>")
                file.write("<Expanded>1</Expanded><List>")
                # Activities belonging to Activity group
                while i+k < len(project_object.activities) and len(project_object.activities[i+k].wbs_id) == 3:
                        file.write("<Child><Type>2</Type><Data>")
                        id = str(project_object.activities[i+k].activity_id)
                        file.write(id)
                        file.write("</Data></Child>")
                        k+=1
                file.write("</List>")
                file.write("</Child>")
        file.write("</List>")
        file.write("</OutlineList>")

        ### Resources ###
        file.write("""<Resources><Name>Resources</Name><UniqueID>-1</UniqueID><UserID>0</UserID></Resources>""")
        file.write("<Resources>")
        for resource in project_object.resources:
            file.write("<Resource>")
            # resource ID
            file.write("<FIELD0>")
            res_id=str(resource.resource_id)
            file.write(res_id)
            file.write("</FIELD0>")
            #
            file.write("<FIELD1>")
            res_name=str(XMLParser.xml_escape(resource.name))
            file.write(res_name)
            file.write("</FIELD1>")
            file.write("<FIELD768></FIELD768>")
            # Availability
            file.write("<FIELD778>")
            availability_string="#"
            availability_string+=str(resource.availability)
            file.write(availability_string)
            file.write("</FIELD778>")
            # Renewable
            file.write("<FIELD769>")
            if resource.resource_type == ResourceType.RENEWABLE:
                ren_string="1"
            else:
                ren_string="0"
            file.write(ren_string)
            file.write("</FIELD769>")
            # Cost per use
            file.write("<FIELD770>")
            costperuse_string= str(resource.cost_use)
            file.write(costperuse_string)
            file.write("</FIELD770>")
            # Cost per unit
            file.write("<FIELD771>")
            costperunit_string=str(resource.cost_unit)
            file.write(costperunit_string)
            file.write("</FIELD771>")
            # Total cost ? not in resource information
            #file.write("<FIELD776></FIELD776>")
            # Availability
            file.write("<FIELD780>")
            file.write(str(resource.availability))
            file.write("</FIELD780>")
            # ?
            file.write("""<FIELD779>-1</FIELD779>""")
            #<FIELD781></FIELD781><FIELD782></FIELD782><FIELD783></FIELD783>""")
            file.write("</Resource>")
            file.write("<DATEPERCSERIE><DEFAULTAVAILABILITY>")
            file.write(str(resource.availability))
            file.write("</DEFAULTAVAILABILITY><DATEPERCBREAKPOINTS/></DATEPERCSERIE>")
        file.write("</Resources>")


        ### Resource Assignment ###
        file.write("""<ResourceAssignments><Name>ResourceAssignments</Name><UniqueID>-1</UniqueID><UserID>0</UserID></ResourceAssignments><ResourceAssignments>""")
        for activity in project_object.activities:
            if len(activity.resources) !=0:
                for resourceTuple in activity.resources:
                    file.write("<ResourceAssignment>")
                    # Resources needed
                    file.write("<FIELD1026>")
                    res_needed=str(resourceTuple[1])
                    file.write(res_needed)
                    file.write("</FIELD1026>")
                    # Unimportant
                    file.write("<FIELD1027>0</FIELD1027>")
                    # Activity ID
                    file.write("<FIELD1024>")
                    activity_id=str(activity.activity_id)
                    file.write(activity_id)
                    file.write("</FIELD1024>")
                    # Resource ID
                    file.write("<FIELD1025>")
                    res_id=str(resourceTuple[0].resource_id)
                    file.write(res_id)
                    file.write("</FIELD1025>")
                    file.write("</ResourceAssignment>")
        file.write("</ResourceAssignments>")



        ### Sort activities based on ID
        activity_list_id=list2=sorted(project_object.activities, key=lambda x: x.activity_id)
        project_object.activities=activity_list_id

        ### TrackingPeriods ####
        file.write("""<TrackingList><Name>TrackingList</Name><UniqueID>-1</UniqueID><UserID>0</UserID></TrackingList>""")
        file.write("<TrackingList>")
        TP_count=0
        for TP in project_object.tracking_periods:
            ## Tracking Period Info
            TP_count+=1
            file.write("""<TrackingPeriod><Abreviation></Abreviation><Description/>""")
            # Enddate
            file.write("<EndDate>")
            enddate_string=self.get_date_string(TP.tracking_period_statusdate,dateformat)
            file.write(enddate_string)
            file.write("</EndDate>")
            # Name
            file.write("<Name>")
            # Weird bug, XML won't be correctly formatted if TP name is written => replace unwanted characters!
            TP_name=str(TP.tracking_period_name)
            file.write(XMLParser.xml_escape(TP_name))
            file.write("</Name>")
            file.write("<PredictiveLogic>0</PredictiveLogic>")
            # Unique ID
            file.write("<UniqueID>")
            file.write(str(TP_count))
            file.write("</UniqueID>")
            file.write("<UserID>0</UserID>")
            file.write("</TrackingPeriod>")
            file.write("<TProTrackActivities-1>")
            # Activity Tracking
            # Needs to be sorted on ID?
            for activity in project_object.activities:

                # No Activity groups
                if len(activity.wbs_id) == 3:

                    # Activity ID
                    file.write("<Activity>")
                    file.write(str(activity.activity_id))
                    file.write("</Activity>")
                    file.write("<TProTrackActivityTracking-1>")
                    for ATR in TP.tracking_period_records:
                        if activity.activity_id == ATR.activity.activity_id:
                            #Actual Start
                            file.write("<ActualStart>")

                            if ATR.actual_start != None:
                                actual_start_str=self.get_date_string(ATR.actual_start,dateformat)
                                file.write(actual_start_str)
                            else:
                                file.write("0")

                            file.write("</ActualStart>")
                            #Actual Duration
                            file.write("<ActualDuration>")
                            file.write(str(ATR.actual_duration.days*8))
                            file.write("</ActualDuration>")
                            #Actual Cost Dev
                            file.write("<ActualCostDev>")
                            file.write(str(ATR.deviation_pac))
                            file.write("</ActualCostDev>")
                            # Remaining Duration
                            file.write("<RemainingDuration>")
                            if ATR.remaining_duration != None:
                                file.write(str(ATR.remaining_duration.days*8))
                            else:
                                file.write("0")
                            file.write("</RemainingDuration>")
                            # Remaining cost dev
                            file.write("<RemainingCostDev>")
                            file.write(str(ATR.deviation_prc))
                            file.write("</RemainingCostDev>")
                            #Percentage Complete
                            file.write("<PercentageComplete>")
                            file.write(str(ATR.percentage_completed/100))
                            file.write("</PercentageComplete>")
                    file.write("</TProTrackActivityTracking-1>")
            file.write("</TProTrackActivities-1>")
        file.write("</TrackingList>")


        ### Distributions ### # TODO Does not work correctly
        # file.write("""<SensitivityDistributions><Name>SensitivityDistributions</Name><UniqueID>-1</UniqueID>
        # <UserID>0</UserID></SensitivityDistributions><SensitivityDistributions>""")
        #
        # ## To many distributions will be created (Standard distrbutions will be recreated)
        # counter=5
        # for counter in range(5,len(project_object.activities)):
        #     notdone=1
        #     for activity in project_object.activities:
        #         distr=activity.risk_analysis
        #         if distr.distr_id == counter and notdone:
        #             #counter+=1
        #             file.write("<TProTrackSensitivityDistribution"+str(counter)+">")
        #             file.write("<UniqueID>"+str(counter)+"</UniqueID>")
        #             file.write("<Name>"+distr.distr_name+"</Name>")
        #             file.write("<StartDuration>"+str(int(distr.optimistic_duration-1))+"</StartDuration>")
        #             file.write("<EndDuration>"+str(int(distr.pessimistic_duration+1))+"</EndDuration>")
        #             file.write("<Style>0</Style>")
        #             file.write("<Distribution>")
        #             file.write("<X>"+str(distr.optimistic_duration)+"</X><Y>0</Y>")
        #             file.write("<X>"+str(distr.probable_duration)+"</X><Y>100</Y>")
        #             file.write("<X>"+str(distr.pessimistic_duration)+"</X><Y>0</Y>")
        #             file.write("</Distribution>")
        #             file.write("</TProTrackSensitivityDistribution"+str(counter)+">")
        #             notdone=0
        # file.write("</SensitivityDistributions>")
        file.write('</Project>')
        file.close()



        return True #print("Write Succesful!")


