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
        current_wbs_list = list(parent_wbs)
        current_wbs_list.append(1)
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
                        grandchild_activityIds = self.set_wbs_read_activities_and_groups(tuple(current_wbs_list), list(grandChildrenNode), activity_dict, activityGroup_dict, 
                                                                                         activityGroupId, activityGroup_to_childActivities_dict)
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
            activityDuration_hours = activity.baseline_schedule.duration.days * agenda.get_working_hours_in_a_day() + activity.baseline_schedule.duration.seconds / 3600
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
                        resource_cost += resource.cost_use + resourceTuple[1] * resource.cost_unit * activityDuration_hours
                else:
                    #resource type is renewable:
                    #add cost_use and variable cost:
                    resource_cost += resourceTuple[1] * (resource.cost_use + resource.cost_unit * activityDuration_hours)
            #endFor adding resource costs

            activity.resource_cost = resource_cost

            activityDuration_hours = activity.baseline_schedule.duration.days * agenda.get_working_hours_in_a_day() + round(activity.baseline_schedule.duration.seconds / 3600)

            # calculate total activity cost:
            activity.baseline_schedule.total_cost = float(activity.resource_cost + activity.baseline_schedule.fixed_cost + activity.baseline_schedule.hourly_cost * activityDuration_hours)
            

    def process_activityTrackingRecord_Node(self, activity, activityTrackingRecord_node, statusdate_datetime, agenda, datetimeFormat):
        """"
        This function processes an acitivityTrackingRecord node from ProTrack p2x file.
        :returns: ActivityTrackingRecord
        """
        # Read Data
        actualStart_field = activityTrackingRecord_node.find('ActualStart').text
        actualStart = self.getdate(actualStart_field, datetimeFormat) if actualStart_field != "0" else datetime.max
        actualDuration_hours = int(activityTrackingRecord_node.find('ActualDuration').text)
        actualDuration = agenda.get_workingDuration_timedelta(duration_hours=actualDuration_hours)
        actualCostDev = float(activityTrackingRecord_node.find('ActualCostDev').text)
        remainingDuration_hours = int(activityTrackingRecord_node.find('RemainingDuration').text)
        remainingDuration = agenda.get_workingDuration_timedelta(duration_hours=remainingDuration_hours)
        remainingCostDev = float(activityTrackingRecord_node.find('RemainingCostDev').text)
        percentageComplete = float(activityTrackingRecord_node.find('PercentageComplete').text)*100
        if percentageComplete >= 100:
            trackingStatus='Finished'
        elif abs(percentageComplete) < 1e-5: # compare float with 0
            trackingStatus='Not Started'
        else:
            trackingStatus='Started'

        # derive remaining data fields:
        #planned_actual_cost = 0.
        #planned_remaining_cost = 0.
        #actual_cost = 0.
        #remaining_cost = 0.
        #earned_value = 0.
        #planned_value = 0.

        actual_cost, earned_value, planned_actual_cost, planned_remaining_cost, planned_value, remaining_cost = ActivityTrackingRecord.calculate_activityTrackingRecord_derived_values(activity,
                                                                                actualCostDev, actualDuration_hours, agenda, percentageComplete, remainingCostDev, remainingDuration_hours, statusdate_datetime, actualStart)
            
        
        return ActivityTrackingRecord(activity, actualStart, actualDuration, planned_actual_cost, float(planned_remaining_cost), remainingDuration, actualCostDev,
                                                        remainingCostDev, actual_cost, remaining_cost, int(round(percentageComplete)), trackingStatus, float(earned_value), float(planned_value), True)


    def to_schedule_object(self, file_path_input):
        tree = ET.parse(file_path_input)
        root = tree.getroot()

        # read dicts from ProTrack file:
        activity_dict = {}
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

                    distributionXPointsList = []
                    distributionYPointsList = []
                    for xPoint in distributionXPoints:
                        distributionXPointsList.append(int(xPoint.text))
                    for yPoint in distributionYPoints:
                        distributionYPointsList.append(int(yPoint.text))

                    if len(distributionXPointsList) != 3 or len(distributionYPointsList) != 3 or not (distributionYPointsList == [0, 100, 0]):
                        # unsupported distribution type: don't fail conversion by this
                        print("XMLparser:to_schedule_object: Only sensitivity distributions defined by 3 points are supported!\n Not with {0} points.".format(len(distributionXPoints)))
                        continue
                    # valid 3 point distribution here:
                    distribution_dict[distributionID] = RiskAnalysisDistribution(distribution_type=DistributionType.MANUAL, distribution_units=distribution_units,
                                                                  optimistic_duration=distributionXPointsList[0],probable_duration=distributionXPointsList[1], pessimistic_duration=distributionXPointsList[2])


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
                baselineDuration = project_agenda.get_workingDuration_timedelta(duration_hours=baselineDuration_hours)
                #BaseLineStart
                baseLineStart = self.getdate(activityNode.find('BaseLineStart').text, dateformat)
                #FixedBaselineCost
                baselineFixedCost = float(activityNode.find('FixedBaselineCost').text)
                #BaselineCostByUnit
                baselineCostByUnit = float(activityNode.find('BaselineCostByUnit').text)

                # calculate derived fields:
                baseline_enddate = project_agenda.get_end_date(baseLineStart, baselineDuration.days, round(baselineDuration.seconds / 3600))
                baseline_var_cost = baselineDuration_hours * baselineCostByUnit
                # activity_total_cost != baselineFixedCost + baseline_var_cost, because resource cost should be incorporated! => calculate it afterwards
                activity_baselineScheduleRecord = BaselineScheduleRecord(baseLineStart, baseline_enddate, baselineDuration, baselineFixedCost, baselineCostByUnit, baseline_var_cost, 0., True)
                # add new activity to the dict of all read activities
                activity_dict[activityID] = Activity(activity_id= activityID, name= activityName, baseline_schedule= activity_baselineScheduleRecord, risk_analysis= activity_distribution, type_check= True)


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
            child_activity_Ids = self.set_wbs_read_activities_and_groups((1,), outlineListChildren, activity_dict, activityGroup_dict, 0, activityGroup_to_childActivities_dict)
            activityGroup_to_childActivities_dict[0] = child_activity_Ids
            

        else:
            print("XMLparser:to_schedule_object: Activity outline list is empty!")
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
        Activity.update_activityGroups_aggregated_values(list(activityGroup_dict.values()), activity_dict, activityGroup_to_childActivities_dict, project_agenda)

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
                        activity_record = self.process_activityTrackingRecord_Node(activity_dict[activityID], activityTrackingRecord_nodes[j], statusdate_datetime, project_agenda, dateformat)
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
                    activityGroup_trackingRecord = ActivityTrackingRecord.construct_activityGroup_trackingRecord(activityGroup_dict[activityGroupId], childActivityIds, currentTrackingPeriod_records_dict, statusdate_datetime, project_agenda)
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

    def generate_ProTrack_projectInfo(self):
        """
        This function generates the default projectInfo node for p2x.
        :returns: xml Elementnode
        """
        defaultFormat = """<ProjectInfo>
	        <LastSavedBy>PMConverter</LastSavedBy>
	        <Name>ProjectInfo</Name> 									
	        <SavedWithMayorBuild>3</SavedWithMayorBuild>
	        <SavedWithMinorBuild>0</SavedWithMinorBuild>
	        <SavedWithVersion>0</SavedWithVersion>
	        <UniqueID>-1</UniqueID>
	        <UserID>0</UserID>
        </ProjectInfo>"""
        return ET.fromstring(defaultFormat)

    def generate_ProTrack_settings(self, projectStartDatetime, projectBaselineEndDatetime, datetimeFormat):
        """
        This function generates the default settings node for p2x.
        :returns: xml Elementnode
        """
        defaultFormat = """<Settings>
            <AbsProjectBuffer>290420150000</AbsProjectBuffer>
            <ActionEndThreshold>100</ActionEndThreshold>
            <ActionStartThreshold>60</ActionStartThreshold>
            <ActiveSensResult>12</ActiveSensResult>
            <ActiveTrackingPeriod>-1</ActiveTrackingPeriod>
            <AllocationMethod>1</AllocationMethod>
            <AutomaticBuffer>1</AutomaticBuffer>
            <ConnectResourceBars>0</ConnectResourceBars>
            <ConstraintHardness>3</ConstraintHardness>
            <CurrencyPrecision>2</CurrencyPrecision>
            <CurrencySymbol>{0}</CurrencySymbol>
            <CurrencySymbolPosition>2</CurrencySymbolPosition>
            <DateTimeFormat>d/MM/yyyy h:mm</DateTimeFormat>
            <DefaultRowBuffer>50</DefaultRowBuffer>
            <DrawRelations>1</DrawRelations>
            <DrawShadow>1</DrawShadow>
            <DurationFormat>2</DurationFormat>
            <DurationLevels>2</DurationLevels>
            <ESSLSSFloat>-99999999</ESSLSSFloat>
            <GanttStartDate>290420150000</GanttStartDate>
            <GanttZoomLevel>1.0</GanttZoomLevel>
            <GroupFilter>0</GroupFilter>
            <HideGraphMarks>0</HideGraphMarks>
            <Name>Settings</Name>
            <PlanningEndThreshold>60</PlanningEndThreshold>
            <PlanningStartThreshold>20</PlanningStartThreshold>
            <PlanningUnit>1</PlanningUnit>
            <ResAllocation1Color>12632256</ResAllocation1Color>
            <ResAllocation2Color>8421504</ResAllocation2Color>
            <ResAvailableColor>15780518</ResAvailableColor>
            <ResourceChartEndDate>060520150000</ResourceChartEndDate>
            <ResourceChartStartDate>290420150000</ResourceChartStartDate>
            <ResOverAllocationColor>255</ResOverAllocationColor>
            <ShowCanEditResultsInHelp>1</ShowCanEditResultsInHelp>
            <ShowCriticalPath>1</ShowCriticalPath>
            <ShowInputModelInfoInHelp>1</ShowInputModelInfoInHelp>
            <SyncGanttAndResourceChart>0</SyncGanttAndResourceChart>
            <UniqueID>-1</UniqueID>
            <UseResourceScheduling>0</UseResourceScheduling>
            <UserID>0</UserID>
            <ViewDateTimeAsUnits>0</ViewDateTimeAsUnits>
        </Settings>""".format(u"\u20AC")  # euro sign in unicode
        settingsNode = ET.fromstring(defaultFormat)

        # perform customizations:
        startDate = self.get_date_string(projectStartDatetime, datetimeFormat) if projectStartDatetime < datetime.max else self.get_date_string(datetime.now(), datetimeFormat)
        endDate = self.get_date_string(projectBaselineEndDatetime, datetimeFormat) if projectBaselineEndDatetime > datetime.min else self.get_date_string(datetime.now(), datetimeFormat)

        XMLParser.find_xmlNode_and_set_text(settingsNode, "DateTimeFormat", datetimeFormat)

        # set some nodes to start date:
        startDateTags_list = ["AbsProjectBuffer", "GanttStartDate", "ResourceChartStartDate"]
        for nodeTag in startDateTags_list:
            XMLParser.find_xmlNode_and_set_text(settingsNode, nodeTag, startDate)
        # set some nodes to end date:
        XMLParser.find_xmlNode_and_set_text(settingsNode, "ResourceChartEndDate", endDate)
        return settingsNode

    def generate_ProTrack_defaults(self):
        """
        This function generates the default Defaults node for p2x.
        :returns: xml Elementnode
        """
        defaultFormat = """<Defaults>
            <DefaultCostPerUnit>50</DefaultCostPerUnit>
            <DefaultDisplayDurationType>0</DefaultDisplayDurationType>
            <DefaultDistributionType>-1</DefaultDistributionType>
            <DefaultDurationInput>0</DefaultDurationInput>
            <DefaultLagTime>0</DefaultLagTime>
            <DefaultNumberOfSimulationRuns>100</DefaultNumberOfSimulationRuns>
            <DefaultNumberOfTrackingPeriodsGeneration>20</DefaultNumberOfTrackingPeriodsGeneration>
            <DefaultNumberOfTrackingPeriodsSimulation>50</DefaultNumberOfTrackingPeriodsSimulation>
            <DefaultRelationType>2</DefaultRelationType>
            <DefaultResourceRenewable>1</DefaultResourceRenewable>
            <DefaultSimulationType>0</DefaultSimulationType>
            <DefaultStartPage>start.html</DefaultStartPage>
            <DefaultTaskDuration>10</DefaultTaskDuration>
            <DefaultTrackingPeriodOffset>50</DefaultTrackingPeriodOffset>
            <DefaultWorkingDaysPerWeek>5</DefaultWorkingDaysPerWeek>
            <DefaultWorkingHoursPerDay>8</DefaultWorkingHoursPerDay>
            <Name>Defaults</Name>
            <UniqueID>-1</UniqueID>
            <UserID>0</UserID>
        </Defaults>"""
        return ET.fromstring(defaultFormat)

    def generate_ProTrack_agenda(self, projectStartdatetime, datetimeFormat, agenda):
        """
        This function generates the default Agenda node for p2x.
        :returns: xml Elementnode
        """
        defaultFormat = """<Agenda>
            <StartDate/>
            <NonWorkingHours/>
            <NonWorkingDays/>
            <Holidays/>
        </Agenda>"""
        agendaNode = ET.fromstring(defaultFormat)

        startDate = self.get_date_string(projectStartdatetime, datetimeFormat) if projectStartdatetime < datetime.max else self.get_date_string(datetime.now(), datetimeFormat)
        XMLParser.find_xmlNode_and_set_text(agendaNode, "StartDate", startDate)
        # Non workinghours:
        for i in range(0,24):
            if agenda.working_hours[i] == False:
                XMLParser.find_xmlNode_and_append_childNode(agendaNode, "NonWorkingHours", "Hour", str(i))
        # non working days:
        for i in range(0,7):
            if agenda.working_days[i] == False:
                XMLParser.find_xmlNode_and_append_childNode(agendaNode, "NonWorkingDays", "Day", str(i))
        # holidays:
        for holiday in agenda.holidays:
            XMLParser.find_xmlNode_and_append_childNode(agendaNode, "Holidays", "Holiday", holiday + "0000")

        return agendaNode

    def generate_ProTrack_activities_and_riskDistributions(self, activities, agenda, datetimeFormat):
        """
        This function generates the default Activities node and the corresponding riskDistributionsNode for p2x.
        NOTE: Only give low-level activities to this function, no activityGroups!
        :returns: tuple of 4 xml Elementnodes: (activitiesPreformatNode, activitiesNode, riskDistributionsPreformatNode, riskDistributionsNode)
        """
        preFormatActivities = """<Activities>
            <Name>Activities</Name>
            <UniqueID>-1</UniqueID>
            <UserID>0</UserID>
        </Activities>"""
        defaultFormatActivities = "<Activities/>"
        activitiesNode = ET.fromstring(defaultFormatActivities)

        preFormatDistributions = """<SensitivityDistributions>
            <Name>SensitivityDistributions</Name>
            <UniqueID>-1</UniqueID>
            <UserID>0</UserID>
        </SensitivityDistributions>"""
        defaultFormatDistributions = "<SensitivityDistributions/>"
        riskDistributionsNode = ET.fromstring(defaultFormatDistributions)

        nextCustomRiskDistribution_id = 5
        standard_riskDistributions_dict = {StandardDistributionUnit.NO_RISK: "1", StandardDistributionUnit.SYMMETRIC: "2", StandardDistributionUnit.SKEWED_LEFT: "3", StandardDistributionUnit.SKEWED_RIGHT: "4"}

        for activity in activities:
            activityDistribution_id = standard_riskDistributions_dict[activity.risk_analysis.distribution_units]  if activity.risk_analysis.distribution_type == DistributionType.STANDARD else nextCustomRiskDistribution_id

            # append new activity to activitiesNode:
            self.generate_ProTrack_activityNode(activitiesNode, activity, activityDistribution_id, agenda, datetimeFormat)

            if activity.risk_analysis.distribution_type == DistributionType.MANUAL:
                # create custom risk distribution:
                self.generate_ProTrack_sensitivityDistribution(riskDistributionsNode, activity, activityDistribution_id)

                # increment risk distribution id for next activities
                nextCustomRiskDistribution_id += 1

        return ET.fromstring(preFormatActivities), activitiesNode, ET.fromstring(preFormatDistributions), riskDistributionsNode

    def generate_ProTrack_activityNode(self, activitiesNode, activity, activityDistribution_id, agenda, datetimeFormat):
        """
        This function generates the default Activity node for p2x and appends it to the given activitiesNode.
        """
        defaultFormat = """<Activity>
	        <BaselineCostByUnit/> 				
	        <BaseLineDuration/>        	 
	        <BaseLineStart/> 			
	        <Constraints> 										 
	            <Direction>0</Direction>                          
	            <DueDateEnd>010101000000</DueDateEnd>   			
	            <DueDateStart>010101000000</DueDateStart>         
	            <LockedTimeEnd>010101000000</LockedTimeEnd>       
	            <LockedTimeStart>010101000000</LockedTimeStart>   
	            <Name/>
	            <ReadyTimeEnd>010101000000</ReadyTimeEnd>         
	            <ReadyTimeStart>010101000000</ReadyTimeStart>     
	            <UniqueID>-1</UniqueID>                           
	            <UserID>0</UserID> 									
	        </Constraints>
	        <Distribution/>							
	        <DurationCPMUnits/>              
	        <FixedBaselineCost/>             
	        <IsMilestone>0</IsMilestone>                          
	        <Name/> 							
	        <StartCPMUnits>0</StartCPMUnits>
	        <UniqueID/>                             
	        <UserID>0</UserID>
        </Activity>"""
        activityNode = ET.fromstring(defaultFormat)

        customValues_dict = {}
        customValues_dict["BaselineCostByUnit"] = str(activity.baseline_schedule.hourly_cost)
        baselineDuration_hours = activity.baseline_schedule.duration.days * agenda.get_working_hours_in_a_day() + int(round(activity.baseline_schedule.duration.seconds / 3600))
        customValues_dict["BaseLineDuration"] = str(baselineDuration_hours)
        customValues_dict["BaseLineStart"] = self.get_date_string(activity.baseline_schedule.start, datetimeFormat)
        customValues_dict["Distribution"] = str(activityDistribution_id)
        customValues_dict["DurationCPMUnits"] = str(baselineDuration_hours)
        customValues_dict["FixedBaselineCost"] =str(activity.baseline_schedule.fixed_cost)
        customValues_dict["Name"] = self.xml_escape(activity.name)
        customValues_dict["UniqueID"] = str(activity.activity_id)

        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(activityNode, key, value)

        activitiesNode.append(activityNode)  # implicit return

    def generate_ProTrack_sensitivityDistribution(self, riskDistributionsNode, activity, distribution_id):
        """
        This function generates the default TrackSensitivityDistribution node for p2x and appends it to the given riskDistributionsNode.
        """
        defaultFormat = """<TProTrackSensitivityDistribution>
	        <UniqueID/>							
	        <Name/>
	        <StartDuration/>			
	        <EndDuration/> 					
	        <Style>0</Style> 								
	        <Distribution/>
        </TProTrackSensitivityDistribution>"""
        distributionNode = ET.fromstring(defaultFormat)
        # change tag name:
        distributionNode.tag += str(distribution_id)

        customValues_dict = {}
        customValues_dict["UniqueID"] = str(distribution_id)
        customValues_dict["Name"] = "Task " +  self.xml_escape(activity.name) + " distribution"
        customValues_dict["StartDuration"] = str(activity.risk_analysis.optimistic_duration)
        customValues_dict["EndDuration"] = str(activity.risk_analysis.pessimistic_duration)

        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(distributionNode, key, value)

        # append the 3 point distribution points:
        distributionPoints_node = distributionNode.find("Distribution")
        tagList = ["X", "Y"]
        valueList = [activity.risk_analysis.optimistic_duration, 0, activity.risk_analysis.probable_duration, 100, activity.risk_analysis.pessimistic_duration, 0]
        for i in range(0, len(valueList)):
            distributionPointNode = ET.Element(tagList[i % 2])
            distributionPointNode.text = str(valueList[i])
            distributionPoints_node.append(distributionPointNode)

        riskDistributionsNode.append(distributionNode)  # implicit return

    def generate_ProTrack_activityGroups(self, activityGroups):
        """
        This function generates the default ActivityGroups node for p2x.
        NOTE: Only give activityGroups to this function, no activities!
        :returns: tuple of 2 xml Elementnodes: (activityGroupsPreformatNode, activityGroupsNode)
        """
        preFormatActivities = """<ActivityGroups>
            <Name>ActivityGroups</Name>
            <UniqueID>-1</UniqueID>
            <UserID>0</UserID>
        </ActivityGroups>"""
        defaultFormatActivities = "<ActivityGroups/>"
        activityGroupsNode = ET.fromstring(defaultFormatActivities)

        for activityGroup in activityGroups:
            self.generate_ProTrack_activityGroupNode(activityGroupsNode, activityGroup)

        return ET.fromstring(preFormatActivities), activityGroupsNode

    def generate_ProTrack_activityGroupNode(self, activityGroupsNode, activityGroup):
        """
        This function generates the default ActivityGroup node for p2x and appends it to the given activityGroupsNode.
        """
        defaultFormat = """<ActivityGroup>
	        <Expanded>1</Expanded>
	        <Name/>
	        <UniqueID/>	
	        <UserID>0</UserID>
        </ActivityGroup>"""
        activityGroupNode = ET.fromstring(defaultFormat)

        customValues_dict = {}
        customValues_dict["Name"] = self.xml_escape(activityGroup.name)
        customValues_dict["UniqueID"] = str(activityGroup.activity_id)

        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(activityGroupNode, key, value)

        activityGroupsNode.append(activityGroupNode) # implicit return

    def generate_ProTrack_relations(self, activities):
        """
        This function generates the default Relations node for p2x.
        NOTE: Only give low-level activities to this function, no activityGroups!
        :returns: tuple of 2 xml Elementnodes: (relationsPreformatNode, relationsNode)
        """
        preFormatRelations = """<Relations>
            <Name>Relations</Name>
            <UniqueID>-1</UniqueID>
            <UserID>0</UserID>
        </Relations>"""
        defaultFormatRelations = "<Relations/>"
        relationsNode = ET.fromstring(defaultFormatRelations)

        relation_id = 1
        for activity in activities:
            for successor in activity.successors:
                self.generate_ProTrack_relation(relationsNode, activity.activity_id ,successor, relation_id)
                relation_id += 1

        return ET.fromstring(preFormatRelations), relationsNode

    def generate_ProTrack_relation(self, relationsNode, predecessor_id, successor, relation_id):
        """
        This function generates the default ActivityGroup node for p2x and appends it to the given relationsNode.
        """
        defaultFormat = """<Relation>
	        <FromTask/>
	        <Lag/> 			
	        <LagKind>0</LagKind> 
	        <LagType/> 
	        <Name>Relation</Name> 
	        <ToTask/> 	
	        <UniqueID/>
	        <UserID>0</UserID>
        </Relation>"""
        relationNode = ET.fromstring(defaultFormat)
        lagType_translation = {"SS": "0", "SF": "1", "FS": "2", "FF": "3"}

        customValues_dict = {}
        customValues_dict["FromTask"] = str(predecessor_id)
        customValues_dict["Lag"] = str(successor[2])
        customValues_dict["LagType"] = lagType_translation[successor[1]]
        customValues_dict["ToTask"] = str(successor[0])
        customValues_dict["UniqueID"] = str(relation_id)

        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(relationNode, key, value)

        relationsNode.append(relationNode) # implicit return

    def generate_ProTrack_outlineList(self, activitiesOnly, activityGroupsOnly):
        """
        This function generates the default OutlineList node for p2x.
        NOTE: Only give activities and activityGroups and not the projectRoot activityGroup with id = 0
        :returns: xml Elementnode: outlineListNode
        """
        defaultFormatOutlineList = """<OutlineList>
            <List/>
        </OutlineList>"""
        outlineListNode = ET.fromstring(defaultFormatOutlineList)
        listNode = outlineListNode.find("List")

        # call recursive function to add all subactivities and subactivityGroups to this list, which is the root list
        self.generate_ProTrack_outlineList_sublist((1,), listNode, activitiesOnly, activityGroupsOnly)

        return outlineListNode

    def generate_ProTrack_outlineList_sublist(self, parent_wbs, parent_listNode, activitiesOnly, activityGroupsOnly):
        """
        This function constructs 1 level of the wbs outline list
        :param parent_wbs: tuple, wbs of parent node:
        :param parent_listNode: xml Element of the parent node
        :param activitiesOnly: list of all activities which are below the parent_wbs
        :param activityGroupsOnly: list of all activityGroups which are below the parent_wbs
        """

        current_wbs_level = len(parent_wbs) + 1
        # extract all activities and activityGroups from this level in the three:
        activities_current_wbs_level = [x for x in activitiesOnly if len(x.wbs_id) == current_wbs_level]
        activityGroups_current_wbs_level = [x for x in activityGroupsOnly if len(x.wbs_id) == current_wbs_level]

        # make 1 list of nodes to add on this wbs level: NOTE: activities_current_wbs_level now also contains the activityGroups!
        activities_current_wbs_level.extend(activityGroups_current_wbs_level)
        currentLevel_activitiesAndGroups_SortedOn_wbs = sorted(activities_current_wbs_level, key= lambda x: x.wbs_id)

        # loop over elements to add:
        for currentChild in currentLevel_activitiesAndGroups_SortedOn_wbs:
            if currentChild in activityGroups_current_wbs_level:
                # current child is an activityGroup
                # create an activityGroup node in the parent_listNode:
                currentNode = self.generate_ProTrack_outlineList_child(1, currentChild.activity_id)
                # extract the list for the subnodes:
                currentNodeList = currentNode.find("List")
                
                # extract all activities and activityGroups with a wbs_id larger than this groupActivity => wbs_id is deeper in tree as groupActivity:

                subActivities = [x for x in activitiesOnly if x.wbs_id > currentChild.wbs_id and x.wbs_id[:current_wbs_level] == currentChild.wbs_id]
                subActivityGroups = [x for x in activityGroupsOnly if x.wbs_id > currentChild.wbs_id and x.wbs_id[:current_wbs_level] == currentChild.wbs_id]

                self.generate_ProTrack_outlineList_sublist(currentChild.wbs_id, currentNodeList, subActivities, subActivityGroups)

            else:
                # current child is an activity:
                currentNode = self.generate_ProTrack_outlineList_child(2, currentChild.activity_id)

            # append currentNode to parent_listNode:
            parent_listNode.append(currentNode)

    def generate_ProTrack_outlineList_child(self, childType, child_id):
        """
        This function generates a child node for the wbs outline list.
        :returns: xml ElementNode, the child node created
        """

        if childType == 1:
            defaultFormat = """<Child>
	            <Type/>
	            <Data/>
                <Expanded>1</Expanded>
                <List/>
            </Child>"""
        else:
            defaultFormat = """<Child>
	            <Type/>
	            <Data/>
            </Child>"""

        childNode = ET.fromstring(defaultFormat)

        customValues_dict = {}
        customValues_dict["Type"] = str(childType)
        customValues_dict["Data"] = str(child_id)

        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(childNode, key, value)

        return childNode

    def generate_ProTrack_resources(self, resources):
        """
        This function generates the default Resources node for p2x.
        :returns: tuple of 2 xml Elementnodes: (resourcesPreformatNode, resourcesNode)
        """
        preFormatResources = """<Resources>
            <Name>Resources</Name>
            <UniqueID>-1</UniqueID>
            <UserID>0</UserID>
        </Resources>"""
        defaultFormatResources = "<Resources/>"
        resourcesNode = ET.fromstring(defaultFormatResources)

        for resource in resources:
            self.generate_ProTrack_resourceNode(resourcesNode, resource)

        return ET.fromstring(preFormatResources), resourcesNode

    def generate_ProTrack_resourceNode(self, resourcesNode, resource):
        """
        This function generates the default resource node for p2x and appends it to the given resourcesNode.
        """
        defaultFormat = """<Resource>
	        <FIELD0/>            
	        <FIELD1/>     
	        <FIELD768></FIELD768> 		  
	        <FIELD778/>
	        <FIELD769/>        
	        <FIELD770/>        
	        <FIELD771/>     
	        <FIELD776/>     
	        <FIELD780/>   
	        <FIELD779>-1</FIELD779>       
	        <FIELD781>-99999999</FIELD781>
	        <FIELD782>-99999999</FIELD782>
	        <FIELD783>-99999999</FIELD783>
        </Resource>"""
        datapercserieFormat = """<DATEPERCSERIE>
	        <DEFAULTAVAILABILITY/>
	        <DATEPERCBREAKPOINTS/> 						
        </DATEPERCSERIE>"""
        resourceNode = ET.fromstring(defaultFormat)
        availabilityNode = ET.fromstring(datapercserieFormat)

        customValues_dict = {}
        customValues_dict["FIELD0"] = str(resource.resource_id)
        customValues_dict["FIELD1"] = self.xml_escape(resource.name)
        customValues_dict["FIELD778"] = "#{0}".format(resource.availability) if resource.resource_type == ResourceType.RENEWABLE else "#Inf"
        customValues_dict["FIELD769"] = "1" if resource.resource_type == ResourceType.RENEWABLE else "0"
        customValues_dict["FIELD770"] = str(resource.cost_use)
        customValues_dict["FIELD771"] = str(resource.cost_unit)
        customValues_dict["FIELD776"] = str(resource.total_resource_cost)
        customValues_dict["FIELD780"] = str(resource.availability) if resource.resource_type == ResourceType.RENEWABLE else "-1"

        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(resourceNode, key, value)

        # set availability node:
        XMLParser.find_xmlNode_and_set_text(availabilityNode, "DEFAULTAVAILABILITY", str(resource.availability) if resource.resource_type == ResourceType.RENEWABLE else "-1")

        resourcesNode.append(resourceNode) # implicit return
        resourcesNode.append(availabilityNode) # implicit return

    def generate_ProTrack_resourceAssignments(self, activitiesOnly):
        """
        This function generates the default ResourceAssignments node for p2x.
        NOTE: Only give low-level activities to this function, no activityGroups!
        :returns: tuple of 2 xml Elementnodes: (resourceAssignmentsPreformatNode, resourceAssignmentsNode)
        """
        preFormatResourceAssignments = """<ResourceAssignments>
            <Name>ResourceAssignments</Name>
            <UniqueID>-1</UniqueID>
            <UserID>0</UserID>
        </ResourceAssignments>"""
        defaultFormatResourceAssignments = "<ResourceAssignments/>"
        resourceAssignmentsNode = ET.fromstring(defaultFormatResourceAssignments)

        for activity in activitiesOnly:
            for resourceAssignment in activity.resources:
                self.generate_ProTrack_resourceAssignmentNode(resourceAssignmentsNode, resourceAssignment, activity.activity_id)

        return ET.fromstring(preFormatResourceAssignments), resourceAssignmentsNode

    def generate_ProTrack_resourceAssignmentNode(self, resourceAssignmentsNode, resourceAssignment, activity_id):
        """
        This function generates the default resourceAssignment node for p2x and appends it to the given resourceAssignmentsNode.
        """
        defaultFormat = """<ResourceAssignment>
	        <FIELD1026/>
	        <FIELD1027/>
	        <FIELD1024/>
	        <FIELD1025/>
        </ResourceAssignment>"""
        resourceAssignmentNode = ET.fromstring(defaultFormat)

        customValues_dict = {}
        customValues_dict["FIELD1026"] = str(resourceAssignment[1])
        customValues_dict["FIELD1027"] = "1" if resourceAssignment[2] else "0"  # fixed assignment or not
        customValues_dict["FIELD1024"] = str(activity_id)
        customValues_dict["FIELD1025"] = str(resourceAssignment[0].resource_id)
        
        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(resourceAssignmentNode, key, value)

        resourceAssignmentsNode.append(resourceAssignmentNode) # implicit return

    def generate_ProTrack_trackingList(self, tracking_periods, datetimeFormat, agenda, activitiesAndGroups):
        """
        This function generates the default TrackingList node for p2x.
        :returns: tuple of 2 xml Elementnodes: (trackingListPreformatNode, trackingListNode)
        """
        preFormatTrackingList = """<TrackingList>
            <Name>TrackingList</Name>
            <UniqueID>-1</UniqueID>
            <UserID>0</UserID>
        </TrackingList>"""
        defaultFormatTrackingList = "<TrackingList/>"
        trackingListNode = ET.fromstring(defaultFormatTrackingList)

        trackingPeriod_id = 1
        for trackingPeriod in tracking_periods:
            self.generate_ProTrack_trackingPeriodNode(trackingListNode, trackingPeriod, trackingPeriod_id, datetimeFormat, agenda, activitiesAndGroups)
            trackingPeriod_id += 1

        return ET.fromstring(preFormatTrackingList), trackingListNode

    def generate_ProTrack_trackingPeriodNode(self, trackingListNode, trackingPeriod, trackingPeriod_id, datetimeFormat, agenda, activitiesAndGroups):
        """
        This function generates the default trackingPeriod node for p2x and appends it to the given trackingListNode.
        """
        defaultFormat = """<TrackingPeriod> 						
	        <Abreviation/>
	        <Description/>
	        <EndDate/>	
	        <Name/>      
	        <PredictiveLogic>0</PredictiveLogic>
	        <UniqueID/>
	        <UserID>0</UserID>
        </TrackingPeriod>"""
        trackingPeriodNode = ET.fromstring(defaultFormat)

        customValues_dict = {}
        customValues_dict["EndDate"] = self.get_date_string(trackingPeriod.tracking_period_statusdate, datetimeFormat)
        customValues_dict["Name"] = self.xml_escape(trackingPeriod.tracking_period_name)
        customValues_dict["UniqueID"] = str(trackingPeriod_id)
        
        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(trackingPeriodNode, key, value)

        # filter: keep only low-level activities:
        trackingPeriod_records = [x for x in trackingPeriod.tracking_period_records if not Activity.is_not_lowest_level_activity(x.activity, activitiesAndGroups)]

        # sort the tracking_period_records on activity_id:
        trackingPeriod_records_sorted = sorted(trackingPeriod_records, key= lambda x: x.activity.activity_id)

        # generate the tracking period activity records
        trackingPeriod_activityRecords = self.generate_ProTrack_trackingPeriod_activityRecords(trackingPeriod_records_sorted, datetimeFormat, agenda)

        trackingListNode.append(trackingPeriodNode) # implicit return
        trackingListNode.append(trackingPeriod_activityRecords) # implicit return

    def generate_ProTrack_trackingPeriod_activityRecords(self, trackingPeriod_records_sorted, datetimeFormat, agenda):
        """
        This function generates a trackingPeriodRecord node for p2x and returns it.
        :returns: xml Element node, "TProTrackActivities-1" node
        """
        defaultFormat = """<TProTrackActivities-1/>"""
        trackingPeriodRecordNode = ET.fromstring(defaultFormat)

        for activityTrackingRecord in trackingPeriod_records_sorted:
            # generate a tracking period activity record node = an activity node + TProTrackActivityTracking-1 node
            self.generate_ProTrack_trackingPeriod_activityRecord_node(trackingPeriodRecordNode, activityTrackingRecord, datetimeFormat, agenda)

        return trackingPeriodRecordNode

    def generate_ProTrack_trackingPeriod_activityRecord_node(self, trackingPeriodRecordNode, activityTrackingRecord, datetimeFormat, agenda):
        """
        This function generates the default trackingPeriodActivityRecord node for p2x and appends it to the given trackingListNode.
        """
        preheaderFormat = "<Activity/>"
        defaultFormat = """<TProTrackActivityTracking-1>
			<ActualStart>0</ActualStart>
			<ActualDuration>0</ActualDuration>
			<ActualCostDev>0</ActualCostDev>
			<RemainingDuration/>
			<RemainingCostDev>0</RemainingCostDev>
			<PercentageComplete>0</PercentageComplete>
		</TProTrackActivityTracking-1>"""
        preheaderNode = ET.fromstring(preheaderFormat)
        trackingPeriodActivityRecordNode = ET.fromstring(defaultFormat)

        preheaderNode.text = str(activityTrackingRecord.activity.activity_id)

        customValues_dict = {}
        if activityTrackingRecord.actual_start < datetime.max:
            customValues_dict["ActualStart"] = self.get_date_string(activityTrackingRecord.actual_start, datetimeFormat)
        if activityTrackingRecord.actual_duration:
            # actual duration > 0
            customValues_dict["ActualDuration"] = str(activityTrackingRecord.actual_duration.days * agenda.get_working_hours_in_a_day() + int(round(activityTrackingRecord.actual_duration.seconds / 3600)))
        if activityTrackingRecord.deviation_pac:
            # PACdev != 0
            customValues_dict["ActualCostDev"] = str(activityTrackingRecord.deviation_pac)
        customValues_dict["RemainingDuration"] = str(activityTrackingRecord.remaining_duration.days * agenda.get_working_hours_in_a_day() + int(round(activityTrackingRecord.remaining_duration.seconds / 3600))) 
        if activityTrackingRecord.deviation_prc:
            # PRCdev != 0
            customValues_dict["RemainingCostDev"] = str(activityTrackingRecord.deviation_prc)
        if activityTrackingRecord.percentage_completed:
            customValues_dict["PercentageComplete"] = str(activityTrackingRecord.percentage_completed / 100)
        
        # set the custom values:
        for key, value in customValues_dict.items():
            XMLParser.find_xmlNode_and_set_text(trackingPeriodActivityRecordNode, key, value)

        trackingPeriodRecordNode.append(preheaderNode) # implicit return
        trackingPeriodRecordNode.append(trackingPeriodActivityRecordNode) # implicit return

    def generate_ProTrack_simulation(self):
        """
        This function generates the default Defaults node for p2x.
        :returns: xml Elementnode
        """
        defaultFormat = """<Simulation>
            <NumberOfRuns>100</NumberOfRuns>
            <SimulationType>1</SimulationType>
            <ScenarioSimulationSettings>
                <AccrueType>0</AccrueType>
                <AllocationMethod>1</AllocationMethod>
                <EndAtStage>100</EndAtStage>
                <EVMAnalyse>1</EVMAnalyse>
                <MaxAccrue>50</MaxAccrue>
                <Measures>[sensCI,sensSI,sensSSI,sensCRIr,sensCRIrho,sensCRItau,sensCRIr_Cost,sensCRIrho_Cost,sensCRItau_Cost,sensCRIr_Res,sensCRIrho_Res,sensCRItau_Res]</Measures>
                <Name>ScenarioSimulationSettings</Name>
                <PercAccrue>100</PercAccrue>
                <Random_Maximum_Deviation>50</Random_Maximum_Deviation>
                <Random_Percentage_Early>50</Random_Percentage_Early>
                <Scenario>1</Scenario>
                <SensitivityAnalyse>1</SensitivityAnalyse>
                <StartAtStage>0</StartAtStage>
                <UniqueID>-1</UniqueID>
                <UseResourceScheduling>0</UseResourceScheduling>
                <UserID>0</UserID>
            </ScenarioSimulationSettings>
            <Res/>
            <Results/>
        </Simulation>"""
        return ET.fromstring(defaultFormat)

    def generate_ProTrack_PSGGame(self, projectEndDate, datetimeFormat, activities_total_number):
        """
        This function generates the PSGGame node for p2x and returns it.
        :returns: xml Element node, "PSGGame" node
        """
        psgActivityFormat = """<PSGActivity>
            <TimeCostTradeOffs>
                <TradeOff>
                    <Duration>10</Duration>
                    <Cost>0</Cost>
                </TradeOff>
            </TimeCostTradeOffs>
            <FIELD82>0</FIELD82>
            <FIELD83>0</FIELD83>
            <FIELD84/>
            <FIELD85>-1</FIELD85>
            <FIELD86>0</FIELD86>
            <FIELD87>0</FIELD87>
            <FIELD88>0</FIELD88>
            <FIELD89>0</FIELD89>
        </PSGActivity>"""
        defaultFormat = """<PSGGame>
            <TargetEnd/>
            <Fine>150</Fine>
            <UncertaintyLevel>0</UncertaintyLevel>
            <NumberOfPeriods>7</NumberOfPeriods>
            <AutoDescisions>0</AutoDescisions>
            <ModifyCostsWhenDelay>0</ModifyCostsWhenDelay>
            <PSGPeriods>
                <Name>PSGPeriods</Name>
                <UniqueID>-1</UniqueID>
                <UserID>0</UserID>
            </PSGPeriods>
            <PSGPeriods/>
            <Activities/>
        </PSGGame>"""
        psgGameNode = ET.fromstring(defaultFormat)
        activitiesNode = psgGameNode.find("Activities")

        projectEndDate_str =  self.get_date_string(projectEndDate, datetimeFormat) if projectEndDate > datetime.min else self.get_date_string(datetime.now(), datetimeFormat)

        XMLParser.find_xmlNode_and_set_text(psgGameNode, "TargetEnd", projectEndDate_str)

        psgActivityNode = ET.fromstring(psgActivityFormat)
        for i in range(0, activities_total_number):
            activitiesNode.append(psgActivityNode)            

        return psgGameNode

    @staticmethod
    def find_xmlNode_and_set_text(parentNode, childTag, value_str):
        "This function searches the parentNode for the childTag and sets its text to value_str"

        childNode = parentNode.find(childTag)
        if childNode is not None:
            childNode.text = value_str
        # return is implicit because parentNode is passed by reference

    @staticmethod
    def find_xmlNode_and_append_childNode(rootNode, parentTag, childTag, child_value_str):
        "This function searches in the rootNode its direct children for the parentTag and appends a new childNode to it."
        parentNode = rootNode.find(parentTag)
        if parentNode is not None:
            childNode = ET.Element(childTag)
            childNode.text = child_value_str
            parentNode.append(childNode)
        # return is implicit because rootNode is passed by reference

    def from_schedule_object(self, project_object, file_path_output="output.xml"):

        projectRootNode = ET.Element("Project")

        nameNode = ET.Element("NAME")
        nameNode.text = XMLParser.xml_escape(project_object.name)

        projectInfoNode = self.generate_ProTrack_projectInfo()

        ## TODO; Read Datetimeformat
        datetimeFormat="d/MM/yyyy h:mm"

        # determine project start datetime and end datetime:
        projectStartDatetime = datetime.max
        projectBaselineEndDatetime = datetime.min

        for activity in project_object.activities:
            if activity.baseline_schedule.start < projectStartDatetime:
                projectStartDatetime = activity.baseline_schedule.start
            if activity.baseline_schedule.end > projectBaselineEndDatetime:
                projectBaselineEndDatetime = activity.baseline_schedule.end
        
        settingsNode = self.generate_ProTrack_settings(projectStartDatetime, projectBaselineEndDatetime, datetimeFormat)
        defaultsNode = self.generate_ProTrack_defaults()
        agendaNode = self.generate_ProTrack_agenda(projectStartDatetime, datetimeFormat, project_object.agenda)

        # get only low-level activities:
        activitiesOnly = [x for x in project_object.activities if not Activity.is_not_lowest_level_activity(x, project_object.activities)]
        activitiesOnly_sortedOn_id = sorted(activitiesOnly, key= lambda x: x.activity_id)

        # extract activityGroups and remove the ProjectRoot activityGroup with id = 0:
        activityGroupsOnly = [x for x in project_object.activities if Activity.is_not_lowest_level_activity(x, project_object.activities) and x.activity_id != 0]
        activityGroupsOnly_sortedOn_id = sorted(activityGroupsOnly, key= lambda x: x.activity_id)

        activitiesPreformatNode, activitiesNode, riskDistributionsPreformatNode, riskDistributionsNode = self.generate_ProTrack_activities_and_riskDistributions(activitiesOnly_sortedOn_id, project_object.agenda, datetimeFormat)

        activityGroupsPreformatNode, activityGroupsNode = self.generate_ProTrack_activityGroups(activityGroupsOnly_sortedOn_id)

        relationsPreformatNode, relationsNode = self.generate_ProTrack_relations(activitiesOnly_sortedOn_id)

        # generate outline list:
        outlineListNode = self.generate_ProTrack_outlineList(activitiesOnly_sortedOn_id, activityGroupsOnly_sortedOn_id)

        # generate resources node:
        resourcesPreformatNode, resourcesNode = self.generate_ProTrack_resources(project_object.resources)

        # generate resource assignments node:
        resourceAssignmentsPreformatNode, resourceAssignmentsNode = self.generate_ProTrack_resourceAssignments(activitiesOnly_sortedOn_id)

        # generate tracking list node:
        # sort tracking periods first on statusdate:
        tracking_periods_sorted = sorted(project_object.tracking_periods, key= lambda x: x.tracking_period_statusdate)

        trackingListPreformatNode, trackingListNode = self.generate_ProTrack_trackingList(tracking_periods_sorted, datetimeFormat, project_object.agenda, project_object.activities)

        # generate default: simulation node:
        defaultSimulationNode = self.generate_ProTrack_simulation()

        # generate default PSGgame nodes:

        defaultPSGGameNode = self.generate_ProTrack_PSGGame(projectBaselineEndDatetime, datetimeFormat, len(activitiesOnly)) # necessary for valid ProTrackFile

        ### append all generated xml nodes in right order to the rootNode:
        projectRootNode.append(nameNode)
        projectRootNode.append(projectInfoNode)
        projectRootNode.append(settingsNode)
        projectRootNode.append(defaultsNode)
        projectRootNode.append(agendaNode)
        projectRootNode.append(activitiesPreformatNode)
        projectRootNode.append(activitiesNode)
        projectRootNode.append(relationsPreformatNode)
        projectRootNode.append(relationsNode)
        projectRootNode.append(activityGroupsPreformatNode)
        projectRootNode.append(activityGroupsNode)
        projectRootNode.append(outlineListNode)
        projectRootNode.append(resourcesPreformatNode)
        projectRootNode.append(resourcesNode)
        projectRootNode.append(resourceAssignmentsPreformatNode)
        projectRootNode.append(resourceAssignmentsNode)
        projectRootNode.append(trackingListPreformatNode)
        projectRootNode.append(trackingListNode)
        projectRootNode.append(defaultSimulationNode)
        projectRootNode.append(riskDistributionsPreformatNode)
        projectRootNode.append(riskDistributionsNode)
        projectRootNode.append(defaultPSGGameNode)

        # convert rootNode to tree:
        xmlTree = ET.ElementTree(projectRootNode)
        xmlTree.write(file_path_output, encoding="UTF-8" ,xml_declaration=False)

        return True # writing to p2x was successful

