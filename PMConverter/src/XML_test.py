import ast #ast.literal_eval
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
from object.activity import Activity
from object.baselineschedule import BaselineScheduleRecord
from object.resource import Resource
from object.riskanalysisdistribution import RiskAnalysisDistribution
from object.activitytracking import ActivityTrackingRecord
from object.trackingperiod import TrackingPeriod
from object.resource import ResourceType
from object.riskanalysisdistribution import DistributionType
from object.riskanalysisdistribution import ManualDistributionUnit
from object.riskanalysisdistribution import StandardDistributionUnit
from object.agenda import Agenda
from object.projectobject import ProjectObject
from convert.XLSXparser import XLSXParser
#todo: extra activity tracking record fields (should already be read), other possibilites Lagtype (FS,...), Agenda (usage)



tree = ET.parse('project.xml')
root = tree.getroot()

activity_list=[]
activity_list_wo_groups=[]
res_list=[]
counter = 0
max=0

## Project name
project_name=root.find("NAME").text

## Dateformat
for settings in root.findall('Settings'):
    dateformat=settings.find('DateTimeFormat').text

def getdate(datestring=""):
    if dateformat == "d/MM/yyyy h:mm":
        if len(datestring) == 12:
            day=int(datestring[:2])
            month=int(datestring[2:4])
            year=int(datestring[4:8])
            hour=int(datestring[8:10])
            minute=int(datestring[10:12])
            date_datetime=datetime(year, month, day, hour, minute)
            return date_datetime
        elif len(datestring) == 8:
            day=int(datestring[:2])
            month=int(datestring[2:4])
            year=int(datestring[4:8])
            date_datetime=datetime(year, month, day)
            return date_datetime
        else:
            return "N/A"
    elif dateformat == "MM/d/yyyy h:mm":
        if len(datestring) == 12:
            month=int(datestring[:2])
            day=int(datestring[2:4])
            year=int(datestring[4:8])
            hour=int(datestring[8:10])
            minute=int(datestring[10:12])
            date_datetime=datetime(year, month, day, hour, minute)
            return date_datetime
        elif len(datestring) == 8:
            month=int(datestring[:2])
            day=int(datestring[2:4])
            year=int(datestring[4:8])
            date_datetime=datetime(year, month, day)
            return date_datetime
        else:
            raise("Warning! Dateformat undefined")
            return "N/A"

## Create Agenda: Working hours, Working days, Holidays
for agenda in root.findall('Agenda'):
    project_agenda=Agenda()
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
    ## No holidays in example xml file: following code is educated guess
    for holidays in agenda.findall('Holidays'):
        for holiday in holidays.findall('Holiday'):
            holiday=int(holiday.text)
            project_agenda.set_holiday(holiday)


## max number activity?
for activities in root.findall('ActivityGroups'):
    for activity in activities.findall('ActivityGroup'):
        #UniqueID
        UniqueID=int(activity.find('UniqueID').text)
        if(UniqueID > max):
            max = UniqueID
for activities in root.findall('Activities'):
    for activity in activities.findall('Activity'):
        #UniqueID
        UniqueID=int(activity.find('UniqueID').text)
        if(UniqueID > max):
            max = UniqueID

#Initialize activity list
activity_list_interim=[None for x in range(max)]
activity_list_wo_groups_interim=[None for x in range(max)]

activity_group_count=0
## Activity group, ID, Name
for activities in root.findall('ActivityGroups'):
    for activity in activities.findall('ActivityGroup'):
        #UniqueID
        UniqueID=int(activity.find('UniqueID').text)
        a_g = Activity(int(UniqueID));
        #Name
        Name=activity.find('Name').text
        a_g.name= Name
        activity_group_count+=1
        activity_list_interim[UniqueID-1]=a_g



## Activity name, ID, Baseline schefule
for activities in root.findall('Activities'):
    for activity in activities.findall('Activity'):
        #UniqueID
        UniqueID=int(activity.find('UniqueID').text)
        a = Activity((UniqueID));
        #Name
        Name=activity.find('Name').text
        a.name= Name

        ##BaseLineSchedule
        #Baseline duration (hours)
        BaselineDuration_hours = int(activity.find('BaseLineDuration').text)
        BaselineDuration=timedelta(hours=BaselineDuration_hours)
        #BaseLineStart
        BaseLineStart=getdate(activity.find('BaseLineStart').text)
        #FixedBaselineCost
        FixedBaselineCost=float(activity.find('FixedBaselineCost').text)
        #BaselineCostByUnit
        BaselineCostByUnit=float(activity.find('BaselineCostByUnit').text)

        BSR = BaselineScheduleRecord()
        BSR.start=BaseLineStart
        BSR.duration=BaselineDuration
        BSR.fixed_cost=FixedBaselineCost
        BSR.hourly_cost=BaselineCostByUnit
        BSR.var_cost = BSR.hourly_cost*BaselineDuration_hours*8
        BSR.total_cost = BSR.var_cost+BSR.fixed_cost   ## Resource cost is missing

        a.baseline_schedule=BSR
        activity_list_interim[UniqueID-1]=a
        activity_list_wo_groups_interim[UniqueID-1]=a


# Removing empty activities from list (list is sorted)
for i in range(0, len(activity_list_interim)):
    if activity_list_interim[i] != None:
        activity_list.append(activity_list_interim[i])
    if activity_list_wo_groups_interim[i] != None:
        activity_list_wo_groups.append(activity_list_wo_groups_interim[i])




## Successor/Predeccessors ##
for relations in root.findall('Relations'):
    for relation in relations.findall('Relation'):
        # Successors
        Predecessor = relation.find('FromTask').text
        Successor = relation.find('ToTask').text
        #Lag
        Lag = relation.find('Lag').text
        LagKind = relation.find('LagKind').text
        LagType = relation.find('LagType').text
        LagString=''
        if LagKind == '0' and LagType == '2':
            LagString='FS'
        else:
            #TODO: Other possibilites
            0;

        SuccessorTuple = (Successor, LagString, Lag)
        PredecessorTuple = (Predecessor, LagString, Lag)

        for activity in activity_list:
            if activity != None:
                if activity.activity_id == int(Predecessor):
                        if len(activity.successors) > 0:
                            (activity.successors).append(SuccessorTuple)
                        else:
                            (activity.successors) = [SuccessorTuple]
                elif activity.activity_id == int(Successor):
                    if len(activity.predecessors) > 0:
                        (activity.predecessors).append(PredecessorTuple)
                    else:
                        (activity.predecessors) = [PredecessorTuple]





## WBS ##
for outline in root.findall('OutlineList'):
    counter1=0
    for list in outline.findall('List'):
        for child in list.findall('Child'):
            counter1+=1
            ID=int(child.find('Data').text)
            #1.x
            for activity in activity_list:
                if activity != None:
                    if activity.activity_id == ID:
                        wbsTuple=(1, counter1)
                        activity.wbs_id=wbsTuple
            #1.x.y
            for list2 in child.findall('List'):
                counter2=0
                for child2 in list2.findall('Child'):
                    counter2+=1
                    ID2=int(child2.find('Data').text)
                    for activity2 in activity_list:
                        if activity2 != None:
                            if activity2.activity_id == ID2:
                                wbsTuple=(1, counter1, counter2)
                                activity2.wbs_id=wbsTuple



###### Resources ######
## Resources (Definition)
for resources in root.findall('Resources'):
    for resource in resources.findall('Resource'):
        res_ID=int(resource.find('FIELD0').text)
        name=resource.find('FIELD1').text
        availability=resource.find('FIELD778').text
        renewable=bool(resource.find('FIELD769').text)
        if renewable == 1:
            res_type= ResourceType.RENEWABLE
        else:
            res_type= ResourceType.CONSUMABLE
        cost_per_use=float(resource.find('FIELD770').text)
        cost_per_unit=float(resource.find('FIELD771').text)
        availability_int=int(resource.find('FIELD780').text)
        total_cost=float(resource.find('FIELD776').text)
        res=Resource(res_ID, name, res_type, availability_int, cost_per_use, cost_per_unit, total_cost)
        res_list.append(res)

## Resources (Assignment)
for resource_assignments in root.findall('ResourceAssignments'):
    for resource_assignment in resource_assignments.findall('ResourceAssignment'):
        res_id=int(resource_assignment.find('FIELD1025').text)
        activity_ID=int(resource_assignment.find('FIELD1024').text)
        res_needed=int(resource_assignment.find('FIELD1026').text)
        for activity in activity_list:
            if activity != None:
               if activity.activity_id == activity_ID:
                   for resource in res_list:
                       resourceTuple=(resource, res_needed)
                       if resource.resource_id == res_id:
                           #print(activity.activity_id,activity.baseline_schedule.duration)
                           activity_duration=activity.baseline_schedule.duration.days * 8 + activity.baseline_schedule.duration.seconds/3600
                           activity.resource_cost=resource.cost_unit*res_needed*activity_duration
                           activity.baseline_schedule.total_cost+=activity.resource_cost
                           if len(activity.resources) >0:
                                activity.resources.append(resourceTuple)
                           else:
                               activity.resources=[resourceTuple]







### Risk Analysis ###
distribution_list =  [0 for x in range(len(activity_list))]  # List with possible distributions
#Standard distributions
distribution_list[1]= RiskAnalysisDistribution(distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.NO_RISK, optimistic_duration=99,
                 probable_duration=100, pessimistic_duration=101)
distribution_list[2]=RiskAnalysisDistribution(distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SYMMETRIC, optimistic_duration=80,
                 probable_duration=100, pessimistic_duration=120)
distribution_list[3]=RiskAnalysisDistribution(distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SKEWED_LEFT, optimistic_duration=80,
                 probable_duration=110, pessimistic_duration=120)
distribution_list[4]=RiskAnalysisDistribution(distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SKEWED_RIGHT, optimistic_duration=80,
                 probable_duration=90, pessimistic_duration=120)
i=0
distr=[0, 0, 0]
for distributions in root.findall('SensitivityDistributions'):
    for x in range(5, len(activity_list)):
        distr_string='TProTrackSensitivityDistribution'+str(x)
        distr_x = distributions.find(distr_string)
        if distr_x != None:
            distribution=distr_x.find('Distribution')
            i=0
            for X in distribution.findall('X'):
                distr[i]=int(X.text)
                i+=1
        distribution_list[x]=(RiskAnalysisDistribution(distribution_type=DistributionType.MANUAL, distribution_units=ManualDistributionUnit.ABSOLUTE,
                                                          optimistic_duration=distr[0],probable_duration=distr[1], pessimistic_duration=distr[2]))

for activities in root.findall('Activities'):
    for activity in activities.findall('Activity'):
        for activity_l in activity_list:
            if activity != None and activity_l!=None:
                if int(activity.find('UniqueID').text) == activity_l.activity_id:
                    distr_number=int(activity.find('Distribution').text)
                    activity_l.risk_analysis=distribution_list[distr_number]




## Activity Tracking
for tracking_list in root.findall('TrackingList'):
    activityTrackingRecord_list=[]
    ## How many Trackng periods?
    count=0
    for tracking_period_info in tracking_list.findall('TrackingPeriod'):
        count+=1
    TP_list=[0 for x in range(count)]
    ATR_matrix=[]
    for tracking_period in tracking_list.findall('TProTrackActivities-1'):
        activity_nr=0

        #ATR_matrix=[None for x in range(count)]

        activityTrackingRecord_list=[]
        for tracking_activity in tracking_period.findall('TProTrackActivityTracking-1'):
            # Read Data
            actualStart=getdate(tracking_activity.find('ActualStart').text)
            actualDuration=tracking_activity.find('ActualDuration').text
            actualCostDev=tracking_activity.find('ActualCostDev').text
            remainingDuration=tracking_activity.find('RemainingDuration').text
            remainingCostDev=tracking_activity.find('RemainingCostDev').text
            percentageComplete=float(tracking_activity.find('PercentageComplete').text)*100
            #Assign Data
            activityTrackingRecord=ActivityTrackingRecord()
            # From xml
            activityTrackingRecord.activity=activity_list_wo_groups[activity_nr]
            activityTrackingRecord.actual_start=actualStart
            activityTrackingRecord.actual_duration=actualDuration
            activityTrackingRecord.deviation_pac=actualCostDev
            activityTrackingRecord.deviation_prc=remainingCostDev
            activityTrackingRecord.remaining_duration=remainingDuration
            activityTrackingRecord.percentage_completed=percentageComplete
            if percentageComplete >= 100:
                trackingStatus='Finished'
            elif percentageComplete == 0:
                trackingStatus='Not Started'
            else:
                trackingStatus='Started'
            activityTrackingRecord.tracking_status=trackingStatus

            # Already read from previous data fields
            #todo
            activityTrackingRecord.actual_cost=0
            activityTrackingRecord.earned_value=0
            activityTrackingRecord.planned_actual_cost=0
            activityTrackingRecord.planned_remaining_cost=0
            activityTrackingRecord.planned_value=0

            if len(activityTrackingRecord_list)>0:
                activityTrackingRecord_list.append(activityTrackingRecord)
            else:
                activityTrackingRecord_list=[activityTrackingRecord]
            activity_nr+=1

        if len(ATR_matrix)>0:
            ATR_matrix.append(activityTrackingRecord_list)
        else:
            ATR_matrix=[activityTrackingRecord_list]


#TrackingPeriod Activity info
count=0
for tracking_period_info in tracking_list.findall('TrackingPeriod'):
    name=tracking_period_info.find('Name').text
    enddate=tracking_period_info.find('EndDate').text
    enddate_datetime=getdate(enddate)
    TP_list[count]=TrackingPeriod(name,enddate, ATR_matrix[count])
    for activityTrackingRecord in ATR_matrix[count]:
        activityTrackingRecord.tracking_period=TP_list[count]
    count+=1


### Sort activities based on WBS

project_activity=Activity(activity_id=0, name=project_name, wbs_id=(1,))
activity_list_wbs=[project_activity]
count1 = 1
count2 = 1

for activity_group2 in activity_list:
    for activity_group in activity_list:
        if activity_group.wbs_id == (1, count1):
            count2 = 1
            activity_list_wbs.append(activity_group)
            for activity2 in activity_list:
                for activity in activity_list:
                    if activity.wbs_id == (1, count1, count2):
                        count2 += 1
                        activity_list_wbs.append(activity)
            count1 += 1

## Subdividing list into Activity groups
count1=1
count2=1
subgroup_count=[0 for x in range(activity_group_count)]
# Counting activity groups
for activity in activity_list_wbs:
    if activity.wbs_id == (1, count1):
        count2=1
        for activity in activity_list_wbs:
            if activity.wbs_id == (1, count1,count2):
                subgroup_count[count1-1]+=1
                count2+=1
        count1+=1


matrix_activity_group=[[] for x in range(activity_group_count)]
matrix_activity_group[0]=activity_list_wbs[1:(2+subgroup_count[0])]

# Making list per activity group: matrix_activity_group[i] contains all activities which are part of activity group i
for i in range(1,len(subgroup_count)):
    matrix_activity_group[i]=activity_list_wbs[1+i+sum(subgroup_count[0:i]):(2+i+sum(subgroup_count[0:i+1]))]

for ag in matrix_activity_group:
    for i in range(1,len(ag)):
        ag[0].baseline_schedule.total_cost += ag[i].baseline_schedule.total_cost
        ag[0].baseline_schedule.fixed_cost += ag[i].baseline_schedule.fixed_cost
#    print('total:', ag[0].baseline_schedule.total_cost, 'fixed:', ag[0].baseline_schedule.fixed_cost)




## Make project object
project_object=ProjectObject(project_name, activity_list_wbs, TP_list, res_list, project_agenda)


## Parse to XLS
xlsx_parser = XLSXParser()
xlsx_parser.from_schedule_object(project_object, "test_alexander.xlsx")