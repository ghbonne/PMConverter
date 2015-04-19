import ast #ast.literal_eval
import xml.etree.ElementTree as ET
from object.activity import Activity
from object.baselineschedule import BaselineScheduleRecord
from object.resource import Resource
from object.riskanalysisdistribution import RiskAnalysisDistribution
from object.activitytracking import ActivityTrackingRecord
from object.trackingperiod import TrackingPeriod

tree = ET.parse('project.xml')
root = tree.getroot()

activity_list=[]
activity_list_wo_groups=[]
res_list=[]
counter = 0
max=0

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

## Activity group, ID, Name
for activities in root.findall('ActivityGroups'):
    for activity in activities.findall('ActivityGroup'):
        #UniqueID
        UniqueID=int(activity.find('UniqueID').text)
        a_g = Activity(int(UniqueID));
        #Name
        Name=activity.find('Name').text
        a_g.name= Name
        activity_list_interim[UniqueID-1]=a_g

## Activity name, ID, BSR
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
        BaselineDuration = activity.find('BaseLineDuration').text
        #BaseLineStart
        BaseLineStart=activity.find('BaseLineStart').text
        #FixedBaselineCost
        FixedBaselineCost=activity.find('FixedBaselineCost').text
        #BaselineCostByUnit
        BaselineCostByUnit=activity.find('BaselineCostByUnit').text

        BSR = BaselineScheduleRecord()
        BSR.start=BaseLineStart
        BSR.duration=BaselineDuration
        BSR.fixed_cost=FixedBaselineCost
        BSR.var_cost=BaselineCostByUnit
        a.baseline_schedule=BSR
        activity_list_interim[UniqueID-1]=a
        activity_list_wo_groups_interim[UniqueID-1]=a
#print(activity_list)

# Removing empty activities from list (list is sorted)
for i in range(0, len(activity_list_interim)):
    if activity_list_interim[i] != None:
        activity_list.append(activity_list_interim[i])
    if activity_list_wo_groups_interim[i] != None:
        activity_list_wo_groups.append(activity_list_wo_groups_interim[i])



## Successor/Predeccessors
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
            #TO DO: Other possibilites
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


## WBS
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




## Resources (Definition)
for resources in root.findall('Resources'):
    for resource in resources.findall('Resource'):
        res_ID=int(resource.find('FIELD0').text)
        name=resource.find('FIELD1').text
        availability=resource.find('FIELD778').text
        renewable=bool(resource.find('FIELD769').text)
        if renewable == 1:
            renewable_string='Renewable'
        cost_per_use=resource.find('FIELD770').text
        cost_per_unit=resource.find('FIELD771').text
        availability_int=int(resource.find('FIELD780').text)
        total_cost=float(resource.find('FIELD776').text)
        res=Resource(res_ID, name, renewable_string, availability, cost_per_use, cost_per_unit, total_cost)
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
                           if len(activity.resources) >0:
                                activity.resources.append(resourceTuple)
                           else:
                               activity.resources=[resourceTuple]

## Risk Analysis
distribution_list =  [0 for x in range(len(activity_list))]  # List with possible distributions
#Standard distributions
distribution_list[1]= RiskAnalysisDistribution(distribution_type="standard", distribution_units="no risk", optimistic_duration=99,
                 probable_duration=100, pessimistic_duration=101)
distribution_list[2]=RiskAnalysisDistribution(distribution_type="standard", distribution_units="symmetric", optimistic_duration=80,
                 probable_duration=100, pessimistic_duration=120)
distribution_list[3]=RiskAnalysisDistribution(distribution_type="standard", distribution_units="skewed left", optimistic_duration=80,
                 probable_duration=110, pessimistic_duration=120)
distribution_list[4]=RiskAnalysisDistribution(distribution_type="standard", distribution_units="skewed right", optimistic_duration=80,
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
        distribution_list[x]=(RiskAnalysisDistribution(distribution_type="manual", distribution_units="absolute",
                                                          optimistic_duration=distr[0],probable_duration=distr[1], pessimistic_duration=distr[2]))

for activities in root.findall('Activities'):
    for activity in activities.findall('Activity'):
        for activity_l in activity_list:
            if activity != None and activity_l!=None:
                if int(activity.find('UniqueID').text) == activity_l.activity_id:
                    distr_number=int(activity.find('Distribution').text)
                    activity_l.risk_analysis=distribution_list[distr_number]

#print(len(activity_list))



i=0
TP_nr=0
## Activity Tracking
for tracking_list in root.findall('TrackingList'):
    activityTrackingRecord_list=[]
    for tracking_period in tracking_list.findall('TProTrackActivities-1'):
        TP_nr+=1
        activity_nr=0
        for tracking_activity in tracking_period.findall('TProTrackActivityTracking-1'):
            # Read Data
            actualStart=tracking_activity.find('ActualStart')
            actualDuration=tracking_activity.find('ActualDuration')
            actualCostDev=tracking_activity.find('ActualCostDev')
            remainingDuration=tracking_activity.find('RemainingDuration')
            remainingCostDev=tracking_activity.find('RemainingCostDev')
            percentageComplete=tracking_activity.find('PercentageComplete')

            #Assign Data
            activityTrackingRecord=ActivityTrackingRecord()
            # From xml
            activityTrackingRecord.activity=activity_list_wo_groups[activity_nr]
            activityTrackingRecord.actual_start=actualStart
            activityTrackingRecord.actual_duration=actualDuration
            activityTrackingRecord.deviation_pac=actualCostDev
            activityTrackingRecord.deviation_prc=remainingCostDev
            activityTrackingRecord.percentage_completed=percentageComplete
            if percentageComplete >= 1:
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
            activityTrackingRecord.remaining_duration=0
            activityTrackingRecord.tracking_period=None

            if len(activityTrackingRecord_list)>0:
                activityTrackingRecord_list.append(activityTrackingRecord)
            else:
                activityTrackingRecord_list=[activityTrackingRecord]
            activity_nr+=1

    #TrackingPeriod Activity info
        count=0
        for tracking_period_info in tracking_list.findall('TrackingPeriod'):
            count+=1
        TP_list=[None for x in range(count)]
        count=0
        for tracking_period_info in tracking_list.findall('TrackingPeriod'):
            name=tracking_period_info.find('Name').text
            enddate=tracking_period_info.find('EndDate').text
            TP_list[count]=TrackingPeriod(name,enddate,)
            count+=1




## Testing
#for activity in activity_list:
 #   for x in range(0,len(activity.resources)):
  #      print(activity.activity_id, activity.resources[x][0].resource_id)

#for distr in distribution_list:
 #   print(distr)