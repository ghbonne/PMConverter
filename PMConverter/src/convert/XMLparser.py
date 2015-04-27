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




class XMLParser(FileParser):

    def __init__(self):
        super().__init__()

    def getdate(self,datestring="", dateformat=""):
            if dateformat == "d/MM/yyyy h:mm" or "d-M-yyyy h:m":
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
                    return datetime.max
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
                    return datetime.max
            else:
                print("Error:" + dateformat)
                raise("Warning! Dateformat undefined" )


    def get_date_string(self,date=datetime.min,dateformat=""):
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
            minute_str=str(day)
        if dateformat == "d/MM/yyyy h:mm":
            return day_str+month_str+year_str+hour_str+minute_str
        elif dateformat == "MM/d/yyyy h:mm":
            return month_str+day_str+year_str+hour_str+minute_str
        else:
            raise "Dateformat undefined"

    def to_schedule_object(self, file_path_input):
        tree = ET.parse(file_path_input)
        root = tree.getroot()

        activity_list=[]
        activity_list_wo_groups=[]
        res_list=[]
        max=0

        ## Project name
        project_name=root.find("NAME").text

        ## Dateformat
        for settings in root.findall('Settings'):
            dateformat=settings.find('DateTimeFormat').text



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
                    holiday=holiday.text[:8]
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
                #agenda=Agenda()
                BaselineDuration=project_agenda.get_duration_working_days(duration_hours=BaselineDuration_hours)
                #BaseLineStart
                BaseLineStart=self.getdate(activity.find('BaseLineStart').text, dateformat)
                #FixedBaselineCost
                FixedBaselineCost=float(activity.find('FixedBaselineCost').text)
                #BaselineCostByUnit
                BaselineCostByUnit=float(activity.find('BaselineCostByUnit').text)

                enddate=project_agenda.get_end_date(BaseLineStart,BaselineDuration_hours/8)
                BSR = BaselineScheduleRecord()
                BSR.start=BaseLineStart
                BSR.end=enddate
                BSR.duration=BaselineDuration
                BSR.fixed_cost=FixedBaselineCost
                BSR.hourly_cost=BaselineCostByUnit
                BSR.var_cost = BSR.hourly_cost*BaselineDuration_hours
                BSR.total_cost = BSR.var_cost+BSR.fixed_cost
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
                if LagKind == '0' and LagType == '2':
                    LagString='FS'
                elif LagKind == '0' and LagType == '0':
                    LagString='SS'
                elif LagKind == '0' and LagType == '1':
                    LagString='SF'
                elif LagKind == '0' and LagType == '3':
                    LagString='FF'
                else:
                    raise 'Lag Undefined'

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
            activity_group_count=0
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
                                activity_group_count+=1
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
                availability_float=float(resource.find('FIELD780').text)
                total_cost=float(resource.find('FIELD776').text)
                res=Resource(res_ID, name, res_type, availability_float, cost_per_use, cost_per_unit, total_cost)
                res_list.append(res)




        ## Resources (Assignment)
        for resource_assignments in root.findall('ResourceAssignments'):
            for resource_assignment in resource_assignments.findall('ResourceAssignment'):
                res_id=int(resource_assignment.find('FIELD1025').text)
                activity_ID=int(resource_assignment.find('FIELD1024').text)
                res_needed=ast.literal_eval(resource_assignment.find('FIELD1026').text)
                for activity in activity_list:
                    if activity != None:
                       if activity.activity_id == activity_ID:
                           for resource in res_list:
                               resourceTuple=(resource, res_needed)
                               if resource.resource_id == res_id:
                                   #print(activity.activity_id,activity.baseline_schedule.duration)
                                   activity_duration=activity.baseline_schedule.duration
                                   activity.resource_cost=resource.cost_unit*res_needed*activity_duration.days*8
                                   activity.baseline_schedule.total_cost+=activity.resource_cost
                                   if len(activity.resources) >0:
                                        activity.resources.append(resourceTuple)
                                   else:
                                       activity.resources=[resourceTuple]





        ### Risk Analysis ###
        distribution_list =  [0 for x in range(len(activity_list)+5)]  # List with possible distributions
        #Standard distributions
        distribution_list[1]= RiskAnalysisDistribution(distr_id=1,distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.NO_RISK, optimistic_duration=99,
                         probable_duration=100, pessimistic_duration=101)
        distribution_list[2]=RiskAnalysisDistribution(distr_id=2,distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SYMMETRIC, optimistic_duration=80,
                         probable_duration=100, pessimistic_duration=120)
        distribution_list[3]=RiskAnalysisDistribution(distr_id=3,distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SKEWED_LEFT, optimistic_duration=80,
                         probable_duration=110, pessimistic_duration=120)
        distribution_list[4]=RiskAnalysisDistribution(distr_id=4,distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.SKEWED_RIGHT, optimistic_duration=80,
                         probable_duration=90, pessimistic_duration=120)
        i=0
        distr=[0, 0, 0]
        for distributions in root.findall('SensitivityDistributions'):
            for x in range(5, len(activity_list)+5):
                distr_string='TProTrackSensitivityDistribution'+str(x)
                distr_x = distributions.find(distr_string)
                if distr_x != None:
                    distribution=distr_x.find('Distribution')
                    name=distr_x.find('Name').text
                    i=0
                    for X in distribution.findall('X'):
                        distr[i]=int(X.text)
                        i+=1


                    distribution_list[x]=(RiskAnalysisDistribution(distr_id=x,distr_name=name,distribution_type=DistributionType.MANUAL, distribution_units=ManualDistributionUnit.ABSOLUTE,
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
                    actualStart=self.getdate(tracking_activity.find('ActualStart').text, dateformat)
                    actualDuration_hours=int(tracking_activity.find('ActualDuration').text)
                    actualDuration=project_agenda.get_duration_working_days(duration_hours=actualDuration_hours)
                    actualCostDev=float(tracking_activity.find('ActualCostDev').text)
                    remainingDuration_days=int(tracking_activity.find('RemainingDuration').text)
                    remainingDuration=project_agenda.get_duration_working_days(duration_hours=remainingDuration_days)
                    remainingCostDev=float(tracking_activity.find('RemainingCostDev').text)
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

                    # Should have been already read from previous data fields
                    activityTrackingRecord.actual_cost=\
                        activity_list_wo_groups[activity_nr].baseline_schedule.total_cost*percentageComplete/100
                    activityTrackingRecord.earned_value=\
                        math.ceil(percentageComplete/100)*(activity_list_wo_groups[activity_nr].baseline_schedule.fixed_cost+\
                                                           (activity_list_wo_groups[activity_nr].resource_cost+\
                                                 activity_list_wo_groups[activity_nr].baseline_schedule.var_cost*\
                                                 activity_list_wo_groups[activity_nr].baseline_schedule.duration.days)
                                                 *percentageComplete/100)

                    activityTrackingRecord.planned_actual_cost=activityTrackingRecord.actual_cost
                    activityTrackingRecord.planned_remaining_cost=\
                        activity_list_wo_groups[activity_nr].baseline_schedule.total_cost * (100-percentageComplete)/100
                    activityTrackingRecord.remaining_cost=activityTrackingRecord.planned_remaining_cost
                    activityTrackingRecord.planned_value=\
                        activity_list_wo_groups[activity_nr].baseline_schedule.total_cost

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
            statusdate=tracking_period_info.find('EndDate').text
            statusdate_datetime=self.getdate(statusdate,dateformat)
            TP_list[count]=TrackingPeriod(name,statusdate_datetime, ATR_matrix[count])

            for activityTrackingRecord in ATR_matrix[count]:
                activityTrackingRecord.tracking_period=TP_list[count]
                enddate=project_agenda.get_end_date(begin_date=activityTrackingRecord.actual_start, days=activityTrackingRecord.actual_duration.days)
                #enddate=self.getdate(enddate,dateformat)
                #print(enddate, activityTrackingRecord.activity.activity_id)

                if statusdate_datetime < enddate and enddate.year < 2030 :
                    t2=project_agenda.get_time_between(activityTrackingRecord.actual_start, enddate).days + \
                       project_agenda.get_time_between(activityTrackingRecord.actual_start, enddate).seconds/8/3600
                    t1=project_agenda.get_time_between(activityTrackingRecord.actual_start, statusdate_datetime).days \
                       +project_agenda.get_time_between(activityTrackingRecord.actual_start, statusdate_datetime).seconds/8/3600
                    percentage=t1/t2
                    ActivityTrackingRecord.planned_value= math.ceil(percentage/100)*\
                                                          (activityTrackingRecord.activity.baseline_schedule.fixed_cost+\
                                                           (activityTrackingRecord.activity.resource_cost+\
                                                 activityTrackingRecord.activity.baseline_schedule.var_cost*\
                                                 activityTrackingRecord.activity.baseline_schedule.duration.days)*percentage/100)
            count+=1


        ### Sort activities based on WBS
        project_activity=Activity(activity_id=0, name=project_name, wbs_id=(1,))
        activity_list_wbs=[project_activity]
        count1 = 1
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

        ## Assigning Data to activity groups
        for ag in matrix_activity_group:
            ## Costs
            total_cost=0
            fixed_cost=0
            min_start=datetime.max
            max_end=datetime.min
            bsr=BaselineScheduleRecord()
            if len(ag) > 1:
                for i in range(1,len(ag)):
                    if ag[i].baseline_schedule.start < min_start:
                        min_start=ag[i].baseline_schedule.start
                    if ag[i].baseline_schedule.end > max_end:
                        max_end=ag[i].baseline_schedule.end
                    total_cost += ag[i].baseline_schedule.total_cost
                    fixed_cost += ag[i].baseline_schedule.fixed_cost
                ag[0].baseline_schedule=bsr
                ag[0].baseline_schedule.start=min_start
                ag[0].baseline_schedule.end=max_end
                ag[0].baseline_schedule.total_cost = total_cost
                ag[0].baseline_schedule.fixed_cost = fixed_cost
                ag[0].baseline_schedule.duration= (max_end -min_start)
            #else:
                # no real activity group (single activity)

        ## Assigning data to total project
        total_cost=0
        fixed_cost=0
        min_start=datetime.max
        max_end=datetime.min
        bsr=BaselineScheduleRecord()
        for ag in matrix_activity_group:
            total_cost+=ag[0].baseline_schedule.total_cost
            fixed_cost+=ag[0].baseline_schedule.fixed_cost
            if ag[0].baseline_schedule.start < min_start:
                min_start=ag[0].baseline_schedule.start
            if ag[0].baseline_schedule.end > max_end:
                max_end=ag[0].baseline_schedule.end
        activity_list_wbs[0].baseline_schedule=bsr
        activity_list_wbs[0].baseline_schedule.start=min_start
        activity_list_wbs[0].baseline_schedule.end=max_end
        activity_list_wbs[0].baseline_schedule.total_cost = total_cost
        activity_list_wbs[0].baseline_schedule.fixed_cost = fixed_cost
        activity_list_wbs[0].baseline_schedule.duration= (max_end -min_start)


        ## Make project object
        project_object=ProjectObject(project_name, activity_list_wbs, TP_list, res_list, project_agenda)
        return project_object


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
        file.write("""<ProjectInfo><LastSavedBy>jbatseli</LastSavedBy><Name>ProjectInfo</Name>
        <SavedWithMayorBuild>3</SavedWithMayorBuild><SavedWithMinorBuild>0</SavedWithMinorBuild>
        <SavedWithVersion>0</SavedWithVersion><UniqueID>-1</UniqueID><UserID>0</UserID></ProjectInfo>""")

        ### Settings (Somewhat important) ###
        ## A lot of this info can (and should probably) deleted
        file.write("""<Settings>
        <AbsProjectBuffer>311220071300</AbsProjectBuffer>
        <ActionEndThreshold>100</ActionEndThreshold>
        <ActionStartThreshold>60</ActionStartThreshold>
        <ActiveSensResult>1</ActiveSensResult>
        <ActiveTrackingPeriod>8</ActiveTrackingPeriod>
        <AllocationMethod>0</AllocationMethod>
        <AutomaticBuffer>0</AutomaticBuffer>
        <ConnectResourceBars>0</ConnectResourceBars>
        <ConstraintHardness>3</ConstraintHardness>
        <CurrencyPrecision>2</CurrencyPrecision>
        <CurrencySymbol></CurrencySymbol>
        <CurrencySymbolPosition>1</CurrencySymbolPosition>""")
        ### Dateformat ###
        ## TODO; Read Datetimeformat
        dateformat="d/MM/yyyy h:mm"
        file.write('<DateTimeFormat>')
        file.write(dateformat)
        file.write("</DateTimeFormat>")

        file.write("""<DefaultRowBuffer>50</DefaultRowBuffer>
        <DrawRelations>1</DrawRelations>
        <DrawShadow>1</DrawShadow>
        <DurationFormat>1</DurationFormat>
        <DurationLevels>2</DurationLevels>
        <ESSLSSFloat>0</ESSLSSFloat>
        <GanttStartDate>080220070800</GanttStartDate>
        <GanttZoomLevel>0.0127138157894736</GanttZoomLevel>
        <GroupFilter>0</GroupFilter>
        <HideGraphMarks>0</HideGraphMarks>
        <Name>Settings</Name>
        <PlanningEndThreshold>60</PlanningEndThreshold>
        <PlanningStartThreshold>20</PlanningStartThreshold>
        <PlanningUnit>1</PlanningUnit>
        <ResAllocation1Color>12632256</ResAllocation1Color>
        <ResAllocation2Color>8421504</ResAllocation2Color>
        <ResAvailableColor>15780518</ResAvailableColor>
        <ResourceChartEndDate>220520071000</ResourceChartEndDate>
        <ResourceChartStartDate>060320060000</ResourceChartStartDate>
        <ResOverAllocationColor>255</ResOverAllocationColor>
        <ShowCanEditResultsInHelp>1</ShowCanEditResultsInHelp>
        <ShowCriticalPath>1</ShowCriticalPath>
        <ShowInputModelInfoInHelp>1</ShowInputModelInfoInHelp>
        <SyncGanttAndResourceChart>0</SyncGanttAndResourceChart>
        <UniqueID>-1</UniqueID>
        <UseResourceScheduling>0</UseResourceScheduling>
        <UserID>0</UserID>
        <ViewDateTimeAsUnits>0</ViewDateTimeAsUnits>""")
        file.write('</Settings>')

        ## Defaults: Obsolete? ###
        file.write("""<Defaults>
        <DefaultCostPerUnit>50</DefaultCostPerUnit>
        <DefaultDisplayDurationType>0</DefaultDisplayDurationType>
        <DefaultDistributionType>2</DefaultDistributionType>
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
        </Defaults>""")

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
        file.write("""<Activities>
        <Name>Activities</Name>
        <UniqueID>-1</UniqueID>
        <UserID>0</UserID>
        </Activities><Activities>""")


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
                file.write("""<Constraints>
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
                </Constraints>""")
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
                name=str(activity.name)
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
        file.write("""<Relations><Name>Relations</Name><UniqueID>-1</UniqueID><UserID>0</UserID></Relations>
        <Relations>""")
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
        file.write("""<ActivityGroups><Name>ActivityGroups</Name><UniqueID>-1</UniqueID><UserID>0</UserID>
        </ActivityGroups><ActivityGroups>""")
        for activitygroup in project_object.activities:
            if len(activitygroup.wbs_id) == 2:
                # Only Activity groups
                file.write("<ActivityGroup>")
                # Unimportant
                file.write("<Expanded>1</Expanded>")
                # Name
                file.write("<Name>")
                name=str(activitygroup.name)
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
            res_name=str(resource.name)
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
        file.write("""<ResourceAssignments><Name>ResourceAssignments</Name><UniqueID>-1</UniqueID><UserID>0</UserID>
        </ResourceAssignments><ResourceAssignments>""")
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
        file.write("""<TrackingList><Name>TrackingList</Name><UniqueID>-1</UniqueID>
        <UserID>0</UserID></TrackingList>""")
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
            # TODO Weird bug, XML won't be correctly formatted if TP name is written
            TP_name=str(TP.tracking_period_name)
            file.write(TP_name)
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



        return print("Write Succesful!")


