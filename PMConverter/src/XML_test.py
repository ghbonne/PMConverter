import ast #ast.literal_eval
import xml.etree.ElementTree as ET
from object.activity import Activity
from object.baselineschedule import BaselineScheduleRecord

tree = ET.parse('project.xml')
root = tree.getroot()

activity_list=[];
counter = 0

for activities in root.findall('ActivityGroups'):
    for activity in activities.findall('ActivityGroup'):
        #UniqueID
        UniqueID=int(activity.find('UniqueID').text)
        a_g = Activity(int(UniqueID));

        #Name
        Name=activity.find('Name').text
        a_g.name= Name
        activity_list.append(a_g)


for activities in root.findall('Activities'):
    for activity in activities.findall('Activity'):
        #UniqueID
        UniqueID=activity.find('UniqueID').text
        a = Activity(int(UniqueID));
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

        activity_list.append(a)

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




## Testing
for activity in activity_list:
    print(activity.activity_id, activity.successors)
    #print(activity.successors)