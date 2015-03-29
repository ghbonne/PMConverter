__author__ = 'PM Group 8'

class Activity(object):
    "The core object a project is constituted of"
     # instance variables:
     # _activityId: unique number indicating an activity
     # _activityName: string describing the activity
     # _wbsId: tuple indicating place of the activity in the Work Breakdown Structure (e.g. 1.1 -> (1,1); 1.2.3 -> (1,2,3); etc.)
     # _relationPredecessorsList: list of tuples (activityPointer, time lag)
     # _relationSuccessorsList: list of tuples (activityPointer, time lag)
     # _baselineSchedule: BaselineScheduleRecord object containing the baseline start, baseline end and duration
     # _resourcesList: list of tuples (resourcePointer, how many units of this resource needed when applicable else -1)
     # _baselineCosts: dictionary containing fields "Fixed Cost", "Cost/Hour", "Variable Cost" and "Total Cost" if nonezero
     # _activityTrackingRecords: list of all different ActivityTrackingRecords found in this project
     # class variables:

    def __init__(self):
        pass


