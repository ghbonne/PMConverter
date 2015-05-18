__author__ = 'PM Group 8'
import datetime
from objects.resource import ResourceType

class ActivityTrackingRecord(object):
    """
    Update about the status of an Activity concerning its actual start, actual duration, actual cost, percentage completed, etc."

    :var activity: Activity, pointer to the concerning Activity
    :var actual_start: datetime  # if not started yet, is set to datetime.max
    :var actual_duration: timedelta
    :var planned_actual_cost: float
    :var planned_remaining_cost: float
    :var remaining_duration: timedelta, expressed in workingHours and workingdays
    :var deviation_pac: float
    :var deviation_prc: float
    :var actual_cost: float
    :var remaining_cost: float
    :var percentage_completed: float (0.0 - 100.0)
    :var tracking_status: string, Not Started/Started/Finished
    :var earned_value: float
    :var planned_value: float
    """

    def __init__(self, activity=None, actual_start=None, actual_duration=None, planned_actual_cost=0.0,
                 planned_remaining_cost=-1.0, remaining_duration=None, deviation_pac=0.0,
                 deviation_prc=0.0, actual_cost=0.0, remaining_cost=-1.0, percentage_completed=0.0,
                 tracking_status='', earned_value=0.0, planned_value=0.0, type_check=True):
        if type_check:
            if not isinstance(planned_actual_cost, float):
                raise TypeError('planned_actual_cost must be a float!')
            if not isinstance(planned_remaining_cost, float):
                raise TypeError('planned_remaining_cost must be a float!')
            if not isinstance(deviation_pac, float):
                raise TypeError('deviation_pac must be a float!')
            if not isinstance(deviation_prc, float):
                raise TypeError('deviation_prc must be a float!')
            if not isinstance(actual_cost, float):
                raise TypeError('actual_cost must be a float!')
            if not isinstance(earned_value, float):
                raise TypeError('earned_value must be a float!')
            if not isinstance(tracking_status, str):
                raise TypeError('tracking_status must be a string!')
            if not isinstance(planned_value, float):
                raise TypeError('planned_value must be a float!')
            if not isinstance(percentage_completed, float) or not (0.0 <= percentage_completed <= 100.0): 
                if not isinstance(percentage_completed, float) or not (abs(percentage_completed) < 1e-10 or abs(percentage_completed - 100) < 1e-10):
                    # percentage completed is not a float or not within allowed margins around bounds
                    raise TypeError('percentage_completed must be a float between 0.0 and 100.0')
                else:
                    # percentage_completed is a float => could be tiny bit out of bounds => round it to the bounds
                    percentage_completed = float(round(percentage_completed))
                    print("ActivityTracking: percentage_completed needed to be rounded to bounds, probably because of float rounding issues.")

        self.activity = activity
        self.actual_start = actual_start
        self.actual_duration = actual_duration
        self.planned_actual_cost = planned_actual_cost
        #if planned_remaining_cost == -1:
        #    self.planned_remaining_cost = self.activity.baseline_schedule.total_cost
        #else:
        self.planned_remaining_cost = planned_remaining_cost
        self.remaining_duration = remaining_duration
        self.deviation_pac = deviation_pac
        self.deviation_prc = deviation_prc
        self.actual_cost = actual_cost
        #if remaining_cost == -1:
         #   self.remaining_cost = self.activity.baseline_schedule.total_cost
        #else:
        self.remaining_cost = remaining_cost
        self.percentage_completed = percentage_completed
        self.tracking_status = tracking_status
        self.earned_value = earned_value
        self.planned_value = planned_value

    @staticmethod
    def calculate_activityTrackingRecord_derived_values(activity, actualCostDev, actualDuration_hours, agenda, percentageComplete, remainingCostDev, remainingDuration_hours, statusdate_datetime, actualStart):
        """This function calculates the derived values of a low-level activity tracking record, given the minimum values of ProTrack.
        :param percentageComplete: in range(0.0, 100.0), with 100 included
        """

        #if actualDuration_hours > 0:
        if actualStart.date() < datetime.datetime.max.date(): # if activity not yet started, its actualStart is set to datetime.max
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
            planned_actual_cost = 0.0
            planned_remaining_cost = activity.baseline_schedule.total_cost
        
        actual_cost = planned_actual_cost + actualCostDev
        remaining_cost = planned_remaining_cost + remainingCostDev
        
        # Calculate EV:
        if remainingDuration_hours == 0:
            # activity finished
            earned_value = activity.baseline_schedule.total_cost
        elif actualDuration_hours == 0:
            # activity not yet started:
            earned_value = 0.
        else:
            # activity running:
            earned_value = activity.baseline_schedule.fixed_cost + percentageComplete/100.0 * (activity.baseline_schedule.var_cost + activity.resource_cost)
        
        # Calculate PV:
        if statusdate_datetime >= activity.baseline_schedule.end:
            # activity should be finished according to baselinschedule
            planned_value = activity.baseline_schedule.total_cost
        elif statusdate_datetime < activity.baseline_schedule.start:
            # activity is not yet started according to baselineschedule
            planned_value = 0.
        else:
            # activity is running
            activityRunningDuration = agenda.get_time_between(activity.baseline_schedule.start, statusdate_datetime)
            activityRunningDuration_workingHours = activityRunningDuration.days * agenda.get_working_hours_in_a_day() + activityRunningDuration.seconds / 3600
        
            planned_value = activity.baseline_schedule.fixed_cost + activityRunningDuration_workingHours * activity.baseline_schedule.hourly_cost
                    
            # add costs of resources until now:
            for resourceTuple in activity.resources:
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
        return actual_cost, earned_value, planned_actual_cost, planned_remaining_cost, planned_value, remaining_cost

    @staticmethod
    def construct_activityGroup_trackingRecord(activityGroup, childActivityIds, currentTrackingPeriod_records_dict, current_status_date, agenda):
        """This function constructs an aggregated trackingRecord of an activityGroup
        :returns: ActivityTrackingRecord of aggregated activityGroup
        """

        total_earned_value = 0.0
        total_planned_value = 0.0
        total_actual_cost = 0.0
        earliest_start = datetime.datetime.max
        latest_finish = datetime.datetime.min

        if childActivityIds:
            for childActivityId in childActivityIds:
                childActivityTrackingRecord = currentTrackingPeriod_records_dict[childActivityId]
                total_earned_value += childActivityTrackingRecord.earned_value
                total_planned_value += childActivityTrackingRecord.planned_value
                total_actual_cost += childActivityTrackingRecord.actual_cost
                if childActivityTrackingRecord.actual_start < earliest_start:
                    earliest_start = childActivityTrackingRecord.actual_start
                if childActivityTrackingRecord.percentage_completed < 100 and abs(childActivityTrackingRecord.percentage_completed - 100) > 1e-10:
                    # activity is definitely not finished yet
                    childActivity_latestDate = current_status_date
                elif childActivityTrackingRecord.actual_start.date() < datetime.datetime.max.date():
                    childActivity_latestDate = agenda.get_end_date(childActivityTrackingRecord.actual_start, childActivityTrackingRecord.actual_duration.days, round(childActivityTrackingRecord.actual_duration.seconds / 3600.0))
                else:
                    childActivity_latestDate = current_status_date

                if childActivity_latestDate > latest_finish:
                    latest_finish = childActivity_latestDate
        
        if earliest_start.date() < datetime.datetime.max.date() and latest_finish > datetime.datetime.min:
            total_actual_duration = agenda.get_time_between(earliest_start, latest_finish)
        else:
            total_actual_duration = datetime.timedelta(0)

        percentage_completed = 100 * total_earned_value / activityGroup.baseline_schedule.total_cost if activityGroup.baseline_schedule.total_cost else 0.0
        if percentage_completed > 100.0: percentage_completed = 100.0 # avoid percentage_completed to be larger than 100.0
        return ActivityTrackingRecord(activity= activityGroup, percentage_completed= percentage_completed, earned_value= total_earned_value, planned_value= total_planned_value,
                                      actual_cost= total_actual_cost, actual_duration= total_actual_duration)

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented