__author__ = 'PM Group 8'
from datetime import datetime, timedelta


class ActivityTrackingRecord(object):
    """
    Update about the status of an Activity concerning its actual start, actual duration, actual cost, percentage completed, etc."

    :var tracking_period: tracking_period, pointer to tracking_period objects where this ActivityTrackingRecord is part of
    :var activity: Activity, pointer to the concerning Activity
    :var actual_start: datetime
    :var actual_duration: timedelta
    :var planned_actual_cost: float
    :var planned_remaining_cost: float
    :var remaining_duration: int, expressed in days
    :var deviation_pac: float
    :var deviation_prc: float
    :var actual_cost: float
    :var remaining_cost: float
    :var percentage_completed: int (0-100)
    :var tracking_status: string, Not Started/Started/Finished
    :var earned_value: float
    :var planned_value: float
    """

    def __init__(self, activity=None, actual_start=None, actual_duration=None, planned_actual_cost=0.0,
                 planned_remaining_cost=-1.0, remaining_duration=None, deviation_pac=0.0,
                 deviation_prc=0.0, actual_cost=0.0, remaining_cost=-1.0, percentage_completed=0,
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
            if not isinstance(percentage_completed, int) or not (0 <= percentage_completed <= 100):
                raise TypeError('percentage_completed must be an integer between 0 and 100')

        self.activity = activity
        self.actual_start = actual_start
        self.actual_duration = actual_duration
        self.planned_actual_cost = planned_actual_cost
        if planned_remaining_cost == -1:
            self.planned_remaining_cost = self.activity.baseline_schedule.total_cost
        else:
            self.planned_remaining_cost = planned_remaining_cost
        self.remaining_duration = remaining_duration
        self.deviation_pac = deviation_pac
        self.deviation_prc = deviation_prc
        self.actual_cost = actual_cost
        if remaining_cost == -1:
            self.remaining_cost = self.activity.baseline_schedule.total_cost
        else:
            self.remaining_cost = remaining_cost
        self.percentage_completed = percentage_completed
        self.tracking_status = tracking_status
        self.earned_value = earned_value
        self.planned_value = planned_value

