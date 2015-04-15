__author__ = 'PM Group 8'
from datetime import datetime, timedelta

class ActivityTrackingRecord(object):
    """
    Update about the status of an Activity concerning its actual start, actual duration, actual cost, percentage completed, etc."

    :var tracking_period: tracking_period, pointer to tracking_period object where this ActivityTrackingRecord is part of
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
    :var tracking_status: Not Started/Started/Finished (enum?)
    :var earned_value: float
    :var planned_value: float
    """

    def __init__(self, tracking_period, activity, actual_start, actual_duration, planned_actual_cost, planned_remaining_cost, 
                 remaining_duration, deviation_pac, deviation_prc, actual_cost, remaining_cost, percentage_completed,
                 tracking_status, earned_value, planned_value):
        self.tracking_period = tracking_period
        self.activity = activity
        self.actual_start = actual_start
        self.actual_duration = actual_duration
        self.planned_actual_cost = planned_actual_cost
        self.planned_remaining_cost = planned_remaining_cost
        self.remaining_duration = remaining_duration
        self.deviation_pac = deviation_pac
        self.deviation_prc = deviation_prc
        self.actual_cost = actual_cost
        self.remaining_cost = remaining_cost
        self.percentage_completed = percentage_completed
        self.tracking_status = tracking_status
        self.earned_value = earned_value
        self.planned_value = planned_value

