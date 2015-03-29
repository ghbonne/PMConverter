__author__ = 'PM Group 8'
from datetime import datetime, timedelta

class ActivityTrackingRecord(object):
    "Update about the status of an Activity concerning its actual start, actual duration, actual cost, percentage completed, etc."
    # instance variables:
    # _trackingPeriod: pointer to TrackingPeriod object where this ActivityTrackingRecord is part of
    # _activity: pointer to the concerning Activity
    # _actualStart: datetime
    # _actualDuration: timedelta
    # _plannedActualCost
    # _plannedRemainingCost
    # _remainingDuration
    # _deviationPAC
    # _deviationPRC
    # _actualCost
    # _remainingCost
    # _percentageCompleted
    # _TrackingStatus: Not Started/Started/Finished
    # _earnedValue
    # _plannedValue
    # class variables:

    def __init__(self):
        pass

