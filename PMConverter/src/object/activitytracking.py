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

    def __init__(self, trackingPeriod=None, activity=None, actualStart=None, actualDuration=None, plannedActualCost=0, plannedRemainingCost=0,
                 remainingDuration=None, deviationPAC=None,deviationPRC=None, actualCost=0, remainingCost=0, percentageCompleted=0,
                 trackingStatus=0, earnedValue=0, plannedValue=0):
        self.trackingPeriod=trackingPeriod
        self.activity=activity
        self.actualStart=actualStart
        self.actualDuration=actualDuration
        self.plannedActualCost=plannedActualCost
        self.plannedRemainingCost=plannedRemainingCost
        self.remainingDuration=remainingDuration
        self.deviationPAC=deviationPAC
        self.deviationPRC=deviationPRC
        self.actualCost=actualCost
        self.remainingCost=remainingCost
        self.percentageCompleted=percentageCompleted
        self.trackingStatus=trackingStatus
        self.earnedValue=earnedValue
        self.plannedValue=plannedValue

