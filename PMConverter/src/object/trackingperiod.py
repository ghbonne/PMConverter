__author__ = 'PM Group 8'
from datetime import datetime

class TrackingPeriod(object):
    "Bundles some Activity Tracking Records in time"
    # instance variables:
    # _trackingPeriodName: string
    # _trackingPeriodStatusDate: datetime
    # _trackingPeriodRecords: List of ActivityTrackingRecords which are nonezero or contain changes w.r.t. previous tracking period
    # class variables:

    def __init__(self, trackingPeriodName='', trackingPeriodStatusDate='', trackingPeriodRecords=''):
        # TODO: Typechecking?
        self.trackingPeriodName= trackingPeriodName
        self.trackingPeriodStatusDate=trackingPeriodStatusDate
        self.trackingPeriodRecords=trackingPeriodRecords

