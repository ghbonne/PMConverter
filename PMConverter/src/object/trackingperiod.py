__author__ = 'PM Group 8'
from datetime import datetime, date
#from object.activitytracking import ActivityTrackingRecord


class TrackingPeriod(object):
    """
    Bundles some Activity Tracking Records in time"

    :var tracking_period_name: string
    :var tracking_period_statusdate: date
    :var tracking_period_records: List of ActivityTrackingRecords, which are nonzero or contain changes w.r.t. previous tracking period
    """

    def __init__(self, tracking_period_name=" ", tracking_period_statusdate=date.now(), tracking_period_records=[], type_check = True):
        if type_check:
            if not isinstance(tracking_period_name, str):
                raise TypeError('tracking_period_name should be a string!')
            if not isinstance(tracking_period_statusdate, date):
                raise TypeError('tracking_period_statusdate should be a datetime!')
#            if not isinstance(tracking_period_records, list) or not all(isinstance(element, ActivityTrackingRecord) for element in tracking_period_records) :
#                raise TypeError('tracking_period_records should be a list of ActivityTrackingRecords!')

        self.tracking_period_name = tracking_period_name
        self.tracking_period_statusdate = tracking_period_statusdate
        self.tracking_period_records = tracking_period_records

