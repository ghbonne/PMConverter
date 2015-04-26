from objects.activitytracking import ActivityTrackingRecord

__author__ = 'PM Group 8'
from datetime import datetime


class TrackingPeriod(object):
    """
    Bundles some Activity Tracking Records in time"

    :var tracking_period_name: string
    :var tracking_period_statusdate: datetime
    :var tracking_period_records: List of ActivityTrackingRecords, which are nonzero or contain changes w.r.t. previous tracking period
    :var spi
    :var cpi
    :var spi_t
    :var p_factor
    :var sv_t
    """

    def __init__(self, tracking_period_name="", tracking_period_statusdate=datetime.now(),
                 tracking_period_records=[], type_check=True):
        if type_check:
            if not isinstance(tracking_period_name, str):
                raise TypeError('tracking_period_name should be a string!')
            if not isinstance(tracking_period_statusdate, datetime):
                raise TypeError('tracking_period_statusdate should be a datetime!')
            if not isinstance(tracking_period_records, list) or not \
                    all(isinstance(element, ActivityTrackingRecord) for element in tracking_period_records):
                raise TypeError('tracking_period_records should be a list of ActivityTrackingRecords!')

        self.tracking_period_name = tracking_period_name
        self.tracking_period_statusdate = tracking_period_statusdate
        self.tracking_period_records = tracking_period_records
        self.spi = 0
        self.cpi = 0
        self.spi_t = 0
        self.p_factor = 0
        self.sv_t = 0


