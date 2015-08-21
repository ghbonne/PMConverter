from objects.activitytracking import ActivityTrackingRecord

__author__ = 'Project management group 8, Ghent University 2015'
from datetime import datetime


class TrackingPeriod(object):
    """
    Bundles some Activity Tracking Records in time"

    :var tracking_period_name: string
    :var tracking_period_statusdate: datetime
    :var tracking_period_records: List of ActivityTrackingRecords, which are nonzero or contain changes w.r.t. previous tracking period

    extra variables only for excel visualisations purposes: their correct value is only set in the XLSXParser.from_schedule_object
    :var spi
    :var cpi
    :var spi_t
    :var p_factor
    :var sv_t : duration in workinghours
    """

    def __init__(self, tracking_period_name="", tracking_period_statusdate=datetime.now(),
                 tracking_period_records=None, type_check=True):
        # avoid mutable default parameters!
        if tracking_period_records is None: tracking_period_records = []

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

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented


