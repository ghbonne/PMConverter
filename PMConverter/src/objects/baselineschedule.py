import datetime

__author__ = 'PM Group 8'


class BaselineScheduleRecord(object):
    """
    Record used in the baseline schedule. Contains the start, end (and calculated from this, the duration),
    the fixed costs, cost per hour and variable costs for every activity

    :var start: datetime
    :var end: datetime
    :var duration: timedelta
    :var fixed_cost: float
    :var hourly_cost: float
    :var var_cost: float
    :var total_cost : float
    """

    def __init__(self, start=datetime.datetime.now(), end=datetime.datetime.now() + datetime.timedelta(days=10),
                 duration=datetime.timedelta(days=10), fixed_cost=0.0, hourly_cost=0.0, var_cost=None, total_cost=0.0,
                 type_check=True):
        if type_check:
            if not isinstance(start, datetime.datetime):
                raise TypeError('BaselineScheduleRecord: start should be a datetime')
            if not isinstance(end, datetime.datetime):
                raise TypeError('BaselineScheduleRecord: end should be a datetime')
            if not isinstance(duration, datetime.timedelta):
                raise TypeError('BaselineScheduleRecord: duration should be a timedelta')
            if not isinstance(fixed_cost, float):
                raise TypeError('BaselineScheduleRecord: fixed_cost should be a float')
            if not isinstance(hourly_cost, float):
                raise TypeError('BaselineScheduleRecord: hourly_cost should be a float')
            # NOTE: var_cost is allowed to be None => indicating that this is not an activity of the lowest level, else it should be a float
            if var_cost is not None:
                # typecheck:
                if not isinstance(var_cost, float):
                    raise TypeError('BaselineScheduleRecord: var_cost should be a float if it is not part of an activityGroup')
            if not isinstance(total_cost, float):
                raise TypeError('BaselineScheduleRecord: total_cost should be a float')
        self.start = start
        self.end = end
        self.duration = duration
        self.fixed_cost = fixed_cost
        self.hourly_cost = hourly_cost
        self.var_cost = var_cost
        self.total_cost = total_cost

    #def get_duration_string(self):
    #    """
    #    First we calculate the difference between the end and the start, then this is converted to a date format.

    #    :return difference between end and start in a date format
    #    """

    #    if self._duration.seconds > 0:
    #        return "{0}d {1}h".format(self._duration.days, round(self._duration.seconds / 3600.0))
    #    else:
    #        return "{0}d".format(self._duration.days)

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented