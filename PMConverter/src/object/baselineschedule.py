from datetime import datetime, timedelta

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

    def __init__(self, start=datetime.now(), end=datetime.now() + timedelta(days=10),
                 duration=timedelta(days=10), fixed_cost=0.0, hourly_cost=0.0, var_cost=0.0, total_cost=0.0, type_check = True):
        """
        Initialize a Baselineschedule record. The data types of the parameters must be the same as the properties of an BaselineScheduleRecord.

        :raises TypeError: one of the parameters is not the right type.
        """
        if type_check:
            if not isinstance(start, datetime):
                raise TypeError('start should be a datetime')
            if not isinstance(end, datetime):
                raise TypeError('end should be a datetime')
            if not isinstance(duration, timedelta):
                raise TypeError('duration should be a timedelta')
            if not isinstance(fixed_cost, float):
                raise TypeError('fixed_cost should be a float')
            if not isinstance(hourly_cost, float):
                raise TypeError('hourly_cost should be a float')
            if not isinstance(var_cost, float):
                raise TypeError('var_cost should be a float')
            if not isinstance(total_cost, float):
                raise TypeError('total_cost should be a float')
        self.start = start
        self.end = end
        self.duration = duration
        self.fixed_cost = fixed_cost
        self.hourly_cost = hourly_cost
        self.var_cost = var_cost
        self.total_cost = total_cost

    def get_duration_string(self):
        """
        First we calculate the difference between the end and the start, then this is converted to a date format.

        :return difference between end and start in a date format
        """

        if self.duration.seconds > 0:
            return "{0}d {1}h".format(self.duration.days, self.duration.seconds//3600)
        else:
            return "{0}d".format(self.duration.days)

