import datetime

__author__ = 'PM Group 8'


class BaselineScheduleRecord(object):
    """
    Record used in the baseline schedule. Contains the start, end (and calculated from this, the duration),
    the fixed costs, cost per hour and variable costs for every activity

    :var start: datetime
    :var duration: datetime
    :var fixed_cost: float
    :var hourly_cost: float
    :var var_cost: float
    """

    def __init__(self, start=datetime.datetime.now(),
                 duration=datetime.datetime.now() + datetime.timedelta(days=10), fixed_cost=0.0,
                 hourly_cost=0.0, var_cost=0.0):
        # TODO: Typechecking?
        self.start = start
        self.duration = duration
        self.fixed_cost = fixed_cost
        self.hourly_cost = hourly_cost
        self.var_cost = var_cost

    def get_duration_string(self):
        """
        First we calculate the difference between the end and the start, then this is converted to a date format.

        :return difference between end and start in a date format
        """

        if self._duration.seconds > 0:
            return "{0}d {1}h".format(self._duration.days, self._duration.seconds//3600)
        else:
            return "{0}d".format(self._duration.days)
