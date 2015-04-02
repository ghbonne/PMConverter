__author__ = 'PM Group 8'


class BaselineScheduleRecord(object):
    """
    Record used in the baseline schedule. Contains the start, end (and calculated from this, the duration),
    the fixed costs, cost per hour and variable costs for every activity

    :var start: String (date: dd/mm/yyyy)
    :var end: String (date: dd/mm/yyyy)
    :var fixed_cost: float
    :var cost_hourly: float
    :var var_cost: float
    """

    def __init__(self, start="01/01/2001", end="01/01/2001", fixed_cost=0.0, cost_hourly=0.0, var_cost=0.0):
        self.start = start
        self.end = end
        self.fixed_cost = fixed_cost
        self.cost_hourly = cost_hourly
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
