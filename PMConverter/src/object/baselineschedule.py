__author__ = 'PM Group 8'
from datetime import datetime, timedelta

class BaselineScheduleRecord(object):
    "Contains a baseline start, baseline end and duration of an activity"
    # instance variables:
    # _baselineStart: datetime indicating the scheduled start
    # _baselineEnd: datetime indicating the scheduled end
    # _duration: timedelta = time between _baselineEnd and _baselineStart

    def __init__(self):
        pass

    def getDurationString(self):
        "Return the duration in desired string format"

        if self._duration.seconds > 0:
            return "{0}d {1}h".format(self._duration.days, self._duration.seconds//3600)
        else:
            return "{0}d".format(self._duration.days)
