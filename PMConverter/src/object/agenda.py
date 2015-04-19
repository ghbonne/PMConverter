from datetime import datetime, timedelta
import copy

__author__ = 'Eveline'


class Agenda(object):

    """
    :var working_hours: list of booleans with length 24: True = working hour
                                                        False = no working hour
    :var working_days: list of booleans with length 7: True = working day
                                                       False = no working day
    :var holidays: list of holidays
    """

    def __init__(self):
        self.working_hours = [True] * 24
        self.working_days = [True] * 7
        self.holidays = []

    def is_working_hour(self, hour):
        """
        :param hour: integer value between 0 and 23
        :return: bool
        """
        return self.working_hours[hour]

    def get_last_working_hour(self):
        for i in range(1, 25):
            if self.is_working_hour(24 - i):
                return 24 - i

    def get_first_working_hour(self):
        for i in range(0, 24):
            if self.is_working_hour(i):
                return i

    def set_non_working_hour(self, hour):
        """
        :param hour: integer value between 0 and 23
        """
        self.working_hours[hour] = False

    def is_working_day(self, day):
        """
        :param day: integer value between 0 and 6
        :return: bool
        """
        return self.working_days[day]

    def set_non_working_day(self, day):
        """
        :param day: integer value between 0 and 6
        """
        self.working_days[day] = False

    def is_holiday(self, holiday):
        return holiday in self.holidays

    def set_holiday(self, holiday):
        """
        :param holiday: string; format: 27122010 (=27/12/2010)
        :return:
        """
        self.holidays.append(holiday)

    def get_end_date(self, begin_date, days, hours=0):
        """
        :param begin_date: datetime
        :param days: integer
        :param hours: integer; smaller then the max working hours in a day
        :return: calculated end date, based on working days, holidays and working hours
        """
        end_date = copy.deepcopy(begin_date)

        # calculate when they stop working
        end_hour_of_a_day = self.get_last_working_hour() + 1
        if end_hour_of_a_day == 24:
            end_date = end_date.replace(hour=0)
            end_date += timedelta(days=1)
        else:
            # set end hour = first day
            end_date = end_date.replace(hour=end_hour_of_a_day)
        days -= 1

        # rest of the days
        while days > 0:
            weekday = end_date.weekday()
            if self.is_working_day(weekday):
                if not self.is_holiday(end_date.strftime('%d%m%Y')):
                    end_date += timedelta(days=1)
                    days -= 1
                else:
                    end_date += timedelta(days=1)
            else:
                end_date += timedelta(days=1)

        if hours:
            #first set hours back to begin of the day
            first_working_hour = self.get_first_working_hour()
            end_date = end_date.replace(hour=first_working_hour)
            while hours > 0:
                if self.is_working_hour(end_date.hour):
                    end_date += timedelta(hours=1)
                    hours -= 1
                else:
                    end_date += timedelta(hours=1)
        return end_date