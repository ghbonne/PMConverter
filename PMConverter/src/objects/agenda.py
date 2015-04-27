from datetime import datetime, timedelta
import copy

__author__ = 'Eveline'


class Agenda(object):

    """
    :var working_hours: list of booleans with length 24 (index 0 = 00:00 - 01:00): True = working hour
                                                                                   False = no working hour
    :var working_days: list of booleans with length 7 (index 0 = Monday): True = working day
                                                                          False = no working day
    :var holidays: list of holidays
    """

    def __init__(self, working_hours=[0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0],
                 working_days=[1, 1, 1, 1, 1, 0, 0], holidays=[]):
        if len(working_hours) != 24 or not all(working_hours[i] in [0, 1] for i in range(0, len(working_hours))):
            raise TypeError("Working hours should be a list of 24 bits")
        if len(working_days) != 7 or not all(working_days[i] in [0, 1] for i in range(0, len(working_days))):
            raise TypeError("Working days should be a list of 7 bits")

        self.working_hours = working_hours
        self.working_days = working_days
        self.holidays = holidays

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

    def get_number_of_working_days(self):
        counter = 0
        for i in range(0, 7):
            if self.working_days[i]:
                counter += 1
        return counter

    def get_working_hours_in_a_day(self):
        counter = 0
        for i in range(0, 24):
            if self.is_working_hour(i):
                counter += 1
        return counter

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
        if hours > self.get_working_hours_in_a_day():
            raise Exception("Agenda: Wrong parameters for get_end_date: hours can't be bigger than max working hours in a day")
        if self.is_holiday(begin_date.strftime('%d%m%Y')):
            raise Exception("Agenda: Wrong parameters for get_end_date: begin_date can't be a holiday")
        if not self.is_working_day(begin_date.weekday()):
            raise Exception("Agenda: Wrong parameters for get_end_date: begin_date can't be a non-working day")
        end_date = copy.deepcopy(begin_date)

        if days:
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
                end_date += timedelta(days=1)
                weekday = end_date.weekday()
                if self.is_working_day(weekday):
                    if not self.is_holiday(end_date.strftime('%d%m%Y')):
                        days -= 1

        if hours:
            #set end_date to firstworking hour of next day
            end_date += timedelta(days=1)
            first_working_hour = self.get_first_working_hour()
            end_date = end_date.replace(hour=first_working_hour)
            while hours > 0:
                if self.is_working_hour(end_date.hour):
                    end_date += timedelta(hours=1)
                    hours -= 1
                else:
                    end_date += timedelta(hours=1)
        return end_date

    def get_next_date(self, begin_date, workingDaysDuration, workingHoursDuration=0):
        """
        :param begin_date: datetime
        :param workingDaysDuration: integer
        :param workingHoursDuration: integer; smaller then the max working hours in a day
        :return: calculated next datetime with a valid working hour, minimum workingDaysDuration + workingHoursDuration after begin_date, based on working days, holidays and working hours
        """

        #if workingHoursDuration > self.get_working_hours_in_a_day():
        #    raise Exception("Agenda: Wrong parameters for get_next_date: hours can't be bigger than max working hours in a day")
        #if self.is_holiday(begin_date.strftime('%d%m%Y')):
        #    raise Exception("Agenda: Wrong parameters for get_next_date: begin_date can't be a holiday")
        #if not self.is_working_day(begin_date.weekday()):
        #    raise Exception("Agenda: Wrong parameters for get_next_date: begin_date can't be a non-working day")
        next_date = copy.deepcopy(begin_date)  # ensures that a new datetime object is returned

        # first go to first working hour >= begin_date
        if self.is_holiday(begin_date.strftime('%d%m%Y')) or not self.is_working_day(begin_date.weekday()):
            # begin_date is not a valid workingday => set time to first working hour of day and move day to first next valid workingday:
            next_date = next_date.replace(hour=self.get_first_working_hour())

            # move day to a valid non-holiday workingday:
            while self.is_holiday(next_date.strftime('%d%m%Y')) or not self.is_working_day(next_date.weekday()):
                next_date += timedelta(days = 1)

            
        elif not self.is_working_hour(begin_date.time().hour):
            # begin_date is a valid workingday but not time

            if begin_date.time().hour < self.get_last_working_hour():
                # increase hours until next valid working hour
                while not self.is_working_hour(next_date.time().hour):
                    next_date += timedelta(hours = 1)

            else:
                # next working hour is on the next day:
                next_date = next_date.replace(hour=self.get_first_working_hour())
                next_date += timedelta(days = 1)

                # move day to a valid non-holiday workingday:
                while self.is_holiday(next_date.strftime('%d%m%Y')) or not self.is_working_day(next_date.weekday()):
                    next_date += timedelta(days = 1)

        #endIf valid starting datetime

        # end_date is now a valid working datetime >= begin_date

        # first add the necessary workingDaysDuration:
        while workingDaysDuration > 0:
            next_date += timedelta(days = 1)
            if (not self.is_holiday(next_date.strftime('%d%m%Y'))) and self.is_working_day(next_date.weekday()):
                # next valid working datetime found
                workingDaysDuration -= 1

        # next add the necessary workingHoursDuration
        next_date_isSet_to_startOfWorkingDay = False
        while workingHoursDuration > 0:
            if next_date.time().hour > self.get_last_working_hour() or self.is_holiday(next_date.strftime('%d%m%Y')) or not self.is_working_day(next_date.weekday()):
                # no next working hours today => move to first working hour on next day:
                next_date = next_date.replace(hour = self.get_first_working_hour())
                next_date += timedelta(days = 1)
                next_date_isSet_to_startOfWorkingDay = True
                continue

            if not next_date_isSet_to_startOfWorkingDay:
                # continue 1 working hour
                next_date += timedelta(hours = 1)
            else:
                next_date_isSet_to_startOfWorkingDay = False

            if self.is_working_hour(next_date.time().hour):
                # next valid working datetime found
                workingHoursDuration -= 1

        return next_date
                
        



    def get_duration_working_days(self, duration_hours=0):
        """
        :param duration_hours: integer; Working hours needed to complete activity
        :return: timedelta; working days
        """
        working_hours_per_day=0
        for working_hour in self.working_hours:
            if working_hour == True:
                working_hours_per_day+=1
        working_days=duration_hours/working_hours_per_day  # NOTE: integer division
        return timedelta(days=working_days)

    def get_time_between(self, begin_date, end_date):
        days = 0
        hours = 0
        temp_date = copy.deepcopy(begin_date)

        # counting hours
        begin_hour = begin_date.hour
        if begin_hour != self.get_first_working_hour():
            while temp_date.hour <= self.get_last_working_hour():
                if self.is_working_hour(temp_date.hour):
                    hours += 1
                temp_date += timedelta(hours=1)
            temp_date += timedelta(days=1)
            temp_date.replace(hour=self.get_first_working_hour())

        # counting days
        while temp_date.strftime('%d%m%Y') != end_date.strftime('%d%m%Y'):
            weekday = temp_date.weekday()
            if self.is_working_day(weekday):
                if not self.is_holiday(temp_date.strftime('%d%m%Y')):
                    days+=1
            temp_date += timedelta(days=1)

        # counting extra hours
        if end_date.hour == self.get_last_working_hour():
            days += 1
        else:
            while temp_date.hour < end_date.hour:
                if self.is_working_hour(temp_date.hour):
                    hours += 1
                temp_date += timedelta(hours=1)

        return timedelta(days= days, hours= hours)

    def get_previous_working_hour(self, date):
        hour = date.hour
        result = copy.deepcopy(date)
        if hour <= self.get_first_working_hour():
            result -= timedelta(days=1)
            result = result.replace(hour=self.get_last_working_hour()+1)
            while (not self.is_working_day(result.weekday())) or self.is_holiday(result.strftime('%d%m%Y')):
                result -= timedelta(days=1)
        else:
            hour -= 1
            while not self.is_working_hour(hour) and hour >= 0:
                hour -= 1
            hour += 1 #because we want the end of the hour
            result = result.replace(hour=hour)
        return result




