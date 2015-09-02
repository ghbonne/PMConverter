from datetime import datetime, timedelta
import copy
import re

__author__ = 'Project management group 8, Ghent University 2015'


class Agenda(object):

    """
    :var working_hours: list of booleans with length 24 (index 0 = 00:00 - 01:00): True = working hour
                                                                                   False = no working hour
    :var working_days: list of booleans with length 7 (index 0 = Monday): True = working day
                                                                          False = no working day
    :var holidays: list of holidays: holiday = date
    """

    def __init__(self, working_hours=None, working_days=None, holidays=None):
        # avoid mutable default parameters!
        if working_hours is None: working_hours = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0]
        if working_days is None: working_days = [1, 1, 1, 1, 1, 0, 0]
        if holidays is None: holidays = []

        if len(working_hours) != 24 or not all(working_hours[i] in [0, 1] for i in range(0, len(working_hours))):
            raise TypeError("Working hours should be a list of 24 bits")
        if len(working_days) != 7 or not all(working_days[i] in [0, 1] for i in range(0, len(working_days))):
            raise TypeError("Working days should be a list of 7 bits")
        # there has to be at least 1 working hour a day and 1 working day a week:
        if 1 not in working_hours or 1 not in working_days:
            raise Exception("Agenda: No workinghours in a week!")

        self.working_hours = working_hours
        self.working_days = working_days
        self.holidays = holidays

    def is_working_hour(self, hour):
        """
        :param hour: integer value between 0 and 23
        :return: bool
        """
        return self.working_hours[round(hour)]

    def get_last_working_hour(self):
        for i in range(1, 25):
            if self.is_working_hour(24 - i):
                return 24 - i

    def get_first_working_hour(self):
        for i in range(0, 24):
            if self.is_working_hour(i):
                return i

    def get_number_of_working_days(self):
        "This function returns the number of working days in a week."
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
        if 1 not in self.working_hours:
            raise Exception("Agenda:set_non_working_hour({0}): No workinghours in a day!".format(hour))

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
        if 1 not in self.working_days:
            raise Exception("Agenda:set_non_working_day({0}): No workingdays in a week!".format(day))

    def is_holiday(self, holiday):
        return holiday in self.holidays

    def set_holiday(self, holiday):
        """
        :param holiday: date
        :return:
        """
        self.holidays.append(holiday)

    def get_end_date(self, begin_date, days, hours=0):
        """
        :param begin_date: datetime
        :param days: integer
        :param hours: integer;
        :return: calculated end date (end of the working hour), based on working days, holidays and working hours
        """
        # ensure days and hours is rounded:
        days = round(days)
        hours = round(hours)

        # get_next_date returns the start of the next valid workinghour after the given duration => give duration - 1 and manually add 1 hour afterwards to get an end date
        if hours >= 1:
            return self.get_next_date(begin_date, days, hours - 1) + timedelta(hours = 1)
        elif days >= 1:
            return self.get_next_date(begin_date, days - 1, self.get_working_hours_in_a_day() - 1) + timedelta(hours = 1)
        else:
            # days == 0 and hours == 0
            # => return the end of the previous workinghour
            return self.get_previous_working_hour_end(begin_date)

    def get_next_date(self, begin_date, workingDaysDuration, workingHoursDuration=0):
        """
        Note: expects rounded dates to the hour
        :param begin_date: datetime
        :param workingDaysDuration: integer
        :param workingHoursDuration: integer; smaller then the max working hours in a day
        :return: calculated next datetime with a valid working hour, minimum workingDaysDuration + workingHoursDuration after begin_date, based on working days, holidays and working hours
        """
        # ensure days and hours is rounded:
        workingDaysDuration = round(workingDaysDuration)
        workingHoursDuration = round(workingHoursDuration)

        # ensure that a new datetime object is returned and drop all time data smaller than an hour:
        next_date = datetime(year= begin_date.year, month= begin_date.month, day= begin_date.day, hour= begin_date.hour)

        # first go to first working hour >= begin_date
        if self.is_holiday(begin_date.date()) or not self.is_working_day(begin_date.weekday()):
            # begin_date is not a valid workingday => set time to first working hour of day and move day to first next valid workingday:
            next_date = next_date.replace(hour=self.get_first_working_hour())

            # move day to a valid non-holiday workingday:
            while self.is_holiday(next_date.date()) or not self.is_working_day(next_date.weekday()):
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
                while self.is_holiday(next_date.date()) or not self.is_working_day(next_date.weekday()):
                    next_date += timedelta(days = 1)

        #endIf valid starting datetime

        # end_date is now a valid working datetime >= begin_date

        # first add the necessary workingDaysDuration:
        while workingDaysDuration > 0:
            next_date += timedelta(days = 1)
            if (not self.is_holiday(next_date.date())) and self.is_working_day(next_date.weekday()):
                # next valid working datetime found
                workingDaysDuration -= 1

        # next add the necessary workingHoursDuration
        next_date_isSet_to_startOfWorkingDay = False
        while workingHoursDuration > 0:
            if next_date.time().hour > self.get_last_working_hour() or self.is_holiday(next_date.date()) or not self.is_working_day(next_date.weekday()):
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
                
        
    def get_workingDuration_timedelta(self, duration_hours = 0):
        """
        Converts an int of workinghours to a timedelta of workingdays and workinghours
        :param duration_hours: integer; Working hours needed to complete activity
        :return: timedelta; working days + remaining working hours
        """
        working_hours_per_day = self.get_working_hours_in_a_day()
        working_days = int(duration_hours / working_hours_per_day)
        return timedelta(days = working_days, hours = duration_hours - working_days * working_hours_per_day)

    def get_workingDuration_workingHours(self, duration_timedelta):
        """
        Converts a timedelta of workingdays and workinghours to an int of only workinghours.
        :param duration_timedelta: timedelta in workingdays and workinghours
        :return: int, workinghours
        """
        return duration_timedelta.days * self.get_working_hours_in_a_day() + round(duration_timedelta.seconds / 3600.)

    def get_duration_working_days(self, duration_hours=0):
        """
        :param duration_hours: integer; Working hours needed to complete activity
        :return: timedelta; working days without excess hours
        """
        working_hours_per_day=0
        for working_hour in self.working_hours:
            if working_hour == True:
                working_hours_per_day+=1
        working_days=duration_hours/working_hours_per_day  # NOTE: integer division
        return timedelta(days=working_days)

    def get_time_between(self, begin_date, end_date):
        " This function calculates the working time between the given dates and returns a timedelta in workingdays and workinghours"

        startingDate = begin_date if begin_date <= end_date else end_date
        endingDate = end_date if begin_date <= end_date else begin_date

        # ensure that endingDate is the start of a valid workinghour:
        endingDate = self.get_next_date(endingDate, 0,0)
        currentDate = self.get_next_date(startingDate, 0,0)

        workinghours = 0
        while currentDate < endingDate:
            currentDate = self.get_next_date(currentDate, 0, 1)
            workinghours += 1

        # convert time in working_hours to timedelta
        workingHoursInADay = self.get_working_hours_in_a_day()
        days = int(workinghours / workingHoursInADay)
        hours = workinghours - days * workingHoursInADay

        return timedelta(days= days, hours = hours)

    def get_previous_working_hour_end(self, date):
        "This function returns the end of the previous workinghour."
        hour = date.hour
        # ensure that a new datetime object is returned and drop all time data smaller than an hour:
        result = datetime(year= date.year, month= date.month, day= date.day, hour= date.hour)

        if hour <= self.get_first_working_hour():
            result -= timedelta(days=1)
            result = result.replace(hour=self.get_last_working_hour()+1)
            while (not self.is_working_day(result.weekday())) or self.is_holiday(result.date()):
                result -= timedelta(days=1)
        else:
            hour -= 1
            while hour >= 0 and not self.is_working_hour(hour):
                hour -= 1
            hour += 1 #because we want the end of the hour
            result = result.replace(hour=hour)
        return result

    def convert_durationString_to_workingHours(self, duration_string):
        """
        This function converts a duration from a string to an int of workinghours
        :param duration_string: string with working duration formatted in weeks, days and hours
        :return: int, workinghours
        """
        durationMatch = re.match(r"(\s*)?((?P<weeks>\d+)w)?(\s*)?((?P<days>\d+)d)?(\s*)?((?P<hours>\d+)(h|u))?(\s*)?", duration_string)

        if durationMatch is not None:
            weeks = durationMatch.group("weeks")
            days = durationMatch.group("days")
            hours = durationMatch.group("hours")

            workingHoursInDay = self.get_working_hours_in_a_day()

            totalWorkingHours = int(weeks) * self.get_number_of_working_days() * workingHoursInDay if weeks is not None else 0
            totalWorkingHours += int(days) * workingHoursInDay if days is not None else 0
            totalWorkingHours += int(hours) if hours is not None else 0
            return totalWorkingHours 
        else:
            # string is empty or does not match format
            return 0

    def convert_workingHours_to_durationString(self, workingHoursDuration, convertDaysToWeeks= True):
        """
        This function converts a duration in workinghours to a formatted string
        :param workingHoursDuration: int, workinghours
        :param convertDaysToWeeks: bool, sets behaviour to convert days > days in workingweek to a week
        :return: formatted string
        """
        workingHoursInDay = self.get_working_hours_in_a_day()
        workingDaysInWeek = self.get_number_of_working_days()
        totalWorkingDays = int(workingHoursDuration / workingHoursInDay)
        
        weeks = int(totalWorkingDays / workingDaysInWeek) if convertDaysToWeeks else 0
        days = totalWorkingDays - weeks * workingDaysInWeek
        hours = workingHoursDuration - totalWorkingDays * workingHoursInDay

        # build output string:
        durationString = str(weeks) + "w" if weeks else ""
        durationString += (" " if durationString and days else "") + str(days) + "d" if days else ""
        durationString += (" " if durationString and hours else "") + str(hours) + "h" if hours else ""
        
        return durationString

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented


