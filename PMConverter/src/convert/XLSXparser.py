import calendar

__author__ = 'PM Group 8'

import xlsxwriter
import openpyxl
import re
import datetime
import itertools
from operator import attrgetter

from objects.activity import Activity
from objects.activitytracking import ActivityTrackingRecord
from objects.baselineschedule import BaselineScheduleRecord
from objects.projectobject import ProjectObject
from objects.resource import Resource, ResourceType
from objects.riskanalysisdistribution import RiskAnalysisDistribution, DistributionType, ManualDistributionUnit
from objects.trackingperiod import TrackingPeriod
from convert.fileparser import FileParser
from objects.agenda import Agenda
from visual.enums import ExcelVersion


class XLSXParser(FileParser):
    """
    Class to convert ProjectObjects to .xlsx files and vice versa. Shout out to John McNamara for his xlsxwriter library
    and the guys from openpyxl.

    """

    def __init__(self):
        super().__init__()

    def to_schedule_object(self, file_path_input):
        # TODO: determine if we're reading a basic or extended version (#columns)

        workbook = openpyxl.load_workbook(file_path_input)
        project_control_sheets = []
        agenda_sheet = None
        for name in workbook.get_sheet_names():
            if "Baseline Schedule" in name:
                activities_sheet = workbook.get_sheet_by_name(name)
            elif "Resources" in name:
                resource_sheet = workbook.get_sheet_by_name(name)
            elif "Risk Analysis" in name:
                risk_analysis_sheet = workbook.get_sheet_by_name(name)
            elif "TP" in name:
                project_control_sheets.append(workbook.get_sheet_by_name(name))
            elif "Agenda" in name:
                agenda_sheet = workbook.get_sheet_by_name(name)

        if activities_sheet.cell(row=2, column=14).value:  # We know if is an extended version if the baseline schedule contains 14 columns
            # We are processing an extended version
            # First we process the resources sheet, we store them in a dict, with index the resource name, to access them
            # easily later when we are processing the activities.
            resources_dict = self.process_resources(resource_sheet)
            # Then we process the risk analysis sheet, again everything is stored in a dict, now index is activity_id.
            risk_analysis_dict = self.process_risk_analysis(risk_analysis_sheet, ExcelVersion.EXTENDED)
            # Finally, the sheet with activities is processed, using the dicts we created above.
            # Again, a new dict is created, to process all tracking periods more easily
            activities_dict = self.process_baseline_schedule(activities_sheet, resources_dict, risk_analysis_dict, ExcelVersion.EXTENDED)
            tracking_periods = self.process_project_controls(project_control_sheets, activities_dict, ExcelVersion.EXTENDED)
            if agenda_sheet:
                agenda = self.process_agenda(agenda_sheet)
            else:
                agenda = Agenda()

            return ProjectObject(activities=[i[1] for i in sorted(activities_dict.values(), key=lambda x: x[1].wbs_id)],
                                 resources=sorted(resources_dict.values(), key=lambda x: x.resource_id),
                                 tracking_periods=tracking_periods, agenda=agenda)
        else:
            # We are processing the basic version
            resources_dict = self.process_resources(resource_sheet)
            risk_analysis_dict = self.process_risk_analysis(risk_analysis_sheet, ExcelVersion.BASIC)
            activities_dict = self.process_baseline_schedule(activities_sheet, resources_dict, risk_analysis_dict, ExcelVersion.BASIC)
            tracking_periods = self.process_project_controls(project_control_sheets, activities_dict, ExcelVersion.BASIC)
            if agenda_sheet:
                agenda = self.process_agenda(agenda_sheet)
            else:
                agenda = Agenda()

            return ProjectObject(activities=[i[1] for i in sorted(activities_dict.values(), key=lambda x: x[0])],
                                 resources=list(resources_dict.values()), tracking_periods=tracking_periods,
                                 agenda=agenda)

    def process_agenda(self, agenda_sheet):
        working_days = [0]*7
        working_hours = [0]*24
        holidays = []
        for i in range(0, 24):
            if agenda_sheet.cell(row=i+2, column=2).value == "Yes":
                working_hours[i] = 1
        for i in range(0, 7):
            if agenda_sheet.cell(row=i+2, column=5).value == "Yes":
                working_days[i] = 1
        i = 2
        while agenda_sheet.cell(row=i, column=7).value:
            holidays.append(agenda_sheet.cell(row=i, column=7).value)
        return Agenda(working_hours=working_hours, working_days=working_days, holidays=holidays)

    def process_project_controls(self, project_control_sheets, activities_dict, excel_version):
        tracking_periods = []
        if excel_version == ExcelVersion.EXTENDED:
            for project_control_sheet in project_control_sheets:
                if type(project_control_sheet.cell(row=1, column=3).value) is float:
                    tp_date = datetime.datetime.utcfromtimestamp(((project_control_sheet.cell(row=1, column=3)
                                                                   .value - 25569)*86400))
                else:
                    tp_date = project_control_sheet.cell(row=1, column=3).value
                tp_name = project_control_sheet.cell(row=1, column=6).value
                act_track_records = []
                for curr_row in range(self.get_nr_of_header_lines(project_control_sheet), project_control_sheet.get_highest_row()+1):
                    activity_id = int(project_control_sheet.cell(row=curr_row, column=1).value)
                    actual_start = None  # Set a default value in case there is nothing in that cell
                    if project_control_sheet.cell(row=curr_row, column=12).value:
                        if type(project_control_sheet.cell(row=curr_row, column=12).value) is float:
                            actual_start = datetime.datetime.utcfromtimestamp(((project_control_sheet.cell(row=curr_row, column=12)
                                                                                .value - 25569)*86400))  # ugly hack to convert
                        else:
                            actual_start = project_control_sheet.cell(row=curr_row, column=12).value
                    actual_duration = None
                    if project_control_sheet.cell(row=curr_row, column=13).value:
                        actual_duration_split = project_control_sheet.cell(row=curr_row, column=13).value.split("d")
                        actual_duration_days = int(actual_duration_split[0])
                        actual_duration_hours = 0  # We need to set this default value for the next loop
                        if len(actual_duration_split) > 1 and actual_duration_split[1] != '':
                            actual_duration_hours = int(actual_duration_split[1][1:-1])  # first char = " "; last char = "h"
                        actual_duration = datetime.timedelta(days=actual_duration_days, hours=actual_duration_hours)
                    pac = 0.0
                    if project_control_sheet.cell(row=curr_row, column=14).value:
                        pac = float(project_control_sheet.cell(row=curr_row, column=14).value)
                    prc = -1.0
                    if project_control_sheet.cell(row=curr_row, column=15).value:
                        prc = float(project_control_sheet.cell(row=curr_row, column=15).value)
                    remaining_duration = None
                    if project_control_sheet.cell(row=curr_row, column=16).value:
                        remaining_duration_split = project_control_sheet.cell(row=curr_row, column=16).value.split("d")
                        remaining_duration_days = int(remaining_duration_split[0])
                        remaining_duration_hours = 0  # We need to set this default value for the next loop
                        if len(remaining_duration_split) > 1 and remaining_duration_split[1] != '':
                            remaining_duration_hours = int(remaining_duration_split[1][1:-1])  # first char = " "; last char = "h"
                        remaining_duration = datetime.timedelta(days=remaining_duration_days, hours=remaining_duration_hours)
                    pac_dev = 0.0
                    if project_control_sheet.cell(row=curr_row, column=17).value:
                        pac_dev = float(project_control_sheet.cell(row=curr_row, column=17).value)
                    prc_dev = 0.0
                    if project_control_sheet.cell(row=curr_row, column=18).value:
                        prc_dev = float(project_control_sheet.cell(row=curr_row, column=18).value)
                    actual_cost = 0.0
                    if project_control_sheet.cell(row=curr_row, column=19).value:
                        actual_cost = float(project_control_sheet.cell(row=curr_row, column=19).value)
                    remaining_cost = 0.0
                    if project_control_sheet.cell(row=curr_row, column=20).value:
                        remaining_cost = float(project_control_sheet.cell(row=curr_row, column=20).value)
                    percentage_completed_str = project_control_sheet.cell(row=curr_row, column=21).value  # always given
                    # In the test file, some percentages are strings, some are ints, some are floats (#YOLO)
                    if type(percentage_completed_str) != str:
                        percentage_completed_str *= 100
                    if type(percentage_completed_str) == float:
                        percentage_completed_str = int(percentage_completed_str)
                    percentage_completed_str = str(percentage_completed_str)
                    if not percentage_completed_str[-1].isdigit():
                        percentage_completed_str = percentage_completed_str[:-1]
                    if float(percentage_completed_str) < 1:
                        percentage_completed_str = float(float(percentage_completed_str)*100)
                    percentage_completed = int(percentage_completed_str)
                    if percentage_completed > 100:
                        print("Gilles V has made a stupid error by putting some ifs & transformations in XLSXparser.py "
                              "with the variable percentage_completed_str")
                    tracking_status = ''
                    if project_control_sheet.cell(row=curr_row, column=22).value:
                        tracking_status = project_control_sheet.cell(row=curr_row, column=22).value
                    earned_value = float(project_control_sheet.cell(row=curr_row, column=23).value)
                    planned_value = float(project_control_sheet.cell(row=curr_row, column=24).value)
                    act_track_record = ActivityTrackingRecord(activity=activities_dict[activity_id][1],
                                                              actual_start=actual_start, actual_duration=actual_duration,
                                                              planned_actual_cost=pac, planned_remaining_cost=prc,
                                                              remaining_duration=remaining_duration, deviation_pac=pac_dev,
                                                              deviation_prc=prc_dev, actual_cost=actual_cost,
                                                              remaining_cost=remaining_cost,
                                                              percentage_completed=percentage_completed,
                                                              tracking_status=tracking_status, earned_value=earned_value,
                                                              planned_value=planned_value)
                    act_track_records.append(act_track_record)
                tracking_periods.append(TrackingPeriod(tracking_period_name=tp_name, tracking_period_statusdate=tp_date,
                                                       tracking_period_records=act_track_records))
        else:
            for project_control_sheet in project_control_sheets:
                if type(project_control_sheet.cell(row=1, column=3).value) is float:
                    tp_date = datetime.datetime.utcfromtimestamp(((project_control_sheet.cell(row=1, column=3)
                                                                   .value - 25569)*86400))
                else:
                    tp_date = project_control_sheet.cell(row=1, column=3).value
                tp_name = project_control_sheet.cell(row=1, column=5).value
                act_track_records = []
                for curr_row in range(self.get_nr_of_header_lines(project_control_sheet), project_control_sheet.get_highest_row()+1):
                    activity_id = int(project_control_sheet.cell(row=curr_row, column=1).value)
                    actual_start = None  # Set a default value in case there is nothing in that cell
                    if project_control_sheet.cell(row=curr_row, column=3).value:
                        if type(project_control_sheet.cell(row=curr_row, column=3).value) is float:
                            actual_start = datetime.datetime.utcfromtimestamp(((project_control_sheet.cell(row=curr_row, column=3)
                                                                                .value - 25569)*86400))  # ugly hack to convert
                        else:
                            actual_start = project_control_sheet.cell(row=curr_row, column=3).value
                    actual_duration = None
                    if project_control_sheet.cell(row=curr_row, column=4).value:
                        actual_duration_split = project_control_sheet.cell(row=curr_row, column=4).value.split("d")
                        actual_duration_days = int(actual_duration_split[0])
                        actual_duration_hours = 0  # We need to set this default value for the next loop
                        if len(actual_duration_split) > 1 and actual_duration_split[1] != '':
                            actual_duration_hours = int(actual_duration_split[1][1:-1])  # first char = " "; last char = "h"
                        actual_duration = datetime.timedelta(days=actual_duration_days, hours=actual_duration_hours)
                    actual_cost = 0.0
                    if project_control_sheet.cell(row=curr_row, column=5).value:
                        actual_cost = float(project_control_sheet.cell(row=curr_row, column=5).value)
                    percentage_completed_str = project_control_sheet.cell(row=curr_row, column=6).value  # always given
                    # In the test file, some percentages are strings, some are ints, some are floats (#YOLO)
                    if type(percentage_completed_str) != str:
                        percentage_completed_str *= 100
                    if type(percentage_completed_str) == float:
                        percentage_completed_str = int(percentage_completed_str)
                    percentage_completed_str = str(percentage_completed_str)
                    if not percentage_completed_str[-1].isdigit():
                        percentage_completed_str = percentage_completed_str[:-1]
                    if float(percentage_completed_str) < 1:
                        percentage_completed_str = float(float(percentage_completed_str)*100)
                    percentage_completed = int(percentage_completed_str)
                    if percentage_completed > 100:
                        print("Gilles V has made a stupid error by putting some ifs & transformations in XLSXparser.py "
                              "with the variable percentage_completed_str")
                    act_track_record = ActivityTrackingRecord(activity=activities_dict[activity_id][1],
                                                              actual_start=actual_start, actual_duration=actual_duration,
                                                              actual_cost=actual_cost,
                                                              percentage_completed=percentage_completed)
                    act_track_records.append(act_track_record)
                    tracking_periods.append(TrackingPeriod(tracking_period_name=tp_name, tracking_period_statusdate=tp_date,
                                                           tracking_period_records=act_track_records))
        return tracking_periods

    def process_resources(self, resource_sheet):
        # We store the resources  in a dict, with as index the resource name, to access them easily later when we
        # are processing the activities.
        resources_dict = {}
        for curr_row in range(self.get_nr_of_header_lines(resource_sheet), resource_sheet.get_highest_row()+1):
            res_id = int(resource_sheet.cell(row=curr_row, column=1).value)
            res_name = resource_sheet.cell(row=curr_row, column=2).value
            res_type = resource_sheet.cell(row=curr_row, column=3).value
            # Had to cast string -> float -> int (silly Python!)
            res_ava = int(float(resource_sheet.cell(row=curr_row, column=4).value.split(" ")[0]
                          .translate(str.maketrans(",", "."))))
            res_cost_use = float(resource_sheet.cell(row=curr_row, column=5).value)
            res_cost_unit = float(resource_sheet.cell(row=curr_row, column=6).value)
            resources_dict[res_name] = Resource(resource_id=res_id, name=res_name, resource_type=res_type,
                                                availability=res_ava, cost_use=res_cost_use, cost_unit=res_cost_unit)
        return resources_dict

    def process_risk_analysis(self, risk_analysis_sheet, excel_version):
        risk_analysis_dict = {}
        if excel_version == ExcelVersion.EXTENDED:
            col = 4
        else:
            col = 3
        for curr_row in range(self.get_nr_of_header_lines(risk_analysis_sheet), risk_analysis_sheet.get_highest_row()+1):
            if risk_analysis_sheet.cell(row=curr_row, column=col).value is not None:
                risk_ana_dist_type = risk_analysis_sheet.cell(row=curr_row, column=col).value.split(" - ")[0]
                risk_ana_dist_units = risk_analysis_sheet.cell(row=curr_row, column=col).value.split(" - ")[1]
                risk_ana_opt_duration = int(risk_analysis_sheet.cell(row=curr_row, column=col+1).value)
                risk_ana_prob_duration = int(risk_analysis_sheet.cell(row=curr_row, column=col+2).value)
                risk_ana_pess_duration = int(risk_analysis_sheet.cell(row=curr_row, column=col+3).value)
                dict_id = int(risk_analysis_sheet.cell(row=curr_row, column=1).value)
                risk_analysis_dict[dict_id] = RiskAnalysisDistribution(distribution_type=risk_ana_dist_type,
                                                                       distribution_units=risk_ana_dist_units,
                                                                       optimistic_duration=risk_ana_opt_duration,
                                                                       probable_duration=risk_ana_prob_duration,
                                                                       pessimistic_duration=risk_ana_pess_duration)
        return risk_analysis_dict

    def process_baseline_schedule(self, activities_sheet, resources_dict, risk_analysis_dict, excel_version):
        activities_dict = {}
        if excel_version == ExcelVersion.EXTENDED:
            for curr_row in range(self.get_nr_of_header_lines(activities_sheet), activities_sheet.get_highest_row()+1):
                activity_id = int(activities_sheet.cell(row=curr_row, column=1).value)
                activity_name = activities_sheet.cell(row=curr_row, column=2).value
                activity_wbs = ()
                for number in activities_sheet.cell(row=curr_row, column=3).value.split("."):
                    activity_wbs = activity_wbs + (int(number),)
                activity_predecessors = self.process_predecessors(activities_sheet.cell(row=curr_row, column=4).value)
                activity_successors = self.process_successors(activities_sheet.cell(row=curr_row, column=5).value)
                activity_resource_cost = 0.0
                if activities_sheet.cell(row=curr_row, column=10).value:
                    activity_resource_cost = float(activities_sheet.cell(row=curr_row, column=10).value)
                if type(activities_sheet.cell(row=curr_row, column=6).value) is float:
                    baseline_start = datetime.datetime.utcfromtimestamp(((activities_sheet.cell(row=curr_row, column=6).value -
                                                                          25569)*86400))  # Convert to correct date
                else:
                    baseline_start = activities_sheet.cell(row=curr_row, column=6).value
                if type(activities_sheet.cell(row=curr_row, column=7).value) is float:
                    baseline_end = datetime.datetime.utcfromtimestamp(((activities_sheet.cell(row=curr_row, column=7).value -
                                                                        25569)*86400))  # Convert to correct date
                else:
                    baseline_end = activities_sheet.cell(row=curr_row, column=7).value
                baseline_duration_split = activities_sheet.cell(row=curr_row, column=8).value.split("d")
                baseline_duration_days = int(baseline_duration_split[0])
                baseline_duration_hours = 0  # We need to set this default value for the next loop
                if baseline_duration_split[1] != '':
                    baseline_duration_hours = int(baseline_duration_split[1][1:-1])  # first char = " "; last char = "h"
                baseline_duration = datetime.timedelta(days=baseline_duration_days, hours=baseline_duration_hours)
                baseline_fixed_cost = 0.0
                if activities_sheet.cell(row=curr_row, column=11).value:
                    baseline_fixed_cost = float(activities_sheet.cell(row=curr_row, column=11).value)
                baseline_hourly_cost = 0.0
                if activities_sheet.cell(row=curr_row, column=12).value:
                    baseline_hourly_cost = float(activities_sheet.cell(row=curr_row, column=12).value)
                baseline_var_cost = None
                if activities_sheet.cell(row=curr_row, column=13).value is not None:
                    baseline_var_cost = float(activities_sheet.cell(row=curr_row, column=13).value)
                baseline_total_cost = 0.0
                if activities_sheet.cell(row=curr_row, column=14).value:
                    baseline_total_cost = float(activities_sheet.cell(row=curr_row, column=14).value)
                activity_baseline_schedule = BaselineScheduleRecord(start=baseline_start, end=baseline_end,
                                                                    duration=baseline_duration,
                                                                    fixed_cost=baseline_fixed_cost,
                                                                    hourly_cost=baseline_hourly_cost,
                                                                    var_cost=baseline_var_cost,
                                                                    total_cost=baseline_total_cost)
                activity_risk_analysis = None
                if activity_id in risk_analysis_dict:
                    activity_risk_analysis = risk_analysis_dict[activity_id]
                activity_resources = []
                if activities_sheet.cell(row=curr_row, column=9).value:
                    for activity_resource in activities_sheet.cell(row=curr_row, column=9).value.split(';'):
                        activity_resource_name = resources_dict[activity_resource.split("[")[0]]
                        activity_resource_demand = 1
                        if len(activity_resource.split("[")) > 1:
                            activity_resource_demand = int(float(activity_resource.split("[")[1].split(" ")[0].translate(str.maketrans(",", "."))))
                        activity_resources.append((activity_resource_name, activity_resource_demand))
                activities_dict[activity_id] = (curr_row, Activity(activity_id, name=activity_name, wbs_id=activity_wbs,
                                                   predecessors=activity_predecessors, successors=activity_successors,
                                                   baseline_schedule=activity_baseline_schedule,
                                                   resource_cost=activity_resource_cost,
                                                   risk_analysis=activity_risk_analysis, resources=activity_resources))
        else:
            for curr_row in range(self.get_nr_of_header_lines(activities_sheet), activities_sheet.get_highest_row()+1):
                activity_id = int(activities_sheet.cell(row=curr_row, column=1).value)
                activity_name = activities_sheet.cell(row=curr_row, column=2).value
                activity_predecessors = self.process_predecessors(activities_sheet.cell(row=curr_row, column=3).value)
                activity_successors = self.process_successors(activities_sheet.cell(row=curr_row, column=4).value)
                if type(activities_sheet.cell(row=curr_row, column=5).value) is float:
                    baseline_start = datetime.datetime.utcfromtimestamp(((activities_sheet.cell(row=curr_row, column=5).value -
                                                                          25569)*86400))  # Convert to correct date
                else:
                    baseline_start = activities_sheet.cell(row=curr_row, column=5).value
                if type(activities_sheet.cell(row=curr_row, column=6).value) is float:
                    baseline_end = datetime.datetime.utcfromtimestamp(((activities_sheet.cell(row=curr_row, column=6).value -
                                                                        25569)*86400))  # Convert to correct date
                else:
                    baseline_end = activities_sheet.cell(row=curr_row, column=6).value
                baseline_duration_split = activities_sheet.cell(row=curr_row, column=7).value.split("d")
                baseline_duration_days = int(baseline_duration_split[0])
                baseline_duration_hours = 0  # We need to set this default value for the next loop
                if baseline_duration_split[1] != '':
                    baseline_duration_hours = int(baseline_duration_split[1][1:-1])  # first char = " "; last char = "h"
                baseline_duration = datetime.timedelta(days=baseline_duration_days, hours=baseline_duration_hours)
                baseline_fixed_cost = 0.0
                if activities_sheet.cell(row=curr_row, column=9).value:
                    baseline_fixed_cost = float(activities_sheet.cell(row=curr_row, column=9).value)
                baseline_hourly_cost = 0.0
                if activities_sheet.cell(row=curr_row, column=10).value:
                    baseline_hourly_cost = float(activities_sheet.cell(row=curr_row, column=10).value)
                baseline_var_cost = None  # This is a hack to determine if an act. in basic version is from lowest lvl
                if activities_sheet.cell(row=curr_row, column=11).value is not None:
                    baseline_var_cost = float(activities_sheet.cell(row=curr_row, column=11).value)
                activity_baseline_schedule = BaselineScheduleRecord(start=baseline_start, end=baseline_end,
                                                                    duration=baseline_duration,
                                                                    fixed_cost=baseline_fixed_cost,
                                                                    hourly_cost=baseline_hourly_cost,
                                                                    var_cost=baseline_var_cost)
                activity_risk_analysis = None
                if activity_id in risk_analysis_dict:
                    activity_risk_analysis = risk_analysis_dict[activity_id]
                activity_resources = []
                if activities_sheet.cell(row=curr_row, column=8).value:
                    for activity_resource in activities_sheet.cell(row=curr_row, column=8).value.split(';'):
                        activity_resource_name = resources_dict[activity_resource.split("[")[0]]
                        activity_resource_demand = 1
                        if len(activity_resource.split("[")) > 1:
                            activity_resource_demand = int(float(activity_resource.split("[")[1].split(" ")[0].translate(str.maketrans(",", "."))))
                        activity_resources.append((activity_resource_name, activity_resource_demand))
                activities_dict[activity_id] = (curr_row, Activity(activity_id, name=activity_name,
                                                   predecessors=activity_predecessors, successors=activity_successors,
                                                   baseline_schedule=activity_baseline_schedule,
                                                   risk_analysis=activity_risk_analysis, resources=activity_resources))

        return activities_dict

    @staticmethod
    def process_predecessors(predecessors):
        if predecessors:
            activity_predecessors = []
            for predecessor in predecessors.split(";"):
                temp = re.split('\-|\+', predecessor)
                # The last two characters are the relation type (e.g. xFS or xxFS), activity can be variable
                # number of digits
                predecessor_activity = int(temp[0][0:-2])
                predecessor_relation = temp[0][-2:]
                if len(temp) == 2:
                    # Was it a + or -?
                    minus_plus = predecessor.split("-")
                    if minus_plus[1]:
                        # It was a -
                        predecessor_lag = -minus_plus[1]
                    else:
                        # It was a +
                        predecessor_lag = minus_plus[1]
                else:
                    predecessor_lag = 0
                activity_predecessors.append((predecessor_activity, predecessor_relation, predecessor_lag))
            return activity_predecessors
        return []

    @staticmethod
    def process_successors(successors):
        if successors:
            activity_successors = []
            for successor in successors.split(";"):
                temp = re.split('\-|\+', successor)
                # The first two characters are the relation type (e.g. xFS or xxFS), activity can be variable
                # number of digits
                successor_activity = int(temp[0][2:])
                successor_relation = temp[0][0:2]
                if len(temp) == 2:
                    # Was it a + or -?
                    minus_plus = successor.split("-")
                    if minus_plus[1]:
                        # It was a -
                        successor_lag = -minus_plus[1]
                    else:
                        # It was a +
                        successor_lag = minus_plus[1]
                else:
                    successor_lag = 0
                activity_successors.append((successor_activity, successor_relation, successor_lag))
            return activity_successors
        return []

    @staticmethod
    def get_nr_of_header_lines(sheet):
        header_lines = 1
        while not(sheet.cell(row=header_lines, column=1).value is not None
                  and (type(sheet.cell(row=header_lines, column=1).value) is int
                       or sheet.cell(row=header_lines, column=1).value.isdigit())):
            header_lines += 1
        return header_lines

    def calculate_aggregated_ac_per_wp(self, tracking_period):
        for item in tracking_period.tracking_period_records:
            if len(item.activity.wbs_id) == 2: #workpackage
                parent = item.activity
                sum_ac = 0
                for atr in tracking_period.tracking_period_records:
                    if len(atr.activity.wbs_id) == 3 and (atr.activity.wbs_id[0], atr.activity.wbs_id[1]) == parent.wbs_id:
                        sum_ac += atr.actual_cost
                item.actual_cost = round(sum_ac,2)

    def calculate_aggregated_ad_per_wp(self, tracking_period, agenda):
        for item in tracking_period.tracking_period_records:
            if len(item.activity.wbs_id) == 2: #workpackage
                parent = item.activity
                latest_end = None
                earliest_start = None
                completed = True
                for atr in tracking_period.tracking_period_records:
                    if len(atr.activity.wbs_id) == 3 and (atr.activity.wbs_id[0], atr.activity.wbs_id[1]) == parent.wbs_id:
                        if atr.actual_start:
                            if not earliest_start or earliest_start > atr.actual_start:
                                earliest_start = atr.actual_start
                            if atr.percentage_completed < 100:
                                completed = False
                                latest_end = tracking_period.tracking_period_statusdate
                            if completed:
                                actual_end = agenda.get_end_date(atr.actual_start, atr.actual_duration.days, atr.actual_duration.seconds/3600)
                                if not latest_end or latest_end < actual_end:
                                    latest_end = actual_end
                if earliest_start and latest_end:
                    ad = agenda.get_time_between(earliest_start, latest_end)
                    item.actual_duration = ad

    def calculate_aggregated_ac(self, tracking_period):
        sum_ac = 0
        for atr in tracking_period.tracking_period_records:
            if not self.is_not_lowest_level_activity(atr.activity):
                sum_ac += atr.actual_cost
        return sum_ac

    def calculate_aggregated_pv(self, tracking_period):
        sum_pv = 0
        for atr in tracking_period.tracking_period_records:
            if not self.is_not_lowest_level_activity(atr.activity):
                sum_pv += atr.planned_value
        return sum_pv

    def calculate_aggregated_ev(self, tracking_period):
        sum_ev = 0
        for atr in tracking_period.tracking_period_records:
            if not self.is_not_lowest_level_activity(atr.activity):
                sum_ev += atr.earned_value
        return sum_ev

    def calculate_aggregated_rc(self, tracking_period):
        sum_rc = 0
        for atr in tracking_period.tracking_period_records:
            if not self.is_not_lowest_level_activity(atr.activity):
                sum_rc += atr.remaining_cost
        return sum_rc

    def get_bac(self, tracking_period):
        pv = None
        for atr in tracking_period.tracking_period_records:
            if len(atr.activity.wbs_id) == 1:
                pv = atr.planned_value
                break
        if pv:
            return pv + self.calculate_aggregated_rc(tracking_period)
        else:
            return 0

    def calculate_eac(self, ac, bac, ev, pf):
        if pf == 0:
            return 0
        return ac + (bac - ev)/pf

    def get_pv(self, tracking_period):
        for atr in tracking_period.tracking_period_records:
            if len(atr.activity.wbs_id) == 1:
                return atr.planned_value
        return None

    #def get_pvs(self, tracking_periods):
    #    pvs = []
    #    for tracking_period in tracking_periods:
    #        pvs.append(self.get_pv(tracking_period))
    #    return pvs

    #def calculate_es(self, tracking_period, tracking_periods):
    #    #TODO: reimplement
    #    ev = self.calculate_aggregated_ev(tracking_period)
    #    pvs = self.get_pvs(tracking_periods)
    #    for i in range(0, len(pvs)-1):
    #        if pvs[i] <= ev < pvs[i+1]:
    #            x = pvs[i+1] - ev

    def calculate_es(self, project_object, PVcurve, current_EV, currentTime):
        """
        This function calculates the ES datetime based on given PVcurve, current EV value and current time

        :param project_object: ProjectObject, the project object corresponding to the PVcurve
        :param PVcurve: list of tuples, (PV cumsum value, datetime of this PV cumsum value) as calculated by calculate_PVcurve
        :param current_EV: float, Earned value for which to search the ES
        :param currentTime: datetime, statusdate
        """
        # algorithm:
        # Find t such that EV ? PV(t) and EV < PV(t+1)
        # ES = t
        ## according to PMKnowledgecenter the following line should be used, but an ES in the middle of non-working datetimes is not desirable!
        ## ES = t + (EV - PV(t)) / (PV(t+1) - PV(t)) * (next_t - t)

        t = min([activity.baseline_schedule.start for activity in project_object.activities])  # projectBaselineStartDate
        lowerPV = -1
        pointFound = False

        # search first PV which is larger than EV
        for i in range(1, len(PVcurve)):
            if PVcurve[i][0] > current_EV:
                # first PV point found which is larger than the given EV value
                t = PVcurve[i-1][1]
                lowerPV = PVcurve[i-1][0]
                pointFound = True
                break
        #endFor searching larger PV point

        if not pointFound:
            t = max([activity.baseline_schedule.end for activity in project_object.activities])  # projectBaselineEndDate

        # NOTE: wanted behaviour?
        # Correct ES to statusDate if PVcumsum has the same value there:
        timesWithSamePV = [x[1] for x in PVcurve if abs(x[0] - lowerPV) < 1e-5] 
        # currentTime can be the start of a workingInterval or the end of a working interval
        if currentTime in timesWithSamePV or project_object.agenda.get_previous_working_hour(currentTime) in timesWithSamePV:
            return project_object.agenda.get_previous_working_hour(currentTime)

        return t

    def calculate_PVcurve(self, project_object):
        "This function generates the PV curve of the baseline schedule of the given project."

        # retrieve only lowest level activities and sort them on their baseline start date and end date
        lowestLevelActivities = [activity for activity in project_object.activities if not XLSXParser.is_not_lowest_level_activity(activity)]
        lowestLevelActivitiesSorted = sorted(lowestLevelActivities, key=attrgetter('baseline_schedule.start', 'baseline_schedule.end'))

        projectBaselineStartDate = min([activity.baseline_schedule.start for activity in project_object.activities])
        projectBaselineEndDate = max([activity.baseline_schedule.end for activity in project_object.activities])

        # initial starting point
        generated_PVcurve = [(0, projectBaselineStartDate)]

        # search end of first working hour of this project
        currentDatetime = project_object.agenda.get_next_date(projectBaselineStartDate, 0, 0) + datetime.timedelta(hours = 1)

        # loop variables
        index_NextActivityToAdd = 0
        currentPVcumsumValue = 0
        runningActivities = []
        consumableResourceIdsAlreadyUsed = []  # to only inquire once the cost/use of a consumable resource

        # DEBUG:
        traversedHours = 0
        
        # traverse the complete project baseline duration
        while currentDatetime <= projectBaselineEndDate:
            # update list of running activities at this time:
            # remove ended activities: activities that finished before the start of the evaluating interval hour that ended on currentDatetime
            for activity in runningActivities.copy():
                if activity.baseline_schedule.end < currentDatetime:
                    # activity ended maximally on the start of the currently evaluating interval hour of currentDatetime
                    runningActivities.remove(activity)

            # add activities that are just started:
            while (index_NextActivityToAdd < len(lowestLevelActivitiesSorted)) and lowestLevelActivitiesSorted[index_NextActivityToAdd].baseline_schedule.start < currentDatetime:
                if len(runningActivities) > 0:
                    runningActivities.append(lowestLevelActivitiesSorted[index_NextActivityToAdd])
                else:
                    runningActivities = [lowestLevelActivitiesSorted[index_NextActivityToAdd]]

                # add starting fixed cost of the newly added activity to the PVcumsumValue:
                currentPVcumsumValue += lowestLevelActivitiesSorted[index_NextActivityToAdd].baseline_schedule.fixed_cost
                # add starting use cost of used resources by this activity:
                for resourceTuple in lowestLevelActivitiesSorted[index_NextActivityToAdd].resources:
                    resource = resourceTuple[0]
                    if resource.resource_type == ResourceType.CONSUMABLE:
                        # only add once the cost for its use!
                        if resource.resource_id not in consumableResourceIdsAlreadyUsed:
                            # add cost for its use:
                            currentPVcumsumValue += resource.cost_use
                            if len(consumableResourceIdsAlreadyUsed) > 0:
                                consumableResourceIdsAlreadyUsed.append(resource.resource_id)
                            else:
                                consumableResourceIdsAlreadyUsed = [resource.resource_id]
                    else:
                        # resource type is renewable:
                        # add cost_use per demanded resource unit
                        currentPVcumsumValue += resourceTuple[1] * resource.cost_use
                #endFor adding resource use costs
                index_NextActivityToAdd += 1
            #endWhile adding new activities that started

            # update PVcurve with one timestemp its variable added value:
            for activity in runningActivities:
                currentPVcumsumValue += activity.baseline_schedule.hourly_cost
                # add also variable cost of the use of its resources
                for resourceTuple in activity.resources:
                    currentPVcumsumValue += resourceTuple[0].cost_unit * resourceTuple[1]  # cost/(hour * unit) * demanded units

            # save the currentPVcumsumValue at its new timestep:
            generated_PVcurve.append((currentPVcumsumValue, currentDatetime))

            # advance to next working hour its end
            currentDatetime = project_object.agenda.get_next_date(currentDatetime, 0, 0) + datetime.timedelta(hours = 1)
            # DEBUG: record how many hours processed, should match with project baseline duration
            traversedHours += 1
            
        return generated_PVcurve

    def calculate_SVt(self, project_object, ES, currentTime):
        """This function calculates the SV(t) in hours, based on the given ES and currentTime.
        returns: (SV(t) in workinghours, string representation of SV(t))
        """
        # determine the time between ES and currentTime:
        currentTimeValidDatetime = project_object.agenda.get_next_date(currentTime, 0,0)
        ESvalidDate = project_object.agenda.get_next_date(ES, 0,0)

        if ES >= currentTime:
            timeBetween = project_object.agenda.get_time_between(currentTimeValidDatetime, ESvalidDate)
            return (timeBetween.days * project_object.agenda.get_working_hours_in_a_day() + int(timeBetween.seconds / 3600), XLSXParser.get_duration_str(timeBetween, False))
        else:
            timeBetween = project_object.agenda.get_time_between(ESvalidDate, currentTimeValidDatetime)
            return (- timeBetween.days * project_object.agenda.get_working_hours_in_a_day() - int(timeBetween.seconds / 3600), XLSXParser.get_duration_str(timeBetween, True))

    def calculate_SPIt(self, project_object, ES, currentTime):
        "This fuction calculates the SPI(t) value = ES / AT"
        ESvalidDate = project_object.agenda.get_next_date(ES, 0,0)
        currentTimeValidDatetime = project_object.agenda.get_next_date(currentTime, 0,0)

        projectBaselineStartDate = min([activity.baseline_schedule.start for activity in project_object.activities])
        durationWorkingTimeES = project_object.agenda.get_time_between(projectBaselineStartDate , ESvalidDate)
        durationWorkingTimeAT = project_object.agenda.get_time_between(projectBaselineStartDate , currentTimeValidDatetime)

        if durationWorkingTimeAT.total_seconds() <= 0:
            return 0

        return (durationWorkingTimeES.days * project_object.agenda.get_working_hours_in_a_day() + (durationWorkingTimeES.seconds / 3600)) / (durationWorkingTimeAT.days * project_object.agenda.get_working_hours_in_a_day() + (durationWorkingTimeAT.seconds / 3600.0))

    def calculate_p_factor(self, project_object, tracking_period, ES):
        "This function calculates the p-factor for the given tracking_period and ES datetime"
        # using algorithm:
        # p-factor = sum(i?N)[ min(PV(i,ES), EV(i,AT))] / sum(i?N)[ PV(i,ES)]
        # with 
        # N: The set of all activities of the project network
        # PV(i,ES): The planned value of activity i at time ES 
        # EV(i,AT): The earned value of activity i at time AT

        # DEBUG:
        print("XLSXParser:calculate_p_factor: given tracking_period_records length = {0}".format(len(tracking_period.tracking_period_records)))

        # sort the tracking_period_records according to its activity baseline schedule start and end date => can figure out which activity uses first a consumable resource

        # retrieve only lowest level activities and sort them on their baseline start date and end date
        lowestLevelActivitiesTrackingRecords = [atr for atr in tracking_period.tracking_period_records if not XLSXParser.is_not_lowest_level_activity(atr.activity)]
        lowestLevelActivitiesTrackingRecordsSorted = sorted(lowestLevelActivitiesTrackingRecords, key=attrgetter('activity.baseline_schedule.start', 'activity.baseline_schedule.end'))

        numerator = 0
        denominator = 0
        consumableResourceIdsAlreadyUsed = []  # to only inquire once the cost/use of a consumable resource
        ESdatetime = project_object.agenda.get_next_date(ES, 0,0)

        for atr in lowestLevelActivitiesTrackingRecordsSorted:
            # determine The planned value of activity i at time ES:
            PV_i_ES = 0
            if ESdatetime <= atr.activity.baseline_schedule.start:
                # activity is not yet started according to plan => no planned value
                pass
            elif ESdatetime >= atr.activity.baseline_schedule.end:
                # activity is already finished according to plan => PVi = total_cost
                PV_i_ES = atr.activity.baseline_schedule.total_cost
                # check if to add resosurce type to consumableResourceIdsAlreadyUsed:
                for resourceTuple in atr.activity.resources:
                    resource = resourceTuple[0]
                    if resource.resource_type == ResourceType.CONSUMABLE:
                        # only add once the cost for its use!
                        if resource.resource_id not in consumableResourceIdsAlreadyUsed:
                            if len(consumableResourceIdsAlreadyUsed) > 0:
                                consumableResourceIdsAlreadyUsed.append(resource.resource_id)
                            else:
                                consumableResourceIdsAlreadyUsed = [resource.resource_id]
                #endFor adding resource use costs
            else:
                # activity is still running => calculate intermediate PV value:
                PV_i_ES = atr.activity.baseline_schedule.fixed_cost

                # add starting use cost of used resources by this activity:
                for resourceTuple in atr.activity.resources:
                    resource = resourceTuple[0]
                    if resource.resource_type == ResourceType.CONSUMABLE:
                        # only add once the cost for its use!
                        if resource.resource_id not in consumableResourceIdsAlreadyUsed:
                            # add cost for its use:
                            PV_i_ES += resource.cost_use
                            if len(consumableResourceIdsAlreadyUsed) > 0:
                                consumableResourceIdsAlreadyUsed.append(resource.resource_id)
                            else:
                                consumableResourceIdsAlreadyUsed = [resource.resource_id]
                    else:
                        # resource type is renewable:
                        # add cost_use per demanded resource unit
                        PV_i_ES += resourceTuple[1] * resource.cost_use
                #endFor adding resource use costs

                # add variable costs of this activity according to duration that this activity is running:
                # determine running duration:
                actStartDatetime = project_object.agenda.get_next_date(atr.activity.baseline_schedule.start, 0,0)
                
                runningTimedelta = project_object.agenda.get_time_between(actStartDatetime, ESdatetime)
                runningWorkingHours = runningTimedelta.days * project_object.agenda.get_working_hours_in_a_day() + int(runningTimedelta.seconds / 3600)

                # activity variable cost:
                PV_i_ES += atr.activity.baseline_schedule.hourly_cost * runningWorkingHours
                # activity its resources variable costs:
                for resourceTuple in atr.activity.resources:
                    PV_i_ES += resourceTuple[0].cost_unit * resourceTuple[1] * runningWorkingHours # cost/(hour * unit) * demanded units * running hours

            #endIF calculating PV_i_ES
            numerator += min(PV_i_ES, atr.earned_value)
            denominator += PV_i_ES
        #endFor all activity tracking records

        if denominator < 1e-10:
            return 0.0
        else:
            return numerator / denominator



    def from_schedule_object(self, project_object, file_path_output, excel_version):        
        """
        This is just a lot of writing to excel code, it is ugly..

        """
        workbook = xlsxwriter.Workbook(file_path_output)

        # Lots of formats
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        yellow_cell = workbook.add_format({'bg_color': 'yellow', 'text_wrap': True, 'border': 1, 'font_size': 8})
        cyan_cell = workbook.add_format({'bg_color': '#D9EAF7', 'text_wrap': True, 'border': 1, 'font_size': 8})
        green_cell = workbook.add_format({'bg_color': '#9BBB59', 'text_wrap': True, 'border': 1, 'font_size': 8})
        red_cell = workbook.add_format({'bg_color': 'red', 'text_wrap': True, 'border': 1, 'font_size': 8})
        gray_cell = workbook.add_format({'bg_color': '#D4D0C8', 'text_wrap': True, 'border': 1, 'font_size': 8})
        date_cyan_cell = workbook.add_format({'bg_color': '#D9EAF7', 'text_wrap': True, 'border': 1,
                                              'num_format': 'dd/mm/yyyy H:MM', 'font_size': 8})
        date_green_cell = workbook.add_format({'bg_color': '#C4D79B', 'text_wrap': True, 'border': 1,
                                              'num_format': 'dd/mm/yyyy H:MM', 'font_size': 8})
        date_lime_cell = workbook.add_format({'bg_color': '#9BBB59', 'text_wrap': True, 'border': 1,
                                              'num_format': 'dd/mm/yyyy H:MM', 'font_size': 8})
        date_gray_cell = workbook.add_format({'bg_color': '#D4D0C8', 'text_wrap': True, 'border': 1,
                                              'num_format': 'dd/mm/yyyy H:MM', 'font_size': 8})
        money_cyan_cell = workbook.add_format({'bg_color': '#D9EAF7', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00' + u"\u20AC", 'font_size': 8})
        money_green_cell = workbook.add_format({'bg_color': '#C4D79B', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00' + u"\u20AC", 'font_size': 8})
        money_lime_cell = workbook.add_format({'bg_color': '#9BBB59', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00' + u"\u20AC", 'font_size': 8})
        money_navy_cell = workbook.add_format({'bg_color': '#D4D0C8', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00' + u"\u20AC", 'font_size': 8})
        money_yellow_cell = workbook.add_format({'bg_color': 'yellow', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00' + u"\u20AC", 'font_size': 8})
        money_gray_cell = workbook.add_format({'bg_color': '#D4D0C8', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00' + u"\u20AC", 'font_size': 8})

        # Worksheets
        bsch_worksheet = workbook.add_worksheet("Baseline Schedule")
        res_worksheet = workbook.add_worksheet("Resources")
        ra_worksheet = workbook.add_worksheet("Risk Analysis")

        # Write the Baseline Schedule Worksheet

        # Set the width of the columns
        if excel_version == ExcelVersion.EXTENDED:
            bsch_worksheet.set_column(0, 0, 3)
            bsch_worksheet.set_column(1, 1, 25)
            bsch_worksheet.set_column(2, 2, 5)
            bsch_worksheet.set_column(3, 4, 16)
            bsch_worksheet.set_column(5, 6, 12)
            bsch_worksheet.set_column(7, 7, 8)
            bsch_worksheet.set_column(8, 8, 25)
            bsch_worksheet.set_column(9, 9, 10)
            bsch_worksheet.set_column(10, 11, 15)
            bsch_worksheet.set_column(13, 13, 12)
        else:
            bsch_worksheet.set_column(0, 0, 3)
            bsch_worksheet.set_column(1, 1, 25)
            bsch_worksheet.set_column(2, 3, 16)
            bsch_worksheet.set_column(6, 6, 8)
            bsch_worksheet.set_column(7, 7, 25)

        # Set the height of rows
        bsch_worksheet.set_row(1, 30)

        # Write header cells (using the header format, and by merging some cells)
        if excel_version == ExcelVersion.EXTENDED:
            bsch_worksheet.merge_range('A1:C1', "General", header)
            bsch_worksheet.merge_range('D1:E1', "Relations", header)
            bsch_worksheet.merge_range('F1:H1', "Baseline", header)
            bsch_worksheet.merge_range('I1:J1', "Resource Demand", header)
            bsch_worksheet.merge_range('K1:N1', "Baseline Costs", header)

            bsch_worksheet.write('A2', "ID", header)
            bsch_worksheet.write('B2', "Name", header)
            bsch_worksheet.write('C2', "WBS", header)
            bsch_worksheet.write('D2', "Predecessors", header)
            bsch_worksheet.write('E2', "Successors", header)
            bsch_worksheet.write('F2', "Baseline Start", header)
            bsch_worksheet.write('G2', "Baseline End", header)
            bsch_worksheet.write('H2', "Duration", header)
            bsch_worksheet.write('I2', "Resource Demand", header)
            bsch_worksheet.write('J2', "Resource Cost", header)
            bsch_worksheet.write('K2', "Fixed Cost", header)
            bsch_worksheet.write('L2', "Cost/Hour", header)
            bsch_worksheet.write('M2', "Variable Cost", header)
            bsch_worksheet.write('N2', "Total Cost", header)
        else:
            bsch_worksheet.merge_range('A1:B1', "General", header)
            bsch_worksheet.merge_range('C1:D1', "Relations", header)
            bsch_worksheet.merge_range('E1:G1', "Baseline", header)
            bsch_worksheet.write('H1', "Resource Demand", header)
            bsch_worksheet.merge_range('I1:K1', "Baseline Costs", header)

            bsch_worksheet.write('A2', "ID", header)
            bsch_worksheet.write('B2', "Name", header)
            bsch_worksheet.write('C2', "Predecessors", header)
            bsch_worksheet.write('D2', "Successors", header)
            bsch_worksheet.write('E2', "Baseline Start", header)
            bsch_worksheet.write('F2', "Baseline End", header)
            bsch_worksheet.write('G2', "Duration", header)
            bsch_worksheet.write('H2', "Resource Demand", header)
            bsch_worksheet.write('I2', "Fixed Cost", header)
            bsch_worksheet.write('J2', "Cost/Hour", header)
            bsch_worksheet.write('K2', "Variable Cost", header)

        # Now we run through all activities to get the required information
        counter = 2
        for activity in project_object.activities:
            if not self.is_not_lowest_level_activity(activity):
                # Write activity of lowest level
                if excel_version == ExcelVersion.EXTENDED:
                    bsch_worksheet.write_number(counter, 0, activity.activity_id, green_cell)
                    bsch_worksheet.write(counter, 1, str(activity.name), green_cell)
                    self.write_wbs(bsch_worksheet, counter, 2, activity.wbs_id, gray_cell) ####
                    self.write_predecessors(bsch_worksheet, counter, 3, activity.predecessors, green_cell)
                    self.write_successors(bsch_worksheet, counter, 4, activity.successors, green_cell)
                    bsch_worksheet.write_datetime(counter, 5, activity.baseline_schedule.start, date_lime_cell)
                    bsch_worksheet.write_datetime(counter, 6, activity.baseline_schedule.end, date_green_cell)
                    bsch_worksheet.write(counter, 7, self.get_duration_str(activity.baseline_schedule.duration), green_cell)
                    self.write_resources(bsch_worksheet, counter, 8, activity.resources, yellow_cell)
                    bsch_worksheet.write_number(counter, 9, activity.resource_cost, money_navy_cell) ####
                    bsch_worksheet.write_number(counter, 10, activity.baseline_schedule.fixed_cost, money_green_cell)
                    bsch_worksheet.write_number(counter, 11, activity.baseline_schedule.hourly_cost, money_lime_cell)
                    if activity.baseline_schedule.var_cost is not None:
                        bsch_worksheet.write_number(counter, 12, activity.baseline_schedule.var_cost, money_green_cell)
                    else:
                        bsch_worksheet.write_number(counter, 12, 0, money_green_cell)
                    bsch_worksheet.write_number(counter, 13, activity.baseline_schedule.total_cost, money_navy_cell) ####
                else:
                    bsch_worksheet.write_number(counter, 0, activity.activity_id, green_cell)
                    bsch_worksheet.write(counter, 1, str(activity.name), green_cell)
                    self.write_predecessors(bsch_worksheet, counter, 2, activity.predecessors, green_cell)
                    self.write_successors(bsch_worksheet, counter, 3, activity.successors, green_cell)
                    bsch_worksheet.write_datetime(counter, 4, activity.baseline_schedule.start, date_lime_cell)
                    bsch_worksheet.write_datetime(counter, 5, activity.baseline_schedule.end, date_green_cell)
                    bsch_worksheet.write(counter, 6, self.get_duration_str(activity.baseline_schedule.duration), green_cell)
                    self.write_resources(bsch_worksheet, counter, 7, activity.resources, yellow_cell)
                    bsch_worksheet.write_number(counter, 8, activity.baseline_schedule.fixed_cost, money_green_cell)
                    bsch_worksheet.write_number(counter, 9, activity.baseline_schedule.hourly_cost, money_lime_cell)
                    if activity.baseline_schedule.var_cost is not None:
                        bsch_worksheet.write_number(counter, 10, activity.baseline_schedule.var_cost, money_green_cell)
                    else:
                        bsch_worksheet.write_number(counter, 10, 0, money_green_cell)
            else:
                # Write aggregated activity
                if excel_version == ExcelVersion.EXTENDED:
                    bsch_worksheet.write_number(counter, 0, activity.activity_id, yellow_cell)
                    bsch_worksheet.write(counter, 1, str(activity.name), yellow_cell)
                    self.write_wbs(bsch_worksheet, counter, 2, activity.wbs_id, cyan_cell)
                    self.write_predecessors(bsch_worksheet, counter, 3, activity.predecessors, cyan_cell)
                    self.write_successors(bsch_worksheet, counter, 4, activity.successors, cyan_cell)
                    bsch_worksheet.write_datetime(counter, 5, activity.baseline_schedule.start, date_cyan_cell)
                    bsch_worksheet.write_datetime(counter, 6, activity.baseline_schedule.end, date_cyan_cell)
                    bsch_worksheet.write(counter, 7, self.get_duration_str(activity.baseline_schedule.duration), cyan_cell)
                    self.write_resources(bsch_worksheet, counter, 8, activity.resources, cyan_cell)
                    bsch_worksheet.write(counter, 9, "", money_cyan_cell)
                    bsch_worksheet.write_number(counter, 10, activity.baseline_schedule.fixed_cost, money_cyan_cell)
                    bsch_worksheet.write(counter, 11, "", money_cyan_cell)
                    bsch_worksheet.write(counter, 12, "", money_cyan_cell)
                    bsch_worksheet.write_number(counter, 13, activity.baseline_schedule.total_cost, money_cyan_cell)
                else:
                    bsch_worksheet.write_number(counter, 0, activity.activity_id, yellow_cell)
                    bsch_worksheet.write(counter, 1, str(activity.name), yellow_cell)
                    self.write_predecessors(bsch_worksheet, counter, 2, activity.predecessors, cyan_cell)
                    self.write_successors(bsch_worksheet, counter, 3, activity.successors, cyan_cell)
                    bsch_worksheet.write_datetime(counter, 4, activity.baseline_schedule.start, date_cyan_cell)
                    bsch_worksheet.write_datetime(counter, 5, activity.baseline_schedule.end, date_cyan_cell)
                    bsch_worksheet.write(counter, 6, self.get_duration_str(activity.baseline_schedule.duration), cyan_cell)
                    self.write_resources(bsch_worksheet, counter, 7, activity.resources, cyan_cell)
                    bsch_worksheet.write_number(counter, 8, activity.baseline_schedule.fixed_cost, money_cyan_cell)
                    bsch_worksheet.write(counter, 9, "", money_cyan_cell)
                    bsch_worksheet.write(counter, 10, "", money_cyan_cell)

            counter += 1

        # Write the resources sheet

        # Some small adjustments to rows and columns in the resource sheet
        res_worksheet.set_row(1, 25)
        res_worksheet.set_column(1, 1, 15)
        res_worksheet.set_column(6, 6, 40)

        # Write header cells (using the header format, and by merging some cells)
        res_worksheet.merge_range('A1:D1', "General", header)
        res_worksheet.merge_range('E1:F1', "Resource Cost", header)
        if excel_version == ExcelVersion.EXTENDED:
            res_worksheet.merge_range('G1:H1', "Resource Demand", header)

        res_worksheet.write('A2', "ID", header)
        res_worksheet.write('B2', "Name", header)
        res_worksheet.write('C2', "Type", header)
        res_worksheet.write('D2', "Availability", header)
        res_worksheet.write('E2', "Cost/Use", header)
        res_worksheet.write('F2', "Cost/Unit", header)
        if excel_version == ExcelVersion.EXTENDED:
            res_worksheet.write('G2', "Assigned To", header)
            res_worksheet.write('H2', "Total Cost", header)

        counter = 2
        for resource in project_object.resources:
            res_worksheet.write_number(counter, 0, resource.resource_id, yellow_cell)
            res_worksheet.write(counter, 1, resource.name, yellow_cell)
            res_worksheet.write(counter, 2, resource.resource_type.value, yellow_cell)
            # God knows why we write the availability twice, it was like that in the template
            useless_availability_string = str(resource.availability) + " #" + str(resource.availability)
            res_worksheet.write(counter, 3, useless_availability_string, yellow_cell)
            res_worksheet.write(counter, 4, resource.cost_use, money_yellow_cell)
            res_worksheet.write(counter, 5, resource.cost_unit, money_yellow_cell)
            if excel_version == ExcelVersion.EXTENDED:
                self.write_resource_assign_cost(res_worksheet, counter, 6, resource, project_object.activities, cyan_cell,
                                            money_cyan_cell, project_object)
            counter += 1

        # Write the risk analysis sheet

        # Adjust some column widths
        ra_worksheet.set_column(0, 0, 3)
        ra_worksheet.set_column(1, 1, 18)
        ra_worksheet.set_column(3, 3, 15)
        ra_worksheet.set_column(4, 6, 12)

        # Write the headers
        if excel_version == ExcelVersion.EXTENDED:
            ra_worksheet.merge_range('A1:B1', "General", header)
            ra_worksheet.write('C1', "Baseline", header)
            ra_worksheet.merge_range('D1:G1', "Activity Duration Distribution Profiles", header)

            ra_worksheet.write('A2', "ID", header)
            ra_worksheet.write('B2', "Name", header)
            ra_worksheet.write('C2', "Duration", header)
            ra_worksheet.write('D2', "Description", header)
            ra_worksheet.write('E2', "Optimistic", header)
            ra_worksheet.write('F2', "Most Probable", header)
            ra_worksheet.write('G2', "Pessimistic", header)
        else:
            ra_worksheet.merge_range('A1:B1', "General", header)
            ra_worksheet.merge_range('C1:F1', "Activity Duration Distribution Profiles", header)

            ra_worksheet.write('A2', "ID", header)
            ra_worksheet.write('B2', "Name", header)
            ra_worksheet.write('C2', "Description", header)
            ra_worksheet.write('D2', "Optimistic", header)
            ra_worksheet.write('E2', "Most Probable", header)
            ra_worksheet.write('F2', "Pessimistic", header)


        # Write the rows by iterating through the activities (since they are linked to it)
        counter = 2
        for activity in project_object.activities:
            if self.is_not_lowest_level_activity(activity):
                if excel_version == ExcelVersion.EXTENDED:
                    ra_worksheet.write_number(counter, 0, activity.activity_id, cyan_cell)
                    ra_worksheet.write(counter, 1, str(activity.name), cyan_cell)
                    ra_worksheet.write(counter, 2, self.get_duration_str(activity.baseline_schedule.duration), cyan_cell)
                    ra_worksheet.write(counter, 3, "", cyan_cell)
                    ra_worksheet.write(counter, 4, "", cyan_cell)
                    ra_worksheet.write(counter, 5, "", cyan_cell)
                    ra_worksheet.write(counter, 6, "", cyan_cell)
                else:
                    ra_worksheet.write_number(counter, 0, activity.activity_id, cyan_cell)
                    ra_worksheet.write(counter, 1, str(activity.name), cyan_cell)
                    ra_worksheet.write(counter, 2, "", cyan_cell)
                    ra_worksheet.write(counter, 3, "", cyan_cell)
                    ra_worksheet.write(counter, 4, "", cyan_cell)
                    ra_worksheet.write(counter, 5, "", cyan_cell)
            else:
                if excel_version == ExcelVersion.EXTENDED:
                    ra_worksheet.write_number(counter, 0, activity.activity_id, gray_cell)
                    ra_worksheet.write(counter, 1, str(activity.name), gray_cell)
                    ra_worksheet.write(counter, 2, self.get_duration_str(activity.baseline_schedule.duration), gray_cell)
                    description = str(activity.risk_analysis.distribution_type.value) + " - " \
                                      + str(activity.risk_analysis.distribution_units.value)
                    ra_worksheet.write(counter, 3, description, yellow_cell)
                    ra_worksheet.write_number(counter, 4, activity.risk_analysis.optimistic_duration, yellow_cell)
                    ra_worksheet.write_number(counter, 5, activity.risk_analysis.probable_duration, yellow_cell)
                    ra_worksheet.write_number(counter, 6, activity.risk_analysis.pessimistic_duration, yellow_cell)
                else:
                    ra_worksheet.write_number(counter, 0, activity.activity_id, gray_cell)
                    ra_worksheet.write(counter, 1, str(activity.name), gray_cell)
                    description = str(activity.risk_analysis.distribution_type.value) + " - " \
                                      + str(activity.risk_analysis.distribution_units.value)
                    ra_worksheet.write(counter, 2, description, yellow_cell)
                    ra_worksheet.write_number(counter, 3, activity.risk_analysis.optimistic_duration, yellow_cell)
                    ra_worksheet.write_number(counter, 4, activity.risk_analysis.probable_duration, yellow_cell)
                    ra_worksheet.write_number(counter, 5, activity.risk_analysis.pessimistic_duration, yellow_cell)
            counter += 1

        # Write the tracking periods, same drill, multiple sheets, too many bloody columns
        for i in range(0, len(project_object.tracking_periods)):
            if i == 0:
                tracking_period_worksheet = workbook.add_worksheet("Project Control - TP1")
            else:
                tracking_period_worksheet = workbook.add_worksheet("TP" + str(i+1))

            # Set column widths and create headers
            if excel_version == ExcelVersion.EXTENDED:
                tracking_period_worksheet.set_column(0, 0, 3)
                tracking_period_worksheet.set_column(1, 1, 18)
                tracking_period_worksheet.set_column(2, 3, 12)
                tracking_period_worksheet.set_column(5, 5, 22)
                tracking_period_worksheet.set_column(11, 11, 12)
                tracking_period_worksheet.set_column(12, 12, 6)
                tracking_period_worksheet.set_column(21, 21, 8)
                tracking_period_worksheet.set_row(3, 30)

                tracking_period_worksheet.write('B1', 'TP Status Date', header)
                tracking_period_worksheet.write('E1', 'TP Name', header)
                tracking_period_worksheet.merge_range('A3:B3', "General", header)
                tracking_period_worksheet.merge_range('C3:E3', "Baseline", header)
                tracking_period_worksheet.merge_range('F3:G3', "Resource Demand", header)
                tracking_period_worksheet.merge_range('H3:K3', "Baseline Costs", header)
                tracking_period_worksheet.merge_range('L3:X3', "Tracking", header)

                tracking_period_worksheet.write('A4', 'ID', header)
                tracking_period_worksheet.write('B4', 'Name', header)
                tracking_period_worksheet.write('C4', 'Baseline Start', header)
                tracking_period_worksheet.write('D4', 'Baseline End', header)
                tracking_period_worksheet.write('E4', 'Duration', header)
                tracking_period_worksheet.write('F4', 'Resource Demand', header)
                tracking_period_worksheet.write('G4', 'Resource Cost', header)
                tracking_period_worksheet.write('H4', 'Fixed Cost', header)
                tracking_period_worksheet.write('I4', 'Cost/Hour', header)
                tracking_period_worksheet.write('J4', 'Variable Cost', header)
                tracking_period_worksheet.write('K4', 'Total Cost', header)
                tracking_period_worksheet.write('L4', 'Actual Start', header)
                tracking_period_worksheet.write('M4', 'Actual Duration', header)
                tracking_period_worksheet.write('N4', 'PAC', header)
                tracking_period_worksheet.write('O4', 'PRC', header)
                tracking_period_worksheet.write('P4', 'Remaining Duration', header)
                tracking_period_worksheet.write('Q4', 'PAC Dev', header)
                tracking_period_worksheet.write('R4', 'PRC Dev', header)
                tracking_period_worksheet.write('S4', 'Actual Cost', header)
                tracking_period_worksheet.write('T4', 'Remaining Cost', header)
                tracking_period_worksheet.write('U4', 'Percentage Completed', header)
                tracking_period_worksheet.write('V4', 'Tracking', header)
                tracking_period_worksheet.write('W4', 'Earned Value (EV)', header)
                tracking_period_worksheet.write('X4', 'Planned Value (PV)', header)
            else:
                tracking_period_worksheet.set_column(0, 0, 3)
                tracking_period_worksheet.set_column(1, 1, 18)
                tracking_period_worksheet.set_column(2, 2, 12)
                tracking_period_worksheet.set_column(3, 3, 9)
                tracking_period_worksheet.set_column(5, 5, 9)
                tracking_period_worksheet.set_row(3, 30)

                tracking_period_worksheet.write('B1', 'TP Status Date', header)
                tracking_period_worksheet.write('D1', 'TP Name', header)
                tracking_period_worksheet.merge_range('A3:B3', "General", header)
                tracking_period_worksheet.merge_range('C3:F3', "Tracking", header)

                tracking_period_worksheet.write('A4', 'ID', header)
                tracking_period_worksheet.write('B4', 'Name', header)
                tracking_period_worksheet.write('C4', 'Actual Start', header)
                tracking_period_worksheet.write('D4', 'Actual Duration', header)
                tracking_period_worksheet.write('E4', 'Actual Cost', header)
                tracking_period_worksheet.write('F4', 'Percentage Completed', header)

            if excel_version == ExcelVersion.EXTENDED:
                # Write the data
                tracking_period_worksheet.write_datetime('C1', project_object.tracking_periods[i].tracking_period_statusdate
                                                         , date_green_cell)
                tracking_period_worksheet.write('F1', project_object.tracking_periods[i].tracking_period_name,
                                                         yellow_cell)
                counter = 4
                self.calculate_aggregated_ac_per_wp(project_object.tracking_periods[i])
                self.calculate_aggregated_ad_per_wp(project_object.tracking_periods[i], project_object.agenda)
                for atr in project_object.tracking_periods[i].tracking_period_records:  # atr = ActivityTrackingRecord
                    if self.is_not_lowest_level_activity(atr.activity):
                        tracking_period_worksheet.write_number(counter, 0, atr.activity.activity_id, cyan_cell)
                        tracking_period_worksheet.write(counter, 1, atr.activity.name, cyan_cell)
                        tracking_period_worksheet.write_datetime(counter, 2, atr.activity.baseline_schedule.start, date_cyan_cell)
                        tracking_period_worksheet.write_datetime(counter, 3, atr.activity.baseline_schedule.end, date_cyan_cell)
                        tracking_period_worksheet.write(counter, 4, self.get_duration_str(atr.activity.baseline_schedule.duration), cyan_cell)
                        self.write_resources(tracking_period_worksheet, counter, 5, atr.activity.resources, cyan_cell)
                        tracking_period_worksheet.write_number(counter, 6, atr.activity.resource_cost, money_cyan_cell)
                        tracking_period_worksheet.write_number(counter, 7, atr.activity.baseline_schedule.fixed_cost, money_cyan_cell)
                        tracking_period_worksheet.write_number(counter, 8, atr.activity.baseline_schedule.hourly_cost, money_cyan_cell)
                        tracking_period_worksheet.write(counter, 9, "", money_cyan_cell)
                        tracking_period_worksheet.write_number(counter, 10, atr.activity.baseline_schedule.total_cost, money_cyan_cell)
                        if atr.actual_start:
                            tracking_period_worksheet.write_datetime(counter, 11, atr.actual_start, date_cyan_cell)
                        else:
                            tracking_period_worksheet.write(counter, 11, '', cyan_cell)
                        tracking_period_worksheet.write(counter, 12, self.get_duration_str(atr.actual_duration), cyan_cell)
                        tracking_period_worksheet.write_number(counter, 13, atr.planned_actual_cost, money_cyan_cell)
                        tracking_period_worksheet.write_number(counter, 14, atr.planned_remaining_cost, money_cyan_cell)
                        tracking_period_worksheet.write(counter, 15, self.get_duration_str(atr.remaining_duration), cyan_cell)
                        tracking_period_worksheet.write_number(counter, 16, atr.deviation_pac, money_cyan_cell)
                        tracking_period_worksheet.write_number(counter, 17, atr.deviation_prc, money_cyan_cell)

                        tracking_period_worksheet.write_number(counter, 18, atr.actual_cost, money_cyan_cell)
                        tracking_period_worksheet.write_number(counter, 19, atr.remaining_cost, money_cyan_cell)
                        percentage_completed = str(atr.percentage_completed) + "%"
                        tracking_period_worksheet.write(counter, 20, percentage_completed, cyan_cell)
                        tracking_period_worksheet.write(counter, 21, atr.tracking_status, cyan_cell)
                        tracking_period_worksheet.write_number(counter, 22, atr.earned_value, money_cyan_cell)
                        tracking_period_worksheet.write_number(counter, 23, atr.planned_value, money_cyan_cell)
                    else:
                        tracking_period_worksheet.write_number(counter, 0, atr.activity.activity_id, gray_cell)
                        tracking_period_worksheet.write(counter, 1, atr.activity.name, gray_cell)
                        tracking_period_worksheet.write_datetime(counter, 2, atr.activity.baseline_schedule.start, date_gray_cell)
                        tracking_period_worksheet.write_datetime(counter, 3, atr.activity.baseline_schedule.end, date_gray_cell)
                        tracking_period_worksheet.write(counter, 4, self.get_duration_str(atr.activity.baseline_schedule.duration), gray_cell)
                        self.write_resources(tracking_period_worksheet, counter, 5, atr.activity.resources, gray_cell)
                        tracking_period_worksheet.write_number(counter, 6, atr.activity.resource_cost, money_gray_cell)
                        tracking_period_worksheet.write_number(counter, 7, atr.activity.baseline_schedule.fixed_cost, money_gray_cell)
                        tracking_period_worksheet.write_number(counter, 8, atr.activity.baseline_schedule.hourly_cost, money_gray_cell)
                        tracking_period_worksheet.write_number(counter, 9, atr.activity.baseline_schedule.var_cost, money_gray_cell)
                        tracking_period_worksheet.write_number(counter, 10, atr.activity.baseline_schedule.total_cost, money_gray_cell)
                        if atr.actual_start:
                            tracking_period_worksheet.write_datetime(counter, 11, atr.actual_start, date_lime_cell)
                        else:
                            tracking_period_worksheet.write(counter, 11, '', green_cell)
                        tracking_period_worksheet.write(counter, 12, self.get_duration_str(atr.actual_duration), green_cell)
                        tracking_period_worksheet.write_number(counter, 13, atr.planned_actual_cost, money_gray_cell)
                        tracking_period_worksheet.write_number(counter, 14, atr.planned_remaining_cost, money_gray_cell)
                        tracking_period_worksheet.write(counter, 15, self.get_duration_str(atr.remaining_duration), gray_cell)
                        tracking_period_worksheet.write_number(counter, 16, atr.deviation_pac, money_gray_cell)
                        tracking_period_worksheet.write_number(counter, 17, atr.deviation_prc, money_gray_cell)
                        tracking_period_worksheet.write_number(counter, 18, atr.actual_cost, money_lime_cell)
                        tracking_period_worksheet.write_number(counter, 19, atr.remaining_cost, money_gray_cell)
                        percentage_completed = str(atr.percentage_completed) + "%"
                        tracking_period_worksheet.write(counter, 20, percentage_completed, green_cell)
                        tracking_period_worksheet.write(counter, 21, atr.tracking_status, gray_cell)
                        tracking_period_worksheet.write_number(counter, 22, atr.earned_value, money_gray_cell)
                        tracking_period_worksheet.write_number(counter, 23, atr.planned_value, money_gray_cell)
                    counter += 1
            else:
                # Write the data
                tracking_period_worksheet.write_datetime('C1', project_object.tracking_periods[i].tracking_period_statusdate
                                                         , date_green_cell)
                tracking_period_worksheet.write('E1', project_object.tracking_periods[i].tracking_period_name,
                                                         yellow_cell)
                counter = 4
                self.calculate_aggregated_ac_per_wp(project_object.tracking_periods[i])
                self.calculate_aggregated_ad_per_wp(project_object.tracking_periods[i], project_object.agenda)
                for atr in project_object.tracking_periods[i].tracking_period_records:  # atr = ActivityTrackingRecord
                    if self.is_not_lowest_level_activity(atr.activity):
                        tracking_period_worksheet.write_number(counter, 0, atr.activity.activity_id, cyan_cell)
                        tracking_period_worksheet.write(counter, 1, atr.activity.name, cyan_cell)
                        if atr.actual_start:
                            tracking_period_worksheet.write_datetime(counter, 2, atr.actual_start, date_cyan_cell)
                        else:
                            tracking_period_worksheet.write(counter, 2, '', cyan_cell)
                        tracking_period_worksheet.write(counter, 3, self.get_duration_str(atr.actual_duration), cyan_cell)
                        tracking_period_worksheet.write_number(counter, 4, atr.actual_cost, money_cyan_cell)
                        percentage_completed = str(atr.percentage_completed) + "%"
                        tracking_period_worksheet.write(counter, 5, percentage_completed, cyan_cell)
                    else:
                        tracking_period_worksheet.write_number(counter, 0, atr.activity.activity_id, gray_cell)
                        tracking_period_worksheet.write(counter, 1, atr.activity.name, gray_cell)
                        if atr.actual_start:
                            tracking_period_worksheet.write_datetime(counter, 2, atr.actual_start, date_lime_cell)
                        else:
                            tracking_period_worksheet.write(counter, 2, '', green_cell)
                        tracking_period_worksheet.write(counter, 3, self.get_duration_str(atr.actual_duration), green_cell)
                        tracking_period_worksheet.write_number(counter, 4, atr.actual_cost, money_lime_cell)
                        percentage_completed = str(atr.percentage_completed) + "%"
                        tracking_period_worksheet.write(counter, 5, percentage_completed, green_cell)
                    counter += 1

        # Write the agenda
        agenda_worksheet = workbook.add_worksheet("Agenda")
        agenda_worksheet.set_column(0, 0, 8)
        agenda_worksheet.set_column(3, 3, 8)
        agenda_worksheet.set_column(6, 6, 12)
        agenda_worksheet.merge_range('A1:B1', 'Working Hours', header)
        agenda_worksheet.merge_range('D1:E1', 'Working Days', header)
        agenda_worksheet.write('G1', 'Holidays', header)
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        for i in range(0, 24):
            hour_string = str(i) + ":00 - " + str((i+1)%24) + ":00"
            agenda_worksheet.write(i+1, 0, hour_string, yellow_cell)
            if project_object.agenda.working_hours[i]:
                agenda_worksheet.write(i+1, 1, "Yes", green_cell)
            else:
                agenda_worksheet.write(i+1, 1, "No", red_cell)
        for i in range(0, 7):
            agenda_worksheet.write(i+1, 3, days[i], yellow_cell)
            if project_object.agenda.working_days[i]:
                agenda_worksheet.write(i+1, 4, "Yes", green_cell)
            else:
                agenda_worksheet.write(i+1, 4, "No", red_cell)
        counter = 1
        for holiday in project_object.agenda.holidays:
            agenda_worksheet.write(1, 6, holiday, yellow_cell)
            counter += 1

        # Write the tracking overview
        overview_worksheet = workbook.add_worksheet("Tracking Overview")
        overview_worksheet.set_column(0, 13, 15)
        overview_worksheet.set_row(1, 30)
        overview_worksheet.merge_range('A1:C1', 'General', header)
        overview_worksheet.merge_range('D1:G1', 'EVM Performance Measures', header)
        if excel_version == ExcelVersion.EXTENDED:
            overview_worksheet.merge_range('H1:N1', 'EVM Forecasting', header)
        else:
            overview_worksheet.merge_range('H1:AE1', 'EVM Forecasting', header)

        overview_worksheet.write('A2', "Name", header)
        overview_worksheet.write('B2', "Start Tracking Period", header)
        overview_worksheet.write('C2', "Status date", header)
        overview_worksheet.write('D2', "Planned Value (PV)", header)
        overview_worksheet.write('E2', "Earned Value (EV)", header)
        overview_worksheet.write('F2', "Actual Cost (AC)", header)
        overview_worksheet.write('G2', "Earned Schedule (ES)", header)
        overview_worksheet.write('H2', "Schedule Variance (SV)", header)
        overview_worksheet.write('I2', "Schedule Performance Index (SPI)", header)
        overview_worksheet.write('J2', "Cost Variance (CV)", header)
        overview_worksheet.write('K2', "Cost Performance Index (CPI)", header)
        overview_worksheet.write('L2', "Schedule Variance (SV(t))", header)
        overview_worksheet.write('M2', "Schedule Performance Index (SPI(t))", header)
        overview_worksheet.write('N2', "p-factor", header)
        if excel_version == ExcelVersion.EXTENDED:
            overview_worksheet.write('O2', "EAC(t)-PV (PF=1)", header)
            overview_worksheet.write('P2', "EAC(t)-PV (PF=SPI)", header)
            overview_worksheet.write('Q2', "EAC(t)-PV (PF=SCI)", header)
            overview_worksheet.write('R2', "EAC(t)-ED (PF=1)", header)
            overview_worksheet.write('S2', "EAC(t)-ED (PF=SPI)", header)
            overview_worksheet.write('T2', "EAC(t)-ED (PF=SPI)", header)
            overview_worksheet.write('U2', "EAC(t)-ES (PF=1)", header)
            overview_worksheet.write('V2', "EAC(t)-ES (PF=SPI(t))", header)
            overview_worksheet.write('W2', "EAC(t)-ES (PF=SCI(t))", header)
            overview_worksheet.write('X2', "EAC (PF=1)", header)
            overview_worksheet.write('Y2', "EAC (PF=CPI)", header)
            overview_worksheet.write('Z2', "EAC (PF=SPI)", header)
            overview_worksheet.write('AA2', "EAC (PF=SPI(t))", header)
            overview_worksheet.write('AB2', "EAC (PF=SCI)", header)
            overview_worksheet.write('AC2', "EAC (PF=SCI(t))", header)
            overview_worksheet.write('AD2', "EAC (PF=0.8*CPI+0.2*SPI)", header)
            overview_worksheet.write('AE2', "EAC (PF=0.8*CPI+0.2*SPI(t))", header)

        # generate PV curve:
        generatedPVcurve = self.calculate_PVcurve(project_object)

        counter = 2
        for tracking_period in project_object.tracking_periods:
            overview_worksheet.write(counter, 0, tracking_period.tracking_period_name, green_cell)
            index = project_object.tracking_periods.index(tracking_period)
            if index == 0:
                overview_worksheet.write_datetime(counter, 1, min([atr.activity.baseline_schedule.start for atr in tracking_period.tracking_period_records]), date_green_cell)
            else:
                overview_worksheet.write_datetime(counter, 1, project_object.tracking_periods[index-1].tracking_period_statusdate, date_green_cell)
            overview_worksheet.write_datetime(counter, 2, tracking_period.tracking_period_statusdate, date_green_cell)
            overview_worksheet.write_number(counter, 3, self.calculate_aggregated_pv(tracking_period), money_green_cell)
            # calculate EV
            EV = self.calculate_aggregated_ev(tracking_period)
            overview_worksheet.write_number(counter, 4, EV, money_green_cell)
            overview_worksheet.write_number(counter, 5, self.calculate_aggregated_ac(tracking_period), money_green_cell)
            # calculate ES
            ES = self.calculate_es(project_object, generatedPVcurve, EV, tracking_period.tracking_period_statusdate)
            overview_worksheet.write_datetime(counter, 6, ES, date_green_cell)
            sv = self.calculate_aggregated_ev(tracking_period) - self.calculate_aggregated_pv(tracking_period)
            overview_worksheet.write_number(counter, 7, sv, money_green_cell)
            if not self.calculate_aggregated_pv(tracking_period):
                spi = 0
            else:
                spi = self.calculate_aggregated_ev(tracking_period)/self.calculate_aggregated_pv(tracking_period)
            # save spi value also in tracking_period for visualisations:
            tracking_period.spi = spi
            overview_worksheet.write(counter, 8, str(round(spi * 100)) + "%", green_cell)
            cv = self.calculate_aggregated_ev(tracking_period) - self.calculate_aggregated_ac(tracking_period)
            overview_worksheet.write_number(counter, 9, cv, money_green_cell)
            if not self.calculate_aggregated_ac(tracking_period):
                cpi = 0
            else:
                cpi = self.calculate_aggregated_ev(tracking_period)/self.calculate_aggregated_ac(tracking_period)
            # save cpi value also in tracking_period for visualisations:
            tracking_period.cpi = cpi
            overview_worksheet.write(counter, 10, str(round(cpi * 100)) +"%", green_cell)

            # calculate SV(t)
            sv_t, sv_t_str = self.calculate_SVt(project_object, ES, tracking_period.tracking_period_statusdate)
            # save SV(t) value also in tracking_period for visualisations:
            tracking_period.sv_t = sv_t
            overview_worksheet.write(counter, 11, sv_t_str, green_cell)

            # calculate SPI(t)
            spi_t = self.calculate_SPIt(project_object, ES, tracking_period.tracking_period_statusdate)
            # save spi_t value also in tracking_period for visualisations:
            tracking_period.spi_t = spi_t
            overview_worksheet.write(counter, 12, str(round(spi_t * 100)) + "%", green_cell)

            # calculate p-factor
            p_factor = self.calculate_p_factor(project_object, tracking_period, ES)
            overview_worksheet.write(counter, 13, str(round(p_factor * 100)) + "%", green_cell)
            # save p_factor value also in tracking_period for visualisations:
            tracking_period.p_factor = p_factor

            # TODO: more metrics

            if excel_version == ExcelVersion.EXTENDED:
                overview_worksheet.write_number(counter, 23, self.calculate_eac(self.calculate_aggregated_ac(tracking_period),
                                                                                self.get_bac(tracking_period), self.calculate_aggregated_ev(tracking_period),
                                                                                1), money_green_cell)
                if cpi != 0:
                    cpi = self.calculate_aggregated_ev(tracking_period)/self.calculate_aggregated_ac(tracking_period)
                    overview_worksheet.write_number(counter, 24, self.calculate_eac(self.calculate_aggregated_ac(tracking_period),
                                                                                    self.get_bac(tracking_period), self.calculate_aggregated_ev(tracking_period),
                                                                                    cpi), money_green_cell)
                else:
                    overview_worksheet.write_number(counter, 24, 0, money_green_cell)
                if spi != 0:
                    spi = self.calculate_aggregated_ev(tracking_period)/self.calculate_aggregated_pv(tracking_period)
                    overview_worksheet.write_number(counter, 25, self.calculate_eac(self.calculate_aggregated_ac(tracking_period),
                                                                                    self.get_bac(tracking_period), self.calculate_aggregated_ev(tracking_period),
                                                                                    spi), money_green_cell)
                else:
                    overview_worksheet.write_number(counter, 25, 0, money_green_cell)

                if cpi != 0 or spi != 0:
                    cpi = self.calculate_aggregated_ev(tracking_period)/self.calculate_aggregated_ac(tracking_period)
                    spi = self.calculate_aggregated_ev(tracking_period)/self.calculate_aggregated_pv(tracking_period)
                    overview_worksheet.write_number(counter, 29, self.calculate_eac(self.calculate_aggregated_ac(tracking_period),
                                                                                    self.get_bac(tracking_period), self.calculate_aggregated_ev(tracking_period),
                                                                                    0.8*cpi+0.2*spi), money_green_cell)
                else:
                    overview_worksheet.write_number(counter, 29, 0, money_green_cell)

            # TODO: more metrics

            counter += 1

        return workbook

    @staticmethod
    def get_duration_str(delta, negativeValue=False):
        if delta:
            # Writing a duration requires some converting..
            if delta.days != 0 and delta.seconds != 0:
                if not negativeValue:
                    duration = str(delta.days) + "d " + str(int(delta.seconds / 3600)) + "h"
                else:
                    duration = "-" + str(delta.days) + "d " + str(int(delta.seconds / 3600)) + "h"
            elif delta.seconds != 0:
                if not negativeValue:
                    duration = str(int(delta.seconds / 3600)) + "h"
                else:
                    duration = "-" + str(int(delta.seconds / 3600)) + "h"
            else:
                if not negativeValue:
                    duration = str(delta.days) + "d"
                else:
                    duration = "-" + str(delta.days) + "d"
            return duration
        return "0"

    @staticmethod
    def write_wbs(worksheet, row, column, wbs, format):
        # Instead of writing (x, y, z) write x.y.z
        to_write = ''
        if len(wbs) > 0:
            for i in range(0, len(wbs)-1):
                to_write += str(wbs[i]) + '.'
            to_write += str(wbs[-1])
        worksheet.write(row, column, to_write, format)

    @staticmethod
    def write_successors(worksheet, row, column, successors, format):
        # Write in format of FSxx;SSxx..
        to_write = ''
        if len(successors) > 0:
            for i in range(0, len(successors)-1):
                to_write += str(successors[i][1]) + str(successors[i][0])
                if successors[i][2] != 0:
                    to_write += "+" + str(successors[i][2])
                to_write += ";"
            to_write += str(successors[-1][1]) + str(successors[-1][0])
            if successors[-1][2] != 0:
                to_write += "+" + str(successors[-1][2])
        worksheet.write(row, column, to_write, format)

    @staticmethod
    def write_predecessors(worksheet, row, column, predecessors, format):
        # Write in format of xxSF;xxFF..
        to_write = ''
        if len(predecessors) > 0:
            for i in range(0, len(predecessors)-1):
                to_write += str(predecessors[i][0]) + str(predecessors[i][1])
                if predecessors[i][2] != 0:
                    to_write += "+" + str(predecessors[i][2])
                to_write += ";"
            to_write += str(predecessors[-1][0]) + str(predecessors[-1][1])
            if predecessors[-1][2] != 0:
                to_write += "+" + str(predecessors[-1][2])
        worksheet.write(row, column, to_write, format)

    @staticmethod
    def is_not_lowest_level_activity(activity):
        # Decide whether an activity is not of the lowest level or not.
        if activity.baseline_schedule.var_cost is not None:
            return False
        else:
            return True

    @staticmethod
    def write_resources(worksheet, row, column, resources, format):
        # Write out the resources in a specific format
        to_write = ''
        if len(resources) > 0:
            for i in range(0, len(resources)-1):
                to_write += resources[i][0].name
                if resources[i][1] > 1:
                    to_write += "[" + str(float(resources[i][1])) + " #" + str(resources[i][0].availability) + "]; "
            to_write += resources[-1][0].name
            if resources[-1][1] > 1:
                to_write += "[" + str(float(resources[-1][1])) + " #" + str(resources[-1][0].availability) + "]"
        worksheet.write(row, column, to_write, format)

    @staticmethod
    def write_resource_assign_cost(worksheet, row, column, resource, activities, format, moneyformat, project_object):
        # For every resource, we check to which activities it is assigned and what the total cost is
        to_write = ''
        cost = 0
        for activity in activities:
            for _resource in activity.resources:
                if resource == _resource[0]:
                    if _resource[1] == 1:
                        to_write += str(activity.activity_id) + ';'
                    else:
                        to_write += str(activity.activity_id) + '[' + str(_resource[1]) + ' #' \
                                    + str(resource.availability) + '];'
                    cost += (activity.baseline_schedule.duration.days*project_object.agenda.get_working_hours_in_a_day() +
                             activity.baseline_schedule.duration.seconds/3600)*(_resource[1]*resource.cost_unit)
        worksheet.write(row, column, to_write, format)
        worksheet.write_number(row, column+1, cost, moneyformat)

