import copy
import re
import datetime
import time
import math
from object.activity import Activity
from object.baselineschedule import BaselineScheduleRecord
from object.projectobject import ProjectObject
from object.resource import Resource
from object.riskanalysisdistribution import RiskAnalysisDistribution

__author__ = 'PM Group 8'

import xlsxwriter
import openpyxl
from convert.fileparser import FileParser


class XLSXParser(FileParser):

    def __init__(self):
        super().__init__()

    def to_schedule_object(self, file_path_input):
        workbook = openpyxl.load_workbook(file_path_input)
        project_control_sheets = []
        for name in workbook.get_sheet_names():
            if "Baseline Schedule" in name:
                activities_sheet = workbook.get_sheet_by_name(name)
            elif "Resources" in name:
                resource_sheet = workbook.get_sheet_by_name(name)
            elif "Risk Analysis" in name:
                risk_analysis_sheet = workbook.get_sheet_by_name(name)
            elif "TP" in name:
                project_control_sheets.append(workbook.get_sheet_by_name(name))

        # First we process the resources sheet, we store them in a dict, with index the resource name, to access them
        # easily later when we are processing the activities.
        resources_dict = self.process_resources(resource_sheet)
        # Then we process the risk analysis sheet, again everything is stored in a dict, now the index is activity_id.
        risk_analysis_dict = self.process_risk_analysis(risk_analysis_sheet)
        # Finally, the sheet with activities is processed, using the dicts we created above.
        activities = self.process_baseline_schedule(activities_sheet, resources_dict, risk_analysis_dict)
        return ProjectObject(activities=activities, resources=list(resources_dict.values()))

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

    def process_risk_analysis(self, risk_analysis_sheet):
        risk_analysis_dict = {}
        for curr_row in range(self.get_nr_of_header_lines(risk_analysis_sheet), risk_analysis_sheet.get_highest_row()+1):
            if risk_analysis_sheet.cell(row=curr_row, column=4).value is not None:
                risk_ana_dist_type = risk_analysis_sheet.cell(row=curr_row, column=4).value.split(" - ")[0]
                risk_ana_dist_units = risk_analysis_sheet.cell(row=curr_row, column=4).value.split(" - ")[1]
                risk_ana_opt_duration = int(risk_analysis_sheet.cell(row=curr_row, column=5).value)
                risk_ana_prob_duration = int(risk_analysis_sheet.cell(row=curr_row, column=6).value)
                risk_ana_pess_duration = int(risk_analysis_sheet.cell(row=curr_row, column=7).value)
                dict_id = int(risk_analysis_sheet.cell(row=curr_row, column=1).value)
                risk_analysis_dict[dict_id] = RiskAnalysisDistribution(distribution_type=risk_ana_dist_type,
                                                                       distribution_units=risk_ana_dist_units,
                                                                       optimistic_duration=risk_ana_opt_duration,
                                                                       probable_duration=risk_ana_prob_duration,
                                                                       pessimistic_duration=risk_ana_pess_duration)
        return risk_analysis_dict

    def process_project_control(self, sheet):
        pass

    def process_baseline_schedule(self, activities_sheet, resources_dict, risk_analysis_dict):
        activities = []
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
            baseline_start = datetime.datetime.utcfromtimestamp(((activities_sheet.cell(row=curr_row, column=6).value -
                                                                  25569)*86400))  # Convert to correct date
            baseline_end = datetime.datetime.utcfromtimestamp(((activities_sheet.cell(row=curr_row, column=7).value -
                                                                25569)*86400))  # Convert to correct date
            baseline_duration_split = activities_sheet.cell(row=curr_row, column=8).value.split("d")
            baseline_duration_days = int(baseline_duration_split[0])
            baseline_duration_hours = 0  # We need to set this default value for the next loop
            if baseline_duration_split[1] != '':
                baseline_duration_hours = int(baseline_duration_split[1][1:-1])  # first char = " "; last char = "h"
            baseline_duration = datetime.timedelta(days=baseline_duration_days, hours=baseline_duration_hours)
            baseline_fixed_cost = float(activities_sheet.cell(row=curr_row, column=11).value)
            baseline_hourly_cost = 0.0
            if activities_sheet.cell(row=curr_row, column=12).value:
                baseline_hourly_cost = float(activities_sheet.cell(row=curr_row, column=12).value)
            baseline_var_cost = 0.0
            if activities_sheet.cell(row=curr_row, column=13).value:
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
            activities.append(Activity(activity_id, name=activity_name, wbs_id=activity_wbs,
                                       predecessors=activity_predecessors, successors=activity_successors,
                                       baseline_schedule=activity_baseline_schedule,
                                       resource_cost=activity_resource_cost,
                                       risk_analysis=activity_risk_analysis, resources=activity_resources))
        return activities

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
        while not sheet.cell(row=header_lines, column=1).value.isdigit():
            header_lines += 1
        return header_lines

    def from_schedule_object(self, project_object, file_path_output):
        """
        This is just a lot of writing to excel code, it is ugly..

        """
        workbook = xlsxwriter.Workbook(file_path_output)

        # Lots of formats
        header = workbook.add_format({'bold': True, 'bg_color': 'blue', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        yellow_cell = workbook.add_format({'bg_color': 'yellow', 'text_wrap': True, 'border': 1, 'font_size': 8})
        cyan_cell = workbook.add_format({'bg_color': 'cyan', 'text_wrap': True, 'border': 1, 'font_size': 8})
        green_cell = workbook.add_format({'bg_color': 'green', 'text_wrap': True, 'border': 1, 'font_size': 8})
        navy_cell = workbook.add_format({'bg_color': '#D4D0C8', 'text_wrap': True, 'border': 1, 'font_size': 8})
        date_cyan_cell = workbook.add_format({'bg_color': 'cyan', 'text_wrap': True, 'border': 1,
                                              'num_format': 'mm/dd/yyyy H:MM', 'font_size': 8})
        date_green_cell = workbook.add_format({'bg_color': 'green', 'text_wrap': True, 'border': 1,
                                              'num_format': 'mm/dd/yyyy H:MM', 'font_size': 8})
        date_lime_cell = workbook.add_format({'bg_color': 'lime', 'text_wrap': True, 'border': 1,
                                              'num_format': 'mm/dd/yyyy H:MM', 'font_size': 8})
        money_cyan_cell = workbook.add_format({'bg_color': 'cyan', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00 €', 'font_size': 8})
        money_green_cell = workbook.add_format({'bg_color': 'green', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00 €', 'font_size': 8})
        money_lime_cell = workbook.add_format({'bg_color': 'lime', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00 €', 'font_size': 8})
        money_navy_cell = workbook.add_format({'bg_color': '#D4D0C8', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00 €', 'font_size': 8})
        money_yellow_cell = workbook.add_format({'bg_color': 'yellow', 'text_wrap': True, 'border': 1,
                                              'num_format': '#,##0.00 €', 'font_size': 8})

        # Worksheets
        bsch_worksheet = workbook.add_worksheet("Baseline Schedule")
        res_worksheet = workbook.add_worksheet("Resources")
        ra_worksheet = workbook.add_worksheet("Risk Analysis")

        # Write the Baseline Schedule Worksheet

        # Set the width of the columns
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

        # Set the height of rows
        bsch_worksheet.set_row(1, 30)

        # Write header cells (using the header format, and by merging some cells)
        bsch_worksheet.merge_range('A1:C1',
                                   "General", header)
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

        # Now we run through all activities to get the required information
        counter = 2
        for activity in project_object.activities:
            if not self.is_not_lowest_level_activity(project_object.activities, activity):
                # Write activity of lowest level
                bsch_worksheet.write_number(counter, 0, activity.activity_id, green_cell)
                bsch_worksheet.write(counter, 1, str(activity.name), green_cell)
                self.write_wbs(bsch_worksheet, counter, 2, activity.wbs_id, navy_cell)
                self.write_predecessors(bsch_worksheet, counter, 3, activity.predecessors, green_cell)
                self.write_successors(bsch_worksheet, counter, 4, activity.successors, green_cell)
                bsch_worksheet.write_datetime(counter, 5, activity.baseline_schedule.start, date_green_cell)
                bsch_worksheet.write_datetime(counter, 6, activity.baseline_schedule.end, date_lime_cell)
                if activity.baseline_schedule.duration.seconds != 0:
                    duration = str(activity.baseline_schedule.duration.days) + "d " \
                               + str(int(activity.baseline_schedule.duration.seconds / 3600)) + "h"
                else:
                    duration = str(activity.baseline_schedule.duration.days) + "d "
                bsch_worksheet.write(counter, 7, duration, green_cell)
                self.write_resources(bsch_worksheet, counter, 8, activity.resources, yellow_cell)
                bsch_worksheet.write_number(counter, 9, activity.resource_cost, money_navy_cell)
                bsch_worksheet.write_number(counter, 10, activity.baseline_schedule.fixed_cost, money_green_cell)
                bsch_worksheet.write_number(counter, 11, activity.baseline_schedule.hourly_cost, money_lime_cell)
                bsch_worksheet.write_number(counter, 12, activity.baseline_schedule.var_cost, money_green_cell)
                bsch_worksheet.write_number(counter, 13, activity.baseline_schedule.total_cost, money_navy_cell)
            else:
                # Write aggregated activity
                bsch_worksheet.write_number(counter, 0, activity.activity_id, yellow_cell)
                bsch_worksheet.write(counter, 1, str(activity.name), yellow_cell)
                self.write_wbs(bsch_worksheet, counter, 2, activity.wbs_id, cyan_cell)
                self.write_predecessors(bsch_worksheet, counter, 3, activity.predecessors, cyan_cell)
                self.write_successors(bsch_worksheet, counter, 4, activity.successors, cyan_cell)
                bsch_worksheet.write_datetime(counter, 5, activity.baseline_schedule.start, date_cyan_cell)
                bsch_worksheet.write_datetime(counter, 6, activity.baseline_schedule.end, date_cyan_cell)
                if activity.baseline_schedule.duration.seconds != 0:
                    duration = str(activity.baseline_schedule.duration.days) + "d " \
                               + str(int(activity.baseline_schedule.duration.seconds / 3600)) + "h"
                else:
                    duration = str(activity.baseline_schedule.duration.days) + "d "
                bsch_worksheet.write(counter, 7, duration, cyan_cell)
                self.write_resources(bsch_worksheet, counter, 8, activity.resources, cyan_cell)
                bsch_worksheet.write_number(counter, 9, activity.resource_cost, money_cyan_cell)
                bsch_worksheet.write_number(counter, 10, activity.baseline_schedule.fixed_cost, money_cyan_cell)
                bsch_worksheet.write_number(counter, 11, activity.baseline_schedule.hourly_cost, money_cyan_cell)
                bsch_worksheet.write_number(counter, 12, activity.baseline_schedule.var_cost, money_cyan_cell)
                bsch_worksheet.write_number(counter, 13, activity.baseline_schedule.total_cost, money_cyan_cell)

            counter += 1

        # Write the resources sheet

        # Some small adjustments to rows and columns in the resource sheet
        res_worksheet.set_row(1, 25)
        res_worksheet.set_column(1, 1, 15)
        res_worksheet.set_column(6, 6, 40)

        # Write header cells (using the header format, and by merging some cells)
        res_worksheet.merge_range('A1:D1', "General", header)
        res_worksheet.merge_range('E1:F1', "Resource Cost", header)
        res_worksheet.merge_range('G1:H1', "Resource Demand", header)

        res_worksheet.write('A2', "ID", header)
        res_worksheet.write('B2', "Name", header)
        res_worksheet.write('C2', "Type", header)
        res_worksheet.write('D2', "Availability", header)
        res_worksheet.write('E2', "Cost/Use", header)
        res_worksheet.write('F2', "Cost/Unit", header)
        res_worksheet.write('G2', "Assigned To", header)
        res_worksheet.write('H2', "Total Cost", header)

        counter = 2
        for resource in project_object.resources:
            res_worksheet.write_number(counter, 0, resource.resource_id, yellow_cell)
            res_worksheet.write(counter, 1, resource.name, yellow_cell)
            res_worksheet.write(counter, 2, resource.resource_type, yellow_cell)
            # God knows why we write the availability twice, it was like that in the template
            useless_availability_string = str(resource.availability) + " #" + str(resource.availability)
            res_worksheet.write(counter, 3, useless_availability_string, yellow_cell)
            res_worksheet.write(counter, 4, resource.cost_use, money_yellow_cell)
            res_worksheet.write(counter, 5, resource.cost_unit, money_yellow_cell)
            self.write_resource_assign_cost(res_worksheet, counter, 6, resource, project_object.activities, cyan_cell,
                                            money_cyan_cell)
            counter += 1

        # Write the risk analysis sheet

        # Adjust some column widths
        ra_worksheet.set_column(0, 0, 3)
        ra_worksheet.set_column(1, 1, 18)
        ra_worksheet.set_column(3, 3, 15)
        ra_worksheet.set_column(4, 6, 12)

        # Write the headers
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

        # Write the rows by iterating through the activities (since they are linked to it)
        counter = 2
        for activity in project_object.activities:
            if self.is_not_lowest_level_activity(project_object.activities, activity):
                ra_worksheet.write_number(counter, 0, activity.activity_id, cyan_cell)
                ra_worksheet.write(counter, 1, str(activity.name), cyan_cell)
                if activity.baseline_schedule.duration.seconds != 0:
                    duration = str(activity.baseline_schedule.duration.days) + "d " \
                               + str(int(activity.baseline_schedule.duration.seconds / 3600)) + "h"
                else:
                    duration = str(activity.baseline_schedule.duration.days) + "d "
                ra_worksheet.write(counter, 2, duration, cyan_cell)
                ra_worksheet.write(counter, 3, "", cyan_cell)
                ra_worksheet.write(counter, 4, "", cyan_cell)
                ra_worksheet.write(counter, 5, "", cyan_cell)
                ra_worksheet.write(counter, 6, "", cyan_cell)
            else:
                ra_worksheet.write_number(counter, 0, activity.activity_id, navy_cell)
                ra_worksheet.write(counter, 1, str(activity.name), navy_cell)
                if activity.baseline_schedule.duration.seconds != 0:
                    duration = str(activity.baseline_schedule.duration.days) + "d " \
                               + str(int(activity.baseline_schedule.duration.seconds / 3600)) + "h"
                else:
                    duration = str(activity.baseline_schedule.duration.days) + "d "
                ra_worksheet.write(counter, 2, duration, navy_cell)
                description = str(activity.risk_analysis.distribution_type) + " - " \
                              + str(activity.risk_analysis.distribution_units)
                ra_worksheet.write(counter, 3, description, yellow_cell)
                ra_worksheet.write_number(counter, 4, activity.risk_analysis.optimistic_duration, yellow_cell)
                ra_worksheet.write_number(counter, 5, activity.risk_analysis.probable_duration, yellow_cell)
                ra_worksheet.write_number(counter, 6, activity.risk_analysis.pessimistic_duration, yellow_cell)
            counter += 1

        workbook.close()

    @staticmethod
    def write_wbs(worksheet, row, column, wbs, format):
        # Instead of writing (x, y, z) write x.y.z
        to_write = ''
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
            to_write += str(predecessors[-1][1]) + str(predecessors[-1][0])
            if predecessors[-1][2] != 0:
                to_write += "+" + str(predecessors[-1][2])
        worksheet.write(row, column, to_write, format)

    @staticmethod
    def is_not_lowest_level_activity(activities, activity):
        # Decide whether an activity is not of the lowest level or not. TODO: A possible optimization is just looking
        # at the Activity next to 'activity' since the list is sorted on wbs_id
        for other_activity in activities:
            if other_activity.wbs_id != activity.wbs_id and len(other_activity.wbs_id) > len(activity.wbs_id):
                for i in range(0, len(activity.wbs_id)):
                    if activity.wbs_id[i] != other_activity.wbs_id[i]:
                        break
                if i == len(activity.wbs_id)-1:
                    return True
        return False

    @staticmethod
    def write_resources(worksheet, row, column, resources, format):
        to_write = ''
        if len(resources) > 0:
            for i in range(0, len(resources)-1):
                to_write += resources[i][0].name
                if resources[i][1] > 1:
                    to_write += " [" + str(float(resources[i][1])) + " #" + str(resources[i][0].availability) + "]; "
            to_write += resources[-1][0].name
            if resources[-1][1] > 1:
                to_write += " [" + str(float(resources[-1][1])) + " #" + str(resources[-1][0].availability) + "]"
        worksheet.write(row, column, to_write, format)

    @staticmethod
    def write_resource_assign_cost(worksheet, row, column, resource, activities, format, moneyformat):
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
                    cost += (activity.baseline_schedule.duration.days*8 +
                             activity.baseline_schedule.duration.seconds/3600)*(_resource[1]*resource.cost_unit)
        worksheet.write(row, column, to_write, format)
        worksheet.write_number(row, column+1, cost, moneyformat)

