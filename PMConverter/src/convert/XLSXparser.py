import re
import datetime
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
        return ProjectObject(activities=self.process_baseline_schedule(activities_sheet, resource_sheet,
                                                                       risk_analysis_sheet))

    def process_baseline_schedule(self, activities_sheet=None, resource_sheet=None, risk_analysis_sheet=None):
        # First we process the resources sheet, we store them in a dict, with index the resource name, to access them
        # easily later when we are processing the activities.
        resources_dict = {}
        for curr_row in range(self.get_nr_of_header_lines(resource_sheet), resource_sheet.get_highest_row()-1):
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

        print(resources_dict)

        # Then we process the risk analysis sheet, again everything is stored in a dict, now the index is activity_id.
        risk_analysis_dict = {}
        for curr_row in range(self.get_nr_of_header_lines(risk_analysis_sheet), risk_analysis_sheet.get_highest_row()-1):
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
        print(risk_analysis_dict)

        # Finally, the sheet with activities is processed.
        activities = []
        for curr_row in range(self.get_nr_of_header_lines(activities_sheet), activities_sheet.get_highest_row()-1):
            activity_id = int(activities_sheet.cell(row=curr_row, column=1).value)
            activity_name = activities_sheet.cell(row=curr_row, column=2).value
            activity_wbs = ()
            for number in activities_sheet.cell(row=curr_row, column=3).value.split("."):
                activity_wbs = activity_wbs + (int(number),)
            activity_predecessors = self.process_predecessors(activities_sheet.cell(row=curr_row, column=4).value)
            activity_successors = self.process_successors(activities_sheet.cell(row=curr_row, column=5).value)
            baseline_start = activities_sheet.cell(row=curr_row, column=6).value
            #baseline_duration = activities_sheet.cell(row=curr_row, column=8).value
            #baseline_start = datetime.datetime.strptime(activities_sheet.cell(row=curr_row, column=6).value,
            #                                            '%m/%d/%Y %H:%M')
            baseline_duration_split = activities_sheet.cell(row=curr_row, column=8).value.split("d")
            baseline_duration_days = int(baseline_duration_split[0])
            baseline_duration_hours = 0
            if baseline_duration_split[1] != '':
                baseline_duration_hours = int(baseline_duration_split[1][1:-1])
            baseline_duration = datetime.timedelta(days=baseline_duration_days, hours=baseline_duration_hours)
            baseline_fixed_cost = float(activities_sheet.cell(row=curr_row, column=11).value)
            baseline_hourly_cost = 0
            if activities_sheet.cell(row=curr_row, column=12).value:
                baseline_hourly_cost = float(activities_sheet.cell(row=curr_row, column=12).value)
            baseline_var_cost = 0
            if activities_sheet.cell(row=curr_row, column=13).value:
                baseline_var_cost = float(activities_sheet.cell(row=curr_row, column=13).value)
            activity_baseline_schedule = BaselineScheduleRecord(start=baseline_start, duration=baseline_duration,
                                                                fixed_cost=baseline_fixed_cost,
                                                                hourly_cost=baseline_hourly_cost,
                                                                var_cost=baseline_var_cost)
            activity_risk_analysis = None
            if activity_id in risk_analysis_dict:
                activity_risk_analysis = risk_analysis_dict[activity_id]
            # TODO: what happens if there is more than 1 resource?
            activity_resources = []
            if activities_sheet.cell(row=curr_row, column=9).value:
                activity_resource_name = resources_dict[activities_sheet.cell(row=curr_row, column=9).value
                                                        .split("[")[0]]
                activity_resource_demand = 1
                if len(activities_sheet.cell(row=curr_row, column=9).value.split("[")) > 1:
                    activity_resource_demand = int(float(activities_sheet.cell(row=curr_row, column=9).value
                                                         .split("[")[1].split(" ")[0]
                                                         .translate(str.maketrans(",", "."))))
                activity_resources.append((activity_resource_name, activity_resource_demand))
            print(activity_predecessors)
            print(activity_successors)
            activities.append(Activity(activity_id, name=activity_name, wbs_id=activity_wbs,
                                       predecessors=activity_predecessors, successors=activity_successors,
                                       baseline_schedule=activity_baseline_schedule,
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

    def process_resources(self, sheet):
        pass

    def process_risk_analysis(self, sheet):
        pass

    def process_project_control(self, sheet):
        pass

    @staticmethod
    def get_nr_of_header_lines(sheet):
        header_lines = 1
        while not sheet.cell(row=header_lines, column=1).value.isdigit():
            header_lines += 1
        return header_lines

    def from_schedule_object(self, project_object, file_path_output):
        workbook = xlsxwriter.Workbook(file_path_output)

        # Formats
        bold = workbook.add_format({'bold': 1})

        # Worksheets
        bsch_worksheet = workbook.add_worksheet("Baseline Schedule")
        res_worksheet = workbook.add_worksheet("Resources")
        ra_worksheet = workbook.add_worksheet("Risk Analysis")

        # Write the Baseline Schedule Worksheet
        # First, the header
        bsch_worksheet.set_column(0, 0, 3)
        bsch_worksheet.set_column(1, 1, 15)
        bsch_worksheet.set_column(2, 2, 8)
        bsch_worksheet.set_column(3, 6, 12)
        bsch_worksheet.set_column(7, 7, 8)
        bsch_worksheet.set_column(8, 8, 25)
        bsch_worksheet.set_column(9, 9, 10)
        bsch_worksheet.set_column(10, 11, 15)

        bsch_worksheet.merge_range('A1:C1',
                                   "General", bold)
        bsch_worksheet.merge_range('D1:E1', "Relations", bold)
        bsch_worksheet.merge_range('F1:H1', "Baseline", bold)
        bsch_worksheet.merge_range('I1:J1', "Resource Demand", bold)
        bsch_worksheet.merge_range('K1:N1', "Baseline Costs", bold)

        bsch_worksheet.write('A2', "ID", bold)
        bsch_worksheet.write('B2', "Name", bold)
        bsch_worksheet.write('C2', "WBS", bold)
        bsch_worksheet.write('D2', "Predecessors", bold)
        bsch_worksheet.write('E2', "Successors", bold)
        bsch_worksheet.write('F2', "Baseline Start", bold)
        bsch_worksheet.write('G2', "Baseline End", bold)
        bsch_worksheet.write('H2', "Duration", bold)
        bsch_worksheet.write('I2', "Resource Demand", bold)
        bsch_worksheet.write('J2', "Resource Cost", bold)
        bsch_worksheet.write('K2', "Fixed Cost", bold)
        bsch_worksheet.write('L2', "Cost/Hour", bold)
        bsch_worksheet.write('M2', "Variable Cost", bold)
        bsch_worksheet.write('N2', "Total Cost", bold)

        # Now we run through all activities to get the required information
        counter = 2
        for activity in project_object.activities:
            bsch_worksheet.write_number(counter, 0, activity.activity_id)
            bsch_worksheet.write(counter, 1, str(activity.name))
            bsch_worksheet.write(counter, 2, str(activity.wbs_id))
            bsch_worksheet.write(counter, 3, str(activity.predecessors))
            bsch_worksheet.write(counter, 4, str(activity.successors))
            bsch_worksheet.write_datetime(counter, 5, activity.baseline_schedule.start)
            bsch_worksheet.write(counter, 6, self.get_end_date(activity))
            bsch_worksheet.write_datetime(counter, 7, activity.baseline_schedule.duration)
            bsch_worksheet.write(counter, 8, self.list_resources(activity))
            bsch_worksheet.write_number(counter, 9, self.get_resource_cost(activity))
            bsch_worksheet.write_number(counter, 10, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write_number(counter, 11, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write_number(counter, 12, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write_number(counter, 13, self.get_total_cost(activity))
            counter += 1

        workbook.close()

    def get_end_date(self, activity):
        return "0"

    def get_resource_cost(self, activity):
        return 0

    def list_resources(self, activity):
        return "0"

    def get_total_cost(self, activity):
        return 0