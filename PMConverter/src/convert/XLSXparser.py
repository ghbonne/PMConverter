from line_chart_example import workbook
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
        self.process_baseline_schedule(activities_sheet, resource_sheet, risk_analysis_sheet)

    def process_baseline_schedule(self, activities_sheet=None, resource_sheet=None, risk_analysis_sheet=None):
        # First we process the resources sheet, we store them in a dict, with index the activity_id, to access them
        # easily later when we are processing the activities.

        # Then we process the risk analysis sheet, again everything is stored in a dict.
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

        # Finally, the sheet with activities is processed.
        for curr_row in range(self.get_nr_of_header_lines(activities_sheet), activities_sheet.get_highest_row()-1):
            activity_id = activities_sheet.cell(row=curr_row, column=1).value
            activity_name = activities_sheet.cell(row=curr_row, column=2).value
            activity_wbs = activities_sheet.cell(row=curr_row, column=3).value
            activity_predecessors = activities_sheet.cell(row=curr_row, column=4).value
            activity_successors = activities_sheet.cell(row=curr_row, column=5).value
            baseline_start = activities_sheet.cell(row=curr_row, column=6).value
            baseline_duration = activities_sheet.cell(row=curr_row, column=8).value
            baseline_fixed_cost = activities_sheet.cell(row=curr_row, column=11).value
            baseline_hourly_cost = activities_sheet.cell(row=curr_row, column=12).value
            baseline_var_cost = activities_sheet.cell(row=curr_row, column=13).value
            print([activity_id, activity_name, activity_wbs, activity_predecessors, activity_successors,
                   baseline_start, baseline_duration, baseline_fixed_cost, baseline_hourly_cost, baseline_var_cost])

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