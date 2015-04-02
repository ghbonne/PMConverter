__author__ = 'PM Group 8'

import xlsxwriter

from convert.fileparser import FileParser


class XMLParser(FileParser):

    def __init__(self):
        super().__init__()

    def to_schedule_object(self, file_path_input):
        pass

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

        bsch_worksheet.merge_range('A1:C1', "General", bold)
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
            bsch_worksheet.write_datetime(counter, 6, activity.baseline_schedule.end)
            bsch_worksheet.write(counter, 7, self.get_duration(activity))
            bsch_worksheet.write(counter, 8, self.list_resources(activity))
            bsch_worksheet.write_number(counter, 9, self.get_resource_cost(activity))
            bsch_worksheet.write_number(counter, 10, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write_number(counter, 11, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write_number(counter, 12, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write_number(counter, 13, self.get_total_cost(activity))
            counter += 1

        workbook.close()

    def get_duration(self, activity):
        return "0"

    def get_resource_cost(self, activity):
        return 0

    def list_resources(self, activity):
        return "0"

    def get_total_cost(self, activity):
        return 0