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
        bsch_worksheet.write('A1', "ID", bold)
        bsch_worksheet.write('B1', "Name", bold)
        bsch_worksheet.write('C1', "WBS", bold)
        bsch_worksheet.write('D1', "Predecessors", bold)
        bsch_worksheet.write('E1', "Successors", bold)
        bsch_worksheet.write('F1', "Baseline Start", bold)
        bsch_worksheet.write('G1', "Baseline End", bold)
        bsch_worksheet.write('H1', "Duration", bold)
        bsch_worksheet.write('I1', "Resource Demand", bold)
        bsch_worksheet.write('J1', "Resource Cost", bold)
        bsch_worksheet.write('K1', "Fixed Cost", bold)
        bsch_worksheet.write('L1', "Cost/Hour", bold)
        bsch_worksheet.write('M1', "Variable Cost", bold)
        bsch_worksheet.write('N1', "Total Cost", bold)

        # Now we run through all activities to get the required information
        counter = 1
        for activity in project_object.activities:
            bsch_worksheet.write(counter, 0, str(activity.activity_id))
            bsch_worksheet.write(counter, 1, str(activity.name))
            bsch_worksheet.write(counter, 2, str(activity.wbs_id))
            bsch_worksheet.write(counter, 3, str(activity.predecessors))
            bsch_worksheet.write(counter, 4, str(activity.successors))
            bsch_worksheet.write(counter, 5, str(activity.baseline_schedule.start))
            bsch_worksheet.write(counter, 6, str(activity.baseline_schedule.end))
            bsch_worksheet.write(counter, 7, activity.get_duration())
            bsch_worksheet.write(counter, 8, activity.list_resources())
            bsch_worksheet.write(counter, 9, activity.get_resource_cost())
            bsch_worksheet.write(counter, 10, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write(counter, 11, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write(counter, 12, activity.baseline_schedule.fixed_cost)
            bsch_worksheet.write(counter, 13, activity.get_total_cost())
            counter += 1

        workbook.close()