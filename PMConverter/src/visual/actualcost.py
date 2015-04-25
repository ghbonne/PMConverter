__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import DataType, LevelOfDetail
from visual.charts.barchart import BarChart


class ActualCost(Visualization):
    """
    :var level_of_detail
    :var data_type
    """

    def __init__(self):
        self.title = "PC vs AC"
        self.description = ""
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES],
                           "data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}

    def draw(self, workbook, worksheet, project_object, tp):
        if not self.level_of_detail:
            raise Exception("Please first set var level_of_detail")
        if not self.data_type:
            raise Exception("Please first set var data_type")

        # first calculate some values before drawing
        self.calculate_values(workbook, worksheet, project_object, tp)

        activities = project_object.activities
        sh_name = worksheet.get_name()
        names = "=("
        baseline_cost = "=("
        actual_cost = "=("
        percentage_completed = "=("
        height = 150  # calculate 20 pixels per element in barchart

        i = 0
        start = 5
        while i < len(activities):
            if self.data_type == DataType.ABSOLUTE:
                if (len(activities[i].wbs_id) == 2 and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or (len(activities[i].wbs_id) == 3 and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                    names += "'" + sh_name + "'!$B$" + str(start+i) + ","
                    baseline_cost += "'" + sh_name + "'!$K$" + str(start+i) + ","
                    actual_cost += "'" + sh_name + "'!$S$" + str(start+i) + ","
                    percentage_completed += "'" + sh_name + "'!$AD$" + str(start+i) + ","
                    height += 20
                    i += 1
                else:  # 1 = project level, >3 not supported yet
                    i += 1
            elif self.data_type == DataType.RELATIVE:
                if (len(activities[i].wbs_id) == 2 and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or (len(activities[i].wbs_id) == 3 and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                    names += "'" + sh_name + "'!$B$" + str(start+i) + ","
                    baseline_cost += "'" + sh_name + "'!$AD$" + str(start+i) + ","
                    actual_cost += "'" + sh_name + "'!$AE$" + str(start+i) + ","
                    percentage_completed += "'" + sh_name + "'!$AF$" + str(start+i) + ","
                    height += 20
                    i += 1
                else:  # 1 = project level, >3 not supported yet
                    i += 1

        # remove last ';' and add ')'
        names = names[:-1] + ")"
        baseline_cost = baseline_cost[:-1] + ")"
        actual_cost = actual_cost[:-1] + ")"
        percentage_completed = percentage_completed[:-1] + ")"

        data_series = [
            ["Baseline cost",
             names,
             baseline_cost
             ],
            ["Actual cost",
             names,
             actual_cost
             ],
            ["Percentage completed",
             names,
             percentage_completed
             ]
        ]

        chart = BarChart(self.title, ["Euro", self.level_of_detail.value], data_series)

        options = {'height': height, 'width': 650}
        position = "K" + str(start + i + 1)
        chart.draw(workbook, worksheet, position, None, options)


    def calculate_values(self, workbook, worksheet, project_object, tp):
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        if self.data_type == DataType.RELATIVE:
            worksheet.merge_range('AD3:AF3', "Costs", header)
            worksheet.write('AD4', 'Relative baseline duration', header)
            worksheet.write('AE4', 'Relative actual duration', header)
            worksheet.write('AF4', 'Percentage completed', header)
        else:
            worksheet.write('AD3', 'Costs', header)
            worksheet.write('AD4', 'Absolute completed', header)

        counter = 4

        for atr in project_object.tracking_periods[tp].tracking_period_records:  # atr = ActivityTrackingRecord
            if (len(atr.activity.wbs_id) == 2 and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or (len(atr.activity.wbs_id) == 3 and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                if self.data_type == DataType.RELATIVE:
                    # relative baseline duration
                    worksheet.write(counter, 29, 1, calculation)
                    # relative actual duration
                    bc = atr.activity.baseline_schedule.total_cost
                    ac = atr.actual_cost
                    if ac:
                        ac_rel = ac/bc
                        worksheet.write(counter, 30, ac_rel, calculation)
                    # percentage completed
                    worksheet.write(counter, 31, atr.percentage_completed/100, calculation)
                elif self.data_type == DataType.ABSOLUTE:
                    pc = atr.actual_cost
                    # absolute completed
                    percent = atr.percentage_completed
                    abs_completed = pc * percent / 100
                    worksheet.write(counter, 29, abs_completed, calculation)
            else:
                pass
            counter += 1