__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import DataType, LevelOfDetail
from visual.charts.barchart import BarChart
from objects.activity import Activity


class ActualCost(Visualization):
    """
    Implements drawings for actual cost (type = Bar chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user

    Settings:
    :var level_of_detail: LevelOfDetail, graph can be shown for workpackages or activities
    :var data_type: DataType, values expressed absolute (euro) or relative (%)
    """

    def __init__(self):
        self.title = "BC vs AC"
        self.description = "Bar chart made for each tracking period that compares the work packages\' or activities\' actual cost with their baseline cost, expressed in monetary units or as relative percentages of the baseline costs. "\
                            +"The percentage completed of each activity at the considered tracking period is also shown."
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES],
                           "data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}
        self.level_of_detail = None
        self.data_type = None
        self.tp = 0

    def draw(self, workbook, worksheet, project_object):
        if not self.level_of_detail:
            raise Exception("Please first set var level_of_detail")
        if not self.data_type:
            raise Exception("Please first set var data_type")

        # first calculate some values before drawing
        self.calculate_values(workbook, worksheet, project_object, self.tp)

        activities = project_object.activities
        sh_name = worksheet.get_name()
        names = "=("
        baseline_cost = "=("
        actual_cost = "=("
        percentage_completed = "=("
        height = 150  # calculate 20 pixels per element in barchart

        i = 1
        start = 5
        while i < len(activities):
            isActivityGroup = Activity.is_not_lowest_level_activity(activities[i], activities)
            if (isActivityGroup and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or ((not isActivityGroup) and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                names += "'" + sh_name + "'!$B$" + str(start+i) + ","
                baseline_cost += "'" + sh_name + "'!$AD$" + str(start+i) + ","
                actual_cost += "'" + sh_name + "'!$AE$" + str(start+i) + ","
                percentage_completed += "'" + sh_name + "'!$AF$" + str(start+i) + ","
                height += 20
                i += 1
            else:
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

        if self.data_type == DataType.ABSOLUTE:
            chartAxisLabels = ["Cost (Euro)", self.level_of_detail.value]
        elif self.data_type == DataType.RELATIVE:
            chartAxisLabels = ["Cost relative to baseline schedule (%)", self.level_of_detail.value]
        chart = BarChart(self.title, chartAxisLabels, data_series)

        options = {'height': height, 'width': 650}
        position = "K" + str(start + i + 1)
        chart.draw(workbook, worksheet, position, None, options)

    """
    Private methods
    """
    def calculate_values(self, workbook, worksheet, project_object, tp):
        """

        :param workbook: Workbook
        :param worksheet: Worksheet
        :param project_object: ProjectObject
        :param tp: int, value of tracking period
        """
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        if self.data_type == DataType.RELATIVE:
            worksheet.merge_range('AD3:AF3', "Costs", header)
            worksheet.write('AD4', 'Relative baseline cost', header)
            worksheet.write('AE4', 'Relative actual cost', header)
            worksheet.write('AF4', 'Percentage completed', header)
        else:
            worksheet.merge_range('AD3:AF3', "Costs", header)
            worksheet.write('AD4', 'Baseline cost', header)
            worksheet.write('AE4', 'Actual cost', header)
            worksheet.write('AF4', 'Absolute percentage completed', header)

        counter = 4

        for atr in project_object.tracking_periods[tp].tracking_period_records:  # atr = ActivityTrackingRecord
            if (atr.activity.baseline_schedule.var_cost is None and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or (atr.activity.baseline_schedule.var_cost is not None and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                if self.data_type == DataType.RELATIVE:
                    # relative baseline cost
                    worksheet.write_number(counter, 29, 100, calculation)
                    # relative actual cost
                    bc = atr.activity.baseline_schedule.total_cost
                    ac = atr.actual_cost
                    if ac and bool(bc):
                        ac_rel = ac/bc * 100
                        worksheet.write_number(counter, 30, ac_rel, calculation)
                    # percentage completed
                    worksheet.write_number(counter, 31, atr.percentage_completed, calculation)
                elif self.data_type == DataType.ABSOLUTE:
                    # absolute baseline cost
                    worksheet.write_number(counter, 29, atr.activity.baseline_schedule.total_cost, calculation)
                    # absolute actual cost
                    ac = atr.actual_cost
                    worksheet.write_number(counter, 30, ac, calculation)
                    # absolute completed
                    percent = atr.percentage_completed
                    bc = atr.activity.baseline_schedule.total_cost
                    abs_completed = bc * percent / 100.0
                    worksheet.write_number(counter, 31, abs_completed, calculation)
            else:
                pass
            counter += 1
