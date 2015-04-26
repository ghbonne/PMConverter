__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import LevelOfDetail, DataType
from visual.charts.barchart import BarChart


class ActualDuration(Visualization):
    """
    :var level_of_detail
    :var data_type

    data: baseline duration, actual duration, percentage completed
    """

    def __init__(self):
        self.title = "PD vs AD"
        self.description = ""
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES],
                           "data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}
        self.level_of_detail = None
        self.data_type = None
        self.tp = 0

    def set_tracking_period(self, tp):
        self.tp = tp

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
        baseline_duration = "=("
        actual_duration = "=("
        percentage_completed = "=("
        height = 150  # calculate 20 pixels per element in barchart

        i = 0
        start = 5
        while i < len(activities):
            if (len(activities[i].wbs_id) == 2  and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or (len(activities[i].wbs_id) == 3 and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                names += "'" + sh_name + "'!$B$" + str(start+i) + ","
                baseline_duration += "'" + sh_name + "'!$AA$" + str(start+i) + ","
                actual_duration += "'" + sh_name + "'!$AB$" + str(start+i) + ","
                percentage_completed += "'" + sh_name + "'!$AC$" + str(start+i) + ","
                height += 20
                i += 1
            else:  # 1 = project level, >3 not supported yet
                i += 1

        # remove last ';' and add ')'
        names = names[:-1] + ")"
        baseline_duration = baseline_duration[:-1] + ")"
        actual_duration = actual_duration[:-1] + ")"
        percentage_completed = percentage_completed[:-1] + ")"

        data_series = [
            ["Baseline duration",
             names,
             baseline_duration
             ],
            ["Actual duration",
             names,
             actual_duration
             ],
            ["Percentage completed",
             names,
             percentage_completed
             ]
        ]

        chart = BarChart(self.title, ["Hours", self.level_of_detail.value], data_series)

        options = {'height': height, 'width': 650}
        position = "B" + str(start + i + 1)
        chart.draw(workbook, worksheet, position, None, options)


    def calculate_values(self, workbook, worksheet, project_object, tp):
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        worksheet.merge_range('AA3:AC3', "Time", header)
        if self.data_type == DataType.RELATIVE:
            worksheet.write('AA4', 'Relative baseline duration', header)
            worksheet.write('AB4', 'Relative actual duration', header)
            worksheet.write('AC4', 'Percentage completed', header)
        else:
            worksheet.write('AA4', 'Absolute baseline duration', header)
            worksheet.write('AB4', 'Absolute actual duration', header)
            worksheet.write('AC4', 'Absolute completed', header)

        counter = 4

        for atr in project_object.tracking_periods[tp].tracking_period_records:  # atr = ActivityTrackingRecord
            if (len(atr.activity.wbs_id) == 2 and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or (len(atr.activity.wbs_id) == 3 and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                if self.data_type == DataType.RELATIVE:
                    # relative baseline duration
                    worksheet.write(counter, 26, 1, calculation)
                    # relative actual duration
                    bd = self.get_duration(atr.activity.baseline_schedule.duration)
                    ad = self.get_duration(atr.actual_duration)
                    if ad:
                        ad_rel = ad/bd
                        worksheet.write(counter, 27, ad_rel, calculation)
                    # percentage completed
                    worksheet.write(counter, 28, atr.percentage_completed/100, calculation)
                elif self.data_type == DataType.ABSOLUTE:
                    # absolute baseline duration (float)
                    pd = self.get_duration(atr.activity.baseline_schedule.duration)
                    worksheet.write(counter, 26, pd, calculation)
                    # absolute actual duration
                    ad = self.get_duration(atr.actual_duration)
                    worksheet.write(counter, 27, ad, calculation)
                    # absolute completed
                    percent = atr.percentage_completed
                    abs_completed = pd * percent / 100
                    worksheet.write(counter, 28, abs_completed, calculation)
            else:
                pass
            counter += 1


    @staticmethod
    def get_duration(delta):
        if delta:
            if delta.seconds != 0:
                days = delta.days
                hours = delta.seconds / (3600 * 24) #hours expressed in days
                duration = days + hours
            else:
                duration = delta.days
            return duration








