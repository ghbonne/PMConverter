__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import LevelOfDetail, DataType, ExcelVersion
from visual.charts.barchart import BarChart
from objects.activity import Activity


class ActualDuration(Visualization):
    """
    Implements drawings for actual duration (type = Bar chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user
    :var supported: list of ExcelVersion, containing the version that are supported

    Settings:
    :var level_of_detail: LevelOfDetail, graph can be shown for workpackages or activities
    :var data_type: DataType, values expressed absolute (hours) or relative (%)
    """

    def __init__(self):
        self.title = "BD vs AD"
        self.description = "A bar chart is generated on every tracking period tab indicating for each work package or activity its actual duration w.r.t. its baseline duration. "\
                            +"Also the percentage completed of the tasks at that tracking period moment is indicated."
        self.parameters = {"level_of_detail": [LevelOfDetail.WORK_PACKAGES, LevelOfDetail.ACTIVITIES],
                           "data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}
        self.level_of_detail = None
        self.data_type = None
        self.tp = 0
        self.support = [ExcelVersion.EXTENDED, ExcelVersion.BASIC]

    def draw(self, workbook, worksheet, project_object, excel_version):
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

        i = 1
        start = 5
        while i < len(activities):
            isActivityGroup = Activity.is_not_lowest_level_activity(activities[i], activities)
            if (isActivityGroup  and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or ((not isActivityGroup) and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                names += "'" + sh_name + "'!$B$" + str(start+i) + ","
                baseline_duration += "'" + sh_name + "'!$AA$" + str(start+i) + ","
                actual_duration += "'" + sh_name + "'!$AB$" + str(start+i) + ","
                percentage_completed += "'" + sh_name + "'!$AC$" + str(start+i) + ","
                height += 20
                i += 1
            else:
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

        if self.data_type == DataType.ABSOLUTE:
            chartAxisLabels = ["Duration (Hours)", self.level_of_detail.value]
        elif self.data_type == DataType.RELATIVE:
            chartAxisLabels = ["Duration relative to baseline duration (%)", self.level_of_detail.value]
        chart = BarChart(self.title, chartAxisLabels, data_series)

        options = {'height': height, 'width': 650}
        position = "B" + str(start + i + 1)
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
            if (atr.activity.baseline_schedule.var_cost is None and self.level_of_detail == LevelOfDetail.WORK_PACKAGES) or (atr.activity.baseline_schedule.var_cost is not None and self.level_of_detail == LevelOfDetail.ACTIVITIES):
                if self.data_type == DataType.RELATIVE:
                    # relative baseline duration
                    worksheet.write(counter, 26, 100, calculation)
                    # relative actual duration
                    bd = self.get_duration(atr.activity.baseline_schedule.duration, project_object.agenda)
                    ad = self.get_duration(atr.actual_duration, project_object.agenda)
                    if ad and bd:
                        ad_rel = ad/bd * 100
                        worksheet.write(counter, 27, ad_rel, calculation)
                    # percentage completed
                    worksheet.write(counter, 28, atr.percentage_completed, calculation)
                elif self.data_type == DataType.ABSOLUTE:
                    # absolute baseline duration (float)
                    pd = self.get_duration(atr.activity.baseline_schedule.duration, project_object.agenda)
                    worksheet.write(counter, 26, pd, calculation)
                    # absolute actual duration
                    ad = self.get_duration(atr.actual_duration, project_object.agenda)
                    worksheet.write(counter, 27, ad, calculation)
                    # absolute completed
                    percent = atr.percentage_completed
                    abs_completed = pd * percent / 100.0
                    worksheet.write(counter, 28, abs_completed, calculation)
            else:
                pass
            counter += 1

    @staticmethod
    def get_duration(delta, agenda):
        """
        Convert from timedelta to a number, this number is the time expressed in working days and excess working hours are added relative to a workingday.
        :param delta: timedelta
        :return: float
        """
        if delta:
            if delta.seconds != 0:
                days = delta.days
                hours = delta.seconds / (3600.0 * agenda.get_working_hours_in_a_day()) #hours expressed in workingdays
                duration = days + hours
            else:
                duration = delta.days
            return duration
        #else no delta:
        return 0.0








