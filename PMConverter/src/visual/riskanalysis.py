__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.enums import LevelOfDetail, DataType
from visual.charts.barchart import BarChart
from objects.riskanalysisdistribution import DistributionType


class RiskAnalysis(Visualization):

    """
    :var level_of_detail

    data: naam, baseline duration, optimistic, most probable, pessimistic
    """

    def __init__(self):
        self.title = "Risk analysis"
        self.description = ""
        self.parameters = {"data_type": [DataType.ABSOLUTE, DataType.RELATIVE]}
        self.level_of_detail = None
        self.data_type = None

    def draw(self, workbook, worksheet, project_object):
        if not self.level_of_detail:
            raise Exception("Please first set var level_of_detail")

        self.calculate_values(workbook, worksheet, project_object)

        activities = project_object.activities
        names = "=("
        optimistic = "=("
        most_probable = "=("
        pessimistic = "=("
        height = 150  # calculate 20 pixels per element in barchart

        i = 0
        start = 3
        while i < len(activities):
            if len(activities[i].wbs_id) == 3:  # activity level
                if self.level_of_detail == LevelOfDetail.ACTIVITIES:
                    names += "'Risk Analysis'!$B$" + str(start+i) + ","
                    optimistic += "'Risk Analysis'!$W$" + str(start+i) + ","
                    most_probable += "'Risk Analysis'!$X$" + str(start+i) + ","
                    pessimistic += "'Risk Analysis'!$Y$" + str(start+i) + ","
                    height += 20
                i += 1
            else:  # not supported
                i += 1

        # remove last ';' and add ')'
        names = names[:-1] + ")"
        optimistic = optimistic[:-1] + ")"
        most_probable = most_probable[:-1] + ")"
        pessimistic = pessimistic[:-1] + ")"

        data_series = [
            ["Optimistic",
             names,
             optimistic
             ],
            ["Most probable",
             names,
             most_probable
             ],
            ["Pessimistic",
             names,
             pessimistic
             ]
        ]

        chart = BarChart(self.title, ["Hours", self.level_of_detail.value], data_series)

        options = {'height': height, 'width': 800}
        chart.draw(workbook, worksheet, 'I1', None, options)

    def calculate_values(self, workbook, worksheet, project_object):
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        worksheet.write('W2', 'Optimistic', header)
        worksheet.write('X2', 'Most probable', header)
        worksheet.write('Y2', 'Pessimistic', header)

        counter = 2
        for activity in project_object.activities:
            ra = activity.risk_analysis
            if ra:
                if ra.distribution_type == DistributionType.MANUAL and self.data_type == DataType.RELATIVE: #calculate relative values
                    dur = self.get_hours(activity.baseline_schedule.duration, project_object.agenda)
                    worksheet.write(counter, 22, int((ra.optimistic_duration/dur)*100), calculation)
                    worksheet.write(counter, 23, int((ra.probable_duration/dur)*100), calculation)
                    worksheet.write(counter, 24, int((ra.pessimistic_duration/dur)*100), calculation)
                elif ra.distribution_type == DistributionType.STANDARD and self.data_type == DataType.ABSOLUTE: #calculate absolute values
                    dur = self.get_hours(activity.baseline_schedule.duration, project_object.agenda)
                    worksheet.write(counter, 22, int((ra.optimistic_duration/100)*dur), calculation)
                    worksheet.write(counter, 23, int((ra.probable_duration/100)*dur), calculation)
                    worksheet.write(counter, 24, int((ra.pessimistic_duration/100)*dur), calculation)
                else:
                    worksheet.write(counter, 22, ra.optimistic_duration, calculation)
                    worksheet.write(counter, 23, ra.probable_duration, calculation)
                    worksheet.write(counter, 24, ra.pessimistic_duration, calculation)
            counter += 1



    def get_hours(self, delta, agenda):
        hours_in_a_day = agenda.get_working_hours_in_a_day()
        days_to_hour = delta.days * hours_in_a_day
        hours = delta.seconds / 3600
        return days_to_hour + hours





