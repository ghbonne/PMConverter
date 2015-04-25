__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.charts.granttchart import GanttChart


class BaselineSchedule(Visualization):
    """
    type gantt chart
    bereik? sheet: baseline schedule, kolom B
            baseline timing

    """

    def __init__(self):
        self.title = "Baseline Schedule"
        self.description = ""
        self.parameters = None

    def draw(self, workbook, worksheet, project_object):
        size = self.calculate_values(workbook, worksheet, project_object)

        chartsheet = workbook.add_worksheet("Gantt chart")

        names = "='" + worksheet.get_name() + "'!$B$3:$B$" + str(size)
        baseline_start = "='" + worksheet.get_name() + "'!$F$3:$F$" + str(size)
        baseline_duration = "='" + worksheet.get_name() + "'!$Q$3:$Q$" + str(size)

        data_series = [
            ["Baseline start",
             names,
             baseline_start,
             ],
            ["Actual duration",
             names,
             baseline_duration,
             ]
        ]

        min_date = project_object.activities[0].baseline_schedule.start
        max_date = self.get_max_date(project_object)
        chart = GanttChart(self.title, ["Date", ""], data_series, min_date, max_date)

        options = {'height': 150 + size*20, 'width': 1000}
        chart.draw(workbook, chartsheet, options)

    def calculate_values(self, workbook, worksheet, project_object):
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        counter = 2

        for activity in project_object.activities:
            begin_date = activity.baseline_schedule.start
            end_date = activity.baseline_schedule.end
            delta = self.get_duration(end_date-begin_date)
            worksheet.write(counter, 16, delta, calculation)
            counter += 1

        return counter

    def get_max_date(self, project_object):
        max = project_object.activities[0].baseline_schedule.end
        for activity in project_object.activities:
            end_date = activity.baseline_schedule.end
            if end_date > max:
                max = end_date
        return max


    @staticmethod
    def get_duration(delta):
        if delta:
            if delta.seconds != 0:
                days = delta.days
                hours = delta.seconds / (3600 * 24)
                duration = days + hours
            else:
                duration = delta.days
            return duration