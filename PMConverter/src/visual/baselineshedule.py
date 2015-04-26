__author__ = 'Eveline'
from visual.visualization import Visualization
from visual.charts.granttchart import GanttChart
from visual.enums import ExcelVersion


class BaselineSchedule(Visualization):
    """
    Implements drawings for baseline schedule (type = gannt chart)
    """

    def __init__(self):
        self.title = "Baseline Schedule"
        self.description = "Baseline Schedule is a gannt chart where every activity / workpackage is shown." \
                           +"In this chart you can clearly see when activities are planned"
        self.parameters = {}
        self.support = [ExcelVersion.EXTENDED, ExcelVersion.BASIC]

    def draw(self, workbook, worksheet, project_object, excel_version):
        size = self.calculate_values(workbook, worksheet, project_object)

        chartsheet = workbook.add_worksheet("Gantt chart")
        self.change_order(workbook)

        names = "='" + worksheet.get_name() + "'!$B$3:$B$" + str(size)
        if excel_version == ExcelVersion.EXTENDED:
            baseline_start = "='" + worksheet.get_name() + "'!$F$3:$F$" + str(size)
        else:
            baseline_start = "='" + worksheet.get_name() + "'!$E$3:$F$" + str(size)
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

    """
    Private methods
    """
    def calculate_values(self, workbook, worksheet, project_object):
        """
        Write duration expressed in days to column Q
        :param workbook: Workbook
        :param worksheet: Worksheet
        :param project_object: ProjectObject
        """
        header = workbook.add_format({'bold': True, 'bg_color': '#316AC5', 'font_color': 'white', 'text_wrap': True,
                                      'border': 1, 'font_size': 8})
        calculation = workbook.add_format({'bg_color': '#FFF2CC', 'text_wrap': True, 'border': 1, 'font_size': 8})

        worksheet.write('Q2', 'Baseline duration (in days)', header)

        counter = 2

        for activity in project_object.activities:
            begin_date = activity.baseline_schedule.start
            end_date = activity.baseline_schedule.end
            delta = self.get_duration(end_date-begin_date)
            worksheet.write(counter, 16, delta, calculation)
            counter += 1

        return counter

    def get_max_date(self, project_object):
        """
        Calculated the latest date
        :param project_object: ProjectObject
        :return: datetime
        """
        max = project_object.activities[0].baseline_schedule.end
        for activity in project_object.activities:
            end_date = activity.baseline_schedule.end
            if end_date > max:
                max = end_date
        return max


    @staticmethod
    def get_duration(delta):
        """
        Convert from timedelta to a number, this number is the time expressed in days
        :param delta: timedelta
        :return: int
        """
        if delta:
            if delta.seconds != 0:
                days = delta.days
                hours = delta.seconds / (3600 * 24) # hours expressed in days
                duration = days + hours
            else:
                duration = delta.days
            return duration

    def change_order(self, workbook):
        """
        Changes the order of the worksheets such that the Gantt chart is de second worksheet
        :param workbook:
        """
        worksheets = workbook.worksheets_objs
        bs_chart_sheet = worksheets[-1]
        worksheets.insert(1, bs_chart_sheet)
        worksheets = worksheets[:-1]
        workbook.worksheets_objs = worksheets