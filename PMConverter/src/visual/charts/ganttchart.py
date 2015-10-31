__author__ = 'Eveline'
from visual.charts.chart import Chart


class GanttChart(Chart):
    def __init__(self, title, labels, data_series, isWorkpackage_list, min_date, max_date):
        self.title = title
        self.labels = labels
        self.data_series = data_series
        self.isWorkpackage_list = isWorkpackage_list
        self.min_date = min_date
        self.max_date = max_date


    def draw(self, workbook, worksheet, size={}):
        """
        add a barchart visualisation to the Excel-file
        :param workbook: xlsxworkbook
        :return:
        """
        # Create a new chart object.
        chart = workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})

        # Define the formatting of the Workpackages:
        workpackage_format = {'fill': {'color': '#4F81BD'}}
        activity_format = {'fill': {'color': '#C0504D'}}

        # Configure the series.
        for row in self.data_series:
            x_values = row[1]
            y_values = row[2]
            data = {
                'name':       row[0],
                'categories': x_values,
                'values':     y_values,
            }
            if row[0] == "Baseline start":
                data['fill'] = {'none': True}
            else:
                # format the workpackages:        
                data['points'] = [workpackage_format if x == True else activity_format for x in self.isWorkpackage_list]
            #endIf formatting bars

            chart.add_series(data)

        # Add a chart title and axis labels.
        chart.set_title({'name': self.title})
        chart.set_x_axis({
            'name': self.labels[0],
            'date_axis': True,
            'min': self.min_date,
            'max': self.max_date,
            'num_format': 'dd/mm/yyyy',
        })
        chart.set_y_axis({
            'name': self.labels[1],
            'reverse': True,
            'interval_unit': 1
        })
        chart.set_legend({'none': True})

        #set size of chart
        if size:
            chart.set_size(size)

        worksheet.insert_chart('A1', chart)