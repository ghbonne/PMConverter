__author__ = 'PM Group 8'

from visual.visualization import Visualization
import xlsxwriter


class LineChart(Visualization):

    """
    :var title: title of the chart
    :var labels: labels on the x and y-axis. Format: [x-axis,y-axis]
    :var data_series: data
        Format: [
                    [heading1:String, #first line
                        [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int], #x-values
                        [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int]  #y-values
                    ],
                    [heading2:String,
                        [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int], #x-values
                        [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int]  #y-values
                    ],
                    ...
                ]
    """

    def __init__(self, title, labels, data_series):
        self.title = title
        self.labels = labels
        self.data_series = data_series

    def visualize(self, workbook):
        """
        add a linechart visualisation to the Excel-file
        :param workbook: xlsxworkbook
        :return:
        """
        worksheet = workbook.add_worksheet()

        # Create a new chart objects.
        chart = workbook.add_chart({'type': 'line'})

        # Configure the series.
        for row in self.data_series:
            x_values = row[1]
            y_values = row[2]
            chart.add_series({
                'name':       row[0],
                'categories': x_values,
                'values':     y_values,
            })

        # Add a chart title and axis labels.
        chart.set_title({'name': self.title})
        chart.set_x_axis({'name': self.labels[0]})
        chart.set_y_axis({'name': self.labels[1]})

        # insert chart
        worksheet.insert_chart('A1', chart)