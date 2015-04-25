__author__ = 'PM Group 8'

from visual.charts.chart import Chart


class PieChart(Chart):
    """
    :var title: title of the chart
    :var data_series: data
        Format: [
                    [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int], #labels
                    [Sheet: String, first_row: int, first_col: int, last_row: int, last_col: int]  #values
                ]
    """

    def __init__(self, title, data_series, relative=True):
        self.title = title
        self.data_series = data_series
        self.relative = relative

    def draw(self, workbook, worksheet, location, options=None, size=None):
        """
        add a piechart to the excel workbook
        :param workbook: xlsxworkbook
        :return:
        """
        #worksheet = workbook.add_worksheet()

        # Create a new chart object
        chart = workbook.add_chart({'type': 'pie'})

        # Configure the series.
        labels = self.data_series[0]
        values = self.data_series[1]
        chart.add_series({
            'name':       "data",
            'categories': labels,
            'values':     values,
            'data_labels': {'value': not self.relative, 'percentage': self.relative},
        })

        # Add a chart title and axis labels.
        chart.set_title({'name': self.title})

        # Set chart style
        chart.set_style(10)

        #set size of chart
        if size:
            chart.set_size(size)

        # Insert chart
        worksheet.insert_chart(location, chart, options)