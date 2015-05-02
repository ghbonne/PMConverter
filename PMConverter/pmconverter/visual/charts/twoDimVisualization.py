__author__ = 'Eveline'

from pmconverter.visual.charts import Chart


class TwoDimChart(Chart):

    def __init__(self, type, subtype=None):
        self.type = type
        self.subtype = subtype

    def draw(self, workbook, worksheet, position, options=None, size=None):
        # Create a new chart object.
        if self.subtype:
            chart = workbook.add_chart({'type': self.type, 'subtype': self.subtype.value})
        else:
            chart = workbook.add_chart({'type': self.type})

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
        if self.type == "bar":
            chart.set_y_axis({'name': self.labels[1], 'reverse': True})
        else:
            chart.set_y_axis({'name': self.labels[1]})

        #set size of chart
        if size:
            chart.set_size(size)

        # insert chart
        if options:
            worksheet.insert_chart(position, chart, options)
        else:
            worksheet.insert_chart(position, chart)