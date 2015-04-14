__author__ = 'Eveline'

from visual.visualization import Visualization

class TwoDimVisualization(Visualization):

    def __init__(self, type, subtype=None):
        self.type = type
        self.subtype = subtype


    def visualize(self, workbook):
        worksheet = workbook.add_worksheet()#todo: debug, remove!

        # Create a new chart object.
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
        chart.set_y_axis({'name': self.labels[1]})

        # insert chart
        worksheet.insert_chart('A1', chart)