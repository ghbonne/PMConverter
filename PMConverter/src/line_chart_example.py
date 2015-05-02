__author__ = 'Eveline'

import os

import xlsxwriter

from visual.charts.linechart import LineChart
from visual.columnchart import ColumnChart
from visual.columnchart import ColumnSubType
from visual.barchart import BarChart
from visual.barchart import BarSubType


file = os.path.join(os.path.dirname(__file__), 'output'+os.sep+'chart_test.xlsx')

workbook = xlsxwriter.Workbook(file)
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})

headings = ['Number', 'Batch 1', 'Batch 2']
data = [
    [2, 3, 4, 5, 6, 7],
    [10, 40, 50, 20, 10, 50],
    [30, 60, 70, 50, 40, 30],
]

worksheet.write_row('A1', headings, bold)
worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])
worksheet.write_column('C2', data[2])

data_series = [
    [headings[1],
     ['Sheet1', 1, 0, 6, 0],
     ['Sheet1', 1, 1, 6, 1]
     ],
    [headings[2],
     ['Sheet1', 1, 0, 6, 0],
     ['Sheet1', 1, 2, 6, 2]
     ]
]

labels = ['x', 'y']

chart = LineChart('test', labels, data_series)

chart.draw(workbook)

chart1 = ColumnChart('test', labels, data_series, ColumnSubType.stacked)

chart1.draw(workbook)

chart2 = BarChart('test', labels, data_series, BarSubType.stacked)

chart2.draw(workbook)

#Pie chart
# worksheet = workbook.add_worksheet()
# headings = ['Category', 'Values']
# data = [
#     ['Apple', 'Cherry', 'Pecan'],
#     [60, 30, 10],
# ]
#
# worksheet.write_row('A1', headings, bold)
# worksheet.write_column('A2', data[0])
# worksheet.write_column('B2', data[1])
#
#
# data_series2 = [
#     ['Sheet4', 1, 0, 3, 0],
#     ['Sheet4', 1, 1, 3, 1]
# ]
#
# chart2 = PieChart('PieChart', data_series2)
#
# chart2.visualize(workbook)


try:
    workbook.close()
except PermissionError:
    print("Permission denied. Please first close the excel file and try again.")

os.system("start excel.exe " + file)
