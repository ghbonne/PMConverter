__author__ = 'Eveline'

import xlsxwriter
import os
from visual.linechart import LineChart

file = 'output'+os.sep+'chart_test.xlsx'

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

chart.visualize(workbook)

try:
    workbook.close()
except PermissionError:
    print("Permission denied. Please first close the excel file and try again.")

os.system("start excel.exe " + file)

