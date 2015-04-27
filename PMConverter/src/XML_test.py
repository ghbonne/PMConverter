# -*- coding: utf-8 -*-

from convert.XLSXparser import XLSXParser
from convert.XMLparser import XMLParser
import os
from datetime import datetime
from src.visual.enums import ExcelVersion


file="C2012-12 Railway Station Ghent.p2x"

## Parse XML File
xml_parser=XMLParser()
project_object=xml_parser.to_schedule_object(file)

## Parse from XLS
xlsx_parser = XLSXParser()
wb2 = xlsx_parser.from_schedule_object(project_object, "test_alexander.xlsx", ExcelVersion.EXTENDED)
wb2.close()

for activity in project_object.activities:
    if activity.activity_id == 5:
        print(activity.baseline_schedule.start)


#po = xlsx_parser.to_schedule_object(os.path.join(os.path.dirname(__file__),
 #                                           "../administration/2_Project data input sheet_extended.xlsx"))


## Write XML file
file_output_XML="writeTest_XML.xml"
xml_parser.from_schedule_object(project_object, file_output_XML)

file_output_xls="writeTest_Excell.p2x"
#xml_parser.from_schedule_object(po, file_output_xls)

#TODO First 2 rows Tracking Period are missing
#TODO Dateformat TrackingPeriod
#TODO PRC, PAC, AC, RC, EV PV (TrackingPeriod)

