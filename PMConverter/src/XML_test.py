# -*- coding: utf-8 -*-

from convert.XLSXparser import XLSXParser
from convert.XMLparser import XMLParser
from datetime import datetime
#todo: extra activity tracking record fields (should already be read), other possibilites Lagtype (FS,...)



file="project.xml"

## Parse XML File
xml_parser=XMLParser()
project_object=xml_parser.to_schedule_object(file)

## Parse from XLS
xlsx_parser = XLSXParser()
#wb2 = xlsx_parser.from_schedule_object(project_object, "test_alexander.xlsx", True)
#wb2.close()


## Write XML file
file_output="writeTest.xml"
xml_parser.from_schedule_object(project_object,file_output)


