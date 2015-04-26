__author__ = 'Eveline'
from enum import Enum


class XAxis(Enum):
    TRACKING_PERIOD = "Tracking period"
    DATE = "Date"


class DataType(Enum):
    ABSOLUTE = "absolute"
    RELATIVE = "relative"


class LevelOfDetail(Enum):
    WORK_PACKAGES = "work packages (super level)"
    ACTIVITIES = "activities (detailed level)"


class ExcelVersion(Enum):
    EXTENDED = "extended"
    BASIC = "basic"