from enum import Enum

__author__ = 'PM Group 8'


class DistributionType(Enum):
    MANUAL = 'manual'
    STANDARD = 'standard'


class ManualDistributionUnit(Enum):
    ABSOLUTE = 'absolute'
    RELATIVE = 'relative'


class StandardDistributionUnit(Enum):
    NO_RISK = 'no risk'
    SYMMETRIC = 'symmetric'
    SKEWED_LEFT = 'skewed left'
    SKEWED_RIGHT = 'skewed right'


class RiskAnalysisDistribution(object):
    """
    Risk Distribution of an Activity

    :var distribution_type: "manual" or "standard"
    :var distribution_units: "absolute" or "relative" if distribution_type="manual";
                             "no risk", "symmetric", "skewed left", "skewed right" if distribution_type="standard"
    :var optimistic_duration: int
    :var probable_duration: int
    :var pessimistic_duration: int
    """

    def __init__(self, distribution_type="manual", distribution_units="absolute", optimistic_duration=0,
                 probable_duration=0, pessimistic_duration=0, type_check=True):
        if type_check:
            if distribution_type == "manual":
                distribution_type = DistributionType.MANUAL
            if distribution_type == "standard":
                distribution_type = DistributionType.STANDARD
            if not isinstance(distribution_type, DistributionType):
                raise TypeError('distribution_type should be an element of the DistributionType enum')
            if distribution_type == DistributionType.MANUAL:
                if distribution_units == "relative":
                    distribution_units = ManualDistributionUnit.RELATIVE
                if distribution_units == "absolute":
                    distribution_units = ManualDistributionUnit.ABSOLUTE
                if not isinstance(distribution_units, ManualDistributionUnit):
                    raise TypeError('Since distribution_type = MANUAL, distribution_units should be an element of ManualDistributionUnit enum')
            if distribution_type == DistributionType.STANDARD:
                if distribution_units == "no risk":
                    distribution_units = StandardDistributionUnit.NO_RISK
                if distribution_units == "symmetric":
                    distribution_units = StandardDistributionUnit.SYMMETRIC
                if distribution_units == "skewed left":
                    distribution_units = StandardDistributionUnit.SKEWED_LEFT
                if distribution_units == "skewed right":
                    distribution_units = StandardDistributionUnit.SKEWED_RIGHT
                if not isinstance(distribution_units, StandardDistributionUnit):
                    raise TypeError('Since distribution_type = STANDARD, distribution_units should be an element of StandardDistributionUnit enum')
            if not isinstance(optimistic_duration, int):
                raise TypeError('optimistic_duration should be an integer')
            if not isinstance(probable_duration, int):
                raise TypeError('probable_duration should be an integer')
            if not isinstance(pessimistic_duration, int):
                raise TypeError('pessimistic_duration should be an integer')
        self.distribution_type = distribution_type
        self.distribution_units = distribution_units
        self.optimistic_duration = optimistic_duration
        self.probable_duration = probable_duration
        self.pessimistic_duration = pessimistic_duration




