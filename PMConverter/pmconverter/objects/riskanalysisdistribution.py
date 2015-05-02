__author__ = 'Project Management Group 8 - 2015 Ghent University'

from enum import Enum


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

    def __init__(self, distribution_type=DistributionType.STANDARD, distribution_units=StandardDistributionUnit.NO_RISK, optimistic_duration=0,
                 probable_duration=0, pessimistic_duration=0, type_check=True):
        if type_check:
            if type(distribution_type) is str:
                if distribution_type.lower() == "manual":
                    distribution_type = DistributionType.MANUAL
                elif distribution_type.lower() == "standard":
                    distribution_type = DistributionType.STANDARD
            if not isinstance(distribution_type, DistributionType):
                raise TypeError('RiskAnalysisDistribution: distribution_type should be an element of the DistributionType enum')
            if distribution_type == DistributionType.MANUAL:
                if type(distribution_units) is str:
                    if distribution_units.lower() == "relative":
                        distribution_units = ManualDistributionUnit.RELATIVE
                    elif distribution_units.lower() == "absolute":
                        distribution_units = ManualDistributionUnit.ABSOLUTE
                    if not isinstance(distribution_units, ManualDistributionUnit):
                        raise TypeError('RiskAnalysisDistribution: Since distribution_type = MANUAL, distribution_units should be an element of '
                                        'ManualDistributionUnit enum')
            if distribution_type == DistributionType.STANDARD:
                if type(distribution_units) is str:
                    if distribution_units.lower() == "no risk":
                        distribution_units = StandardDistributionUnit.NO_RISK
                    elif distribution_units.lower() == "symmetric":
                        distribution_units = StandardDistributionUnit.SYMMETRIC
                    elif distribution_units.lower() == "skewed left":
                        distribution_units = StandardDistributionUnit.SKEWED_LEFT
                    elif distribution_units.lower() == "skewed right":
                        distribution_units = StandardDistributionUnit.SKEWED_RIGHT
                    if not isinstance(distribution_units, StandardDistributionUnit):
                        raise TypeError('RiskAnalysisDistribution: Since distribution_type = STANDARD, distribution_units should be an element of StandardDistributionUnit enum')
            if not isinstance(optimistic_duration, int):
                raise TypeError('RiskAnalysisDistribution: optimistic_duration should be an integer')
            if not isinstance(probable_duration, int):
                raise TypeError('RiskAnalysisDistribution: probable_duration should be an integer')
            if not isinstance(pessimistic_duration, int):
                raise TypeError('RiskAnalysisDistribution: pessimistic_duration should be an integer')
        self.distribution_type = distribution_type
        self.distribution_units = distribution_units
        self.optimistic_duration = optimistic_duration
        self.probable_duration = probable_duration
        self.pessimistic_duration = pessimistic_duration

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented


