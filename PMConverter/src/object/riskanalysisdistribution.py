__author__ = 'PM Group 8'


class RiskAnalysisDistribution(object):
    """
    Risk Distribution of an Activity

    :var distribution_type: "Manual" or "Standard"
    :var distribution_units: "Absolute" or "Relative"
    :var optimistic_duration: int
    :var probable_duration: int
    :var pessimistic_duration: int
    """

    def __init__(self, distribution_type="Manual", distribution_units="Absolute", optimistic_duration=0,
                 probable_duration=0, pessimistic_duration=0):
        if distribution_type != "Manual" and distribution_type != "Standard":
            # TODO: maybe write own exception class?
            raise TypeError()
        if distribution_units != "Absolute" and distribution_units != "Relative":
            # TODO: maybe write own exception class?
            raise TypeError()
        self.distribution_type = distribution_type
        self.distribution_units = distribution_units
        self.optimistic_duration = optimistic_duration
        self.probable_duration = probable_duration
        self.pessimistic_duration = pessimistic_duration


