__author__ = 'PM Group 8'


class RiskAnalysisDistribution(object):
    """
    Risk Distribution of an Activity

    :var distribution_type: "manual" or "standard"
    :var distribution_units: "absolute" or "relative" if distribution_type="manual"; "no risk", "symmetric",
                             "skewed left", "skewed right" if distribution_type="standard"
    :var optimistic_duration: int
    :var probable_duration: int
    :var pessimistic_duration: int
    """

    def __init__(self, distribution_type="manual", distribution_units="absolute", optimistic_duration=0,
                 probable_duration=0, pessimistic_duration=0):
        # TODO: Typechecking?
        if distribution_type != "manual" and distribution_type != "standard":
            raise TypeError()
        if ((distribution_type == "manual" and distribution_units != "absolute" and distribution_units != "relative")
            or (distribution_type == "standard" and distribution_units != "no risk"
                and distribution_units != "symmetric" and distribution_units != "skewed left"
                and distribution_units != "skewed right")):
            raise TypeError()
        self.distribution_type = distribution_type
        self.distribution_units = distribution_units
        self.optimistic_duration = optimistic_duration
        self.probable_duration = probable_duration
        self.pessimistic_duration = pessimistic_duration




