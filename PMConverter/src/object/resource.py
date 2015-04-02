__author__ = 'PM Group 8'


class Resource(object):
    """
    A type of resource described in the project

    :var resource_id: int
    :var name: String
    :var resource_type: "Renewable" or "Consumable"
    :var availability: int, number of units available
    :var cost_use: float, one-time cost that is incurred every time that the resource is used by an activity
    :var cost_unit: float, cost per unit
    """

    def __init__(self, resource_id, name="", resource_type="Renewable", availability="", cost_use=0.0, cost_unit=0.0):
        self.resource_id = resource_id
        self.name = name
        if resource_type != "Renewable" and resource_type != "Consumable":
            # TODO: maybe write own exception class?
            raise TypeError()
        self.resource_type = resource_type
        self.availability = availability
        self.cost_use = cost_use
        self.cost_unit = cost_unit


