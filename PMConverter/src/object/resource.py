__author__ = 'PM Group 8'


class Resource(object):
    """
    A type of resource described in the project

    :var resource_id: int
    :var name: String
    :var type: "Renewable" or "Consumable"
    :var availability: int, number of units available
    :var cost_use: float, one-time cost that is incurred every time that the resource is used by an activity
    :var cost_unit: float, cost per unit
    """
    "A type of resource described in the project"
    # instance variables:
    # _resourceId: unique number indicating a resource type
    # _resourceName: string indicating a resource type
    # _type
    # _availibility
    # _resourceCost: tuple (cost/use, cost/unit)
    # _resourceDemand: list of tuples (activityPointer, assigned resource units) describing the usage of this resource type in the project
    # _resourceTotalCost: number indicating total cost attributed to this resource
    # class variables:

    def __init__(self, resource_id, name="", type="", availability="", cost_use=0.0, cost_unit=0.0):
        pass


