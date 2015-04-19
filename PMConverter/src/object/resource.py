from enum import Enum

__author__ = 'PM Group 8'


class ResourceType(Enum):
    RENEWABLE = 'Renewable'
    CONSUMABLE = 'Consumable'


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

    def __init__(self, resource_id, name="", resource_type=ResourceType.RENEWABLE, availability="", cost_use=0.0, cost_unit=0.0, type_check = True):
        if type_check:
            if not isinstance(resource_id, int):
                raise TypeError('resource_id should be an integer')
            if not isinstance(name, str):
                raise TypeError('name should be an string')
            if not isinstance(resource_type, ResourceType):
                raise TypeError('resource_type should be an element of ResourceType enum')
            if not isinstance(availability, int):
                raise TypeError('availability should be an integer')
            if not isinstance(cost_use, float):
                raise TypeError('cost_use should be an float')
            if not isinstance(cost_unit, float):
                raise TypeError('cost_unit should be an float')

        self.resource_id = resource_id
        self.name = name
        self.resource_type = resource_type
        self.availability = availability
        self.cost_use = cost_use
        self.cost_unit = cost_unit
        self.total_cost = total_cost


