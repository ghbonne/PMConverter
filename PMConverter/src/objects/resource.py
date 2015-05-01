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
    :var availability: int, number of units available OR float, percentage of resource available ELSE -1 = infinity for consumable resource 
    :var cost_use: float, one-time cost that is incurred every time that the resource is used by an activity
    :var cost_unit: float, cost per unit
    :var total_resource_cost: float or int, total value spent at this resource in the project  # TODO: implement 
    """

    def __init__(self, resource_id, name="", resource_type=ResourceType.RENEWABLE, availability=0, cost_use=0.0,
                 cost_unit=0.0, total_resource_cost=0, type_check = True):
        if type_check:
            if not isinstance(resource_id, int):
                raise TypeError('Resource: resource_id should be an integer')
            if not isinstance(name, str):
                raise TypeError('Resource: name should be an string')
            if type(resource_type) is str:
                if resource_type.lower() == 'RENEWABLE'.lower():
                    resource_type = ResourceType.RENEWABLE
                elif resource_type.lower() == 'CONSUMABLE'.lower():
                    resource_type = ResourceType.CONSUMABLE
            if not isinstance(resource_type, ResourceType):
                raise TypeError('Resource: resource_type should be an element of ResourceType enum')
            if not isinstance(availability, (float, int)):
                raise TypeError('Resource: availability should be a float or integer')
            if not isinstance(cost_use, float):
                raise TypeError('Resource: cost_use should be an float')
            if not isinstance(cost_unit, float):
                raise TypeError('Resource: cost_unit should be an float')
            if not isinstance(total_resource_cost, (float, int)):
                raise TypeError('Resource: total_resource_cost should be an integer or float')

        self.resource_id = resource_id
        self.name = name
        self.resource_type = resource_type
        self.availability = availability
        self.cost_use = cost_use
        self.cost_unit = cost_unit
        self.total_resource_cost = total_resource_cost

    @staticmethod
    def calculate_resource_assignment_cost(resource, demand, isFixed, duration_hours):
        """This function calculates the cost of a resource assignment for a given demand and duration in workinghours.
        :returns: float, resource assignment cost
        """
        cost = 0.0
        if resource.resource_type == ResourceType.CONSUMABLE:
            ## only add once the cost for its use!
            # check if fixed resource assignment:
            if isFixed:
                # fixed resource assignment => variable cost is not multiplied by activity duration:
                cost += resource.cost_use + demand * resource.cost_unit
            else:
                # non fixed resource assignment:
                cost += resource.cost_use + demand * resource.cost_unit * duration_hours
        else:
            #resource type is renewable:
            #add cost_use and variable cost:
            cost += demand * (resource.cost_use + resource.cost_unit * duration_hours)

        return cost

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return NotImplemented

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return not self.__eq__(other)
        return NotImplemented
