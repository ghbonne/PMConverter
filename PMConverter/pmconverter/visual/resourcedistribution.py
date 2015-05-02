__author__ = 'Eveline'

from pmconverter.visual import Visualization, DataType, ExcelVersion
from pmconverter.visual.charts import PieChart


class ResourceDistribution(Visualization):

    """
    Implements drawings for resource distribution (type = Pie chart)

    Common:
    :var title: str, title of the graph
    :var description, str description of the graph
    :var parameters: dict, the present keys indicate which parameters should be available for the user
    :var supported: list of ExcelVersion, containing the version that are supported

    Settings:
    :var data_type: DataType, labels expressed in absolute(euro) or relative(%) values
    """

    def __init__(self):
        self.title = "Resource costs"
        self.description = ""
        self.parameters = {'data_type': [DataType.ABSOLUTE, DataType.RELATIVE]}
        self.data_type = None
        self.support = [ExcelVersion.EXTENDED]

    def draw(self, workbook, worksheet, project_object, excel_version=ExcelVersion.EXTENDED):
        if not self.data_type:
            raise Exception("Please first set var data_type")

        res_size = len(project_object.resources)

        data_series = [
            ['Resources', 2, 1, (1+res_size), 1],
            ['Resources', 2, 7, (1+res_size), 7]
        ]

        if self.data_type == DataType.RELATIVE:
            relative = True
        else:
            relative = False

        chart2 = PieChart('Resources', data_series, relative)
        size = {'height': 150 + res_size*20}
        options = {'x_offset': 25, 'y_offset': 10}
        chart2.draw(workbook, worksheet, 'I1', options, size)
