__author__ = 'Project Management Group 8 - 2015 Ghent University'


class FileParser(object):

    def __init__(self):
        pass

    def to_schedule_object(self, file_path_input):
        raise NotImplementedError("This method is not implemented!")

    def from_schedule_object(self, project_object, file_path_output):
        raise NotImplementedError("This method is not implemented!")