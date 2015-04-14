__author__ = 'gilles'


class ParseError(Exception):
    def __init__(self, message="An error occurred while parsing!"):
        super(ParseError, self).__init__(message)


class XLSXParseError(ParseError):
    def __init__(self, message="An error occurred while parsing the XLSX file!"):
        super(XLSXParseError, self).__init__(message)


class XMLParseError(ParseError):
    def __init__(self, message="An error occurred while parsing the XML file!"):
        super(XMLParseError, self).__init__(message)
