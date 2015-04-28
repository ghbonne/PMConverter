__author__ = 'ghbonne'
__license__ = "GPL"

from PyQt4.QtGui import *
from PyQt4.QtCore import *

class Custom_QItemDelegate(QItemDelegate):
    """Custom delegate for QComboBox to make grouped items"""

    def __init__(self, parent= None):
        super(Custom_QItemDelegate, self).__init__(parent)

    def paint(self, painter, option, index):
        type = str(index.data(Qt.AccessibleDescriptionRole))
        if type == "parent":
            parentOption = option
            parentOption.state |= Qt.ItemIsEnabled
            super(Custom_QItemDelegate, self).paint(painter, parentOption, index)
        elif type == "child":
            childOption = option
            indent = option.fontMetrics.width("    ")
            childOption.rect.adjust(indent, 0,0,0)
            childOption.textElideMode = Qt.ElideNone
            super(Custom_QItemDelegate, self).paint(painter, childOption, index)
        else:
            super(Custom_QItemDelegate, self).paint(painter, option, index)
