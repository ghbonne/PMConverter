__author__ = 'ghbonne'
__license__ = "GPL"

from PyQt4.QtGui import *
from PyQt4.QtCore import *
from pmconverter.view.customWidgets.custom_QItemDelegate import Custom_QItemDelegate

class Custom_QComboBox(QComboBox):
    """description of class"""

    def __init__(self, parent= None):
        super(Custom_QComboBox, self).__init__(parent)
        self.customDelegate = Custom_QItemDelegate()
        self.setItemDelegate(self.customDelegate)

    def addSeperator(self):
        self.insertSeparator(self.count())

    def addParentItem(self, text):
        item = QStandardItem(text)
        item.setFlags(item.flags() & ~(Qt.ItemIsEnabled | Qt.ItemIsSelectable))
        item.setData("parent", Qt.AccessibleDescriptionRole)

        font = item.font()
        font.setBold(True)
        item.setFont(font)

        # append new item
        itemModel = self.model()
        itemModel.appendRow(item)

    def addChildItem(self, text):
        item = QStandardItem(text)
        item.setData("child", Qt.AccessibleDescriptionRole)

        # append new item
        itemModel = self.model()
        itemModel.appendRow(item)