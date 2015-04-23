__author__ = 'PM Group 8'

#from processor.processor import Processor
from PyQt4.QtCore import *
from PyQt4.QtGui import *
import ui_UIView
import sys
import ntpath
import os

class UIView(QDialog, ui_UIView.Ui_UIView):
    """This is the main dialog window of the GUI controlling the PMConverter application."""

    def __init__(self, parent= None):
        super(UIView, self).__init__(parent)
        # init generated GUI
        self.setupUi(self)

        # custom GUI inits
        self.inputFiletypes = {"Excel": ".xlsx", "ProTrack": ".p2x"}

        # custom Application inits
        self.inputFilename = ""


    @pyqtSlot("bool")
    def on_cmdStep0_Start_clicked(self, clicked):
        "This function handles clicked events on the start button of the GUI"
        self.btnStep1_InputFile.setDefault(True)
        self.chkStep1_ExcelExtendedVersion.setVisible(False)

        self.pagesMain.setCurrentIndex(1)

    @pyqtSlot("bool")
    def on_btnStep1_InputFile_clicked(self, clicked):
        "This function handles clicked events on the input file button of Step 1"
        fileWindow = QFileDialog(parent = self, caption = "Choose an input file...")
        fileWindow.setFileMode(QFileDialog.ExistingFile)
        fileWindow.setNameFilter("ProTrack files (*.p2x);;Excel files (*.xlsx)")
        fileWindow.setViewMode(QFileDialog.List)

        if(fileWindow.exec_()):
            self.inputFilename = fileWindow.selectedFiles()[0]
            filenamePath, fileExtension = os.path.splitext(self.inputFilename)
            filename = ntpath.basename(filenamePath)
            self.lineEditStep1_InputFilename.setText(filename)
            width = self.lineEditStep1_InputFilename.fontMetrics().boundingRect(self.lineEditStep1_InputFilename.text()).width()
            if width > 300:
                self.lineEditStep1_InputFilename.setFixedWidth(width + 10 if width < self.width() - 200 else self.width() - 200)
            else:
                self.lineEditStep1_InputFilename.setFixedWidth(300)

        
            # if a file is chosen: enable conversion options
            self.ddlStep1_InputFormat.clear()
            self.ddlStep1_InputFormat.addItems([type for type in self.inputFiletypes.keys() if self.inputFiletypes[type] == fileExtension])
            self.ddlStep1_InputFormat.setEnabled(True)

            self.ddlStep1_OutputFormat.clear()
            self.ddlStep1_OutputFormat.addItems([type for type in self.inputFiletypes.keys() if self.inputFiletypes[type] != fileExtension])
            self.ddlStep1_OutputFormat.setEnabled(True)

            #if output file format == Excel => enable chkbox
            self.chkStep1_ExcelExtendedVersion.setVisible(self.ddlStep1_OutputFormat.currentText() == "Excel")
            self.chkStep1_ExcelExtendedVersion.setEnabled(self.ddlStep1_OutputFormat.currentText() == "Excel")

            self.cmdStep1_Next.setEnabled(True)

            #inputFiletype = filename.ending(

            #self.lineEditStep1_InputFilename.setSizePolicy(QSizePolicy.MinimumExpanding)
            #self.ddlStep1_InputFormat.setFocus()
            #print("Selected: {0}".format(self.inputFilename))

        @pyqtSlot("int")
        def on_ddlStep1_OutputFormat_currentIndexChanged(self, index):
            "This function handles index changed events from combobox outputFormat from step 1"
            self.chkStep1_ExcelExtendedVersion.setVisible(self.ddlStep1_OutputFormat.currentText() == "Excel")
            self.chkStep1_ExcelExtendedVersion.setEnabled(self.ddlStep1_OutputFormat.currentText() == "Excel")


if __name__ == '__main__':
    # This is where the main function comes
    #proc = Processor()
    
    app = QApplication(sys.argv)
    window = UIView()
    window.show()
    app.exec_()