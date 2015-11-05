__author__ = 'ghbonne'
__license__ = "GPL"

__version__ = "1.1.1"

from PyQt4.QtCore import *
from PyQt4.QtGui import *
from view.ui_UIView import Ui_UIView
import sys
import ntpath
import os
from processor.processor import Processor
from visual.enums import *
import xml.etree.ElementTree as ET
import ast # ast.literal_eval(node_or_string)
import time

class UIView(QDialog, Ui_UIView):
    """
    This is the main dialog window of the GUI controlling the PMConverter application.

    :var inputFilename: string, contains the filename of the file to process
    :var possibleVisualisations: dict, keys are the visualisation names and the corresponding value is an object instance of that visualisation type
    :var possibleVisualisationsStructured: dict, translation dict of visualisation to its corresponding header, contains item.title as keys and their header as a value
    :var chosenVisualisations: list, list of names of the visualisations that are added
    :var processor: Processor object, backend of PMConverter
    :var inputFiletypes: dict, links an application name (Excel, ProTrack) with its file extension
    
    """

    def __init__(self, parent= None):
        super(UIView, self).__init__(parent)
        # init generated GUI
        self.setupUi(self)
        self.lineEditStep1_InputFilename.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred)
        self.ddlStep2_VisualisationType.blockSignals(True)  # avoid signalling when initing GUI
        self.listStep2_ChosenVisualisations.setDragDropMode(QAbstractItemView.InternalMove)  # enable dragging and dropping of list items
        # connect signal of listStep2_ChosenVisualisations to notify when order of items changed:
        list_model = self.listStep2_ChosenVisualisations.model()
        list_model.layoutChanged.connect(self.on_listStep2_ChosenVisualisations_LayoutChanged)

        self.lblStep2_ParamsSaved.setVisible(False)  # Message label to show when parameters are saved
        self.lblFinished_OutputFilename.setWordWrap(True)
        self.lblStep2_VisualisationDescription.setWordWrap(True)
        #self.connect(self.cmdFinished_End, SIGNAL("clicked(bool)"), self.done)  # finish PMConvert program

        self.loadingAnimation = QMovie("view/Loading.gif")
        if not self.loadingAnimation.isValid():
            print("Supported formats by QMovie on this system = {0}".format(QMovie.supportedFormats()))
            print("UIView: Could not load loading animation")
        self.loadingAnimation.setScaledSize(QSize(50, 50))
        self.lblConverting_WaitingSpinner.setMovie(self.loadingAnimation)

         # custom Application inits
        self.inputFilename = ""
        self.processor = Processor()

        # connect signals from Processor with GUI
        self.connect(self.processor, self.processor.conversionSucceeded, self.ConversionSucceeded)
        self.connect(self.processor, self.processor.conversionFailedErrorMessage, self.ShowErrorMessage)

        self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.StartPage))
        #DEBUG
        #self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageStep2))
        #self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageConverting))
        #self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageFinished))

        # custom GUI inits
        self.inputFiletypes = self.processor.inputFiletypes

        # construct visualisations dropdownlist
        ddlVisualisationsModel = QStandardItemModel()
        self.possibleVisualisations = {}
        self.chosenVisualisations = []
        self.possibleVisualisationsStructured = {}  # translation dict, contains item.title as keys and their header as a value
        self.ddlStep2_VisualisationType.setModel(ddlVisualisationsModel)

        # clear visualization list
        self.listStep2_ChosenVisualisations.clear()

        # actual filling the ddl is done in handler which chooses to use conversion to excel       

        # enable blocked signals again:
        self.ddlStep2_VisualisationType.blockSignals(False)


    @pyqtSlot("bool")
    def on_cmdStep0_Start_clicked(self, clicked):
        "This function handles clicked events on the start button of the GUI"
        self.btnStep1_InputFile.setDefault(True)

        self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageStep1))

    @pyqtSlot("bool")
    def on_btnStep1_InputFile_clicked(self, clicked):
        "This function handles clicked events on the input file button of Step 1"
        #fileWindow = QFileDialog(parent = self, caption = "Choose an input file...")
        #fileWindow.setFileMode(QFileDialog.ExistingFile)
        #fileWindow.setNameFilter("ProTrack files (*.p2x);;Excel files (*.xlsx)")
        #fileWindow.setViewMode(QFileDialog.List)

        #if(fileWindow.exec_()):
            #self.inputFilename = fileWindow.selectedFiles()[0]
        filename = QFileDialog.getOpenFileName(self, "Choose an input file...", filter = "ProTrack files (*.p2x);;Excel files (*.xlsx)") # EVT: make filter dynamic with self.inputFiletypes

        if filename:
            self.inputFilename = filename
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
            # allow to convert to any format:
            possibleInputTypes = [type for type in self.inputFiletypes.keys()]
            self.ddlStep1_OutputFormat.addItems(possibleInputTypes)
            # auto enable the first other format:
            otherFormats = [type for type in possibleInputTypes if self.inputFiletypes[type] != fileExtension]
            self.ddlStep1_OutputFormat.setCurrentIndex(possibleInputTypes.index(otherFormats[0]))
            # Append combinations:
            possibleOutputCombinations = []
            for i in range(0,len(possibleInputTypes)):
                combination = possibleInputTypes[i] + " and "
                for j in range(i + 1, len(possibleInputTypes)):
                    possibleOutputCombinations.append(combination + possibleInputTypes[j])
            self.ddlStep1_OutputFormat.addItems(possibleOutputCombinations)

            
            self.ddlStep1_OutputFormat.setEnabled(True)

            self.cmdStep1_Next.setEnabled(True)

            self.lineEditStep1_InputFilename.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred)

            self.btnStep1_InputFile.setDefault(False)
            self.cmdStep1_Next.setDefault(True)
            self.cmdStep1_Next.setFocus()

    #@pyqtSlot("int")
    #def on_ddlStep1_OutputFormat_currentIndexChanged(self, index):
    #    "This function handles index changed events from combobox outputFormat from step 1"
    #    self.chkStep1_ExcelExtendedVersion.setVisible(self.ddlStep1_OutputFormat.currentText() == "Excel")
    #    self.chkStep1_ExcelExtendedVersion.setEnabled(self.ddlStep1_OutputFormat.currentText() == "Excel")

    @pyqtSlot("bool")
    def on_cmdStep1_Next_clicked(self, clicked):
        "This function handles click events on the next button of step 1."
        if "Excel" in self.ddlStep1_OutputFormat.currentText():
            # Conversion contains one to Excel
            
            # construct ddlStep2_VisualisationType, only if this is not yet constructed in a previous run of the application
            if self.ddlStep2_VisualisationType.count() == 0:
                self.ddlStep2_VisualisationType.blockSignals(True)
                currentHeader = "TOPHEADER"
                for item in self.processor.get_supported_visualisations():
                    if type(item) == str:
                        # insert header here
                        self.ddlStep2_VisualisationType.addParentItem(item)
                        currentHeader = item
                    else:
                        # insert selectable item here:
                        self.ddlStep2_VisualisationType.addChildItem(item.title)
                        self.possibleVisualisationsStructured[item.title] = currentHeader
                        self.possibleVisualisations[item.title] = item

                # update parameter fields for currently preselected visualisation type
                self.ddlStep2_VisualisationType.setCurrentIndex(1)  # set current index to 1 because 0 is a header item
                self.on_ddlStep2_VisualisationType_currentIndexChanged(1)
            
                # re-enable signalling
                self.ddlStep2_VisualisationType.blockSignals(False)            
            #endIf construct ddlStep2_VisualisationType

            self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageStep2))
            #self.listStep2_ChosenVisualisations.clear()
            self.btnStep2_ImportSettings.setDefault(True)
        else:
            self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageConverting))
            self.loadingAnimation.start()
            # start conversion to ProTrack!
            self.processor.setConversionSettings(self.ddlStep1_InputFormat.currentText(), self.ddlStep1_OutputFormat.currentText(), self.inputFilename)
            self.processor.start()

    @pyqtSlot("int")
    def on_ddlStep2_VisualisationType_currentIndexChanged(self, index):
        "This function handles index change events from the dropdown list of step 2 to change the visualisation type"
        # hide under every circumstance the user message:
        self.lblStep2_ParamsSaved.setVisible(False)

        chosenVisualisationTypeName = self.ddlStep2_VisualisationType.currentText()
        chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]

        if chosenVisualisationTypeObject is None:
            raise Exception("UIView:on_ddlStep2_VisualisationType_currentIndexChanged: Chosen visualisation type {0} is not found in available visualisations.".format(chosenVisualisationTypeName))
            return

        # update description groupbox
        self.grpbxStep2_VisualisationDescription.setTitle(chosenVisualisationTypeObject.title)
        self.lblStep2_VisualisationDescription.setText(chosenVisualisationTypeObject.description)

        chosenVisualisationTypePresetValues = {}
        # update parameter fields
        if chosenVisualisationTypeName in self.chosenVisualisations:
            # visualisation already chosen => load its parameters as preset:
            chosenVisualisationTypePresetValues = chosenVisualisationTypeObject.__dict__
            # select the corresponding item in the chosen list:
            foundItems = self.listStep2_ChosenVisualisations.findItems(chosenVisualisationTypeObject.title, Qt.MatchExactly)
            if len(foundItems) != 1:
                raise Exception("UIView:on_ddlStep2_VisualisationType_currentIndexChanged: Multiple items in the chosen visualisatios list with title = {0}".format(chosenVisualisationTypeObject.title))
            else:
                self.listStep2_ChosenVisualisations.blockSignals(True)  # avoid loop of signalling
                self.listStep2_ChosenVisualisations.setItemSelected(foundItems[0], True)
                self.listStep2_ChosenVisualisations.blockSignals(False)
            # set states of parameter handle buttons:
            self.btnStep2_AddVisualisation.setEnabled(False)
            self.btnStep2_DeleteVisualisation.setEnabled(True)
        else:
            # visualisation is not yet added to chosen visualisations list
            self.listStep2_ChosenVisualisations.blockSignals(True)  # avoid loop of signalling
            self.listStep2_ChosenVisualisations.clearSelection()
            self.listStep2_ChosenVisualisations.blockSignals(False)
            # set states of parameter handle buttons:
            self.btnStep2_AddVisualisation.setEnabled(True)
            self.btnStep2_DeleteVisualisation.setEnabled(False)
        
        # show parameter fields:
        # disable auto-save signalling:
        self.ddlStep2_Param_LevelOfDetail.blockSignals(True)
        self.chkStep2_Param_RelativeValues.blockSignals(True)
        self.chkStep2_Param_AddThreshold.blockSignals(True)
        self.dblspnbxStep2_Param_StartThresholdValue.blockSignals(True)
        self.dblspnbxStep2_Param_EndThresholdValue.blockSignals(True)

        # top ddl for choosing the level of detail
        level_of_detail_param_found = "level_of_detail" in chosenVisualisationTypeObject.parameters
        x_axis_param_found = "x_axis" in chosenVisualisationTypeObject.parameters

        if level_of_detail_param_found and x_axis_param_found:
            raise NotImplementedError("GUI is currently not designed to display 2 dropdownlists to choose level of detail and x-axis scale.")
        elif level_of_detail_param_found or x_axis_param_found:
            # object needs to show ddl with choice of level_of_detail or x-axis scale
            self.lblStep2_Param_LevelOfDetail.setVisible(True)
            self.ddlStep2_Param_LevelOfDetail.clear()
            self.ddlStep2_Param_LevelOfDetail.addItems([option.value for option in chosenVisualisationTypeObject.parameters["level_of_detail" if level_of_detail_param_found else "x_axis"]])
            
            if chosenVisualisationTypeName in self.chosenVisualisations:
                # if visualisation type already chosen: set last value:
                chosenIndex = self.ddlStep2_Param_LevelOfDetail.findText(chosenVisualisationTypePresetValues["level_of_detail" if level_of_detail_param_found else "x_axis"].value, Qt.MatchExactly)
                self.ddlStep2_Param_LevelOfDetail.setCurrentIndex(chosenIndex)

            self.ddlStep2_Param_LevelOfDetail.setVisible(True)
            self.ddlStep2_Param_LevelOfDetail.setEnabled(True)

        else:
            # object does not need to show ddl with choice of level_of_detail or x-axis scale
            self.lblStep2_Param_LevelOfDetail.setVisible(False)
            self.ddlStep2_Param_LevelOfDetail.setVisible(False)
            self.ddlStep2_Param_LevelOfDetail.setEnabled(False)

        # checkbox to choose the datatype
        if "data_type" in chosenVisualisationTypeObject.parameters:
            # enable checkbox
            if chosenVisualisationTypeName in self.chosenVisualisations:
                # if visualisation type already chosen: set last value:
                self.chkStep2_Param_RelativeValues.setChecked(chosenVisualisationTypeObject.data_type == DataType.RELATIVE)
            else:
                self.chkStep2_Param_RelativeValues.setChecked(False)

            self.chkStep2_Param_RelativeValues.setEnabled(True)
            self.chkStep2_Param_RelativeValues.setVisible(True)
        else:
            # disable checkbox
            self.chkStep2_Param_RelativeValues.setEnabled(False)
            self.chkStep2_Param_RelativeValues.setVisible(False)

        # fields to set threshold
        if "threshold" in chosenVisualisationTypeObject.parameters:
            # enable checkbox to set threshold
            enableSpinBoxes = False
            if chosenVisualisationTypeName in self.chosenVisualisations:
                # if visualisation type already chosen: set last value:
                enableSpinBoxes = chosenVisualisationTypeObject.threshold
                self.chkStep2_Param_AddThreshold.setChecked(chosenVisualisationTypeObject.threshold)
                self.dblspnbxStep2_Param_StartThresholdValue.setValue(chosenVisualisationTypeObject.thresholdValues[0])
                self.dblspnbxStep2_Param_EndThresholdValue.setValue(chosenVisualisationTypeObject.thresholdValues[1])

            else:
                self.chkStep2_Param_AddThreshold.setChecked(False)
                self.dblspnbxStep2_Param_StartThresholdValue.setValue(0)
                self.dblspnbxStep2_Param_EndThresholdValue.setValue(0)


            self.chkStep2_Param_AddThreshold.setEnabled(True)
            self.chkStep2_Param_AddThreshold.setVisible(True)
            self.lblStep2_Param_StartingThreshold.setVisible(True)
            self.dblspnbxStep2_Param_StartThresholdValue.setEnabled(enableSpinBoxes)
            self.dblspnbxStep2_Param_StartThresholdValue.setVisible(True) 
            self.lblStep2_Param_EndingThreshold.setVisible(True)
            self.dblspnbxStep2_Param_EndThresholdValue.setEnabled(enableSpinBoxes)
            self.dblspnbxStep2_Param_EndThresholdValue.setVisible(True)

        else:
            # disable checkbox to set threshold
            self.chkStep2_Param_AddThreshold.setEnabled(False)
            self.chkStep2_Param_AddThreshold.setVisible(False)
            self.lblStep2_Param_StartingThreshold.setVisible(False)
            self.dblspnbxStep2_Param_StartThresholdValue.setEnabled(False)
            self.dblspnbxStep2_Param_StartThresholdValue.setVisible(False)
            self.lblStep2_Param_EndingThreshold.setVisible(False)
            self.dblspnbxStep2_Param_EndThresholdValue.setEnabled(False)
            self.dblspnbxStep2_Param_EndThresholdValue.setVisible(False)

        # show "No parameters" label if no parameters available for selected visualisation type
        self.lblStep2_Param_NoParams.setVisible(not bool(chosenVisualisationTypeObject.parameters))

        # enable auto-save signalling:
        self.ddlStep2_Param_LevelOfDetail.blockSignals(False)
        self.chkStep2_Param_RelativeValues.blockSignals(False)
        self.chkStep2_Param_AddThreshold.blockSignals(False)
        self.dblspnbxStep2_Param_StartThresholdValue.blockSignals(False)
        self.dblspnbxStep2_Param_EndThresholdValue.blockSignals(False)

        return

    @pyqtSlot("int")
    def on_chkStep2_Param_AddThreshold_stateChanged(self, value):
        "This function handles events of checking/unchecking of the AddThreshold parameter and auto-save if the visualisation is added."
        booleanValue = value == 2  # 2 = checkbox checked
        self.dblspnbxStep2_Param_StartThresholdValue.setEnabled(booleanValue)
        self.dblspnbxStep2_Param_EndThresholdValue.setEnabled(booleanValue)

        chosenVisualisationTypeName = self.ddlStep2_VisualisationType.currentText()

        if chosenVisualisationTypeName in self.chosenVisualisations:
            chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]
            chosenVisualisationTypeObject.threshold = booleanValue

            # show user message:
            self.lblStep2_ParamsSaved.setVisible(True)
            QTimer.singleShot(2000, self.remove_SavedChanges_Message)

        #else: return

    @pyqtSlot()
    def on_listStep2_ChosenVisualisations_itemSelectionChanged(self):
        """
        This function handles selection changed events from the chosen visualisations list.
        Logic to change parameters to show is handled by the event handler of the ddlStep2_VisualisationType
        """
        selectedItems = self.listStep2_ChosenVisualisations.selectedItems()
        if len(selectedItems) != 1:
            raise Exception("UIView:on_listStep2_ChosenVisualisations_itemSelectionChanged: Not 1 item selected but {0}".format(len(selectedItems)))
        else:
            # set ddl of visualisation types to the correct index
            chosenDllIndex = self.ddlStep2_VisualisationType.findText(selectedItems[0].text(), Qt.MatchExactly)
            self.ddlStep2_VisualisationType.setCurrentIndex(chosenDllIndex)
    
        return

    @pyqtSlot()
    def on_listStep2_ChosenVisualisations_LayoutChanged(self):
        """
        This function handles the change of order of the list of chosen visualisations by the user such that the output Excel file its tabs follow the same order.
        """
        # copy the new order of chosen visualisations to the chosenVisualisations list
        self.chosenVisualisations = [self.listStep2_ChosenVisualisations.item(row).text() for row in range(self.listStep2_ChosenVisualisations.count())]
        return

    @pyqtSlot("bool")
    def on_btnStep2_AddVisualisation_clicked(self, clicked):
        "This function handles click events of the Add button of step 2"
        chosenVisualisationTypeName = self.ddlStep2_VisualisationType.currentText()
        chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]

        # set chosen parameters in the visualisation object
        for key, value in chosenVisualisationTypeObject.parameters.items():
            if key == "level_of_detail":
                selectedText = self.ddlStep2_Param_LevelOfDetail.currentText()
                chosenVisualisationTypeObject.level_of_detail = LevelOfDetail._value2member_map_[selectedText]
            elif key == "x_axis":
                selectedText = self.ddlStep2_Param_LevelOfDetail.currentText()
                chosenVisualisationTypeObject.x_axis = XAxis._value2member_map_[selectedText]
            elif key == "data_type":
                chosenVisualisationTypeObject.data_type = DataType.RELATIVE if self.chkStep2_Param_RelativeValues.isChecked() else DataType.ABSOLUTE
            elif key == "threshold":
                # also add threshold values
                chosenVisualisationTypeObject.threshold = self.chkStep2_Param_AddThreshold.isChecked()
                if chosenVisualisationTypeObject.threshold:
                    # create tuple of start and end threshold
                    chosenVisualisationTypeObject.thresholdValues = (self.dblspnbxStep2_Param_StartThresholdValue.value(), self.dblspnbxStep2_Param_EndThresholdValue.value())
                else:
                    chosenVisualisationTypeObject.thresholdValues = (0,0)
        #endFor setting possible parameters in object

        # add visualisation to chosen visualisations list
        self.listStep2_ChosenVisualisations.addItem(chosenVisualisationTypeName)
        self.listStep2_ChosenVisualisations.setItemSelected(self.listStep2_ChosenVisualisations.item(self.listStep2_ChosenVisualisations.count() - 1), True)

        # add visualisation to chosen visualisations dict
        self.chosenVisualisations.append(chosenVisualisationTypeName)

        # disable adding and enable deleting:
        self.btnStep2_AddVisualisation.setEnabled(False)
        self.btnStep2_DeleteVisualisation.setEnabled(True)

        # show user message:
        self.lblStep2_ParamsSaved.setVisible(True)
        QTimer.singleShot(2000, self.remove_SavedChanges_Message)
        return

    @pyqtSlot("bool")
    def on_btnStep2_DeleteVisualisation_clicked(self, clicked):
        "This function handles click events on the Delete button of step 2"
        chosenVisualisationTypeName = self.ddlStep2_VisualisationType.currentText()
        chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]

        # delete item from chosen visualisations list
        #self.listStep2_ChosenVisualisations.removeItemWidget(self.listStep2_ChosenVisualisations.currentItem())
        listItem = self.listStep2_ChosenVisualisations.indexFromItem(self.listStep2_ChosenVisualisations.findItems(chosenVisualisationTypeName, Qt.MatchExactly)[0])
        self.listStep2_ChosenVisualisations.blockSignals(True)
        self.listStep2_ChosenVisualisations.takeItem(listItem.row())
        self.listStep2_ChosenVisualisations.blockSignals(False)

        # delete item from chosen visualisations list
        self.chosenVisualisations.remove(chosenVisualisationTypeObject.title)

        # reset parameter fields
        self.on_ddlStep2_VisualisationType_currentIndexChanged(self.ddlStep2_VisualisationType.currentIndex())

         # show user message:
        self.lblStep2_ParamsSaved.setVisible(True)
        QTimer.singleShot(2000, self.remove_SavedChanges_Message)
        return

    def remove_SavedChanges_Message(self):
        "This function hides the saved changes message"
        self.lblStep2_ParamsSaved.setVisible(False)
        return
    
    # slots to auto-save changes of parameters when this visualisation is saved
    @pyqtSlot("int")
    def on_ddlStep2_Param_LevelOfDetail_currentIndexChanged(self, index):
        "This function handles changes to the LevelOfDetail combobox to auto-save if visualisation is added."
        
        chosenVisualisationTypeName = self.ddlStep2_VisualisationType.currentText()

        if chosenVisualisationTypeName in self.chosenVisualisations:
            chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]

            if "level_of_detail" in chosenVisualisationTypeObject.parameters:
                selectedText = self.ddlStep2_Param_LevelOfDetail.currentText()
                chosenVisualisationTypeObject.level_of_detail = LevelOfDetail._value2member_map_[selectedText]
            elif "x_axis" in chosenVisualisationTypeObject.parameters:
                selectedText = self.ddlStep2_Param_LevelOfDetail.currentText()
                chosenVisualisationTypeObject.x_axis = XAxis._value2member_map_[selectedText]

            # show user message:
            self.lblStep2_ParamsSaved.setVisible(True)
            QTimer.singleShot(2000, self.remove_SavedChanges_Message)

        # else: return

    @pyqtSlot("int")
    def on_chkStep2_Param_RelativeValues_stateChanged(self, value):
        "This function handles changes to the data type checkbox to auto-save if visualisation is added."

        chosenVisualisationTypeName = self.ddlStep2_VisualisationType.currentText()

        if chosenVisualisationTypeName in self.chosenVisualisations:
            chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]
            booleanValue = value == 2  # 2 = checkbox checked
            chosenVisualisationTypeObject.data_type = DataType.RELATIVE if booleanValue else DataType.ABSOLUTE

            # show user message:
            self.lblStep2_ParamsSaved.setVisible(True)
            QTimer.singleShot(2000, self.remove_SavedChanges_Message)

        # else: return
    
    @pyqtSlot("double")
    def on_dblspnbxStep2_Param_StartThresholdValue_valueChanged(self, value):
        "This function handles changes to the startThresholdValue to auto-save if visualisation is added."

        chosenVisualisationTypeName = self.ddlStep2_VisualisationType.currentText()

        if chosenVisualisationTypeName in self.chosenVisualisations:
            chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]

            if chosenVisualisationTypeObject.threshold:
                chosenVisualisationTypeObject.thresholdValues = (value, chosenVisualisationTypeObject.thresholdValues[1])

            # show user message:
            self.lblStep2_ParamsSaved.setVisible(True)
            QTimer.singleShot(2000, self.remove_SavedChanges_Message)

        # else: return
        return

    @pyqtSlot("double")
    def on_dblspnbxStep2_Param_EndThresholdValue_valueChanged(self, value):
        "This function handles changes to the endThresholdValue to auto-save if visualisation is added."

        chosenVisualisationTypeName = self.ddlStep2_VisualisationType.currentText()

        if chosenVisualisationTypeName in self.chosenVisualisations:
            chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]

            if chosenVisualisationTypeObject.threshold:
                chosenVisualisationTypeObject.thresholdValues = (chosenVisualisationTypeObject.thresholdValues[0], value)

            # show user message:
            self.lblStep2_ParamsSaved.setVisible(True)
            QTimer.singleShot(2000, self.remove_SavedChanges_Message)

        # else: return

    @pyqtSlot("bool")
    def on_cmdStep2_Convert_clicked(self, clicked):
        "This function handles click events of the Convert button of Step 2"
        # make structured dict of the wanted visualisations:
        wantedVisualisations = {}

        for chosenVisualisationTypeName in self.chosenVisualisations:
            # add each chosen visualisation to the correct headerlist in wantedVisualisations:
            chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]

            chosenVisualisationTypeHeader = self.possibleVisualisationsStructured[chosenVisualisationTypeName]

            if chosenVisualisationTypeHeader not in wantedVisualisations:
                wantedVisualisations[chosenVisualisationTypeHeader] = [chosenVisualisationTypeObject]
            else:
                wantedVisualisations[chosenVisualisationTypeHeader].append(chosenVisualisationTypeObject)

        #endFor structuring chosen visualisations

        # append the empty headers to wantedVisualisations:
        allPossibleHeaders = [x for x in self.processor.visualizations if type(x) == str]
        for header in allPossibleHeaders:
            if header not in wantedVisualisations:
                wantedVisualisations[header] = []

        # start conversion here to excel (and ProTrack)
        # start thread here
        self.processor.setConversionSettings(self.ddlStep1_InputFormat.currentText(), self.ddlStep1_OutputFormat.currentText(), self.inputFilename, wantedVisualisations)
        self.processor.start()

        self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageConverting))
        self.loadingAnimation.start()
        
    @pyqtSlot("bool")
    def on_btnStep2_ExportSettings_clicked(self, clicked):
        """
        This function handles click events on the Export button of Step 2.
        The currently chosen visualisations with their settings will be exported to a settings xml file.
        """
        filename = QFileDialog.getSaveFileName(self, "Save current visualisation settings", filter = "PMConverter settingfile (*-PMConverterSettingsFile.xml)")
        if filename:
            # file chosen to export to
            filenamePath, fileExtension = os.path.splitext(filename)
            # check if overwriting previous PMConverterSettingsFile => avoid appending custom file ending after custom file ending
            if filenamePath.endswith("-PMConverterSettingsFile"):
                filenamePath = filenamePath[:-24]
            outputFilename = filenamePath + "-PMConverterSettingsFile.xml"

            # write visualisation settings
            self.ExportChosenVisualisationSettings(outputFilename)

            # change focus to export button:
            self.btnStep2_ImportSettings.setDefault(False)
            self.cmdStep2_Convert.setDefault(True)
            self.cmdStep2_Convert.setFocus()

        else:
            # no file chosen to export
            # do nothing
            #print("No file chosen to export")
            return

    def ExportChosenVisualisationSettings(self, outputFilename):
        "This function exports the currently chosen visualisation settings in the GUI to a file."

        xmlRoot = ET.Element("PMConverter_Chosen_Visualisation_Settings", {"TimeOfSaving": "{0}".format(time.strftime("%d-%m-%Y_%H-%M-%S")), "PMConverterVersion": __version__})

        for chosenVisualisationTypeName in self.chosenVisualisations:
            # create a new visualisation item node for each chosen visualisation
            visualisationItemNode = ET.Element("VisualisationItem")

            chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]

            tempEl = ET.Element("Name")
            tempEl.text = chosenVisualisationTypeName
            visualisationItemNode.append(tempEl)

            for key, value in chosenVisualisationTypeObject.parameters.items():
                tempEl = ET.Element(key)

                if key == "level_of_detail":
                    tempEl.text = chosenVisualisationTypeObject.level_of_detail.value
                elif key == "x_axis":
                    tempEl.text = chosenVisualisationTypeObject.x_axis.value
                elif key == "data_type":
                    tempEl.text = chosenVisualisationTypeObject.data_type.value
                elif key == "threshold":
                    tempEl.set("value", str(chosenVisualisationTypeObject.threshold))

                    # also add threshold values
                    subNode = ET.Element("thresholdValues")
                    subNode.text = str(chosenVisualisationTypeObject.thresholdValues)
                    tempEl.append(subNode)

                visualisationItemNode.append(tempEl)
            #endFor setting possible parameters in object

            # append visualisation item node to root
            xmlRoot.append(visualisationItemNode)
        #endFor chosenVisualisations

        # write xmlRoot to file
        xmlTree = ET.ElementTree(xmlRoot)
        xmlTree.write(outputFilename)
        return

    @pyqtSlot("bool")
    def on_btnStep2_ImportSettings_clicked(self, clicked):
        """
        This function handles click events on the Import button of Step 2.
        The currently chosen visualisations with their settings will be imported from a settings xml file.
        """
        filename = QFileDialog.getOpenFileName(self, "Load current visualisation settings", filter= "PMConverter settingfile (*-PMConverterSettingsFile.xml)")

        if filename:
            # file chosen to import from
            #print("Chosen file = {0}".format(filename))
            # load visualisation settings
            self.ImportChosenVisualisationSettings(filename)

            # change selected button to convert after having imported a settings file:
            self.btnStep2_ImportSettings.setDefault(False)
            self.cmdStep2_Convert.setDefault(True)
            self.cmdStep2_Convert.setFocus()

        else:
            # no file chosen to import
            # do nothing
            #print("No file chosen to import")
            return

    def ImportChosenVisualisationSettings(self, inputFilename):
        "This function imports chosen visualisation settings in the GUI of a file."

        savedTree = None

        try:
            with open(inputFilename, "r") as inputFile:
                savedTree = ET.parse(inputFile)
        except:
            print("UIView:ImportChosenVisualisationSettings: Could not load from file {0}".format(inputFilename))
            print("UIView:ImportChosenVisualisationSettings: Unhandled Exception occurred of type: {0}".format(sys.exc_info[0]))
            print("UIView:ImportChosenVisualisationSettings: Unhandled Exception value = {0}".format(sys.exc_info[1] if sys.exc_info[1] is not None else "None"))

        if savedTree is None:
            print("UIView:ImportChosenVisualisationSettings: No visualisation settings found to import.")
            return

        rootNode = savedTree.getroot()

        if rootNode.tag != "PMConverter_Chosen_Visualisation_Settings":
            print("UIView:ImportChosenVisualisationSettingsl: Unexpected rootnode found with tag = {0}".format(rootNode.tag))
            return

        # start overwriting current settings with imported settings
        self.chosenVisualisations.clear()
        self.listStep2_ChosenVisualisations.blockSignals(True)
        self.listStep2_ChosenVisualisations.clear()

        visualisationItemNodesList = list(rootNode)

        if len(visualisationItemNodesList) == 0:
            print("UIView:ImportChosenVisualisationSettings: Imported visualisation settings file does not contain any chosen visualisations.")

        for visualisationItemNode in visualisationItemNodesList:
            nameNode = visualisationItemNode.find("Name")

            if nameNode is None:
                # no valid visualisationItemNode
                continue

            chosenVisualisationTypeName = nameNode.text

            if chosenVisualisationTypeName not in self.possibleVisualisations:
                # no possible visualisation type
                continue

            chosenVisualisationTypeObject = self.possibleVisualisations[chosenVisualisationTypeName]
            # append imported visualisation type name to chosenVisualisations
            self.chosenVisualisations.append(chosenVisualisationTypeName)
            self.listStep2_ChosenVisualisations.addItem(chosenVisualisationTypeName)

            # check if to set any settings in the chosenVisualisationTypeObject
            for key, value in chosenVisualisationTypeObject.parameters.items():
                # search for key in imported node:
                propertyNode = visualisationItemNode.find(key)
                if propertyNode is not None:
                    if key == "level_of_detail":
                        chosenVisualisationTypeObject.level_of_detail = LevelOfDetail._value2member_map_[propertyNode.text]
                    elif key == "x_axis":
                        chosenVisualisationTypeObject.x_axis = XAxis._value2member_map_[propertyNode.text]
                    elif key == "data_type":
                        chosenVisualisationTypeObject.data_type = DataType._value2member_map_[propertyNode.text]
                    elif key == "threshold":
                        # also add threshold values
                        chosenVisualisationTypeObject.threshold = ast.literal_eval(propertyNode.get("value"))
                        if chosenVisualisationTypeObject.threshold:
                            # create tuple of start and end threshold
                            subNode = propertyNode.find("thresholdValues")
                            chosenVisualisationTypeObject.thresholdValues = ast.literal_eval(subNode.text)
                        else:
                            chosenVisualisationTypeObject.thresholdValues = (0,0)
                else:
                    print("UIView:ImportChosenVisualisationSettings: Could not find parameter {0} in imported visualisationItem {1}".format(key, chosenVisualisationTypeName))
        #endFor visualisationItemNodesList

        # re-enable signals from list of chosenVisualisations
        self.listStep2_ChosenVisualisations.blockSignals(False)
        # set ddl of visualisationtype selection to start:
        self.ddlStep2_VisualisationType.setCurrentIndex(1)
        return

    @pyqtSlot("QString")
    def ConversionSucceeded(self, outputFilename):
        "This function handles calls from backend that conversion was successful."
        self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageFinished))
        self.loadingAnimation.stop()

        self.lblFinished_OutputFilename.setText(outputFilename)
        self.cmdRestart.setDefault(True)

    @pyqtSlot("QString")
    def ShowErrorMessage(self, errorMessage):
        "This function handles calls from backend that conversion failed due to the specified error message."
        self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageConvertingError))
        self.loadingAnimation.stop()

        self.textEdit_errorMessagebox.clear()
        self.textEdit_errorMessagebox.setTextColor(QColor("red"))
        self.textEdit_errorMessagebox.setFontWeight(QFont.Bold)

        self.textEdit_errorMessagebox.append(str(errorMessage))
        
    @pyqtSlot("bool")
    def on_cmdRestart_clicked(self, clicked):
        "This function handles click events on the restart button of finished page."
        # custom Application inits
        self.inputFilename = ""
        self.processor = Processor()
        # connect signals from Processor with GUI
        self.connect(self.processor, self.processor.conversionSucceeded, self.ConversionSucceeded)
        self.connect(self.processor, self.processor.conversionFailedErrorMessage, self.ShowErrorMessage)

        # inits for page step 1
        self.lineEditStep1_InputFilename.clear()
        self.ddlStep1_InputFormat.clear()
        self.ddlStep1_OutputFormat.clear()
        self.cmdStep1_Next.setDefault(False)
        self.cmdStep1_Next.setEnabled(False)
        self.btnStep1_InputFile.setDefault(True)
        self.btnStep1_InputFile.setFocus()

        # copy settings of previous visualisations:
        old_possibleVisualisations = self.possibleVisualisations
        
        # only copy visualisations when already defined:
        if len(old_possibleVisualisations) > 0:
            for item in self.processor.get_supported_visualisations():
                if type(item) != str:
                    # copy previous visualisation item settings:
                    for key, value in item.parameters.items():
                        if key == "level_of_detail":
                            item.level_of_detail = old_possibleVisualisations[item.title].level_of_detail
                        elif key == "x_axis":
                            item.x_axis = old_possibleVisualisations[item.title].x_axis
                        elif key == "data_type":
                            item.data_type = old_possibleVisualisations[item.title].data_type
                        elif key == "threshold":
                            item.threshold = old_possibleVisualisations[item.title].threshold
                            item.thresholdValues = old_possibleVisualisations[item.title].thresholdValues
                    #endFor setting possible parameters in object

                    # update visualisation:
                    self.possibleVisualisations[item.title] = item
            #endFor constructing new possibleVisualisations

        self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageStep1))
 

if __name__ == '__main__':
    # This is where the main function comes
    
    app = QApplication(sys.argv)
    
    try:
        window = UIView()
        window.show()
        app.exec_()
    except:
        print("Unhandled Exception occurred in PMConverter of type: {0}\n".format(sys.exc_info()[0]))
        print("Unhandled Exception value = {0}\n".format(sys.exc_info()[1] if sys.exc_info()[1] is not None else "None"))
        print("Unhandled Exception traceback = {0}\n".format(sys.exc_info()[2] if sys.exc_info()[2] is not None else "None"))
        print("\PMConverter CRASHED\n")