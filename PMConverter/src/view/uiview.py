__author__ = 'PM Group 8'

#from processor.processor import Processor
from PyQt4.QtCore import *
from PyQt4.QtGui import *
import ui_UIView
import sys
import ntpath
import os
from processor.processor import Processor
from visual.enums import *

#from processor import Processor

class UIView(QDialog, ui_UIView.Ui_UIView):
    """
    This is the main dialog window of the GUI controlling the PMConverter application.

    :var possibleVisualisations: dict, keys are the visualisation names and the corresponding value is an object instance of that visualisation type
    :var chosenVisualisations: list, list of names of the visualisations that are added
    
    """

    def __init__(self, parent= None):
        super(UIView, self).__init__(parent)
        # init generated GUI
        self.setupUi(self)
        self.lineEditStep1_InputFilename.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred)
        self.ddlStep2_VisualisationType.blockSignals(True)  # avoid signalling when initing GUI
        self.listStep2_ChosenVisualisations.setDragDropMode(QAbstractItemView.InternalMove)  # enable dragging and dropping of list items
        self.lblStep2_ParamsSaved.setVisible(False)  # Message label to show when parameters are saved

        #self.loadingAnimation = QMovie("Loading.gif", QByteArray(), self)
        #self.loadingAnimation = QMovie("F:/Gilles/Universiteit/5e Master 2/Projectmanagement/Project/PMConverter/PMConverter/src/view/Loading.gif")
        self.loadingAnimation = QMovie("view/Loading.gif")
        if not self.loadingAnimation.isValid():
            print("Supported formats by QMovie on this system = {0}".format(QMovie.supportedFormats()))
            print("UIView: Could not load loading animation")
        self.loadingAnimation.setScaledSize(QSize(50, 50))
        self.lblConverting_WaitingSpinner.setMovie(self.loadingAnimation)

         # custom Application inits
        self.inputFilename = ""
        self.processor = Processor()

        #DEBUG
        #self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageStep2))
        #self.loadingAnimation.start()
        #self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageConverting))

        # custom GUI inits
        self.inputFiletypes = {"Excel": ".xlsx", "ProTrack": ".p2x"}
        

        # construct visualisations dropdownlist
        ddlVisualisationsModel = QStandardItemModel()
        self.possibleVisualisations = {}
        self.chosenVisualisations = []

        for item in self.processor.visualizations:
            if type(item) == str:
                # TODO: insert header here
                pass
            else:
                newModelItem = QStandardItem()
                newModelItem.setText(item.title)
                ddlVisualisationsModel.appendRow(newModelItem)
                self.possibleVisualisations[item.title] = item
        
        self.ddlStep2_VisualisationType.setModel(ddlVisualisationsModel)
        #self.ddlStep2_VisualisationType.setView(ddlVisualisationView)

        # update parameter fields for currently preselected visualisation type
        self.on_ddlStep2_VisualisationType_currentIndexChanged(0)

        # enable blocked signals again:
        self.ddlStep2_VisualisationType.blockSignals(False)


    @pyqtSlot("bool")
    def on_cmdStep0_Start_clicked(self, clicked):
        "This function handles clicked events on the start button of the GUI"
        self.btnStep1_InputFile.setDefault(True)
        self.chkStep1_ExcelExtendedVersion.setVisible(False)

        self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageStep1))

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

            self.lineEditStep1_InputFilename.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred)

            self.btnStep1_InputFile.setDefault(False)
            self.cmdStep1_Next.setDefault(True)
            self.cmdStep1_Next.setFocus()

    @pyqtSlot("int")
    def on_ddlStep1_OutputFormat_currentIndexChanged(self, index):
        "This function handles index changed events from combobox outputFormat from step 1"
        self.chkStep1_ExcelExtendedVersion.setVisible(self.ddlStep1_OutputFormat.currentText() == "Excel")
        self.chkStep1_ExcelExtendedVersion.setEnabled(self.ddlStep1_OutputFormat.currentText() == "Excel")

    @pyqtSlot("bool")
    def on_cmdStep1_Next_clicked(self, clicked):
        "This function handles click events on the next button of step 1."
        if self.ddlStep1_OutputFormat.currentText() == "Excel":
            self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageStep2))
            self.listStep2_ChosenVisualisations.clear()
            self.btnStep2_ImportSettings.setDefault(True)
        else:
            self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageConverting))
            self.loadingAnimation.start()
            #TODO: start conversion to ProTrack!

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
        #TODO start conversion here to excel

        self.pagesMain.setCurrentIndex(self.pagesMain.indexOf(self.pageConverting))
        self.loadingAnimation.start()
        




 

if __name__ == '__main__':
    # This is where the main function comes
    #proc = Processor()
    
    app = QApplication(sys.argv)
    window = UIView()
    window.show()
    try:
        app.exec_()
    except:
        print("Unhandled Exception occurred in PMConverter of type: {0}\n".format(sys.exc_info()[0]))
        print("Unhandled Exception value = {0}\n".format(sys.exc_info()[1] if sys.exc_info()[1] is not None else "None"))
        print("Unhandled Exception traceback = {0}\n".format(sys.exc_info()[2] if sys.exc_info()[2] is not None else "None"))
        print("\PMConverter CRASHED\n")