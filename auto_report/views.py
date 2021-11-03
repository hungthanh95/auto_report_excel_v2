# -*- coding: utf-8 -*-
# auto_report/views.py


"""This module provides the Auto report main window."""

import os
from pathlib import Path
from collections import deque
from PyQt5.QtWidgets import QWidget, QFileDialog, QColorDialog
from PyQt5.QtCore import QThread, QCoreApplication
from .my_ui.window import Ui_MyGui
from .handle_files import get_all_images, FileHandler
from .reporter import Reporter
from .custom_message import *


TESSERACT_FOLDER = 'Tesseract-OCR'
TESSERACT_EXEC_FILE = 'tesseract.exe'


class Window(QWidget, Ui_MyGui):
    def __init__(self):
        super().__init__()
        self._files = deque()
        self._filesCount = len(self._files)
        self._dict_chimes = {}
        self._fileReportPath = Path('')
        self._fileReportSavePath = Path('')
        self._folderImagePath = Path('')
        self._tesseractPath = Path('')
        self._sheetTabColor = '00AA00'
        self._is_fill_actual = True
        self._is_delete_all_img = True
        self._report_type = ''
        self._setupUI()

    def _setupUI(self):
        self.setupUi(self)
        self._updateInitState()
        self._connectSignalsSlots()

    def _updateInitState(self):
        self.loadReportButton.setEnabled(True)
        self.loadImagesButton.setEnabled(True)
        self.saveToButton.setEnabled(True)
        self.colorPickerButton.setEnabled(True)
        self.autoReportButton.setEnabled(True)
        self.fileReportPath.clear()
        self.folderImagePath.clear()
        self.saveToPath.clear()
        self.inputListWidget.clear()
        self.outputListWidget.clear()
        self.tesseractPath.setEnabled(True)
        self.checkTesseractPath()

    def _connectSignalsSlots(self):
        self.loadReportButton.clicked.connect(self.loadReportFile)
        self.loadImagesButton.clicked.connect(self.loadImagesFolderButton)
        self.saveToButton.clicked.connect(self.saveReportFile)
        self.colorPickerButton.clicked.connect(self.openColorDialog)
        self.autoReportButton.clicked.connect(self.analyzeImagesButton)
        self.fillActualCheckbox.clicked.connect(self._changeStateTesseractPath)
        self.folderImagePath.editingFinished.connect(self.loadImageWhenTextChanged)
        self.fileReportPath.editingFinished.connect(self.checkReportPathWhenTextChanged)
        self.saveToPath.editingFinished.connect(self.checkReportSavePathWhenTextChanged)
        self.tesseractPath.editingFinished.connect(self.checkTesseractPathWhenTextChanged)


    # check tesseract Path when init ui
    def checkTesseractPath(self):
        current_folder = Path('.')
        tesseract_folder = [f for f in current_folder.iterdir() if (current_folder.is_dir() and f.is_dir() and str(f) == TESSERACT_FOLDER)]
        if len(tesseract_folder) > 0:
            tesseract_exec_file = [f for f in tesseract_folder[0].iterdir() if f.name == TESSERACT_EXEC_FILE]
            if len(tesseract_exec_file) > 0:
                self.tesseractPath.setText(str(tesseract_exec_file[0].absolute()))
                self._tesseractPath = tesseract_exec_file[0].absolute()

    # utilize
    def checkIsExcelFile(self, file):
        file_ext = Path(file).suffixes

        if file_ext and '.xlsx' in file_ext:
            return True
        else:
            fileIncorrectMessage()
            return False


    def loadReportFile(self):
        self.fileReportPath.clear()
        file, filters = QFileDialog.getOpenFileName(
            self,
            "Choose File Report",
            os.getcwd(),
            filter=(
                "Excel File (*.xlsx)"
            )
        )
        if file:
            if self.checkIsExcelFile(file):
                self.fileReportPath.setText(file)
                self._fileReportPath = Path(file)
            else:
                self.fileReportPath.clear()


    def saveReportFile(self):
        self.saveToPath.clear()
        file, filters = QFileDialog.getSaveFileName(
            self,
            "Save report file to...",
            os.getcwd(),
            filter=(
                "Excel File (*.xlsx)"
            )
        )
        if file:
            if self.checkIsExcelFile(file):
                self.saveToPath.setText(file)
                self._fileReportSavePath = Path(file)
            else:
                self.saveToPath.clear()


    def checkPathIsEmpty(self, path):
        if path == '.' or path == '':
            return True
        else:
            return False

    def loadAllImage(self, folder):
        # get all image
        try:
            list_files = get_all_images(folder)
        except:
            folderNotExistMessage()
            return
        self._folderImagePath = Path(folder)
        # clear input list
        self.inputListWidget.clear()
        self._files.clear()
        for file in list_files:
            self._files.append(file)
            self.inputListWidget.addItem(file.name)
        self._filesCount = len(self._files)
        self.inputLabel.setText('Input: {} images'.format(self._filesCount))


    def openColorDialog(self):
        color = QColorDialog.getColor()

        rgbColor = color.getRgb()
        self._sheetTabColor = color.name()[1:]
        self.colorPickerFrame \
            .setStyleSheet(
            "background-color: rgb({}, {}, {});"
                .format(rgbColor[0], rgbColor[1], rgbColor[2])
        )


    # handle when text change
    def checkTesseractPathWhenTextChanged(self):
        returnValue = confirmTextChangeMessage()
        newPath = Path(self.tesseractPath.text())
        if returnValue == QMessageBox.Ok:
            if newPath.is_file() and newPath.name == TESSERACT_EXEC_FILE:
                self._tesseractPath = newPath
                return
            else:
                fileTesseractNotExistMessage()
        self.tesseractPath.setText(str(self._tesseractPath))
  

    def checkReportPathWhenTextChanged(self):
        returnValue = confirmTextChangeMessage()
        newPath = Path(self.fileReportPath.text())
        if returnValue == QMessageBox.Ok:
            if newPath.is_file() and self.checkIsExcelFile(newPath):
                self._fileReportPath = newPath
                return
            else:
                fileNotExistMessage()
        self.fileReportPath.setText(str(self._fileReportPath))


    def checkReportSavePathWhenTextChanged(self):
        returnValue = confirmTextChangeMessage()
        newPath = Path(self.saveToPath.text())
        if returnValue == QMessageBox.Ok:
            if self.checkIsExcelFile(newPath):
                self._fileReportSavePath = newPath
                return
        self.saveToPath.setText(str(self._fileReportSavePath))


    def loadImageWhenTextChanged(self):
        returnValue = confirmTextChangeMessage()
        if returnValue == QMessageBox.Ok:
            folder = self.folderImagePath.text()
            if folder:
                self.loadAllImage(folder)
        else:
            self.folderImagePath.setText(str(self._folderImagePath))

    # handle when button clicked
    def loadImagesFolderButton(self):
        # clear folder image path
        self.folderImagePath.clear()
        folder = QFileDialog.getExistingDirectory(
            self,
            "Choose Folder Contain Images",
            os.getcwd(),
            QFileDialog.ShowDirsOnly
        )
        # if path is not empty
        if folder:
            # set folder image path
            self.loadAllImage(folder)
            self.folderImagePath.setText(folder)
        else:
            self.inputListWidget.clear()
            self._folderImagePath = Path()
            self.inputLabel.setText('Input:')


    def analyzeImagesButton(self):
        if self.inputListWidget.count() > 0 and not self.folderImagePath.size().isEmpty():
            self._is_fill_actual = self.fillActualCheckbox.isChecked()
            self._tesseractPath = Path(self.tesseractPath.text())
            report_type = str(self.oscOrLogicComboBox.currentText())
            if report_type == 'Select report type...':
                unselectReportTypeMessage()
                return
            elif report_type == 'Oscilloscope':
                self._report_type = 'osc'
            elif report_type == 'Logic Analyzer':
                self._report_type = 'logic'
            elif report_type == 'Waveform':
                self._report_type = 'waveform'
            else:
                unselectReportTypeMessage()
                return
            self._runAnalyzeThread()
            self._updateStateWhenAnalyzing()
        else:
            nothingToAnalyzeMessage()


    def runAutoReportButton(self):
        if self.checkPathIsEmpty(self.fileReportPath.text()) or \
                self.checkPathIsEmpty(self.saveToPath.text()):
            pathIsEmptyMessage()
        else:
            imagesHandleFailed = self.inputListWidget.count()
            self._is_delete_all_img = self.deleteImagesCheckbox.isChecked()
            if imagesHandleFailed > 0:
                return_value = imageHandleFailedMessage(imagesHandleFailed)
                if return_value == QMessageBox.Cancel:
                    return
            self._runAutoReportThread()
            self._updateStateWhenRunningReport()

    def _openFileReport(self):
        os.system("start excel.exe {}".format(self._fileReportSavePath))

    # define task run on a thread
    def _runAnalyzeThread(self):
        self._thread = QThread()
        self._file_handler = FileHandler(files=tuple(self._files), 
                                         is_fill_actual_value=self._is_fill_actual,
                                         tesseract_path=str(self.tesseractPath),
                                         report_type=self._report_type)
        self._file_handler.moveToThread(self._thread)
        self._thread.started.connect(self._file_handler.handle_chime_images_v2)
        self._file_handler.handledFileSignal.connect(self._updateFileListWhenHandled)
        self._file_handler.tesseractNotExist.connect(self._updateStateWhenTesseractNotExist)
        self._file_handler.progressed.connect(self._updateProgressBar1)
        self._file_handler.finished.connect(self._updateStateWhenAnalyzeFinished)
        # Clean up
        self._file_handler.finished.connect(self._thread.quit)
        self._file_handler.finished.connect(self._file_handler.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        # Run the thread
        self._thread.start()


    def _runAutoReportThread(self):
        self._thread = QThread()
        self._reporter = Reporter(wb_path=str(self._fileReportPath),
                                  wb_save_path=str(self._fileReportSavePath),
                                  sheettab_color=self._sheetTabColor,
                                  is_fill_actual=self._is_fill_actual,
                                  is_delete_all_img=self._is_delete_all_img,
                                  report_type = self._report_type,
                                  dict=self._dict_chimes)
        self._reporter.moveToThread(self._thread)
        self._thread.started.connect(self._reporter.run_report)
        self._reporter.logSignal.connect(self._updateOutputListWhenReceivedLog)
        self._reporter.progressed.connect(self._updateProgressBar2)
        self._reporter.finished.connect(self._updateStateWhenReportFinished)
        # Clean up
        self._reporter.finished.connect(self._thread.quit)
        self._reporter.finished.connect(self._reporter.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)

        # Run the thread
        self._thread.start()
    



    # Update state
    def _updateProgressBar1(self, fileNum):
        progressPercent = int((fileNum + 1) / len(self._files) * 100)
        self.progressBar.setValue(progressPercent)


    def _updateFileListWhenHandled(self, status):
        takeItem = self.inputListWidget.takeItem(0)
        if status:
            self.outputListWidget.addItem(takeItem)
        else:
            self.inputListWidget.addItem(takeItem)
        self.inputLabel.setText('Input: Remain {} images'.format(self.inputListWidget.count()))
        self.outputLabel.setText('Output: Handled {} images'.format(self.outputListWidget.count()))
        self.outputListWidget.setCurrentRow(self.outputListWidget.count()-1)


    def _updateStateWhenTesseractNotExist(self):
        fileTesseractNotExistMessage()
        QCoreApplication.exit(0)

    def _updateStateWhenAnalyzeFinished(self, dict):
        self._dict_chimes = dict
        self.autoReportButton.disconnect()
        self.autoReportButton.setEnabled(True)
        self.autoReportButton.setText('Run Report')
        self.autoReportButton.clicked.connect(self.runAutoReportButton)

    def _updateStateWhenRunningReport(self):
        self.fileReportPath.setEnabled(False)
        self.loadReportButton.setEnabled(False)
        self.saveToPath.setEnabled(False)
        self.saveToButton.setEnabled(False)
        self.autoReportButton.setEnabled(False)
        self.fillActualCheckbox.setEnabled(False)
        self.deleteImagesCheckbox.setEnabled(False)
        self.tesseractPath.setEnabled(False)
        self.inputLabel.setText('Input:')
        self.inputListWidget.clear()
        self.inputListWidget.addItems([chime for chime in self._dict_chimes])
        self.outputLabel.setText('Output:')
        self.outputListWidget.clear()
        self.progressBar.setValue(0)

    def _updateOutputListWhenReceivedLog(self, str_log):
        self.outputListWidget.addItem(str_log)
        self.outputListWidget.setCurrentRow(self.outputListWidget.count()-1)

    def _updateProgressBar2(self, chimeNum):
        progressPercent = int((chimeNum + 1) / len(self._dict_chimes) * 100)
        self.progressBar.setValue(progressPercent)

    def _updateStateWhenReportFinished(self):
        self.autoReportButton.disconnect()
        self.autoReportButton.setEnabled(True)
        self.autoReportButton.setText('Open Report')
        self.autoReportButton.clicked.connect(self._openFileReport)
    

    def _changeStateTesseractPath(self):
        if self.fillActualCheckbox.isChecked():
           self.tesseractPath.setEnabled(True)
        else:
            self.tesseractPath.setEnabled(False)
    
    
    def _updateStateWhenAnalyzing(self):
        self.folderImagePath.setEnabled(False)
        self.loadImagesButton.setEnabled(False)
        self.autoReportButton.setEnabled(False)
        self.fillActualCheckbox.setEnabled(False)
        self.tesseractPath.setEnabled(False)
        self.oscOrLogicComboBox.setEnabled(False)