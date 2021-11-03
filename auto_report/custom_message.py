from PyQt5.QtWidgets import QMessageBox


def confirmTextChangeMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("Do you want to change path?")
    msg.setIcon(QMessageBox.Warning)
    msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    return msg.exec_()

def fileIncorrectMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("File is not .xlsx file!")
    msg.setIcon(QMessageBox.Warning)
    msg.exec_()

def folderNotExistMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("Folder is not exist!")
    msg.setIcon(QMessageBox.Warning)
    msg.exec_()

def pathEmptyMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("Some path is empty.\nPlease check again!")
    msg.setIcon(QMessageBox.Warning)
    msg.exec_()

def nothingToAnalyzeMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("Nothing to analyze.\nPlease check again!")
    msg.setIcon(QMessageBox.Warning)
    msg.exec_()

def pathIsEmptyMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("Path file excel is empty.\nPlease check again!")
    msg.setIcon(QMessageBox.Warning)
    msg.exec_()

def imageHandleFailedMessage(num):
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("We found {} image(s) file not correct rule name.\nDo you want to continuous?".format(num))
    msg.setIcon(QMessageBox.Warning)
    msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    return msg.exec_()

def fileTesseractNotExistMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("File tesseract.exe is not exist!")
    msg.setIcon(QMessageBox.Warning)
    msg.exec_()

def fileNotExistMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("File is not exist!")
    msg.setIcon(QMessageBox.Warning)
    msg.exec_()

def unselectReportTypeMessage():
    msg = QMessageBox()
    msg.setWindowTitle("Warning")
    msg.setText("Please select report type before analyze!")
    msg.setIcon(QMessageBox.Warning)
    msg.exec_()