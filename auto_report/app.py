# -*- coding: utf-8 -*-
# auto_report/app.py

"""This module provides the Auto report application."""

import sys

from PyQt5.QtWidgets import QApplication
from PyQt5 import QtGui

from .views import Window

def main():
    # Create the application
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setWindowIcon(QtGui.QIcon("auto_report\\resource\\icon.jpg"))
    # Create and show the main window
    win = Window()
    win.show()

    # Run the event loop
    sys.exit(app.exec())