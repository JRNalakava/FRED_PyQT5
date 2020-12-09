import os
import sys
from pathlib import Path

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5 import uic
from PyQt5.QtWidgets import QFileDialog

from gui.date_picker import DatePicker


# This class describes the File Upload widget
# Here, the user is able to enter two different filepaths
# This widget contains validation and dialogs
# For FRED's purpose, file_1 refers to raw data and file_2 refers to
# Termination report
class FileUpload(QtWidgets.QWidget):

    # init function for FileUpload
    # Adds the main UI from QtCreator
    # Sets up variables and button functionality
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Setting up GUI from QtCreator
        ui_path = os.path.dirname(os.path.abspath(__file__))
        uic.loadUi(os.path.join(ui_path, "file_upload_widget.ui"), self)
        # initializing variables
        self.file_1_path = None
        self.file_2_path = None
        self.error_message = ''
        # Setting up button functionality
        self.upload_1_btn.clicked.connect(self.show_dialog_1)
        self.upload_2_btn.clicked.connect(self.show_dialog_2)

    # Functions to show File Picker Dialog for File 1 and File 2
    # Shows dialog for raw data report
    def show_dialog_1(self, file):
        home_dir = str(Path.home())
        fname = QFileDialog.getOpenFileName(self, 'Open file', home_dir)[0]
        self.label_1.setText(fname)
        self.file_1_path = fname

    # Shows dialog for termination report
    def show_dialog_2(self, file):
        home_dir = str(Path.home())
        fname = QFileDialog.getOpenFileName(self, 'Open file', home_dir)[0]
        self.label_2.setText(fname)
        self.file_2_path = fname

    # Function to validate user information
    def validate(self):
        if self.file_1_path is None:
            self.error_message = 'Raw Data Filepath is empty.'
            return False
        if self.file_2_path is None:
            self.error_message = 'Termination Report Filepath is empty.'
            return False
        if self.file_1_path.split('.')[-1] not in ('xlsx', 'xls'):
            self.error_message = 'Raw Data File Extension {} is not supported.'. \
                format(self.file_1_path.split('.')[-1])
            return False
        if self.file_2_path.split('.')[-1] not in ('xlsx', 'xls'):
            self.error_message = 'Termination Report File Extension ({}) is not supported.'. \
                format(self.file_2_path.split('.')[-1])
            return False
        return True
