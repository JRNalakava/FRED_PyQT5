import os, platform, traceback
from pathlib import Path

from PyQt5 import QtWidgets, uic, QtCore
from PyQt5.QtWidgets import QFileDialog
from shutil import copyfile


# This is the class for the ReportDownload widget
# This class holds the UI code, sets up button functionality, and has Functions
# for handling opening and saving the result excel file
class ReportDownload(QtWidgets.QWidget):

    # init function for ReportDownload
    # Adds the main UI from QtCreator
    # Sets up button functionality
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Loads ui code
        ui_path = os.path.dirname(os.path.abspath(__file__))
        uic.loadUi(os.path.join(ui_path, "report_download_widget.ui"), self)
        # Sets up button functionality
        self.open_btn.clicked.connect(self.open_file)
        self.save_btn.clicked.connect(self.save_file)
        self.temp_file_path = ''

    # Function to set the file path of the returned excel file
    # from the app's script
    def set_temp_file_path(self, file_path):
        self.temp_file_path = file_path

    # Function to handle save_btn
    # Function saves result excel file
    def save_file(self):
        # gets url of save file
        home_dir = str(Path.home())
        # Shows dialog to get user's desired filepath
        file_path = QFileDialog.getSaveFileName(self)[0] + '.xlsx'

        # This line copies the temporary result file to the final user
        # desired destination
        copyfile(self.temp_file_path, file_path)

    # Function to handle open_btn
    # Function opens excel file
    def open_file(self):
        # Checks whether user is running mac or Windows
        # If user is on windows
        if platform.system() == 'Windows':
            try:
                os.startfile(self.temp_file_path)
            except Exception:
                traceback.print_exc()
        # If user is on mac
        elif platform.system() == 'Darwin':
            os.system("open -a 'Microsoft Excel.app' '%s'" % self.temp_file_path)