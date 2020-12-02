import os

from PyQt5 import QtWidgets, uic, QtCore

from gui.generate_report import GenerateReport
from gui.report_download import ReportDownload


# This class describes the Date Picker widget
# Here, the user is able to pick a specific year or choose for an all time report
# This widget contains validation and dialogs
class DatePicker(QtWidgets.QWidget):

    # init function for DatePicker
    # Adds the main UI from QtCreator
    # Sets up variables and calendar functionality
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        ui_path = os.path.dirname(os.path.abspath(__file__))
        uic.loadUi(os.path.join(ui_path, "date_picker_widget.ui"), self)
        self.file_1_path = None
        self.file_2_path = None
        self.start_date = -1
        self.end_date = -1
        self.error_message = ''
        self.all_time = False
        self.year_radio.clicked.connect(self.year_report_radio)
        self.alltime_radio.clicked.connect(self.all_time_report)
        self.year_box.hide()
        # setting range
        self.year_box.setRange(-1, 3000)
        # setting value
        self.year_box.setValue(2020)

    # Function to handle all time selection radio button
    def all_time_report(self):
        self.all_time = True
        self.year_box.setValue(-1)
        self.year_box.hide()

    # Function to handle year selection radio button
    def year_report_radio(self):
        self.year_box.show()
        self.year_box.setValue(2020)
        self.all_time = False

    # Function to validate user info
    def validate(self):
        if self.year_box.value() == -1 and not self.all_time:
            self.error_message = 'Start date is empty.'
            return False
        return True