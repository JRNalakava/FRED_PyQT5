import os

from PyQt5 import QtWidgets, uic, QtCore

from gui.generate_report import GenerateReport
from gui.report_download import ReportDownload


class DatePicker(QtWidgets.QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        ui_path = os.path.dirname(os.path.abspath(__file__))
        uic.loadUi(os.path.join(ui_path, "date_picker_widget.ui"), self)
        self.file_1_path = None
        self.file_2_path = None
        self.start_date = None
        self.end_date = None
        self.error_message = ''
        self.calendar_start.clicked[QtCore.QDate].connect(self.register_date_start)
        self.calendar_end.clicked[QtCore.QDate].connect(self.register_date_end)


    def register_date_start(self, date):
        self.label_start.setText(date.toString())
        self.start_date = date

    def register_date_end(self, date):
        self.label_end.setText(date.toString())
        self.end_date = date

    def validate(self):
        if self.start_date is None:
            self.error_message = 'Start date is empty.'
            return False
        if self.end_date is None:
            self.error_message = 'End date is empty.'
            return False
        if self.start_date > self.end_date:
            self.error_message = 'End date must be after start date.'
            return False
        return True