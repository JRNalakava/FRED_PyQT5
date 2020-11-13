import os

from PyQt5 import QtWidgets, uic, QtCore

from gui.report_download import ReportDownload
from logic import sample_script


class GenerateReport(QtWidgets.QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        ui_path = os.path.dirname(os.path.abspath(__file__))
        uic.loadUi(os.path.join(ui_path, "generate_report_widget.ui"), self)

    #TODO fix variable naming with file paths
    def update_info(self, file_1_path, file_2_path, start_date, end_date):
        #todo test this out
        try:
            self.label_end_date.setText(end_date.toString())
            self.label_start_date.setText(start_date.toString())
        except:
            print()
        self.label_filea_path.setText(file_1_path)
        self.label_fileb_path.setText(file_2_path)
