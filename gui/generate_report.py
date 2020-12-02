import os

from PyQt5 import QtWidgets, uic, QtCore

from gui.report_download import ReportDownload
from logic import sample_script


# This is the class for the GenerateReport widget
# This class holds the UI code and an injection function
class GenerateReport(QtWidgets.QWidget):

    # init function for GenerateReport
    # Adds the main UI from QtCreator
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # loads ui from QtCreator
        ui_path = os.path.dirname(os.path.abspath(__file__))
        uic.loadUi(os.path.join(ui_path, "generate_report_widget.ui"), self)


    # This function allows the AppWindow instance to share user info from
    # other widgets to this widget
    def update_info(self, file_1_path, file_2_path, start_date):
        try:
            if start_date != 0:
                self.label_start_date.setText(str(start_date))
            else:
                self.label_start_date.setText('All Time Report')
        except:
            print()
        self.label_filea_path.setText(file_1_path)
        self.label_fileb_path.setText(file_2_path)
