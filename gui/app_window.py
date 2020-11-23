import os
import sys
import traceback
from math import e

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5 import uic
from PyQt5.QtWidgets import QMessageBox

from gui.date_picker import DatePicker
from gui.file_upload import FileUpload
from gui.generate_report import GenerateReport
from gui.report_download import ReportDownload
from logic import sample_script, FRED_script


class AppWindow(QtWidgets.QMainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        ui_path = os.path.dirname(os.path.abspath(__file__))
        self.setWindowIcon(QtGui.QIcon(os.path.join(ui_path, "fred.ico")))
        uic.loadUi(os.path.join(ui_path, "app_window.ui"), self)

        self.setWindowTitle('FRED')
        self.termination_file_path = ''
        self.raw_data_file_path = ''
        self.start_date = ''
        self.end_date = ''
        self.temp_file_path = ''
        self.file_upload_page = FileUpload()
        self.date_picker_page = DatePicker()
        self.confirmation_page = GenerateReport()
        self.save_page = ReportDownload()

        self.stackedWidget.addWidget(self.file_upload_page)
        self.stackedWidget.addWidget(self.date_picker_page)
        self.stackedWidget.addWidget(self.confirmation_page)
        self.stackedWidget.addWidget(self.save_page)

        self.exit_btn.clicked.connect(self.kill)
        self.report_btn.clicked.connect(self.report_process)
        self.continue_btn.clicked.connect(self.continue_process)
        self.continue_btn.hide()
        self.return_btn.clicked.connect(self.return_page)
        self.return_btn.hide()

    def kill(self):
        sys.exit()

    def return_page(self):
        self.stackedWidget.setCurrentIndex(self.stackedWidget.currentIndex() - 1)
        if(self.stackedWidget.currentIndex() == 0):
            self.continue_btn.hide()
            self.return_btn.hide()

    def report_process(self):
        try:
            self.continue_btn.show()
            self.return_btn.show()
            self.continue_process()
        except:
            traceback.print_exc()

    def continue_process(self):
        error_dialog = QtWidgets.QErrorMessage()
        form_is_valid = True
        try:
            # Checks where in the process application is at
            if self.stackedWidget.currentIndex() == 1:
                self.continue_btn.setText('Continue')
                if self.file_upload_page.validate():
                    self.process_file_upload()
                else:
                    self.show_error(self.file_upload_page.error_message)
                    form_is_valid = False
            elif self.stackedWidget.currentIndex() == 2:
                if self.date_picker_page.validate():
                    self.process_date()

                    self.confirmation_page.update_info(file_1_path=self.raw_data_file_path,
                                                       file_2_path=self.termination_file_path,
                                                       start_date=self.start_date,
                                                       end_date=self.end_date)
                    self.continue_btn.setText('Confirm')
                else:
                    self.show_error(self.date_picker_page.error_message)
                    form_is_valid = False
            elif self.stackedWidget.currentIndex() == 3:
                self.temp_file_path = FRED_script.run(self.raw_data_file_path, self.termination_file_path, self.start_date, self.end_date)
                self.continue_btn.hide()
                self.continue_btn.setText('Continue')
                self.save_page.set_temp_file_path(self.temp_file_path)
            elif self.stackedWidget.currentIndex() == 4:
                self.return_btn.hide()

            if form_is_valid:
                self.stackedWidget.setCurrentIndex(self.stackedWidget.currentIndex() + 1)
        except:
            traceback.print_exc()

    def show_error(self, error_message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(error_message)
        msg.setWindowTitle("Error")
        msg.exec_()

    def process_file_upload(self):
        self.raw_data_file_path = self.file_upload_page.file_1_path
        self.termination_file_path = self.file_upload_page.file_2_path

    def process_date(self):
        self.start_date = self.date_picker_page.start_date
        self.end_date = self.date_picker_page.end_date
