import os
import sys
import traceback
from math import e

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5 import uic
from PyQt5.QtWidgets import QMessageBox
from xlsxwriter.exceptions import FileCreateError

from gui.date_picker import DatePicker
from gui.file_upload import FileUpload
from gui.generate_report import GenerateReport
from gui.report_download import ReportDownload
from logic import sample_script, FRED_script


# Class that holds main frame of application
# This class has a StackedWidget which is where most of the action happens
# This class has several widgets that act as steps in the process
# Each widget has its own GUI and its own logic
# The goal was to keep it as OOP as possible
# To change graphical components use QtCreator to edit the
# associated .ui files
class AppWindow(QtWidgets.QMainWindow):

    # init function for AppWindow
    # Adds the main UI from QtCreator and calls load_components
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        ui_path = os.path.dirname(os.path.abspath(__file__))

        # This line sets the icon for the window
        # to change icon, change where it says "fred.ico" to the new filepath
        self.setWindowIcon(QtGui.QIcon(os.path.join(ui_path, "fred.ico")))

        # This line loads the ui from QtCreator
        uic.loadUi(os.path.join(ui_path, "app_window.ui"), self)

        self.load_components()

    # Function to initialize all variables and widgets of the AppWindow
    def load_components(self):
        # This line sets the window title
        self.setWindowTitle('FRED')
        # Following lines initialize all of the system's variables
        self.termination_file_path = ''
        self.raw_data_file_path = ''
        self.start_date = ''
        self.end_date = ''
        self.temp_file_path = ''

        # Following lines initialize all the system's widgets
        self.file_upload_page = FileUpload()
        self.date_picker_page = DatePicker()
        self.confirmation_page = GenerateReport()
        self.save_page = ReportDownload()

        # Following lines add widgets to stackedWidget
        self.stackedWidget.addWidget(self.file_upload_page)
        self.stackedWidget.addWidget(self.date_picker_page)
        self.stackedWidget.addWidget(self.confirmation_page)
        self.stackedWidget.addWidget(self.save_page)

        # Following lines set up button functionality
        # This line sets up the exit button to terminate application
        self.exit_btn.clicked.connect(self.kill)
        # This line sets up the 'Create New Report' button to create new report
        self.report_btn.clicked.connect(self.report_process)
        # This line sets up the 'Continue' button to change stackedWidget currentIndex up
        self.continue_btn.clicked.connect(self.continue_process)
        # This line hides the button because it's not supposed to show up on the menu
        self.continue_btn.hide()
        # This line sets up the 'Continue' button to change stackedWidget currentIndex down
        self.return_btn.clicked.connect(self.return_page)
        # Following lines initialize all the object's variables
        self.return_btn.hide()

    # Function terminates application
    def kill(self):
        sys.exit()

    # Function for handling return_btn
    # Function changes stackedWidget currentIndex to i - 1
    # Function hides continue and return buttons
    def return_page(self):
        self.stackedWidget.setCurrentIndex(self.stackedWidget.currentIndex() - 1)
        if(self.stackedWidget.currentIndex() == 0):
            self.continue_btn.hide()
            self.return_btn.hide()

    # Function for handling report_btn
    # Function shows continue and return buttons
    # Function calls continue_process
    def report_process(self):
        try:
            self.continue_btn.show()
            self.return_btn.show()
            self.continue_process()
        except:
            traceback.print_exc()

    # This function handles whenever the stackedWidget changes currentIndex
    # It checks where in the application process the user is and shows the
    # appropriate widget for that step
    def continue_process(self):
        # Creates error dialog
        error_dialog = QtWidgets.QErrorMessage()
        # Flag to check whether or not all info is valid
        form_is_valid = True

        try:
            # Checks where in the process application is at
            if self.stackedWidget.currentIndex() == 1:
                # Means that user is at FileUpload Widget
                # Checks that all info on FileUpload Widget is correct and
                # uploads file if successful
                # else, show error, and mark flag as false
                self.continue_btn.setText('Continue')
                if self.file_upload_page.validate():
                    self.process_file_upload()
                else:
                    self.show_error(self.file_upload_page.error_message)
                    form_is_valid = False
            elif self.stackedWidget.currentIndex() == 2:
                # Means that user is at DatePicker Widget
                # Checks that all info on DatePicker Widget is correct and
                # uploads data if succesful
                # else, show error, and mark flag as false
                if self.date_picker_page.validate():
                    self.process_date()
                    # this line injects all of the required data into the
                    # GenerateReport Widget
                    self.confirmation_page.update_info(file_1_path=self.raw_data_file_path,
                                                       file_2_path=self.termination_file_path,
                                                       start_date=self.start_date)
                    self.continue_btn.setText('Confirm')
                else:
                    self.show_error(self.date_picker_page.error_message)
                    form_is_valid = False
            elif self.stackedWidget.currentIndex() == 3:
                try:
                    # Means that user is at GenerateReport Widget
                    # Runs script using uploaded information
                    self.temp_file_path = FRED_script.run(self.raw_data_file_path, self.termination_file_path, self.start_date)
                    self.continue_btn.hide()
                    self.continue_btn.setText('Continue')
                    # Injects file path into ReportDownload Widget
                    self.save_page.set_temp_file_path(self.temp_file_path)
                except KeyError:
                    self.show_error('There was an error processing the files.'
                                    ' Please make sure that they are in the correct format.')
                    form_is_valid = False
                except FileCreateError:
                    self.show_error('There was an error creating the report.'
                                    ' Please make sure that no reports are open.')
                    form_is_valid = False

            elif self.stackedWidget.currentIndex() == 4:
                # Means that user is at ReportDownload Widget
                self.return_btn.hide()
            # After all that, if form is valid, move StackedWidget one index
            # forward
            if form_is_valid:
                self.stackedWidget.setCurrentIndex(self.stackedWidget.currentIndex() + 1)
        except:
            traceback.print_exc()

    # Function for handling error message dialogs
    # Takes in a error message and displays a QMessageBox
    def show_error(self, error_message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(error_message)
        msg.setWindowTitle("Error")
        msg.exec_()

    # Function to handle file upload
    # Sets user inputted file paths to object's filepaths
    def process_file_upload(self):
        self.raw_data_file_path = self.file_upload_page.file_1_path
        self.termination_file_path = self.file_upload_page.file_2_path

    # Function to handle date picking
    # Sets user inputted dates to object's dates
    def process_date(self):
        self.start_date = self.date_picker_page.year_box.value()
        if self.date_picker_page.all_time:
            self.start_date = 0
