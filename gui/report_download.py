import os, platform, traceback
from pathlib import Path

from PyQt5 import QtWidgets, uic, QtCore
from PyQt5.QtWidgets import QFileDialog
from shutil import copyfile


class ReportDownload(QtWidgets.QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        ui_path = os.path.dirname(os.path.abspath(__file__))
        uic.loadUi(os.path.join(ui_path, "report_download_widget.ui"), self)
        self.open_btn.clicked.connect(self.open_file)
        self.save_btn.clicked.connect(self.save_file)

    def set_temp_file_path(self, file_path):
        self.temp_file_path = file_path

    def save_file(self):
        # gets url of save file
        home_dir = str(Path.home())
        file_path = QFileDialog.getSaveFileName(self)[0] + '.xlsx'

        copyfile(self.temp_file_path, file_path)

    def open_file(self):
        if platform.system() == 'Windows':
            print('windows:  ', self.temp_file_path)
            try:
                os.startfile(self.temp_file_path)
            except Exception:
                traceback.print_exc()
        elif platform.system() == 'Darwin':
            print('mac')
            os.system("open -a 'Microsoft Excel.app' '%s'" % self.temp_file_path)
