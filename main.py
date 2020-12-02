# This is the program's main file
# The GUI is created in this file and all other logic is handled elsewhere
# For more information on GUI and how to use PyQt5 refer to https://build-system.fman.io/docs/
# Press the green button in the gutter to run the script.
import sys
from PyQt5 import QtWidgets

from gui.app_window import AppWindow

# Main logic of the application
# Every PyQT5 application needs an app
# The AppWindow class holds our main frame
if __name__ == '__main__':
    # Instantiates QApplication
    app = QtWidgets.QApplication(sys.argv)
    # Instantiates AppWindow class and shows it
    window = AppWindow()
    window.show()
    # Runs application
    app.exec_()

