# This is the program's main file
# The GUI is created in this file and all other logic is handled elsewhere
# For more information on GUI and how to use PyQt5 refer to https://build-system.fman.io/docs/
# Press the green button in the gutter to run the script.
import sys
from PyQt5 import QtWidgets

from gui.app_window import AppWindow

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = AppWindow()
    window.show()
    app.exec_()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
