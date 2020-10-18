import getpass
import os
import sys
import traceback
from win32com.client import Dispatch

import shutil
from PyQt5 import QtWidgets
from PyQt5 import QtCore
from PyQt5 import QtGui

# Getting the username and creating the path to the startup folder
user_name = getpass.getuser()
AUTORUN_PATH = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % user_name

FILE_NAME = os.path.basename(sys.argv[0])  # Current file name
ICON_NAME = 'scraper.ico'  # Application icon file name

DEBUG_MOD = True  # Displaying program errors

ORGANIZATION_NAME = 'Example App'
ORGANIZATION_DOMAIN = 'example.com'
APPLICATION_NAME = 'QSettings program'
SETTING_MINIMIZE_TO_TRAY = 'settings/minimize'
SETTING_AUTORUN = 'settings/autorun'


class MainWindow(QtWidgets.QMainWindow):
    """Main window class

        Attributes
        ----------
        tray_icon
            Responsible for the tray icon
        minimize_to_tray
            Checkbox for minimizing the program to tray
        autorun
            Checkbox for adding a program to autorun

        Methods
        -------
        init_ui(self)
            Sets the window interface
        menu(self)
            Creates a window menu
        tray(self)
            Creates a tray icon
        settings(self)
            Defines the interface and functionality of the settings window
        save_minimize_settings(self)
            Saves the checkbox value of minimizing the window to tray
        save_autorun_settings(self)
            Saves the value of the checkbox for adding a program to autorun
        hideEvent(self, event)
            Defines the behavior of the program when minimized
        closeEvent(self, event)
            Defines the behavior of the closing program
        log_uncaught_exceptions(self, ex_cls, ex, tb)
            Catches and prints errors
        """

    tray_icon = None
    minimize_to_tray = None
    autorun = None

    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)

        self.init_ui()
        self.tray()
        self.menu()

    def init_ui(self):
        """Sets the window interface"""
        self.setMinimumSize(QtCore.QSize(480, 80))
        self.setWindowTitle("Sample")
        self.setWindowIcon(QtGui.QIcon(ICON_FILE))

    def menu(self):
        """Creates a window menu"""
        setting_action = QtWidgets.QAction('Settings', self)
        setting_action.triggered.connect(self.settings)

        menu_bar = self.menuBar()
        menu_bar.addAction(setting_action)

    def tray(self):
        """Creates a tray icon with three buttons"""
        self.tray_icon = QtWidgets.QSystemTrayIcon(self)
        self.tray_icon.setIcon(QtGui.QIcon(ICON_FILE))
        self.tray_icon.setToolTip('MyApp')  # Pop up lettering icons

        show_action = QtWidgets.QAction("Show", self)
        hide_action = QtWidgets.QAction("Hide", self)
        quit_action = QtWidgets.QAction("Exit", self)
        show_action.triggered.connect(self.show)
        hide_action.triggered.connect(self.hide)
        quit_action.triggered.connect(self.close)

        tray_menu = QtWidgets.QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

    def settings(self):
        """Creates a window with program settings"""
        setting_win = QtWidgets.QWidget(self, QtCore.Qt.Window)

        setting_win.setWindowModality(QtCore.Qt.WindowModal)
        setting_win.setMinimumSize(QtCore.QSize(450, 200))
        setting_win.setWindowTitle("Settings")

        self.minimize_to_tray = QtWidgets.QCheckBox('Minimizing to tray', setting_win)
        self.minimize_to_tray.move(5, 5)
        self.minimize_to_tray.clicked.connect(self.save_minimize_settings)  # Save checkbox settings

        self.autorun = QtWidgets.QCheckBox('Run with Windows', setting_win)
        self.autorun.move(5, 25)
        self.autorun.clicked.connect(self.save_autorun_settings)  # Save checkbox settings

        settings = QtCore.QSettings()  # Loading previous settings, if any

        # Setting the checkbox to the desired position
        check_state = settings.value(SETTING_MINIMIZE_TO_TRAY, False, type=bool)
        self.minimize_to_tray.setChecked(check_state)

        # Setting the checkbox to the desired position
        check_state = settings.value(SETTING_AUTORUN, False, type=bool)
        self.autorun.setChecked(check_state)

        setting_win.show()

    def save_minimize_settings(self):
        """Saves the checkbox settings for minimizing the application to tray"""
        settings = QtCore.QSettings()
        settings.setValue(SETTING_MINIMIZE_TO_TRAY, self.minimize_to_tray.isChecked())
        settings.sync()

    def save_autorun_settings(self):
        """Saves the checkbox settings for autorun of the program and also adds or removes it from autorun"""
        settings = QtCore.QSettings()
        settings.setValue(SETTING_AUTORUN, self.autorun.isChecked())
        settings.sync()

        # Checking the state of the checkbox
        check_state = settings.value(SETTING_AUTORUN, False, type=bool)
        if check_state:
            file_path = os.path.dirname(os.path.realpath(__file__)) + "\\" + FILE_NAME

            path_to_target = file_path
            path_to_shortcut = file_path.rsplit('.', 1)[0] + '.lnk'
            path_to_work_dir = file_path.rsplit('\\', 1)[0]

            # Create a program shortcut
            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(path_to_shortcut)
            shortcut.Targetpath = path_to_target
            shortcut.WorkingDirectory = path_to_work_dir
            shortcut.save()

            # Moving the program shortcut to the autorun folder
            shutil.move(path_to_shortcut, AUTORUN_PATH)
        else:
            shortcut_path = f"{AUTORUN_PATH}\{FILE_NAME.rsplit('.', 1)[0] + '.lnk'}"

            # Checking the path of the shortcut and removing it from autorun
            if os.path.isfile(shortcut_path):
                os.remove(shortcut_path)

    def hideEvent(self, event):
        """Minimizes the program window to tray if there is a corresponding checkbox in the settings"""
        settings = QtCore.QSettings()
        check_state = settings.value(SETTING_MINIMIZE_TO_TRAY, False, type=bool)
        if check_state:
            self.hide()

    def closeEvent(self, event):
        """Closes the tray icon when the program is closed"""
        self.tray_icon.hide()

    def log_uncaught_exceptions(self, ex_cls, ex, tb):
        """Catches and prints errors if DEBUG_MOD = True"""
        text = '{}: {}:\n'.format(ex_cls.__name__, ex)
        text += ''.join(traceback.format_tb(tb))

        QtWidgets.QMessageBox.critical(self, 'Error', text)
        quit()


if __name__ == "__main__":
    # Finding the path to the icon if the program is packaged in .exe
    try:
        ico_path = sys._MEIPASS
    except AttributeError:
        ico_path = '.'
    ICON_FILE = os.path.join(ico_path, ICON_NAME)

    QtCore.QCoreApplication.setApplicationName(ORGANIZATION_NAME)
    QtCore.QCoreApplication.setOrganizationDomain(ORGANIZATION_DOMAIN)
    QtCore.QCoreApplication.setApplicationName(APPLICATION_NAME)

    app = QtWidgets.QApplication(sys.argv)
    mw = MainWindow()

    # Displaying error messages in debug mode
    if DEBUG_MOD:
        sys.excepthook = mw.log_uncaught_exceptions

    mw.show()
    sys.exit(app.exec())
