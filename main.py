import getpass
import os
import sys
from win32com.client import Dispatch

import shutil
from PyQt5 import QtWidgets
from PyQt5 import QtCore
from PyQt5 import QtGui

user_name = getpass.getuser()
AUTORUN_PATH = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % user_name
FILE_NAME = os.path.basename(sys.argv[0])
ICON_NAME = 'scraper.ico'

ORGANIZATION_NAME = 'Example App'
ORGANIZATION_DOMAIN = 'example.com'
APPLICATION_NAME = 'QSettings program'
SETTING_MINIMIZE_TO_TRAY = 'settings/minimize'
SETTING_AUTORUN = 'settings/autorun'


def log_uncaught_exceptions(ex_cls, ex, tb):
    text = '{}: {}:\n'.format(ex_cls.__name__, ex)
    import traceback
    text += ''.join(traceback.format_tb(tb))

    QtWidgets.QMessageBox.critical(None, 'Error', text)
    quit()


def add_to_startup():
    file_path = os.path.dirname(os.path.realpath(__file__)) + "\\" + FILE_NAME
    create_shortcut(file_path)
    shutil.move(file_path.rsplit('.', 1)[0] + '.lnk', AUTORUN_PATH)


def rm_from_startup():
    shortcut_path = f"{AUTORUN_PATH}\{FILE_NAME.rsplit('.', 1)[0] + '.lnk'}"
    if os.path.isfile(shortcut_path):
        os.remove(shortcut_path)


def create_shortcut(path):
    path_to_target = path
    path_to_shortcut = path.rsplit('.', 1)[0] + '.lnk'
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path_to_shortcut)
    shortcut.Targetpath = path_to_target
    shortcut.WorkingDirectory = path.rsplit('\\', 1)[0]
    shortcut.save()


class MainWindow(QtWidgets.QMainWindow):
    tray_icon = None
    minimize_to_tray = None
    autorun = None

    # Переопределяем конструктор класса
    def __init__(self):
        # Обязательно нужно вызвать метод супер класса
        QtWidgets.QMainWindow.__init__(self)

        self.init_ui()
        self.tray()
        self.menu()

    def init_ui(self):
        self.setMinimumSize(QtCore.QSize(480, 80))  # Устанавливаем размеры
        self.setWindowTitle("Sample")  # Устанавливаем заголовок окна
        self.setWindowIcon(QtGui.QIcon(ICON_FILE))

    def menu(self):
        setting_action = QtWidgets.QAction('Settings', self)
        setting_action.triggered.connect(self.settings)

        menu_bar = self.menuBar()
        menu_bar.addAction(setting_action)

    def tray(self):
        self.tray_icon = QtWidgets.QSystemTrayIcon(self)
        self.tray_icon.setIcon(QtGui.QIcon(ICON_FILE))
        self.tray_icon.setToolTip('MyApp')

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
        setting_win = QtWidgets.QWidget(self, QtCore.Qt.Window)

        setting_win.setWindowModality(QtCore.Qt.WindowModal)
        setting_win.setMinimumSize(QtCore.QSize(450, 200))
        setting_win.setWindowTitle("Settings")

        self.minimize_to_tray = QtWidgets.QCheckBox('Minimizing to tray', setting_win)
        self.minimize_to_tray.move(5, 5)

        self.autorun = QtWidgets.QCheckBox('Run with Windows', setting_win)
        self.autorun.move(5, 25)

        settings = QtCore.QSettings()

        check_state = settings.value(SETTING_MINIMIZE_TO_TRAY, False, type=bool)
        self.minimize_to_tray.setChecked(check_state)
        self.minimize_to_tray.clicked.connect(self.save_minimize_settings)

        check_state = settings.value(SETTING_AUTORUN, False, type=bool)
        self.autorun.setChecked(check_state)
        self.autorun.clicked.connect(self.save_autorun_settings)

        setting_win.show()

    def save_minimize_settings(self):
        settings = QtCore.QSettings()
        settings.setValue(SETTING_MINIMIZE_TO_TRAY, self.minimize_to_tray.isChecked())
        settings.sync()

    def save_autorun_settings(self):
        settings = QtCore.QSettings()
        settings.setValue(SETTING_AUTORUN, self.autorun.isChecked())
        check_state = settings.value(SETTING_AUTORUN, False, type=bool)
        if check_state:
            add_to_startup()
        else:
            rm_from_startup()
        settings.sync()

    def hideEvent(self, event):
        settings = QtCore.QSettings()
        check_state = settings.value(SETTING_MINIMIZE_TO_TRAY, False, type=bool)
        if check_state:
            self.hide()

    def closeEvent(self, event):
        self.tray_icon.hide()


if __name__ == "__main__":
    try:
        ico_path = sys._MEIPASS
    except AttributeError:
        ico_path = '.'
    ICON_FILE = os.path.join(ico_path, ICON_NAME)

    sys.excepthook = log_uncaught_exceptions

    QtCore.QCoreApplication.setApplicationName(ORGANIZATION_NAME)
    QtCore.QCoreApplication.setOrganizationDomain(ORGANIZATION_DOMAIN)
    QtCore.QCoreApplication.setApplicationName(APPLICATION_NAME)

    app = QtWidgets.QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())
