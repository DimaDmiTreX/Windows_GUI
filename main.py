import getpass
import os
import sys

from PyQt5 import QtWidgets
from PyQt5 import QtCore


def add_to_startup(name_bat, file_path=""):
    if file_path == "":
        file_name = __file__.rsplit('.', 1)[0] + '.exe'
        file_path = f'"{os.path.dirname(os.path.realpath(__file__))}\{file_name}"'
    user_name = getpass.getuser()
    bat_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % user_name
    with open(bat_path + '\\' + f"{name_bat}.bat", 'w+') as bat_file:
        bat_file.write(r'start "" %s' % file_path)


def rm_from_startup(name_bat):
    user_name = getpass.getuser()
    bat_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % user_name
    bat_path = bat_path + '\\' + f"{name_bat}.bat"
    if os.path.isfile(bat_path):
        os.remove(bat_path)


class MainWindow(QtWidgets.QMainWindow):
    check_box = None
    tray_icon = None

    # Переопределяем конструктор класса
    def __init__(self):
        # Обязательно нужно вызвать метод супер класса
        QtWidgets.QMainWindow.__init__(self)

        self.setMinimumSize(QtCore.QSize(480, 80))  # Устанавливаем размеры
        self.setWindowTitle("Sample")  # Устанавливаем заголовок окна
        self.setWindowIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DesktopIcon))

        central_widget = QtWidgets.QWidget(self)  # Создаём центральный виджет
        self.setCentralWidget(central_widget)  # Устанавливаем центральный виджет

        grid_layout = QtWidgets.QGridLayout(self)  # Создаём QGridLayout
        central_widget.setLayout(grid_layout)  # Устанавливаем данное размещение в центральный виджет

        # Добавляем чекбокс, от которого будет зависеть поведение программы при закрытии окна
        self.check_box = QtWidgets.QCheckBox('Minimize to Tray')
        grid_layout.addWidget(self.check_box, 0, 0)
        grid_layout.addItem(
            QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding), 2, 0)

        self.tray()
        self.menu()

    def menu(self):
        setting_action = QtWidgets.QAction('Settings', self)
        setting_action.triggered.connect(self.settings_show)

        menu_bar = self.menuBar()
        menu_bar.addAction(setting_action)

    def tray(self):
        # Инициализируем QSystemTrayIcon
        self.tray_icon = QtWidgets.QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_DesktopIcon))
        self.tray_icon.setToolTip('MyApp')

        '''
            Объявим и добавим действия для работы с иконкой системного трея
            show - показать окно
            hide - скрыть окно
            exit - выход из программы
        '''
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

    def settings_show(self):
        settings = QtWidgets.QWidget(self, QtCore.Qt.Window)
        settings.setWindowModality(QtCore.Qt.WindowModal)
        settings.setMinimumSize(QtCore.QSize(480, 80))  # Устанавливаем размеры
        settings.setWindowTitle("Settings")
        settings.show()

    # Переопределение метода closeEvent, для перехвата события закрытия окна
    # Окно будет закрываться только в том случае, если нет галочки в чекбоксе
    def closeEvent(self, event):
        if not self.check_box.isChecked() or not self.isVisible():
            self.tray_icon.hide()
        else:
            event.ignore()
            self.hide()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())
