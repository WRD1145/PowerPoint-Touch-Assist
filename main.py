# ======================  main.py  ======================
import os
import sys
import json
import time
import multiprocessing
from PyQt6.QtWidgets import QApplication, QWidget, QSystemTrayIcon, QMenu
from PyQt6 import uic
from PyQt6.QtCore import Qt, QTimer, QEvent
from PyQt6.QtGui import QCursor, QIcon
import threading
import uuid
from pathlib import Path
from loguru import logger
import conf_file
import conf_ui
from configparser import ConfigParser

# --------------------------------------------------
# ① 更新窗口模块（同级目录）
# --------------------------------------------------
from updater_gui import UpdaterWindow

# --------------------------------------------------
# ② 实例 ID & 日志模板
# --------------------------------------------------
INSTANCE_ID = uuid.uuid4().hex
CONFIG_TEMPLATE = f"""
实例 ID        : {INSTANCE_ID}
""".strip()

print("""
    ____                          ____        _       __      ______                 __          ___              _      __ 
   / __ \____ _      _____  _____/ __ \____  (_)___  / /_    /_  __/___  __  _______/ /_        /   |  __________(_)____/ /_
  / /_/ / __ \ | /| / / _ \/ ___/ /_/ / __ \/ / __ \/ __/_____/ / / __ \/ / / / ___/ __ \______/ /| | / ___/ ___/ / ___/ __/
 / ____/ /_/ / |/ |/ /  __/ /  / ____/ /_/ / / / / / /_/_____/ / / /_/ / /_/ / /__/ / / /_____/ ___ |(__  |__  ) (__  ) /_  
/_/    \____/|__/|__/\___/_/  /_/    \____/_/_/ /_/\__/     /_/  \____/\__,_/\___/_/ /_/     /_/  |_/____/____/_/____/\__/  

钟表的指针周而复始，就像人的困惑、烦恼、软弱…摇摆不停。但最终，人们依旧要前进，就像你的指针，永远落在前方。

""")

conf = ConfigParser()
conf.read('config.ini', encoding='utf-8')
version = conf['About']['version']

# --------------------------------------------------
# ③ PathManager & 日志
# --------------------------------------------------
class PathManager:
    def __init__(self, anchor=None):
        anchor = anchor or __file__
        self._root = Path(anchor).resolve().parent

    def get_log_dir(self, *, create=True) -> Path:
        log_dir = self._root / "logs"
        if create:
            log_dir.mkdir(exist_ok=True)
        return log_dir


path_manager = PathManager()


def configure_logging():
    log_dir = path_manager.get_log_dir()
    logger.add(
        log_dir / f"PowerPointTouchAssist_{{time:YYYY-MM-DD-HH-mm-ss}}.log",
        rotation="5 MB",
        retention="30 days",
        compression="tar.gz",
        backtrace=True,
        diagnose=True,
        catch=True
    )


def log_software_info():
    logger.info(CONFIG_TEMPLATE + "\n日志系统启动成功")
    logger.info("软件启动成功")
    logger.info("作者: WRD1145")


# --------------------------------------------------
# ④ 单实例检测
# --------------------------------------------------
if sys.platform == 'win32':
    import ctypes

    ERROR_ALREADY_EXISTS = 183
    MUTEX_NAME = 'PowerPointTouchAssist_{1BDABE98-2B59-F1AA-CED5-FC3CE6AF205E}'
    kernel32 = ctypes.windll.kernel32
    _mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        logger.info("已经有一个实例了，正在退出此实例")
        sys.exit(0)


def resource(*relative_parts: str) -> str:
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *relative_parts)


# --------------------------------------------------
# ⑤ 主窗口（无边框）
# --------------------------------------------------
class FramelessWindow(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi(resource('main.ui'), self)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.m_flag = False
        QTimer.singleShot(3000, self.hide)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.m_flag = True
            self.m_Position = event.globalPosition().toPoint() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        if Qt.MouseButton.LeftButton and self.m_flag:
            self.move(event.globalPosition().toPoint() - self.m_Position)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.m_flag = False
        self.setCursor(QCursor(Qt.CursorShape.ArrowCursor))


# --------------------------------------------------
# ⑥ 系统托盘
# --------------------------------------------------
class TrayIcon(QSystemTrayIcon):
    def __init__(self, main_window: FramelessWindow, parent=None):
        super().__init__(parent)
        logger.info("开始创建托盘图标")
        self.main_window = main_window
        self.set_icon()
        self.menu = QMenu()

        settings_action = self.menu.addAction('设置')
        settings_action.triggered.connect(self.open_settings)

        # ★ 新增「检查更新」
        self.menu.addSeparator()
        update_action = self.menu.addAction('检查更新')
        update_action.triggered.connect(self.open_updater)

        quit_action = self.menu.addAction('退出')
        quit_action.triggered.connect(self.quit_app)
        self.setContextMenu(self.menu)
        self.show()
        logger.info("创建托盘图标完成")

    def set_icon(self):
        icon_path = resource('icon.png')
        if os.path.isfile(icon_path):
            self.setIcon(QIcon(icon_path))
        else:
            self.setIcon(QIcon())

    def on_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.toggle_window()

    def toggle_window(self):
        if self.main_window.isVisible():
            self.main_window.hide()
        else:
            self.main_window.show()
            self.main_window.move(QCursor.pos() - self.main_window.rect().center())
            self.main_window.raise_()
            self.main_window.activateWindow()

    def open_settings(self):
        logger.info("软件设置被打开")
        conf_ui.main()
        logger.info("软件设置被关闭")

    # ★ 弹出更新窗口
    def open_updater(self):
        logger.info("用户手动打开更新窗口")
        self.updater_win = UpdaterWindow(current_ver= version)  # 版本号可动态读
        self.updater_win.show()

    def quit_app(self):
        logger.info("软件被关闭")
        QApplication.instance().quit()


# --------------------------------------------------
# ⑦ 业务线程（示例）
# --------------------------------------------------
def run_func():
    # 这里放你的核心业务，例如监听 PowerPoint 触屏
    logger.info("业务线程启动")


# --------------------------------------------------
# ⑧ 主入口
# --------------------------------------------------
def main():
    args = sys.argv[1:]
    if 'settings' in args:
        # conf_ui.main()   # 若不存在可注释
        return

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    window = FramelessWindow()
    tray = TrayIcon(window)

    # 首次启动快捷方式 & 设置
    if int(conf_file.read_conf('Miscellaneous', 'InitialStartUp') or 0) == 1:
        conf_file.write_conf('Miscellaneous', 'InitialStartUp', '0')
        # 若 shortcut 模块不存在，下面三行可注释
        # s.add_to_desktop('PowerPoint_TouchAssist.exe')
        # s.add_to_startmenu('PowerPoint_TouchAssist.exe')
        # s.add_to_startmenu('PowerPoint_TouchAssist.exe', name='PowerPoint 触屏辅助 - 设置', args='settings')
        # conf_ui.main()
    else:
        t = threading.Thread(target=run_func, daemon=True)
        t.start()

    window.show()
    logger.info("软件启动")
    sys.exit(app.exec())


# --------------------------------------------------
# ⑨ 启动
# --------------------------------------------------
if __name__ == '__main__':
    configure_logging()
    log_software_info()
    main()