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
import func
import conf_file
import conf_ui
import shortcut as s
from loguru import logger
from pathlib import Path
import uuid
import requests
# ---------------- 实例ID + 固定模板 ----------------
INSTANCE_ID = uuid.uuid4().hex          # 32位

CONFIG_TEMPLATE = f"""
    ____                          ____        _       __      ______                 __          ___              _      __ 
   / __ \____ _      _____  _____/ __ \____  (_)___  / /_    /_  __/___  __  _______/ /_        /   |  __________(_)____/ /_
  / /_/ / __ \ | /| / / _ \/ ___/ /_/ / __ \/ / __ \/ __/_____/ / / __ \/ / / / ___/ __ \______/ /| | / ___/ ___/ / ___/ __/
 / ____/ /_/ / |/ |/ /  __/ /  / ____/ /_/ / / / / / /_/_____/ / / /_/ / /_/ / /__/ / / /_____/ ___ |(__  |__  ) (__  ) /_  
/_/    \____/|__/|__/\___/_/  /_/    \____/_/_/ /_/\__/     /_/  \____/\__,_/\___/_/ /_/     /_/  |_/____/____/_/____/\__/  
                                                                                                                            
================================================================================
PowerPointTouchAssist 启动配置
----------------------------------------------------------------------
实例 ID        : {INSTANCE_ID}
日志文件       : PowerPointTouchAssist_{{time:YYYY-MM-DD-HH-mm-ss}}_{INSTANCE_ID}.log
日志等级       : INFO
rotation       : 5 MB
retention      : 30 days
compression    : tar.gz
================================================================================\
""".strip()   # 末尾加 \ 防止空行缝隙

print("""
    ____                          ____        _       __      ______                 __          ___              _      __ 
   / __ \____ _      _____  _____/ __ \____  (_)___  / /_    /_  __/___  __  _______/ /_        /   |  __________(_)____/ /_
  / /_/ / __ \ | /| / / _ \/ ___/ /_/ / __ \/ / __ \/ __/_____/ / / __ \/ / / / ___/ __ \______/ /| | / ___/ ___/ / ___/ __/
 / ____/ /_/ / |/ |/ /  __/ /  / ____/ /_/ / / / / / /_/_____/ / / /_/ / /_/ / /__/ / / /_____/ ___ |(__  |__  ) (__  ) /_  
/_/    \____/|__/|__/\___/_/  /_/    \____/_/_/ /_/\__/     /_/  \____/\__,_/\___/_/ /_/     /_/  |_/____/____/_/____/\__/  
                                                                                                                            
""")
class PathManager:
    def __init__(self, anchor=None):
        if anchor is None:
            anchor = __file__
        self._root = Path(anchor).resolve().parent

    def _get_app_root(self) -> Path:
        return self._root

    def get_log_dir(self, *, create=True) -> Path:
        log_dir = self._root / "logs"
        if create:
            log_dir.mkdir(exist_ok=True)
        return log_dir


def configure_logging():
    log_dir = path_manager.get_log_dir()
    logger.add(
        log_dir / f"PowerPointTouchAssist_{{time:YYYY-MM-DD-HH-mm-ss}}_{INSTANCE_ID}.log",
        rotation="5 MB",
        retention="30 days",
        compression="tar.gz",
        backtrace=True,
        diagnose=True,
        catch=True
    )


def log_software_info():
    logger.info(
        CONFIG_TEMPLATE + "\n" +
        """Hello,World！日志系统启动成功"""
    )

    logger.info("软件启动成功")
    software_info = {"作者": "WRD1145"}
    for k, v in software_info.items():
        logger.info(f"软件{k}: {v}")


path_manager = PathManager(__file__)

# ====================== 单实例检测 ======================
if sys.platform == 'win32':
    import ctypes
    ERROR_ALREADY_EXISTS = 183
    MUTEX_NAME = 'PowerPointTouchAssist_{1BDABE98-2B59-F1AA-CED5-FC3CE6AF205E}'
    kernel32 = ctypes.windll.kernel32
    _mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        print("已经有一个实例了，正在退出此实例")
        logger.info("已经有一个实例了，正在退出此实例")
        sys.exit(0)


def resource(*relative_parts: str) -> str:
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *relative_parts)


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


class TrayIcon(QSystemTrayIcon):
    def __init__(self, main_window: FramelessWindow, parent=None):
        super().__init__(parent)
        logger.info("开始创建托盘图标")
        self.main_window = main_window
        self.set_icon()
        self.menu = QMenu()
        settings_action = self.menu.addAction('设置')
        settings_action.triggered.connect(self.open_settings)
        self.menu.addSeparator()
        quit_action = self.menu.addAction('退出')
        quit_action.triggered.connect(self.quit_app)
        self.setContextMenu(self.menu)
        self.show()
        logger.info("创建托盘图标完成")

    def event(self, e: QEvent):
        if e.type() == QEvent.Type.ContextMenu:
            logger.info("托盘被右键")
        return super().event(e)

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

    def quit_app(self):
        logger.info("软件被关闭")
        QApplication.instance().quit()

    def right_click_settings(self):
        logger.info("右击菜单的设置被点击")

    def right_click_quit(self):
        logger.info("右击菜单的退出被点击")


def run_func():
    func.main()


def main():
    args = sys.argv[1:]
    if 'settings' in args:
        conf_ui.main()
        return

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    window = FramelessWindow()
    tray = TrayIcon(window)

    if int(conf_file.read_conf('Miscellaneous', 'InitialStartUp') or 0) == 1:
        conf_file.write_conf('Miscellaneous', 'InitialStartUp', '0')
        s.add_to_desktop('PowerPoint_TouchAssist.exe')
        s.add_to_startmenu('PowerPoint_TouchAssist.exe')
        s.add_to_startmenu('PowerPoint_TouchAssist.exe',
                           name='PowerPoint 触屏辅助 - 设置', args='settings')
        conf_ui.main()
    else:
        t = threading.Thread(target=run_func, daemon=True)
        t.start()

    window.show()
    logger.info("软件启动")
    sys.exit(app.exec())

if __name__ == '__main__':
    configure_logging()
    log_software_info()
    main()