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

# 太懒了直接复制main里的了

def main():
    conf_ui.main()
    return

if __name__ == '__main__':
    main()