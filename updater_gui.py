#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyQt6 带 UI 的自更新示例
author : github.com/yourname
"""
import json
import os
import sys
import time
import signal
import platform
from pathlib import Path
from typing import Optional

import requests
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QProgressBar, QTextEdit)
from PyQt6.QtGui import QDesktopServices

# ----------- 配置区 -----------
CURRENT_VERSION = "1.0.0"
VERSION_JSON_URL = "https://wrd1145.dev/version.json"   # 远程版本信息
# 主程序路径（打包后 exe / 源码脚本）
MAIN_PATH = Path(sys.executable) if getattr(sys, 'frozen', False) else Path(__file__).resolve()
# ----------------------------


class DownloadThread(QThread):
    """后台下载线程"""
    progress = pyqtSignal(int)          # 0-100
    finished = pyqtSignal(Path)         # 返回临时文件路径
    error = pyqtSignal(str)

    def __init__(self, url: str):
        super().__init__()
        self.url = url

    def run(self):
        try:
            resp = requests.get(self.url, stream=True, timeout=15)
            resp.raise_for_status()
            total = int(resp.headers.get('content-length', 0))
            tmp = MAIN_PATH.with_suffix('.tmp')
            down_bytes = 0
            with open(tmp, 'wb') as f:
                for chunk in resp.iter_content(8192):
                    if not chunk:
                        continue
                    f.write(chunk)
                    down_bytes += len(chunk)
                    if total:
                        self.progress.emit(int(down_bytes * 100 / total))
            self.finished.emit(tmp)
        except Exception as e:
            self.error.emit(str(e))


class UpdaterWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("在线更新示例")
        self.resize(500, 350)

        self.latest_ver: Optional[str] = None
        self.download_url: Optional[str] = None
        self.changelog: str = ""

        self._init_ui()
        self._check_version()

    # ---------- UI ----------
    def _init_ui(self):
        lay = QVBoxLayout(self)

        self.lb_cur = QLabel(f"当前版本：{CURRENT_VERSION}")
        self.lb_remote = QLabel("远程版本：检查中…")
        self.btn_update = QPushButton("立即更新")
        self.btn_update.setEnabled(False)
        self.btn_update.clicked.connect(self._start_update)

        self.bar = QProgressBar()
        self.bar.setVisible(False)

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        lay.addWidget(self.lb_cur)
        lay.addWidget(self.lb_remote)
        lay.addWidget(self.btn_update)
        lay.addWidget(self.bar)
        lay.addWidget(QLabel("更新日志："))
        lay.addWidget(self.log)

    # ---------- 版本检查 ----------
    def _check_version(self):
        self.log.append("正在检查版本…")
        try:
            resp = requests.get(VERSION_JSON_URL, timeout=10)
            resp.raise_for_status()
            info = resp.json()
            self.latest_ver = info["version"]
            self.download_url = info["url"]
            self.changelog = info.get("changelog", "暂无更新说明")
        except Exception as e:
            self.log.append(f"<font color=red>获取版本信息失败：{e}</font>")
            return

        self.lb_remote.setText(f"远程版本：{self.latest_ver}")
        if CURRENT_VERSION < self.latest_ver:
            self.btn_update.setEnabled(True)
            self.log.append(f"<font color=green>发现新版本 {self.latest_ver}</font>")
            self.log.append(self.changelog)
        else:
            self.log.append("当前已是最新版本。")

    # ---------- 更新 ----------
    def _start_update(self):
        if not self.download_url:
            return
        self.btn_update.setEnabled(False)
        self.bar.setVisible(True)
        self.log.append("开始下载…")
        self.dl_thread = DownloadThread(self.download_url)
        self.dl_thread.progress.connect(self.bar.setValue)
        self.dl_thread.finished.connect(self._on_downloaded)
        self.dl_thread.error.connect(self._on_error)
        self.dl_thread.start()

    def _on_downloaded(self, tmp_path: Path):
        self.log.append("下载完成，正在替换主程序…")
        # 根据平台选择覆盖策略
        if platform.system() == "Windows":
            self._win_delayed_replace(tmp_path)
        else:
            try:
                os.replace(tmp_path, MAIN_PATH)
                self.log.append("替换成功，3 秒后重启…")
                QThread.msleep(1500)
                self._restart()
            except Exception as e:
                self._on_error(f"替换失败：{e}")

    def _win_delayed_replace(self, tmp: Path):
        """Windows 专用：生成 bat，延迟 2 s 后覆盖并重启"""
        bat = MAIN_PATH.with_name("updater.bat")
        bat.write_text(f"""@echo off
timeout /t 2 /nobreak > NUL
move /Y "{tmp}" "{MAIN_PATH}"
start "" "{MAIN_PATH}"
del "{bat}"
""", encoding="gbk")
        os.startfile(str(bat))
        self.log.append("已生成升级脚本，程序即将退出…")
        QApplication.quit()

    def _restart(self):
        """重启自己"""
        QApplication.quit()
        os.execv(sys.executable, [sys.executable] + sys.argv)

    def _on_error(self, msg: str):
        self.log.append(f"<font color=red>错误：{msg}</font>")
        self.bar.setVisible(False)
        self.btn_update.setEnabled(True)


# ----------------- 入口 -----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = UpdaterWindow()
    w.show()
    sys.exit(app.exec())