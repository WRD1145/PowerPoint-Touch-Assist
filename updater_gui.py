#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
updater_gui.py  ——  Windows 风格更新窗口
功能：下载 zip → 解压 → 除 logs & config.ini 外全部覆盖 → 重启
      若主程序被占用（Windows）→ 生成延迟 bat 脚本完成覆盖与重启
"""
import json
import os
import sys
import platform
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Optional

import requests
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QCoreApplication
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QProgressBar, QTextEdit, QApplication)
from PyQt6.QtGui import QIcon

# -------------- 日志 --------------
from loguru import logger

# -------------- 读 config.ini --------------
from configparser import ConfigParser

conf = ConfigParser()
conf.read('config.ini', encoding='utf-8')
CURRENT_VERSION = conf.get('About', 'version', fallback='1.0.0')
DEFAULT_VERSION_URL = "https://wrd1145.dev/version.json"
# 主程序路径
MAIN_PATH = Path(sys.executable) if getattr(sys, 'frozen', False) else Path(__file__).resolve()
APP_DIR = MAIN_PATH.parent


# ------------------------------------------------------------------
# 下载线程
# ------------------------------------------------------------------
class _DownloadThread(QThread):
    progress = pyqtSignal(int)          # 0-100
    finished = pyqtSignal(Path)
    error = pyqtSignal(str)

    def __init__(self, url: str):
        super().__init__()
        self.url = url

    def run(self):
        try:
            resp = requests.get(self.url, stream=True, timeout=15)
            resp.raise_for_status()
            total = int(resp.headers.get('content-length', 0))
            tmp = MAIN_PATH.with_suffix('.tmp.zip')
            down = 0
            with open(tmp, 'wb') as f:
                for chunk in resp.iter_content(8192):
                    if not chunk:
                        continue
                    f.write(chunk)
                    down += len(chunk)
                    if total:
                        self.progress.emit(int(down * 100 / total))
            logger.info(f"下载完成：{tmp}")
            self.finished.emit(tmp)
        except Exception as e:
            logger.error(f"下载失败：{e}")
            self.error.emit(str(e))


# ------------------------------------------------------------------
# 更新窗口
# ------------------------------------------------------------------
class UpdaterWindow(QWidget):
    def __init__(self,
                 current_ver: str = CURRENT_VERSION,
                 version_url: str = DEFAULT_VERSION_URL,
                 parent=None):
        super().__init__(parent)
        self.current_ver = current_ver
        self.version_url = version_url
        self.latest_ver: Optional[str] = None
        self.download_url: Optional[str] = None
        self.changelog: str = ""

        # 普通窗口，保留系统标题栏
        self.setWindowTitle("软件更新")
        self.setFixedSize(750, 420)
        logger.info("更新窗口初始化完成")

        self._init_ui()
        self._check_version()

    # ---------------- UI ----------------
    def _init_ui(self):
        main = QVBoxLayout(self)

        # 顶部标题
        top = QHBoxLayout()
        self.lb_title = QLabel("软件更新")
        self.lb_title.setStyleSheet("font: 22pt 'Microsoft YaHei';font-weight:bold;")
        top.addWidget(self.lb_title)
        top.addStretch()
        self.lb_star = QLabel("*")
        self.lb_star.setStyleSheet("color:#ff5050;font:16pt;")
        top.addWidget(self.lb_star)

        # 中部左右分栏
        mid = QHBoxLayout()
        left = QVBoxLayout()
        right = QVBoxLayout()

        # 左侧
        self.lb_sub1 = QLabel("版本")
        self.lb_sub1.setStyleSheet("font:17pt 'Microsoft YaHei';")
        left.addWidget(self.lb_sub1)
        self.lb_cur = QLabel(f"当前版本：{self.current_ver}")
        self.lb_remote = QLabel("远程版本：检查中…")
        for lb in (self.lb_cur, self.lb_remote):
            lb.setStyleSheet("font:12pt 'Microsoft YaHei';")
            left.addWidget(lb)
        self.bar = QProgressBar()
        self.bar.setVisible(False)
        self.bar.setStyleSheet("""
            QProgressBar{border-radius:4px;background:#555;}
            QProgressBar::chunk{border-radius:4px;background:#ff5050;}
        """)
        left.addWidget(self.bar)

        # 右侧
        self.lb_sub2 = QLabel("日志")
        self.lb_sub2.setStyleSheet("font:17pt 'Microsoft YaHei';")
        right.addWidget(self.lb_sub2)
        self.log = QTextEdit(readOnly=True)
        self.log.setStyleSheet("background:#f5f5f5;border:1px solid #ccc;border-radius:6px;padding:6px;")
        self.log.setMaximumHeight(180)
        right.addWidget(self.log)

        mid.addLayout(left)
        mid.addLayout(right)

        # 底部按钮
        bot = QHBoxLayout()
        self.btn = QPushButton("立即更新")
        self.btn.setEnabled(False)
        self.btn.setFixedSize(160, 40)
        self.btn.setStyleSheet("""
            QPushButton{background:#ff5050;color:white;border-radius:6px;font:12pt;font-weight:bold;}
            QPushButton:hover{background:#ff7070;}
            QPushButton:pressed{background:#cc4040;}
            QPushButton:disabled{background:#ccc;color:#666;}
        """)
        self.btn.clicked.connect(self._do_update)
        bot.addStretch()
        bot.addWidget(self.btn)
        bot.addStretch()

        # 底部标语
        self.lb_foot = QLabel("""就算结局早已注定，但在此之前，在走向结局的路上，人能做的事同样很多。
而结局也会因此展现截然不同的意义。""")
        self.lb_foot.setStyleSheet("color:gray;font:10pt;")
        bot.addWidget(self.lb_foot)

        # logo 行
        foot = QHBoxLayout()
        self.lb_logo = QLabel()
        self.lb_logo.setPixmap(QIcon("img/favicon.ico").pixmap(30, 30))
        self.lb_txt = QLabel("PPT 触屏辅助")
        self.lb_txt.setStyleSheet("color:gray;font:12pt;font-weight:bold;")
        foot.addWidget(self.lb_logo)
        foot.addWidget(self.lb_txt)
        foot.addStretch()

        main.addLayout(top)
        main.addLayout(mid)
        main.addStretch()
        main.addLayout(bot)
        main.addLayout(foot)

    # ---------------- 版本检查 ----------------
    def _check_version(self):
        self.log.append("正在检查版本…")
        try:
            info = requests.get(self.version_url, timeout=10).json()
            self.latest_ver = info["version"]
            self.download_url = info["url"]
            self.changelog = info.get("changelog", "暂无说明")
        except Exception as e:
            logger.error(f"获取版本信息失败：{e}")
            self.log.append(f"<font color=red>检查失败：{e}</font>")
            return

        self.lb_remote.setText(f"远程版本：{self.latest_ver}")
        if self.current_ver < self.latest_ver:
            self.btn.setEnabled(True)
            self.log.append(f"<font color=green>发现新版本 {self.latest_ver}</font>")
            self.log.append(self.changelog)
            logger.info(f"发现新版本：{self.latest_ver}")
        else:
            self.log.append("已是最新版本。")
            logger.info("当前已是最新版本")

    # ---------------- 开始下载 ----------------
    def _do_update(self):
        if not self.download_url:
            return
        self.btn.setEnabled(False)
        self.bar.setVisible(True)
        self.log.append("开始下载…")
        logger.info(f"开始下载更新包：{self.download_url}")
        self.thread = _DownloadThread(self.download_url)
        self.thread.progress.connect(self.bar.setValue)
        self.thread.finished.connect(self._on_downloaded)
        self.thread.error.connect(self._on_error)
        self.thread.start()

    # ---------------- 下载完成 → 覆盖 ----------------
    def _on_downloaded(self, tmp: Path):
        """tmp 是下载到的 zip 路径"""
        logger.info(f"下载完成：{tmp}")
        self.log.append("下载完成，开始解压并覆盖…")

        # 1. 解压到临时目录
        with zipfile.ZipFile(tmp, 'r') as zf:
            temp_dir = Path(tempfile.mkdtemp(prefix='update_'))
            zf.extractall(temp_dir)
            logger.info(f"解压完成：{temp_dir}")

        # 2. 忽略列表（不覆盖）
        ignore_patterns = {'logs', 'config.ini'}

        def ignore_func(src, names):
            return [n for n in names if n in ignore_patterns]

        # 3. 覆盖（除忽略外全部）
        app_dir = MAIN_PATH.parent
        for item in temp_dir.iterdir():
            dst = app_dir / item.name
            try:
                if item.is_dir():
                    if dst.exists():
                        logger.info(f"删除旧目录：{dst}")
                        shutil.rmtree(dst)
                    logger.info(f"复制目录：{item} → {dst}")
                    shutil.copytree(item, dst, ignore=ignore_func)
                else:
                    if item.name not in ignore_patterns:
                        logger.info(f"复制文件：{item} → {dst}")
                        shutil.copy2(item, dst)
            except Exception as e:
                logger.warning(f"复制失败（可能被占用）：{e} → 将使用延迟脚本")
                # 一旦复制失败（Windows 占用）立即切到 bat 方案
                self._win_delayed_replace(tmp, temp_dir)
                return

        # 4. 清理
        shutil.rmtree(temp_dir)
        tmp.unlink(missing_ok=True)
        logger.info("覆盖完成，准备重启")
        self.log.append("覆盖完成，3 秒后重启…")
        self.thread.msleep(1500)
        self._restart()

    # ---------------- Windows 占用时的延迟覆盖 ----------------
    def _win_delayed_replace(self, zip_path: Path, temp_dir: Path):
        logger.info("生成延迟脚本（bat）以解决文件占用")
        bat = APP_DIR / "updater.bat"
        bat.write_text(f"""@echo off
timeout /t 2 /nobreak > NUL
powershell -command "Expand-Archive -Path '{zip_path}' -DestinationPath '{APP_DIR}' -Force"
del "{zip_path}"
start "" "{MAIN_PATH}"
del "{bat}"
""", encoding='gbk')
        logger.info(f"启动延迟脚本：{bat}")
        os.startfile(str(bat))
        QApplication.quit()          # 旧进程立即退出，bat 接管后续

    # ---------------- 重启自己 ----------------
    def _restart(self):
        logger.info("重启主程序")
        QCoreApplication.quit()
        os.execv(sys.executable, [sys.executable] + sys.argv)

    # ---------------- 下载错误 ----------------
    def _on_error(self, msg: str):
        logger.error(f"更新失败：{msg}")
        self.log.append(f"<font color=red>错误：{msg}</font>")
        self.bar.setVisible(False)
        self.btn.setEnabled(True)