"""配置文件读写工具。

将配置文件放在用户的 %APPDATA%/PowerPointTouchAssist 目录，避免打包后写入受限。
提供读写接口：read_conf / write_conf。
"""

import os
import configparser as config
from typing import Optional

# 配置文件目录与路径（放在用户 APPDATA 下）
CONFIG_DIR = os.path.join(os.environ.get('APPDATA', ''), 'PowerPointTouchAssist')
CONFIG_PATH = os.path.join(CONFIG_DIR, 'config.ini')
os.makedirs(CONFIG_DIR, exist_ok=True)

# DPI 缩放映射表（配置项 -> 缩放倍数）
dpi_dict = {
    '0': 1,
    '1': 1.25,
    '2': 1.5,
    '3': 1.75,
    '4': 2
}

# 默认配置
DEFAULTS = {
    'General': {
        'DPI': '0',
        'PPT_Title': 'PowerPoint 幻灯片放映',
        'auto_startup': '0'
    },
    'Miscellaneous': {
        'InitialStartUp': '1'
    }
}


def _ensure_conf():
    """确保配置文件存在；不存在时写入默认配置."""
    if os.path.isfile(CONFIG_PATH):
        return
    cfg = config.ConfigParser()
    cfg.read_dict(DEFAULTS)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        cfg.write(f)


def read_conf(section: str = 'General', key: str = '') -> Optional[str]:
    """读取配置值。

    优先从用户配置文件读取，读取不到时回退到内置的 DEFAULTS。

    参数:
        section: 配置节名
        key: 配置项名

    返回:
        配置值的字符串或 None（如果既不在文件也不在默认配置中）。
    """
    _ensure_conf()
    cfg = config.ConfigParser()
    cfg.read(CONFIG_PATH, encoding='utf-8')
    return cfg.get(section, key, fallback=DEFAULTS.get(section, {}).get(key))


def write_conf(section: str, key: str, value) -> None:
    """写入（或更新）配置项并保存到用户配置文件."""
    _ensure_conf()
    cfg = config.ConfigParser()
    cfg.read(CONFIG_PATH, encoding='utf-8')
    if not cfg.has_section(section):
        cfg.add_section(section)
    cfg.set(section, key, str(value))
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        cfg.write(f)