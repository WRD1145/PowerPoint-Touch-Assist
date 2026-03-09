# telemetry.py
import sentry_sdk
import os
import sys
import json
import atexit
from typing import Optional, Dict, Any
from loguru import logger

try:
    from _version import __version__ as VERSION
except ImportError:
    VERSION = "unknown"

_is_initialized = False
_environment = None


def init_telemetry():
    """
    初始化遥测：可选配置，无DSN时静默跳过
    - 开发：读取 SENTRY_DSN 环境变量（未设置则跳过）
    - 生产：PyInstaller 打包后读取 resources/telemetry.json（不存在则跳过）
    """
    global _is_initialized, _environment
    
    if _is_initialized:
        return True
    
    # 清理代理干扰
    for var in ['HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy']:
        os.environ.pop(var, None)
    
    config = _get_config()
    _environment = config.get("environment", "unknown")
    dsn = config.get("dsn")
    
    # 无DSN时静默跳过（Debug级别记录，不影响正常使用）
    if not dsn:
        logger.debug(f"[Telemetry] 未配置，已跳过 [{_environment}]")
        _is_initialized = True  # 标记为已处理，避免重复检查
        return False
    
    try:
        sentry_sdk.init(
            dsn=dsn,
            environment=_environment,
            release=config.get("release", VERSION),
            traces_sample_rate=config.get("traces_sample_rate", 0.1 if _environment == "production" else 0.0),
            shutdown_timeout=5,
        )
        
        atexit.register(lambda: sentry_sdk.flush(timeout=3))
        _is_initialized = True
        
        logger.info(f"[Telemetry] 已初始化 [{_environment}@{config.get('release', VERSION)}]")
        
        # 验证连接
        try:
            event_id = sentry_sdk.capture_message(f"Start {VERSION}", level="info")
            sentry_sdk.flush(timeout=2)
            logger.debug(f"[Telemetry] 验证事件已发送 [{event_id}]")
        except Exception:
            pass  # 验证失败不影响主程序
            
        return True
        
    except Exception as e:
        # 初始化失败也静默处理，仅debug记录
        logger.debug(f"[Telemetry] 初始化失败: {e}")
        _is_initialized = True
        return False


def _get_config():
    """获取配置：环境变量(开发) > 配置文件(生产) > 无（静默）"""
    
    # 开发环境：检查环境变量
    dsn = os.environ.get("SENTRY_DSN", "").strip()
    if dsn:
        return {
            "dsn": dsn,
            "environment": os.environ.get("ENV", "development").strip(),
            "release": os.environ.get("RELEASE", VERSION).strip() or VERSION,
            "traces_sample_rate": 0.0
        }
    
    # 生产环境（PyInstaller 打包后）
    if getattr(sys, 'frozen', False):
        try:
            base_path = sys._MEIPASS
            config_path = os.path.join(base_path, "resources", "telemetry.json")
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    config.setdefault("environment", "production")
                    return config
        except Exception:
            pass  # 配置文件不存在或损坏也静默处理
    
    # 无配置
    return {"dsn": None, "environment": "unknown"}


def report(error=None, message=None, level="error", tags=None, extra=None):
    """
    上报事件（未初始化时静默跳过）
    """
    if not sentry_sdk.Hub.current.client:
        return None
    
    try:
        with sentry_sdk.push_scope() as scope:
            if tags:
                for k, v in (tags or {}).items():
                    scope.set_tag(k, str(v))
            if extra:
                for k, v in (extra or {}).items():
                    scope.set_extra(k, v)
            
            scope.set_extra("app_version", VERSION)
            
            if error:
                return sentry_sdk.capture_exception(error)
            if message:
                return sentry_sdk.capture_message(message, level=level)
    except Exception:
        return None


def get_environment():
    """获取当前环境: development/production/unknown"""
    return _environment or "unknown"


def is_production():
    return _environment == "production"


def is_development():
    return _environment == "development"