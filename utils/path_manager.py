"""路径管理工具：自动定位应用根目录并提供常用子目录接口（如 logs、data）。"""

from __future__ import annotations
import inspect
import sys
from pathlib import Path
from typing import Union, Optional

StrOrPath = Union[str, Path]

class PathManager:
    """统一的路径管理器。

    功能：
    - 自动定位应用根目录
    - 提供语义化的子目录访问（可选自动创建）
    """

    def __init__(self, anchor: Optional[StrOrPath] = None):
        """初始化 PathManager。

        参数：
            anchor: 可选的锚点路径或文件，默认使用调用者文件所在目录作为根目录。
        """
        if anchor is None:
            # 如果未指定锚点，使用调用者文件的路径作为参考
            frame = inspect.stack()[1]
            anchor = frame.filename
        self._root: Path = Path(anchor).resolve().parent

    # ------------------------------------------------------------------
    # 内部工具
    # ------------------------------------------------------------------
    def _get_app_root(self) -> Path:
        """返回应用根目录（Path 对象）。"""
        return self._root

    # ------------------------------------------------------------------
    # 对外语义化接口
    # ------------------------------------------------------------------
    def get_log_dir(self, *, create: bool = True) -> Path:
        """返回日志目录路径，可选时自动创建目录。"""
        log_dir = self._root / "logs"
        if create:
            log_dir.mkdir(exist_ok=True)
        return log_dir

    def get_data_dir(self, *, create: bool = True) -> Path:
        """返回数据目录路径，可选时自动创建目录。"""
        data_dir = self._root / "data"
        if create:
            data_dir.mkdir(exist_ok=True)
        return data_dir

    # 以后可按需扩展缓存、配置、导出等目录接口