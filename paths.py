import os
import sys


def resource(*relative_parts: str) -> str:
    """返回资源的绝对路径。

    在使用 PyInstaller 打包后的环境里，资源位于运行时生成的临时目录（sys._MEIPASS）。
    在源码运行时，资源相对于当前脚本文件所在目录。

    参数:
        *relative_parts: 与资源文件的相对路径片段（类似 os.path.join 的参数）。

    返回:
        资源文件的绝对路径字符串。
    """
    # 如果应用已被 PyInstaller 打包，使用解压后的临时目录
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS  # PyInstaller 运行时的临时目录
    else:
        # 源码运行时，资源基于当前文件所在目录
        base = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base, *relative_parts)