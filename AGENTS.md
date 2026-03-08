# AGENTS.md - PowerPoint Touch Assist

> 本文档供 AI 编程助手阅读，帮助快速理解本项目架构和开发规范。

## 项目概述

**PowerPoint Touch Assist (PPT触屏辅助)** 是一个 Windows 桌面应用程序，用于解决 PowerPoint 在 Windows 10/11 上全屏放映时无法通过单击翻页的问题。

当用户在触摸屏或触控设备上点击 PPT 放映区域时，程序会自动模拟空格键按键，实现翻页功能。同时会过滤掉点击任务栏或 PPT 菜单栏的误操作。

## 技术栈

- **语言**: Python 3.11+
- **GUI 框架**: PyQt6 / PySide6
- **UI 设计**: Qt Designer (`.ui` 文件)
- **打包工具**: PyInstaller
- **核心依赖**:
  - `pynput` - 全局鼠标事件监听
  - `pyautogui` - 模拟键盘输入
  - `pygetwindow` - 窗口检测
  - `pywin32` / `win32api` - Windows API 调用
  - `loguru` - 日志记录
  - `requests` - 网络请求（用于更新检查）

## 项目结构

```
PowerPoint-Touch-Assist/
├── main.py              # 主入口，无边框窗口、系统托盘、单例检测
├── func.py              # 核心业务逻辑：鼠标监听和翻页触发
├── conf_file.py         # 配置文件读写（存储于 %APPDATA%/PowerPointTouchAssist）
├── conf_ui.py           # 设置界面逻辑
├── settings.py          # 设置模块入口
├── updater_gui.py       # 自动更新功能
├── shortcut.py          # 快捷方式管理（开机自启、桌面、开始菜单）
├── _version.py          # 版本号定义
├── utils/
│   └── path_manager.py  # 路径管理工具
├── img/                 # 图片资源
│   ├── favicon.ico
│   ├── launch_tip.png
│   └── ...
├── logs/                # 日志目录（运行时生成）
├── .github/workflows/   # CI/CD 配置
│   └── Build.yml        # PR 合并后自动打包
├── *.ui                 # Qt Designer 界面文件
├── *.spec               # PyInstaller 打包配置
├── requirements.txt     # Python 依赖
└── config.ini           # 默认配置文件模板
```

## 核心模块说明

### main.py
- **单实例检测**: 使用 Windows Mutex 防止多开
- **无边框窗口**: 启动时显示提示窗口，3秒后自动隐藏
- **系统托盘**: 提供设置、检查更新、退出功能
- **日志配置**: 使用 loguru，日志保存 30 天，自动轮转压缩

### func.py
- **鼠标监听**: 使用 `pynput.mouse.Listener` 监听全局点击事件
- **智能翻页逻辑**:
  - 左键点击且未滑动 → 发送空格键翻页
  - 过滤点击任务栏（屏幕底部 95px 区域）
  - 过滤点击 PPT 菜单栏（左右两侧 95px 区域）
  - 右键点击视为菜单操作，阻止翻页
- **窗口检测**: 检测是否存在指定标题的 PowerPoint 放映窗口

### conf_file.py
- 配置文件存储路径: `%APPDATA%/PowerPointTouchAssist/config.ini`
- 默认配置项:
  - `General/DPI`: 屏幕缩放比例 (0-4 对应 100%-200%)
  - `General/PPT_Title`: PowerPoint 窗口标题匹配字符串
  - `General/auto_startup`: 开机自启开关
  - `Miscellaneous/InitialStartUp`: 首次启动标记

### updater_gui.py
- 从远程服务器获取版本信息（JSON 格式）
- 支持自动下载、解压、覆盖更新
- Windows 文件占用时，使用延迟脚本（bat）处理
- 更新时保留 `logs/` 和 `config.ini`

## 构建流程

### 本地构建

```bash
# 安装依赖
pip install -r requirements.txt

# 打包（使用命令行）
pyinstaller -w -i icon.ico -n PowerPoint-Touch-Assist ^
  --add-data "main.ui;." ^
  --add-data "settings.ui;." ^
  --add-data "img;img" ^
  --add-data "icon.png;." ^
  main.py

# 或使用 .spec 文件
pyinstaller PowerPoint-Touch-Assist.spec
```

打包输出目录: `dist/PowerPoint-Touch-Assist/`

### CI/CD (GitHub Actions)

- **触发条件**: Pull Request 被合并时
- **运行环境**: windows-latest
- **构建步骤**:
  1. 检出合并后的代码
  2. 设置 Python 3.11
  3. 缓存并安装依赖
  4. PyInstaller 打包
  5. 上传构建产物为 Artifact

## 开发规范

### 代码风格
- 使用中文注释和文档字符串
- 函数和类名使用英文（PEP 8）
- 字符串引号: 单引号或双引号均可，保持一致

### 配置管理
- 用户配置存储在 `%APPDATA%`，避免打包后写入权限问题
- 默认配置内嵌在 `conf_file.py` 的 `DEFAULTS` 字典中
- 使用 `configparser` 读写 INI 格式配置

### 日志规范
- 使用 `loguru` 替代标准库 logging
- 日志级别: DEBUG 用于调试信息，INFO 用于关键流程，ERROR 用于异常
- 日志文件: `logs/PowerPointTouchAssist_YYYY-MM-DD-HH.log`
- 自动轮转: 5MB，保留 30 天，压缩为 tar.gz

### 资源路径处理
- 开发环境和打包环境兼容:
```python
if getattr(sys, 'frozen', False):
    base = sys._MEIPASS  # PyInstaller 打包后的临时目录
else:
    base = os.path.dirname(os.path.abspath(__file__))
```

## 测试说明

本项目**暂无自动化测试套件**。测试主要依赖手动验证:

1. **功能测试**: 启动 PPT 放映，点击屏幕验证翻页
2. **边界测试**: 点击任务栏、菜单栏不应触发翻页
3. **配置测试**: 修改 DPI 和窗口标题，验证生效
4. **更新测试**: 手动触发检查更新流程

## 部署说明

- 目标平台: Windows 10/11
- 分发方式: ZIP 压缩包
- 运行方式: 解压后运行 `PowerPoint-Touch-Assist.exe`
- 权限要求: 无需管理员权限（除创建开机自启快捷方式外）

## 安全注意事项

1. **单实例机制**: 使用 Windows Mutex 避免重复运行
2. **全局鼠标监听**: 程序运行期间会监听所有鼠标点击事件
3. **网络请求**: 更新功能会访问外部服务器获取版本信息
4. **文件操作**: 更新时会替换程序目录下的文件，保留日志和配置

## 版本历史

当前版本: `1.3.0` (定义于 `_version.py`)

版本号格式: Semantic Versioning (MAJOR.MINOR.PATCH)

---

*本文档最后更新: 2026-01-30*
