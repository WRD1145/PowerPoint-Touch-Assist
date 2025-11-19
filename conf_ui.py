"""设置界面（Qt）相关逻辑。

通过 `paths.resource` 加载 UI 文件，读取/写入配置并处理开机自启选项。
"""

from PyQt6.QtWidgets import QDialog, QLineEdit, QPushButton, QComboBox, QCheckBox
from PyQt6 import uic
from PyQt6.QtCore import Qt
import paths
import conf_file as config
import shortcut as a


class FramelessWindow(QDialog):
    """无边框设置窗口，绑定 UI 控件并保存配置。"""

    def __init__(self):
        super().__init__()
        # 使用 paths.resource 加载 UI 文件（兼容打包与源码）
        uic.loadUi(paths.resource('settings.ui'), self)

        # 绑定 DPI 组合框
        self.opt1_Combo = self.findChild(QComboBox, 'opt1_Combo')
        self.opt1_Combo.setCurrentIndex(int(config.read_conf('General', 'DPI') or 0))
        self.opt1_Combo.currentIndexChanged.connect(self.opt1_Save)

        # 绑定 PPT 标题输入框
        self.opt2_LineEdit = self.findChild(QLineEdit, 'opt2_LineEdit')
        self.opt2_LineEdit.setText(config.read_conf('General', 'PPT_Title') or 'PowerPoint 幻灯片放映')
        self.opt2_LineEdit.textChanged.connect(self.opt2_Save)

        # 绑定开机自启复选框
        self.opt3_checkBox = self.findChild(QCheckBox, 'opt3_checkBox')
        self.opt3_checkBox.setChecked(int(config.read_conf('General', 'auto_startup') or 0))
        self.opt3_checkBox.stateChanged.connect(self.opt3_Save)

        # 重置按钮
        self.opt2_Reset = self.findChild(QPushButton, 'opt2Reset')
        self.opt2_Reset.clicked.connect(self.opt2_ResetToDefault)

    # ---------- 配置保存相关方法 ----------
    def opt1_Save(self, idx):
        """保存 DPI 设置索引."""
        config.write_conf('General', 'DPI', str(idx))

    def opt2_Save(self, txt):
        """保存 PPT 标题配置."""
        config.write_conf('General', 'PPT_Title', txt)

    def opt2_ResetToDefault(self):
        """将 PPT 标题恢复为默认并更新输入框."""
        config.write_conf('General', 'PPT_Title', 'PowerPoint 幻灯片放映')
        self.opt2_LineEdit.setText('PowerPoint 幻灯片放映')

    def opt3_Save(self, state):
        """保存开机自启选项并根据选中状态添加或移除自启."""
        checked = state == Qt.CheckState.Checked.value
        config.write_conf('General', 'auto_startup', str(int(checked)))
        (a.add_to_startup if checked else a.remove_from_startup)('PowerPoint_TouchAssist.exe')


def main():
    """显示设置对话框并阻塞直到关闭."""
    dlg = FramelessWindow()
    dlg.exec()