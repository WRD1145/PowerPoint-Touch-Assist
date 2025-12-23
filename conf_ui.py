"""设置界面（Qt）相关逻辑"""

from PyQt6.QtWidgets import QDialog, QLineEdit, QPushButton, QComboBox, QCheckBox
from PyQt6 import uic
from PyQt6.QtCore import Qt
from pathlib import Path
import conf_file as config
import shortcut as a

_UI_FILE = Path(__file__).with_name('settings.ui')   


class FramelessWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi(_UI_FILE, self)   

        self.opt1_Combo = self.findChild(QComboBox, 'opt1_Combo')
        self.opt1_Combo.setCurrentIndex(int(config.read_conf('General', 'DPI') or 0))
        self.opt1_Combo.currentIndexChanged.connect(self.opt1_Save)

        self.opt2_LineEdit = self.findChild(QLineEdit, 'opt2_LineEdit')
        self.opt2_LineEdit.setText(config.read_conf('General', 'PPT_Title') or 'PowerPoint 幻灯片放映')
        self.opt2_LineEdit.textChanged.connect(self.opt2_Save)

        self.opt3_checkBox = self.findChild(QCheckBox, 'opt3_checkBox')
        self.opt3_checkBox.setChecked(int(config.read_conf('General', 'auto_startup') or 0))
        self.opt3_checkBox.stateChanged.connect(self.opt3_Save)

        self.opt2_Reset = self.findChild(QPushButton, 'opt2Reset')
        self.opt2_Reset.clicked.connect(self.opt2_ResetToDefault)

    
    def opt1_Save(self, idx):
        config.write_conf('General', 'DPI', str(idx))

    def opt2_Save(self, txt):
        config.write_conf('General', 'PPT_Title', txt)

    def opt2_ResetToDefault(self):
        config.write_conf('General', 'PPT_Title', 'PowerPoint 幻灯片放映')
        self.opt2_LineEdit.setText('PowerPoint 幻灯片放映')

    def opt3_Save(self, state):
        checked = state == Qt.CheckState.Checked.value
        config.write_conf('General', 'auto_startup', str(int(checked)))
        (a.add_to_startup if checked else a.remove_from_startup)('PowerPoint_TouchAssist.exe')


def main():
    dlg = FramelessWindow()
    dlg.exec()