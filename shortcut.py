import os
from win32com.client import Dispatch

program_name = 'PowerPoint 触屏辅助'


def add_to_startup(file_path="", icon_path=""):
    """添加开机自启快捷方式"""
    if file_path == "":
        file_path = os.path.realpath(__file__)
    else:
        file_path = os.path.abspath(file_path)

    if icon_path == "":
        icon_path = file_path

    startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    shortcut_path = os.path.join(startup_folder, f'{program_name}.lnk')

    # 先删再建
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = file_path
    shortcut.WorkingDirectory = os.path.dirname(file_path)
    shortcut.IconLocation = icon_path
    shortcut.save()


def remove_from_startup():
    """移除开机自启快捷方式"""
    startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    shortcut_path = os.path.join(startup_folder, f'{program_name}.lnk')
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)


def add_to_desktop(file_path="", icon_path=""):
    """添加桌面快捷方式"""
    if file_path == "":
        file_path = os.path.realpath(__file__)
    else:
        file_path = os.path.abspath(file_path)

    if icon_path == "":
        icon_path = file_path

    desktop_folder = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    shortcut_path = os.path.join(desktop_folder, f'{program_name}.lnk')

    # 先删再建
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = file_path
    shortcut.WorkingDirectory = os.path.dirname(file_path)
    shortcut.IconLocation = icon_path
    shortcut.save()


def add_to_startmenu(file_path="", icon_path="", name='PowerPoint 触屏辅助', args=''):
    """添加开始菜单快捷方式"""
    if file_path == "":
        file_path = os.path.realpath(__file__)
    else:
        file_path = os.path.abspath(file_path)

    if icon_path == "":
        icon_path = file_path

    menu_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs')
    shortcut_path = os.path.join(menu_folder, f'{name}.lnk')

    # ****** 关键：先删再建 ******
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = file_path
    shortcut.Arguments = args
    shortcut.WorkingDirectory = os.path.dirname(file_path)
    shortcut.IconLocation = icon_path
    shortcut.save()


if __name__ == '__main__':
    add_to_startup('PowerPoint_TouchAssist.exe')