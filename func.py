import pyautogui
import pygetwindow as gw
from pynput import mouse
from win32 import win32api
import conf_file as conf
from loguru import logger
# 屏幕分辨率
screenX = win32api.GetSystemMetrics(0)
screenY = win32api.GetSystemMetrics(1)

# 缩放比率（读取配置，读取失败时默认为 1）
scale_rate = conf.dpi_dict.get(conf.read_conf('General', 'DPI'), 1)

# PowerPoint 全屏窗口标题（可配置）
ppt_window_title = conf.read_conf('General', 'PPT_Title') or 'PowerPoint 幻灯片放映'

# 用于判断是否点击到任务栏或 PPT 菜单区的阈值
tsk_edge = screenY - 95 * scale_rate
ppt_menuWL = 95 * scale_rate
ppt_menuWR = screenX - 95 * scale_rate

# 全局状态变量
is_pressed = False
pressed_menuButton = False
pos_mouse = [0, 0]
count = 0


def is_powerpoint_showing():
    """判断是否存在指定标题的 PowerPoint 放映窗口。

    返回:
        bool: 如果存在匹配窗口则为 True，否则为 False。
    """
    try:
        windows = gw.getWindowsWithTitle(ppt_window_title)
        return len(windows) > 0
    except Exception:
        return False


def is_finger_not_slide(x, y):
    """判断当前位置是否与按下时的位置相同（即未滑动）。"""
    return [x, y] == pos_mouse


def is_not_click_taskbar(mouse_y):
    """判断 y 坐标是否不在任务栏区域内（未点击任务栏）。"""
    return mouse_y <= tsk_edge


def is_not_click_ppt_menubar(mouse_x):
    """判断 x 坐标是否不在 PPT 菜单栏区域内（未点击菜单）。"""
    return ppt_menuWL <= mouse_x <= ppt_menuWR


def on_click(x, y, button, pressed):
    """鼠标点击回调，处理左/右键按下和释放逻辑。

    左键按下：记录按下位置，判断是否点击到菜单或任务栏。
    左键释放：若为点击（未滑动、未触发菜单、且为 PPT 放映），则发送空格键（翻页）。
    右键按下：视作触发菜单按钮，避免误触发翻页。
    """
    global is_pressed, pos_mouse, pressed_menuButton, count
    print(f"X坐标：{x}，Y坐标{y}，按键：{button}，按下：{pressed}\n")
    if button == mouse.Button.right:
        # 右键视作菜单操作，阻止翻页触发
        pressed_menuButton = True
        print("右键")
        logger.info("A.R.O.N.A: 右键点击")

    if button == mouse.Button.left:
        if pressed:
            # 按下：记录位置并判断是否点击菜单或任务栏
            pos_mouse = [x, y]
            if not is_not_click_ppt_menubar(x) or not is_not_click_taskbar(y):
                pressed_menuButton = True
                count = 2
            else:
                is_pressed = True
        else:
            # 释放：若为有效点击且为 PPT 放映，则发送空格翻页
            if is_pressed and not pressed_menuButton and is_finger_not_slide(x, y) and is_powerpoint_showing():
                pyautogui.press('space')
                print("A.R.O.N.A-DEBUG: 翻页")
                logger.info("A.R.O.N.A: 被点击并翻页")
            elif count != 0:
                count -= 1
            if count == 0:
                pressed_menuButton = False
            is_pressed = False


def main():
    """启动鼠标监听，运行主循环."""
    print('A.R.O.N.A-DEBUG: 程序开始运行')
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()


if __name__ == '__main__':
    main()