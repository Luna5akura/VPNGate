import time

import pyautogui
import pywinauto
import win32gui

# 等待窗口出现
from pywinauto import Application

1


def wait_for_window(title):
    while True:
        needed_hwnd = win32gui.FindWindow(None, title)
        if needed_hwnd != 0 and win32gui.GetWindowText(needed_hwnd) == title:
            return needed_hwnd
        time.sleep(0.1)


def main():
    TITLE = "SoftEther VPN Client 管理器"
    PATH = r"C:\Program Files\SoftEther VPN Client\vpncmgr.exe"
    Application(backend="uia").start(f'{PATH}')
    hwnd = wait_for_window("SoftEther VPN Client 管理器")
    main_app = pywinauto.Application().connect(handle=hwnd)
    # 获取主窗口
    app = pywinauto.Application().connect(handle=hwnd)
    dlg = app.window(handle=hwnd)

    # 获取ListView控件对象
    list_view = dlg.child_window(class_name="SysListView32", control_id=1047)

    # 获取第三行第一列的单元格对象
    # cell = list_view.get_item(2, 0)

    dlg.set_focus()

    # 模拟 Ctrl+I 组合键
    dlg.type_keys('^i')
    time.sleep(0.5)

    hwnd = wait_for_window(TITLE)
    # 获取主窗口
    app = pywinauto.Application().connect(handle=hwnd)
    dlg = app.window(handle=hwnd)
    dlg.type_keys('y')
    time.sleep(1)
    main_app.kill()
