import time
import pywinauto
import pyautogui
import win32gui
from pywinauto.application import Application



class FunctionExecution:
    def __init__(self):
        self.stored_function = None

    def execute_and_store(self, func, *args, **kwargs):
        result = func(*args, **kwargs)
        self.stored_function = (func, args, kwargs)
        return result

    def re_execute(self):
        if self.stored_function:
            func, args, kwargs = self.stored_function
            func(*args, **kwargs)
        else:
            print("No function to execute.")


def window_close(title):
    MAX_TRIES = 30
    SLEEP_TIME = 0.1
    execution = FunctionExecution()
    tries = 0
    while True:
        tries += 1
        if tries > MAX_TRIES:
            tries = 0
            execution.re_execute()
        needed_hwnd = win32gui.FindWindow(None, title)
        if needed_hwnd == 0:
            break
        time.sleep(SLEEP_TIME)


def wait_for_window(title, re_execute=True):
    MAX_TRIES = 30
    SLEEP_TIME = 0.1
    execution = FunctionExecution()
    tries = 0
    while True:
        tries += 1
        if re_execute and tries > MAX_TRIES:
            tries = 0
            execution.re_execute()
        needed_hwnd = win32gui.FindWindow(None, title)
        if needed_hwnd != 0 and win32gui.GetWindowText(needed_hwnd) == title:
            return needed_hwnd
        time.sleep(SLEEP_TIME)


def wait_for_either_window(title1, title2):
    MAX_TRIES = 30
    SLEEP_TIME = 0.1

    execution = FunctionExecution()
    tries = 0
    while True:
        tries += 1
        if tries > MAX_TRIES:
            tries = 0
            execution.re_execute()
        needed_hwnd1 = win32gui.FindWindow(None, title1)
        needed_hwnd2 = win32gui.FindWindow(None, title2)
        if needed_hwnd1 != 0 and win32gui.GetWindowText(needed_hwnd1) == title1:
            return True
        elif needed_hwnd2 != 0 and win32gui.GetWindowText(needed_hwnd2) == title2:
            return False
        time.sleep(SLEEP_TIME)


def wait_for_window_close(needed_hwnd):
    MAX_TRIES = 30
    SLEEP_TIME = 0.1
    tries = 0
    while tries < MAX_TRIES:
        if not win32gui.IsWindow(needed_hwnd):
            return True
        time.sleep(SLEEP_TIME)
        tries += 1
    else:
        return False


def get_dlg(title=None, hwnd=None):
    if hwnd == None:
        hwnd = win32gui.FindWindow(None, title)
    app = pywinauto.Application().connect(handle=hwnd)
    return app.window(handle=hwnd)


def click_button(button_name, dlg):
    Button = dlg.child_window(class_name="Button", title=button_name)
    Button.click()


def double_click_cell(listview, rowcount):
    cell = listview.get_item(rowcount, 0)
    try:
        cell.click()
        cell.click()
        pyautogui.press("enter")

    except:
        pass


def get_data(listview, row, column):
    cell = listview.get_item(row, column)
    return cell.text()


def convert_runtime_to_hours(runtime_str):
    if "天" in runtime_str:
        return int(runtime_str.split()[0]) * 60 * 24
    elif "小时" in runtime_str:
        return int(runtime_str.split()[0]) * 60
    else:
        return 0


def convert_speed_to_float(speed_str):
    return float(speed_str.split()[0])


def convert_country_to_id(country_str):
    return country_str[0]


def main(categories="J,K"):
    execution = FunctionExecution()
    TITLE = "SoftEther VPN Client 管理器"
    TITLE2 = "SoftEther VPN 客户端的 VPN Gate 学术试验项目插件"
    CHOICE1 = "选择 VPN 协议来连接"
    CHOICE2 = "正在连接 VPN Gate Connection ..."
    BUTTON1 = "确定(&O)"
    BUTTON2 = "关闭(&C)"
    ERROR = "连接错误 - VPN Gate Connection"
    SUCCESS = '虚拟网络适配器 "VPN" 状态'
    CANCEL = "连接取消"
    PATH = r"C:\Program Files\SoftEther VPN Client\vpncmgr.exe"
    Application(backend="uia").start(fr'{PATH}')
    row_count = 0
    while True:
        dlg = get_dlg(title=TITLE)
        hwnd = win32gui.FindWindow(None, TITLE)
        main_app = pywinauto.Application().connect(handle=hwnd)

        list_view = dlg.child_window(class_name="SysListView32", control_id=1047)
        execution.execute_and_store(double_click_cell, list_view, 1)

        print("正在准备连接新的vpn")

        now_hwnd = wait_for_window(TITLE2)
        now_dlg = get_dlg(hwnd=now_hwnd)

        list_view2 = now_dlg.child_window(class_name="SysListView32")

        runtime_speed_country_id = [
            [convert_runtime_to_hours(get_data(list_view2, i, 3)), convert_speed_to_float(get_data(list_view2, i, 5)),
             convert_country_to_id(get_data(list_view2, i, 2)), i] for i in range(0, list_view2.item_count())]

        filtered_data = [item for item in runtime_speed_country_id if item[2] in categories]

        sorted_data = sorted(filtered_data, key=lambda x: (x[0] - x[1] / 40))
        # print(sorted_data)
        # exit()
        try:
            row_to_click = sorted_data[row_count][3]
        except:
            print(row_count)
            print(len(sorted_data))
            print(sorted_data[row_count])
            exit()
        execution.execute_and_store(double_click_cell, list_view2, row_to_click)
        row_count += 1
        if row_count == 16 or row_count == 19 \
                or row_count == 16 or row_count == 16:
            continue
        print("开始连接")
        if wait_for_either_window(CHOICE1, CHOICE2):
            now_dlg = get_dlg(title=CHOICE1)
            # now_dlg.print_ctrl_ids()
            execution.execute_and_store(click_button, BUTTON1, now_dlg)
            window_close(CHOICE1)
            print("选择了连接方式")

        hwnd3 = wait_for_window(CHOICE2, False)
        # print(hwnd3)
        print("连接中")

        if not wait_for_window_close(hwnd3):

            print("连接超时，更换连接中")
            now_dlg = get_dlg(title=CHOICE2, hwnd=hwnd3)
            click_button("取消", now_dlg)
        else:

            print("连接结果出现")
            if wait_for_either_window(ERROR, SUCCESS):

                now_dlg = get_dlg(title=ERROR)
                execution.execute_and_store(click_button, CANCEL, now_dlg)
                print("更换vpn中")
                wait_for_window(TITLE)
            else:

                print("连接成功！连接国家：{}，连接速度：{}，位置：{}".format(sorted_data[row_count][2],
                                                                        sorted_data[row_count][1], row_count))
                now_hwnd = wait_for_window(SUCCESS)
                now_dlg = get_dlg(hwnd=now_hwnd)

                click_button(BUTTON2, now_dlg)
                main_app.kill()
                exit()


if __name__ == "__main__":
    main("J")
