import win32gui
import win32api
import win32con
import tkinter as tk
import time


def move_it(pos):
    handle = win32gui.WindowFromPoint(pos)
    client_pos = win32gui.ScreenToClient(handle, pos)
    tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
    win32gui.SendMessage(handle, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
    win32gui.SendMessage(handle, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, tmp)


def click_it(pos):  # 第三种，可后台
    handle = win32gui.WindowFromPoint(pos)
    client_pos = win32gui.ScreenToClient(handle, pos)
    tmp = win32api.MAKELONG(client_pos[0], client_pos[1])
    win32gui.SendMessage(handle, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
    win32gui.SendMessage(handle, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, tmp)
    win32gui.SendMessage(handle, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, tmp)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    root = tk.Tk()  # 產生 tkinter 視窗
    width = root.winfo_screenwidth()
    height = root.winfo_screenheight()
    print(width, height)
    root.destroy()  # 關閉視窗
    time.sleep(1)
    click_it((width - 40, height - 20))
    time.sleep(1)
    click_it((width - 40, height - 70))
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
