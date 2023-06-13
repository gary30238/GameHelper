from win32com.client import Dispatch
import os
import win32gui
import win32api
from tkinter import messagebox
import atexit
import time
import tkinter
import threading
import pythoncom
import random


class Operation:

    def __init__(self, dm, hwnd, event):
        self.dm = dm
        self.hwnd = hwnd
        print(self.dm.Reg('按鍵精靈註冊碼', '按鍵精靈附加碼'))
        print(self.dm.Ver())
        self.bind()
        self.reset_z()
        while True:
            temp_str = self.dm_find_memory(14606668)
            if len(temp_str) > 0:
                monster_array = temp_str.split('|')
                for aid, monster in enumerate(monster_array):
                    mid = self.monster_info(0, monster)
                    # 邪惡老公公
                    if mid == 1247:
                        self.dm.SetWindowState(self.hwnd, 1)
                        self.dm.EnableBind(0)
                        while True:
                            if self.dm_read_memory(monster) == 14606668:
                                ro_x = self.read_ro_x()
                                ro_y = self.read_ro_y()
                                mx = self.monster_info(1, monster)
                                my = self.monster_info(2, monster)
                                atk_x = 510 + (mx - ro_x) * 20
                                atk_y = 370 - (my - ro_y) * 25
                                self.dm_click(atk_x, atk_y, 10, 10, 0)
                            else:
                                self.dm.EnableBind(1)
                                break

                            # 結束程式
                            if event.is_set():
                                break

                    elif mid == 2382:
                        # 怪物座標
                        mx = self.monster_info(1, monster)
                        my = self.monster_info(2, monster)
                        # 顯示怪物座標
                        print(f'{mid}, x:{mx}, y:{my}')
                        self.dm.EnableBind(0)
                        messagebox.showinfo('monster', f'{mx}, {my}')
                        self.dm.EnableBind(1)

                    # 結束程式
                    if event.is_set():
                        break

            # 瞬移
            if chkValue.get():
                self.dm.KeyPress(112)
            time.sleep(1)

            # 結束程式
            if event.is_set():
                break

    def bind(self):
        if chkValue.get():
            self.dm.BindWindow(self.hwnd, "dx2", "dx", "dx", 101)
        else:
            self.dm.BindWindow(self.hwnd, "normal", "normal", "normal", 101)
        self.dm.SetSimMode(0)
        self.dm.EnableRealKeypad(1)
        self.dm.EnableRealMouse(2, 2, 45)

    def dm_click(self, click_x, click_y, rnd_x, rnd_y, click_type):
        temp_x = int(rnd_x * random.random())
        temp_y = int(rnd_y * random.random())
        if random.random() < 0.5:
            temp_x = -temp_x
        if random.random() < 0.5:
            temp_y = -temp_y
        self.dm.MoveTo(click_x + temp_x, click_y + temp_y)
        time.sleep(random.uniform(0.1, 0.2))
        if click_type == 0:
            self.dm.LeftClick()
        else:
            self.dm.RightClick()

    def read_ro_x(self):
        return self.dm_read_memory('11B95B0')

    def read_ro_y(self):
        return self.dm_read_memory('11B95B4')

    def read_ro_z(self):
        return self.dm.ReadFloat(self.hwnd, "F41CA4")

    def reset_z(self):
        ro_z = int(self.read_ro_z())
        while ro_z < 594:
            wheel_num = int(abs(ro_z - 594) / 24)
            for i in range(wheel_num):
                if ro_z < 594:
                    self.dm.WheelUp()
                else:
                    self.dm.WheelDown()
                time.sleep(0.3)
            ro_z = int(self.dm.ReadFloat(self.hwnd, "F41CA4"))

    def monster_info(self, monster_info_type, monster_addr):
        match monster_info_type:
            case 0:
                return self.dm_read_memory(f'{monster_addr}+10c')
            case 1:
                return self.dm_read_memory(f'{monster_addr}+16c')
            case 2:
                return self.dm_read_memory(f'{monster_addr}+170')

    def dm_find_memory(self, val):
        return self.dm.FindIntEx(self.hwnd, '01000000-7FFFFFFF', val, val, 0, 4, 1, 1)

    def dm_read_memory(self, addr):
        return self.dm.ReadInt(self.hwnd, addr, 0)


def regsvr():
    pythoncom.CoInitialize()
    try:
        dm_1 = Dispatch('dm.dmsoft')

    except:
        os.system(r'regsvr32 /s %s\dm.dll' % os.getcwd())
        dm_1 = Dispatch('dm.dmsoft')
    print(dm_1.Ver())
    return dm_1


def btn_start():
    global thread, event
    event = threading.Event()
    thread = threading.Thread(target=Operation, args=(dm_main, win_id, event))
    thread.start()


def unbind():
    try:
        event.set()
        thread.join()
    except Exception as ex:
        print(ex)
    finally:
        dm_main.UnBindWindow()


if __name__ == '__main__':
    thread = None
    event = None
    # 正文
    # 綁定視窗
    messagebox.showinfo('Hwnd', '滑鼠移動至待綁定窗口')
    point = win32api.GetCursorPos()
    win_id = win32gui.WindowFromPoint(point)
    rect = win32gui.GetWindowRect(win_id)
    # 註冊
    dm_main = regsvr()
    # 生成視窗
    top = tkinter.Tk()
    top.attributes("-topmost", True)
    top.geometry(f'+{rect[0]}+{rect[1]}')
    btn1 = tkinter.Button(top, text="start", command=btn_start)
    btn2 = tkinter.Button(top, text="end", command=unbind)
    btn1.pack(side=tkinter.LEFT)
    btn2.pack(side=tkinter.LEFT)
    chkValue = tkinter.BooleanVar()
    chkValue.set(True)
    chkF1 = tkinter.Checkbutton(top, text='F1', variable=chkValue)
    chkF1.pack()
    top.mainloop()
    # 結束指令
    atexit.register(unbind)
