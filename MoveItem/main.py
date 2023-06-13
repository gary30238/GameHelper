from win32com.client import Dispatch
import os
import win32gui
import win32api
from tkinter import messagebox
import atexit
import time
import tkinter
import tkinter.ttk as ttk
import threading
import pythoncom
import random


class Operation:

    def __init__(self, dm, hwnd, event):
        self.dm = dm
        self.hwnd = hwnd
        print(self.dm.Reg('按鍵精靈註冊碼', '按鍵精靈附加碼'))
        print(self.dm.Ver())
        # 初始背包數量
        init_bag_num = int(equip_entry.get())
        # 起始座標
        messagebox.showinfo('Hwnd', '滑鼠移動至待點擊位置')
        init_x = 0
        init_y = 0
        init_pos = self.dm.GetCursorPos(init_x, init_y)
        print(f'{init_pos[1]}, {init_pos[2]}')
        # 綁定
        self.bind()
        init_pos = self.dm.ScreenToClient(hwnd, init_pos[1], init_pos[2])
        init_x = init_pos[1]
        init_y = init_pos[2]
        print(f'{init_x}, {init_y}')
        # alt鍵
        self.dm.KeyDown(18)
        while True:
            self.dm_click(init_x, init_y, 3, 3, 1)
            bag_num = self.read_ro_bag() + 20
            # 背包至倉庫
            if bag_num >= 150 or bag_num == init_bag_num:
                self.dm.KeyUp(18)
                break

            # 結束程式
            if event.is_set():
                self.dm.KeyUp(18)
                break

        # 解除綁定
        unbind()

    def bind(self):
        self.dm.BindWindow(self.hwnd, "dx2", "dx", "dx", 101)
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

    def read_ro_bag(self):
        return self.dm_read_memory('11CCEF4')

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


def validate(entry_str):
    print(entry_str)
    if str.isdigit(entry_str) or entry_str == '':
        return True
    else:
        return False


def unbind():
    try:
        event.set()
        thread.join()
    except Exception as ex:
        print(ex)
    finally:
        dm_main.UnBindWindow()


if __name__ == '__main__':
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
    start_btn = tkinter.Button(top, text="start", command=btn_start)
    end_btn = tkinter.Button(top, text="end", command=unbind)
    start_btn.pack(side=tkinter.LEFT)
    end_btn.pack(side=tkinter.LEFT)
    number_register = (top.register(validate), '%P')
    equip_entry = tkinter.Entry(top, validate='key', validatecommand=number_register)
    equip_entry.pack(side=tkinter.LEFT)
    top.mainloop()
    # 結束指令
    atexit.register(unbind)
