from win32com.client import Dispatch
import os
import win32gui
import win32api
from tkinter import messagebox
import atexit
import time
import pythoncom
import pymysql
import re
import random
from datetime import datetime, timedelta

# 資料庫設定
db_settings = {
    "host": "127.0.0.1",
    "port": 3306,
    "user": "root",
    "password": "vu,41i6u04",
    "db": "ro_stock",
    "charset": "utf8"
}


class Operation:

    def __init__(self, dm, hwnd):
        self.dm = dm
        self.hwnd = hwnd
        print(self.dm.Reg('1347528682239add254149076f357bd15a5b80f8b56', 'fmvbhs120'))
        print(self.dm.Ver())
        self.dm.setDict(0, r'd:\User\Desktop\dm_soft.txt')
        self.dm.useDict(0)
        self.bind()
        # 正規化
        pattern = re.compile(r'\d+P')
        # 取得時間
        now_time = datetime.now()
        run_min = '{0:02d}'.format(now_time.minute - (now_time.minute % 5))
        run_time_str = datetime.strftime(now_time, f'%Y-%m-%d %H:{run_min}:00')
        run_time = datetime.strptime(run_time_str, f'%Y-%m-%d %H:%M:%S')
        print(f'time: {run_time_str}')

        while True:

            while True:
                stock_price_array = []

                # 股票市場
                while self.dm.ReadInt(self.hwnd, "F80388", 0) != 9:
                    self.dm_click(630, 555, 10, 10)
                    if self.dm.ReadInt(self.hwnd, "F80388", 0) == 9:
                        break
                    else:
                        time.sleep(5)

                # 讀取股價
                # 第一頁
                money = int(self.dm.Ocr(200, 100, 450, 160, "ff0000-000000", 0.9))
                print(f'money: {money}')
                self.read_price(pattern, stock_price_array)

                # 滾動視窗
                self.dm_wheel(400, 200, 10, 10, 1, 20)
                time.sleep(random.uniform(0.1, 0.2))

                # 第二頁
                self.read_price(pattern, stock_price_array)

                # 關閉視窗
                # 下一步
                self.dm.KeyPress(13)
                time.sleep(random.uniform(0.3, 0.5))
                # 取消
                for i in range(3):
                    self.dm.KeyPress(40)
                    time.sleep(random.uniform(0.1, 0.2))
                # 確認
                for i in range(2):
                    self.dm.KeyPress(13)
                    time.sleep(random.uniform(0.3, 0.5))

                if len(stock_price_array) >= 10:
                    break

            try:
                # 建立Connection物件
                with pymysql.connect(**db_settings) as conn:
                    # 建立Cursor物件
                    with conn.cursor() as cursor:
                        # 資料表相關操作
                        # 查詢資料庫
                        command = "SELECT * FROM price where create_date = %s"
                        cursor.execute(command, run_time)
                        price_table = cursor.fetchone()
                        if price_table is None:
                            # 寫入資料表
                            command = "INSERT INTO price(price, create_date, stock_id)VALUES(%s, %s, %s)"
                            for sid, stock_price in enumerate(stock_price_array):
                                cursor.execute(command, (stock_price, run_time, sid + 1))
                                print(f'sid: {sid + 1} price: {stock_price}')
                            # 儲存變更
                            conn.commit()

            except Exception as ex:
                print(ex)

            # 下次執行時間
            run_time = run_time + timedelta(minutes=5)

            while datetime.now() <= run_time:
                time.sleep(60)

    def bind(self):
        self.dm.BindWindowEx(self.hwnd, "dx2", "dx", "dx", "dx.public.fake.window.min", 101)
        self.dm.SetSimMode(0)
        self.dm.EnableRealKeypad(1)
        self.dm.EnableRealMouse(2, 2, 45)

    def dm_click(self, click_x, click_y, rnd_x, rnd_y):
        temp_x = int(rnd_x * random.random())
        temp_y = int(rnd_y * random.random())
        if random.random() < 0.5:
            temp_x = -temp_x
        if random.random() < 0.5:
            temp_y = -temp_y
        self.dm.MoveTo(click_x + temp_x, click_y + temp_y)
        time.sleep(random.uniform(0.1, 0.2))
        self.dm.LeftClick()

    def dm_wheel(self, click_x, click_y, rnd_x, rnd_y, wheel_type, wheel_num):
        temp_x = int(rnd_x * random.random())
        temp_y = int(rnd_y * random.random())
        if random.random() < 0.5:
            temp_x = -temp_x
        if random.random() < 0.5:
            temp_y = -temp_y
        self.dm.MoveTo(click_x + temp_x, click_y + temp_y)
        time.sleep(random.uniform(0.1, 0.2))
        if wheel_type == 0:
            for i in range(wheel_num):
                self.dm.WheelUp()
                time.sleep(random.uniform(0.01, 0.02))
        else:
            for i in range(wheel_num):
                self.dm.WheelDown()
                time.sleep(random.uniform(0.01, 0.02))

    def read_price(self, read_price_pattern, read_price_array):
        sentence = self.dm.Ocr(200, 100, 350, 250, "000000-000000", 0.9)
        time.sleep(random.uniform(0.5, 1))
        for temp in read_price_pattern.finditer(sentence):
            read_price_array.append(re.sub('P', '', temp.group()))
        time.sleep(1)

    def trading_stock(self, trading_type, trading_stock, trading_num):
        # 股票市場
        while self.dm.ReadInt(self.hwnd, "F80388", 0) != 9:
            self.dm_click(630, 555, 10, 10)
            time.sleep(1)
            if self.dm.ReadInt(self.hwnd, "F80388", 0) == 9:
                break
            else:
                time.sleep(5)

        # 下一步
        self.dm.KeyPress(13)
        time.sleep(random.uniform(0.3, 0.5))

        for i in range(trading_type):
            self.dm.KeyPress(40)
            time.sleep(random.uniform(0.1, 0.2))

        # 確認
        if trading_type == 1:
            for i in range(3):
                self.dm.KeyPress(13)
                time.sleep(random.uniform(0.3, 0.5))
        else:
            self.dm.KeyPress(13)
            time.sleep(random.uniform(0.3, 0.5))

        # 選擇股票
        if trading_stock - 1 > 0:
            for i in range(trading_stock - 1):
                self.dm.KeyPress(40)
                time.sleep(random.uniform(0.1, 0.2))

        # 確認
        self.dm.KeyPress(13)
        time.sleep(random.uniform(0.3, 0.5))
        # 輸入張數
        temp_num = str(trading_num)
        for i in range(len(temp_num)):
            key_num = 48 + int(temp_num[i])
            self.dm.KeyPress(key_num)
            time.sleep(random.uniform(0.1, 0.2))
        # 確認
        for i in range(3):
            self.dm.KeyPress(13)
            time.sleep(random.uniform(0.3, 0.5))
        time.sleep(1)

        # 關閉視窗
        # 下一步
        self.dm.KeyPress(13)
        time.sleep(random.uniform(0.3, 0.5))
        # 取消
        for i in range(3):
            self.dm.KeyPress(40)
            time.sleep(random.uniform(0.1, 0.2))
        # 確認
        for i in range(2):
            self.dm.KeyPress(13)
            time.sleep(random.uniform(0.3, 0.5))


def regsvr():
    pythoncom.CoInitialize()
    try:
        dm_1 = Dispatch('dm.dmsoft')

    except Exception as ex:
        print(ex)
        os.system(r'regsvr32 /s %s\dm.dll' % os.getcwd())
        dm_1 = Dispatch('dm.dmsoft')
    return dm_1


def unbind():
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
    # 執行
    Operation(dm_main, win_id)
    # 結束指令
    atexit.register(unbind)
