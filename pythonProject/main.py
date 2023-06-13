import matplotlib.pyplot as plt
import pymysql

# 資料庫設定
db_settings = {
    "host": "",
    "port": ,
    "user": "",
    "password": "",
    "db": "ro_stock",
    "charset": "utf8"
}

y = [[], [], [], [], [], [], [], [], [], []]

try:
    # 建立Connection物件
    with pymysql.connect(**db_settings) as conn:
        # 建立Cursor物件
        with conn.cursor() as cursor:
            # 資料表相關操作
            # 查詢資料庫
            command = "SELECT * FROM ro_stock.price where " \
                      "create_date >= CAST((NOW() - INTERVAL 4 HOUR) AS DATETIME)"
            cursor.execute(command)
            price_table = cursor.fetchall()
            i = 1
            for pid, price, create_date, stock_id in price_table:
                y[stock_id - 1].append(price)


except Exception as ex:
    print(ex)

x = list(range(1, len(y[0]) + 1))
# 繪製折線圖，顏色「藍色」，線條樣式「-」，線條寬度「2」，標記大小「16」，標記樣式「.」，圖例名稱「Plot 1」
for i in range(10):
    plt.subplot(10, 1, i + 1)
    plt.plot(x, y[i], color='blue', linestyle="-", linewidth="2", markersize="16", marker=".")

# 設定 x 軸標題內容及大小
plt.xlabel('x label', fontsize="10")
# 設定 y 軸標題內容及大小
plt.ylabel('y label', fontsize="10")
# 設定圖表標題內容及大小
plt.title('Plot title', fontsize="18")

plt.show()
