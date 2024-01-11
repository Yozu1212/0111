import pyodbc

# 設定資料庫連接字串
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\User\project\report.accdb;'  # 替換成你的 Access 資料庫路徑
)

# 連接到 Microsoft Access 資料庫
try:
    connection = pyodbc.connect(conn_str)

    if connection:
        print("成功連接到 Microsoft Access 資料庫")

        # 取得資料庫中的所有資料表
        tables = connection.cursor().tables()

        # 輸出所有資料表的名稱
        for table_info in tables:
            print(table_info.table_name)

except pyodbc.Error as err:
    print(f"錯誤: {err}")

#finally:
    # 關閉連接
    #if 'connection' in locals() and connection:
        #connection.close()
        #print("Microsoft Access 連接已關閉")
