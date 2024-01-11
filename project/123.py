from flask import Flask, render_template, request
import pyodbc

app = Flask(__name__)

# 連接到 Access 資料庫
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\User\project\report.accdb;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()


# 根據品項編號查詢相對應的資料
@app.route('/', methods=['GET', 'POST'])
def 庫查():
    if request.method == 'POST':
        item_number = request.form.get('item_number')

        # 使用品項編號進行查詢
        cursor.execute(f"SELECT 品項編號 FROM 品項 WHERE 品項編號 IN ('F001', 'F002', 'F003', 'F004', 'F005')", item_number)
        result = cursor.fetchone()

        if result:
            item_name = result.品項名稱
            item_category = result.品項類別
            item_unit = result.品項單位
            item_spec = result.品項規格
            item_location = result.所屬的庫別
            warehouse_name = result.倉庫名稱
            warehouse_management_unit = result.倉庫的管理單位
            warehouse_location = result.倉庫儲位

            # 透過品項編號查詢倉庫相關資料
            cursor.execute(f"SELECT 品項名稱 FROM 品項 WHERE 品項名稱 IN ('能量棒', '蛋白質粉末', '維生素A', '運動飲料A', '運動飲料B')", item_number)
            item_name_result = cursor.fetchone()

            cursor.execute(f"SELECT 品項類別 FROM 品項 WHERE 品項類別 IN ('能量補給', '蛋白質補給', '維生素補給', '電解質補充', '電解質補充')", item_number)
            item_category_result = cursor.fetchone()


            cursor.execute(f"SELECT 品項單位 FROM 品項 WHERE 品項單位 IN ('隻', '包', '罐', '瓶', '瓶')", item_number)
            item_unit_result = cursor.fetchone()

            cursor.execute(f"SELECT 品項規格 FROM 品項 WHERE 品項規格 IN ('總熱量', '蛋白質含量', 'IU', '電解質', '電解質')", item_number)
            item_spec_result = cursor.fetchone()

            cursor.execute(f"SELECT 所屬的庫別 FROM 品項 WHERE 所屬的庫別 IN ('A館', 'C館', 'D館', 'B館', 'C館')", item_number)
            item_location_result = cursor.fetchone()

            cursor.execute(f"SELECT 倉庫名稱 FROM 倉庫 WHERE 倉庫名稱 IN ('W001', 'W002', 'W003', 'W004', 'W005')", item_number)
            warehouse_name_result = cursor.fetchone()

            cursor.execute(f"SELECT 倉庫的管理單位 FROM 倉庫 WHERE 倉庫的管理單位 IN ('運動科學處', '競技運動處', '教育訓練處', '運動科學處', '教育訓練處')", item_number)
            warehouse_management_unit_result = cursor.fetchone()

            cursor.execute(f"SELECT 倉庫儲位 FROM 倉庫儲位 WHERE 倉庫儲位 IN ('S001', 'S002', 'S003', 'S004', 'S005')", item_number)
            warehouse_location_result = cursor.fetchone()


            return render_template('result.html',
                                   item_name=item_name,
                                   item_category=item_category,
                                   item_unit=item_unit,
                                   item_spec=item_spec,
                                   item_location=item_location,
                                   warehouse_name=warehouse_name,
                                   warehouse_management_unit=warehouse_management_unit,
                                   warehouse_location=warehouse_location)
        else:
            return render_template('error.html', message='找不到該品項編號')

    return render_template('庫查.html')

if __name__ == '__main__':
    app.run(debug=True)