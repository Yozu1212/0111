from flask import Flask, render_template, request
import pyodbc

app = Flask(__name__)

# 連接到 Access 資料庫
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\User\project\report.accdb;'  # 這是我的 Access 資料庫路徑
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

def get_warehouse_data(warehouse_id):
    # 根據倉庫編號查詢相關資料
    cursor.execute(f"SELECT * FROM 倉庫 WHERE 倉庫編號 = '{warehouse_id}'")
    warehouse_result = cursor.fetchone()

    if warehouse_result:
        warehouse_name = warehouse_result.倉庫名稱
        warehouse_management_unit = warehouse_result.倉庫的管理單位

        # 根據倉庫編號查詢倉庫儲位
        cursor.execute(f"SELECT * FROM 倉庫儲位 WHERE 倉庫編號 = '{warehouse_id}'")
        warehouse_location_result = cursor.fetchone()

        warehouse_location = warehouse_location_result.倉庫儲位

        return warehouse_name, warehouse_management_unit, warehouse_location
    else:
        return None, None, None

# 根據品項編號查詢相對應的資料
@app.route('/', methods=['GET', 'POST'])
def 庫查():
    if request.method == 'POST':
        item_number = request.form.get('item_number')

        # 使用品項編號進行查詢
        cursor.execute(f"SELECT * FROM 品項 WHERE 品項編號 = '{item_number}'")
        result = cursor.fetchone()

        if result:
            item_name = result.品項名稱
            item_category = result.品項類別
            item_unit = result.品項單位
            item_spec = result.品項規格
            item_location = result.所屬的庫別

            # 獲取倉庫相關資料
            warehouse_name, warehouse_management_unit, warehouse_location = get_warehouse_data(result.倉庫編號)

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