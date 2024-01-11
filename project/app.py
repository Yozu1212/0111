from flask import Flask, render_template
import pyodbc

app = Flask(__name__)

# 設定資料庫連接字串
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\User\project\report.accdb;'  # 替換成你的 Access 資料庫路徑
)

def get_table_data():
    try:
        connection = pyodbc.connect(conn_str)

        if connection:
            cursor = connection.cursor()

            # 執行 SQL 查詢，取得所有資料表的資訊
            tables = cursor.tables()

            # 取得第一個資料表的名稱
            first_table_name = tables.fetchone().table_name

            # 執行 SQL 查詢，取得第一個資料表的內容
            query = f"SELECT * FROM {first_table_name}"
            cursor.execute(query)

            # 取得查詢結果
            result = cursor.fetchall()

            return result

    except pyodbc.Error as err:
        print(f"錯誤: {err}")
    finally:
        # 關閉連接
        if 'connection' in locals() and connection:
            connection.close()

@app.route('/')
def index():
    table_data = get_table_data()
    return render_template('index.html', table_data=table_data)

if __name__ == '__main__':
    app.run(debug=True)

@app.route('/other_page')
def ina():
    return render_template('ina.html')

@app.route('/other_page')
def inin():
    return render_template('inin.html')

@app.route('/other_page')
def 庫查():
    return render_template('庫查.html')

@app.route('/other_page')
def 採修():
    return render_template('採修.html')

@app.route('/other_page')
def 採刪():
    return render_template('採刪.html')

@app.route('/other_page')
def 採查():
    return render_template('採查.html')

@app.route('/other_page')
def 採申():
    return render_template('採申.html')

@app.route('/other_page')
def 計查():
    return render_template('計查.html')

@app.route('/other_page')
def 領修():
    return render_template('領修.html')

@app.route('/other_page')
def 領取():
    return render_template('領取.html')

@app.route('/other_page')
def 領查():
    return render_template('領查.html')

@app.route('/other_page')
def 領申():
    return render_template('領申.html')








