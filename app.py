from flask import Flask, render_template, session, request, redirect, url_for
import os
from werkzeug.utils import secure_filename
from flask import send_from_directory
import pandas as pd
import openpyxl

#インスタンスの作成
app = Flask(__name__)

#暗号鍵の作成
key = os.urandom(21)
app.secret_key = key

#メイン
@app.route("/")
def index():
    #表示させたいdataを渡す
    return render_template('index.html')

@app.route('/edit', methods=["POST"])
def edit():
    #ファイルを取り出す
    report = request.files['report']

    #ファイルを取り出す
    this_year = request.files["this_year"]

    #今年の売上合計照会のCSVファイルを取得
    this_year_csv = pd.read_csv(this_year,encoding="shift-jis")

    #客数と点数と実粗利の取得
    kyakusuu = this_year_csv.loc[0]["客数"]
    tensuu = this_year_csv.loc[0]["点数"]
    zituarari = this_year_csv.loc[0]["実粗利"]

    #ブックを取得
    book = openpyxl.load_workbook(report)

    #シートを取得
    sheet = book.worksheets[1]

    #シートに書き込む
    sheet["H6"] = kyakusuu
    sheet["J10"] = tensuu
    sheet["D7"] = zituarari

    #パスを作成
    result_file_path = 'report_dir/' + '編集済月度報告書.xlsx'

    #シートを上書き保存する
    report.save(result_file_path)

    return send_from_directory('report_dir','編集済月度報告書.xlsx', as_attachment=True)
