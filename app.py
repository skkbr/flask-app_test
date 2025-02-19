from flask import Flask, request, render_template, redirect, url_for, flash
import sqlite3
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your_secret_key'

DB_NAME = "data.db"

def create_or_replace_table_with_schema(df, table_name):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()

        # テーブルを削除して再作成
        cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
        cursor.execute(f"""
            CREATE TABLE {table_name} (
                No INTEGER PRIMARY KEY,
                車No TEXT,
                車名 TEXT,
                工程 TEXT,
                ステータス TEXT,
                作業開始日 DATETIME,
                作業終了日 DATETIME,
                開始時間 DATETIME,
                終了時間 DATETIME,
                合計時間 DATETIME,
                入庫日 DATE,
                担当者名 TEXT,
                優先順位 INTEGER
            );
        """)

        # データを挿入
        for _, row in df.iterrows():
            cursor.execute(f"""
                INSERT INTO {table_name} (
                    No, 車No, 車名, 工程, ステータス, 作業開始日, 作業終了日, 開始時間, 終了時間, 合計時間, 入庫日, 担当者名, 優先順位
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, tuple(row))

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if "file" not in request.files:
            flash("ファイルが選択されていません。")
            return redirect(request.url)

        file = request.files["file"]
        if file.filename == "":
            flash("有効なファイルを選択してください。")
            return redirect(request.url)

        if file and file.filename.endswith(".xlsx"):
            try:
                # Excelデータの読み込み
                df = pd.read_excel(file)

                # テーブル名を現在の日付に基づいて設定
                table_name = "table_" + datetime.now().strftime("%Y%m%d")

                # SQLiteにテーブルを作成または上書き
                create_or_replace_table_with_schema(df, table_name)

                flash(f"データが正常に登録されました (テーブル名: {table_name})。")
                return redirect(url_for("upload_file"))
            except Exception as e:
                flash(f"エラーが発生しました: {e}")
                return redirect(request.url)
        else:
            flash("有効なExcelファイル（.xlsx）をアップロードしてください。")
            return redirect(request.url)
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
