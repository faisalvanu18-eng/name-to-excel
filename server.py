from flask import Flask, render_template, request, redirect, url_for
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_FILE = "records.xlsx"
SHEET_NAME = "Records"

def init_excel_if_needed():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(["Name", "Date", "Time"])
        wb.save(EXCEL_FILE)

def save_record(name: str):
    init_excel_if_needed()
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")

    wb = load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    ws.append([name, date_str, time_str])
    wb.save(EXCEL_FILE)

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")

@app.route("/submit", methods=["POST"])
def submit():
    name = request.form.get("name", "").strip()
    if not name:
        return redirect(url_for("home"))
    save_record(name)
    return redirect(url_for("success"))

@app.route("/success", methods=["GET"])
def success():
    return render_template("success.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)