from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    jsonify,
)
import os
import pandas as pd
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)

GSHEET_ID = "1hzP9wBwfVv-K3SaD9PeES4lsseSf62-NKdG_nXWEj5I"

GSHEET_XLSX_URL = (
    f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=xlsx"
)

EXCEL_FILE = "data_local.xlsx"

def download_excel_from_gsheet():
    print(">> Mencoba mengunduh Google Sheets terbaru ...")

    try:
        resp = requests.get(GSHEET_XLSX_URL, verify=False, timeout=30)
        resp.raise_for_status()

        with open(EXCEL_FILE, "wb") as f:
            f.write(resp.content)
    except requests.exceptions.RequestException as e:
        if not os.path.exists(EXCEL_FILE):
            raise


def read_excel_sheet(sheet_name=None):
    download_excel_from_gsheet()

    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(
            f"Tidak ada file {EXCEL_FILE} dan tidak bisa download dari Google Sheets."
        )

    xls = pd.ExcelFile(EXCEL_FILE)
    sheet_names = xls.sheet_names

    if sheet_name not in sheet_names:
        sheet_name = sheet_names[0]

    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str).fillna("")
    columns = list(df.columns)
    data = df.values.tolist()

    return columns, data, sheet_names, sheet_name

@app.route("/")
def root():
    if "user" in session:
        return redirect(url_for("home"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")

        # Login sederhana (bisa kamu ganti nanti)
        if username == "admin" and password == "admin":
            session["user"] = username
            return redirect(url_for("home"))

        error = "Username atau password salah."
        return render_template("login.html", error=error)

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/home")
def home():
    if "user" not in session:
        return redirect(url_for("login"))

    sheet = request.args.get("sheet")

    try:
        columns, data, sheets, active_sheet = read_excel_sheet(sheet)
        download_error = None
    except Exception as e:
        # Kalau benar-benar gagal dapat data
        columns, data, sheets, active_sheet = [], [], [], None
        download_error = str(e)

    return render_template(
        "index.html",
        columns=columns,
        data=data,
        sheets=sheets,
        active_sheet=active_sheet,
        download_error=download_error,
    )


@app.route("/save", methods=["POST"])
def save():
    if "user" not in session:
        return jsonify({"status": "error", "message": "Unauthorized"}), 401

    return jsonify(
        {
            "status": "error",
            "message": "Edit data dilakukan di Google Sheets, bukan lewat web.",
        }
    ), 400

if __name__ == "__main__":
    app.run(debug=True)
