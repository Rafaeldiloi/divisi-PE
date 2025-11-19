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

# Matikan warning SSL karena kita pakai verify=False (workaround SSL)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.secret_key = "ganti-ini-dengan-string-random"  # ganti dengan string acak

# ==========================
# KONFIGURASI GOOGLE SHEETS
# ==========================

GSHEET_ID = "1hzP9wBwfVv-K3SaD9PeES4lsseSf62-NKdG_nXWEj5I"

# Export sebagai file Excel (semua sheet)
GSHEET_XLSX_URL = (
    f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=xlsx"
)

# File lokal cache
EXCEL_FILE = "data_local.xlsx"


# ==========================
# FUNGSI BANTUAN
# ==========================

def download_excel_from_gsheet():
    """
    Coba download file Excel terbaru dari Google Sheets.
    Jika gagal (masalah jaringan / SSL / diblok) dan file lokal SUDAH ADA,
    maka tetap pakai file lokal terakhir.
    Jika gagal dan file lokal BELUM ADA, lempar error (tidak ada data sama sekali).
    """
    print(">> Mencoba mengunduh Google Sheets terbaru ...")

    try:
        resp = requests.get(GSHEET_XLSX_URL, verify=False, timeout=30)
        resp.raise_for_status()

        with open(EXCEL_FILE, "wb") as f:
            f.write(resp.content)

        print(">> Download selesai.")
    except requests.exceptions.RequestException as e:
        print(">> GAGAL download dari Google Sheets:", e)

        if not os.path.exists(EXCEL_FILE):
            # Belum ada cache sama sekali -> fatal
            raise
        # Kalau cache sudah ada -> pakai data lama
        print(">> Menggunakan file lokal lama:", EXCEL_FILE)


def read_excel_sheet(sheet_name=None):
    """
    Baca satu sheet dari file Excel lokal.
    Sebelum membaca, COBA download dulu versi terbaru.
    Kalau download gagal, tetap pakai file lokal terakhir.
    """
    # Sinkron dulu (kalau gagal -> pakai file lokal)
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


# ==========================
# ROUTES
# ==========================

@app.route("/")
def root():
    # Kalau sudah login -> langsung ke home
    if "user" in session:
        return redirect(url_for("home"))
    # Kalau belum -> ke halaman login
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
    # Saat ini data dikelola via Google Sheets, jadi fitur save dimatikan
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
