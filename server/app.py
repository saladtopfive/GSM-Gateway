from flask import Flask, request, render_template, jsonify, send_file
from openpyxl import load_workbook
from datetime import datetime
import pytz
import os
import shutil
import tempfile
import re

app = Flask(__name__)

LOCAL_TZ = pytz.timezone("Europe/Warsaw")

UPLOAD_PATH = "/home/kp_rpi_user/GSM-Gateway/stable-release/schedule.xlsx"
PHONEBOOK_PATH = "/home/kp_rpi_user/GSM-Gateway/server/phonebook.txt"

ALLOWED_EXTENSIONS = {"xlsx"}

PHONE_REGEX = re.compile(r"^\+?48?\d{9}$")


# ===== PHONEBOOK =====

def load_phonebook():
    phonebook = {}

    if not os.path.exists(PHONEBOOK_PATH):
        return phonebook

    with open(PHONEBOOK_PATH, "r", encoding="utf-8") as f:
        for line in f:
            if "|" not in line:
                continue
            number, name = line.split("|", 1)
            phonebook[number.strip()] = name.strip()

    return phonebook


def resolve_name(number):
    return load_phonebook().get(number, number)


# ===== HELPERS =====

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def normalize_number(raw):
    if raw is None:
        return None

    raw = str(raw).strip()

    if not PHONE_REGEX.fullmatch(raw):
        return None

    return raw


def load_schedule():
    wb = load_workbook(UPLOAD_PATH)
    ws = wb.active
    rows = []

    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or len(row) < 5:
            continue

        sd, ed, st, et, raw_num = row[:5]

        if not all([sd, ed, st, et, raw_num]):
            continue

        number = normalize_number(raw_num)
        if not number:
            continue

        start = LOCAL_TZ.localize(datetime.combine(sd, st))
        end = LOCAL_TZ.localize(datetime.combine(ed, et))

        rows.append((start, end, number))

    return sorted(rows, key=lambda x: x[0])


# ===== ROUTES =====

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/status")
def status():
    now = datetime.now(LOCAL_TZ)
    schedule = load_schedule()

    current = None
    next_one = None

    for start, end, num in schedule:
        if start <= now <= end:
            current = (num, start, end)
        elif start > now and next_one is None:
            next_one = (num, start, end)

    def fmt(entry):
        if not entry:
            return None
        num, start, end = entry
        return {
            "person": resolve_name(num),
            "number": num,
            "start": start.strftime("%Y-%m-%d %H:%M"),
            "end": end.strftime("%Y-%m-%d %H:%M"),
        }

    return jsonify({
        "current": fmt(current),
        "next": fmt(next_one)
    })


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Brak pliku"}), 400

    file = request.files["file"]

    if not allowed_file(file.filename):
        return jsonify({"error": "Dozwolony jest tylko plik .xlsx"}), 400

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        wb = load_workbook(tmp_path)
        ws = wb.active
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        expected = ["start_date", "end_date", "start_time", "end_time", "number"]

        if headers[:5] != expected:
            raise ValueError("Zła struktura nagłówków")

    except Exception:
        os.unlink(tmp_path)
        return jsonify({"error": "Nieprawidłowy plik Excel"}), 400

    shutil.move(tmp_path, UPLOAD_PATH)
    return jsonify({"status": "ok"})


@app.route("/download")
def download():
    return send_file(UPLOAD_PATH, as_attachment=True, download_name="schedule.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
